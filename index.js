'use strict';

const express    = require('express');
const cors       = require('cors');
const { Pool }   = require('pg');
const { google } = require('googleapis');
const { Readable } = require('stream');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat
} = require('docx');

const app  = express();
const pool = new Pool({ connectionString: process.env.DATABASE_URL });

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ─────────────────────────────────────────────────────────────────────
// HEALTH CHECK
// ─────────────────────────────────────────────────────────────────────
app.get('/health', async (_, res) => {
  try {
    await pool.query('SELECT 1');
    res.json({ ok: true, db: 'conectado', ts: new Date().toISOString() });
  } catch (e) {
    res.status(500).json({ ok: false, erro: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// ROTA PRINCIPAL — POST /aprovar
// ─────────────────────────────────────────────────────────────────────
app.post('/aprovar', async (req, res) => {
  const {
    modulo,
    etapa_funil,
    canal,
    tipo_conteudo,
    geo_ou_humano  = 'humano',
    persona_alvo   = '',
    objetivo       = '',
    titulo,
    corpo_conteudo,
    usuario        = 'time-marketing',
    contexto_brief = '',
    formato_saida  = 'markdown'
  } = req.body;

  // Validação mínima
  const obrigatorios = { modulo, etapa_funil, canal, tipo_conteudo, titulo, corpo_conteudo };
  const faltando = Object.entries(obrigatorios)
    .filter(([, v]) => !v || !v.trim())
    .map(([k]) => k);

  if (faltando.length) {
    return res.status(400).json({
      ok: false,
      erro: `Campos obrigatórios ausentes: ${faltando.join(', ')}`
    });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    // 1. Inserir campanha
    const { rows: [campanha] } = await client.query(`
      INSERT INTO campanhas
        (modulo, etapa_funil, objetivo, persona_alvo,
         usuario_solicitante, contexto_brief, status)
      VALUES ($1, $2, $3, $4, $5, $6, 'aprovado')
      RETURNING id`,
      [modulo, etapa_funil, objetivo, persona_alvo, usuario, contexto_brief]
    );

    // 2. Inserir conteúdo
    const { rows: [conteudo] } = await client.query(`
      INSERT INTO conteudos
        (campanha_id, tipo_conteudo, canal, titulo,
         corpo_conteudo, formato_saida, geo_ou_humano)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING id`,
      [campanha.id, tipo_conteudo, canal, titulo,
       corpo_conteudo, formato_saida, geo_ou_humano]
    );

    // 3. Inserir aprovação (url será atualizada após upload)
    await client.query(`
      INSERT INTO aprovacoes
        (conteudo_id, aprovado_por, aprovado_em, adicionado_repo)
      VALUES ($1, $2, NOW(), false)`,
      [conteudo.id, usuario]
    );

    await client.query('COMMIT');

    // 4. Gerar nome do arquivo
    const nomeArquivo = gerarNomeArquivo(
      modulo, etapa_funil, canal, tipo_conteudo, titulo
    );

    // 5. Gerar .docx
    const docBuffer = await gerarDocx({
      titulo, modulo, etapa_funil, canal, tipo_conteudo,
      geo_ou_humano, persona_alvo, objetivo,
      corpo_conteudo, usuario, contexto_brief,
      campanha_id: campanha.id,
      conteudo_id: conteudo.id
    });

    // 6. Upload para o Google Drive
    let driveUrl = null;
    try {
      driveUrl = await uploadDrive(docBuffer, nomeArquivo);

      // Atualizar URL e flag no banco
      await pool.query(`
        UPDATE aprovacoes
        SET url_publicado = $1, adicionado_repo = true
        WHERE conteudo_id = $2`,
        [driveUrl, conteudo.id]
      );
    } catch (driveErr) {
      console.error('Erro no upload Drive (não bloqueia):', driveErr.message);
    }

    return res.json({
      ok:          true,
      campanha_id: campanha.id,
      conteudo_id: conteudo.id,
      arquivo:     `${nomeArquivo}.docx`,
      drive_url:   driveUrl,
      mensagem:    driveUrl
        ? `Conteúdo aprovado e salvo no Drive: ${nomeArquivo}.docx`
        : `Conteúdo aprovado e salvo no banco. Upload no Drive falhou — verifique as credenciais.`
    });

  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('Erro em /aprovar:', err);
    return res.status(500).json({ ok: false, erro: err.message });
  } finally {
    client.release();
  }
});

// ─────────────────────────────────────────────────────────────────────
// ROTA — GET /campanhas  (consulta para o dashboard)
// ─────────────────────────────────────────────────────────────────────
app.get('/campanhas', async (req, res) => {
  const { modulo, funil, canal, dias = 90 } = req.query;

  let where = [`c.criado_em >= NOW() - INTERVAL '${parseInt(dias, 10)} days'`];
  const params = [];

  if (modulo && modulo !== 'todos') {
    params.push(modulo);
    where.push(`c.modulo = $${params.length}`);
  }
  if (funil && funil !== 'todos') {
    params.push(funil);
    where.push(`c.etapa_funil = $${params.length}`);
  }
  if (canal && canal !== 'todos') {
    params.push(canal);
    where.push(`co.canal = $${params.length}`);
  }

  const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';

  try {
    const { rows } = await pool.query(`
      SELECT
        c.id           AS campanha_id,
        c.modulo,
        c.etapa_funil,
        c.status,
        c.criado_em,
        co.id          AS conteudo_id,
        co.tipo_conteudo,
        co.canal,
        co.titulo,
        co.geo_ou_humano,
        a.aprovado_por,
        a.aprovado_em,
        a.url_publicado,
        a.adicionado_repo
      FROM campanhas c
      LEFT JOIN conteudos co ON co.campanha_id = c.id
      LEFT JOIN aprovacoes a ON a.conteudo_id  = co.id
      ${whereClause}
      ORDER BY c.criado_em DESC
      LIMIT 500`,
      params
    );
    res.json({ ok: true, total: rows.length, dados: rows });
  } catch (err) {
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// ROTA — GET /stats  (resumo para o dashboard)
// ─────────────────────────────────────────────────────────────────────
app.get('/stats', async (req, res) => {
  const dias = parseInt(req.query.dias || 90, 10);
  try {
    const { rows } = await pool.query(`
      SELECT
        c.modulo,
        c.etapa_funil,
        co.canal,
        co.tipo_conteudo,
        co.geo_ou_humano,
        COUNT(c.id)                                     AS total_gerados,
        COUNT(a.id)                                     AS total_aprovados,
        COUNT(a.publicado_em)                           AS total_publicados,
        COUNT(CASE WHEN a.adicionado_repo THEN 1 END)   AS no_drive
      FROM campanhas c
      LEFT JOIN conteudos co ON co.campanha_id = c.id
      LEFT JOIN aprovacoes a ON a.conteudo_id  = co.id
      WHERE c.criado_em >= NOW() - INTERVAL '${dias} days'
      GROUP BY c.modulo, c.etapa_funil, co.canal, co.tipo_conteudo, co.geo_ou_humano
      ORDER BY total_gerados DESC`,
      []
    );
    res.json({ ok: true, dados: rows });
  } catch (err) {
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// HELPER — NOME DO ARQUIVO
// ─────────────────────────────────────────────────────────────────────
function gerarNomeArquivo(modulo, funil, canal, tipo, titulo) {
  const data = new Date().toISOString().slice(0, 10);

  const slug = (titulo || 'sem-titulo')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .slice(0, 50);

  const canalSlug  = (canal  || '').replace(/_/g, '-');
  const moduloSlug = (modulo || '').replace(/\s+/g, '-');
  const funilSlug  = (funil  || '').replace(/\s+/g, '-');

  return `${data}_${moduloSlug}_${funilSlug}_${canalSlug}_${tipo}_${slug}`;
}

// ─────────────────────────────────────────────────────────────────────
// HELPER — GERAR DOCX
// ─────────────────────────────────────────────────────────────────────
async function gerarDocx(dados) {
  const COR_ROXO   = '5B4FBE';
  const COR_TEXTO  = '1A1A2E';
  const COR_CINZA  = 'F7F7F7';
  const COR_BRANCO = 'FFFFFF';
  const COR_BORDA  = 'CCCCCC';

  const borda1 = { style: BorderStyle.SINGLE, size: 1, color: COR_BORDA };
  const bordas = { top: borda1, bottom: borda1, left: borda1, right: borda1 };

  const metaRow = (label, valor) => new TableRow({
    children: [
      new TableCell({
        borders: bordas,
        width: { size: 2600, type: WidthType.DXA },
        shading: { fill: COR_CINZA, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: label, font: 'Arial', size: 20, bold: true, color: COR_ROXO
          })]
        })]
      }),
      new TableCell({
        borders: bordas,
        width: { size: 6426, type: WidthType.DXA },
        shading: { fill: COR_BRANCO, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: String(valor || '—'), font: 'Arial', size: 20, color: COR_TEXTO
          })]
        })]
      })
    ]
  });

  // Converte o corpo em parágrafos, mantendo quebras de linha
  const linhasCorpo = (dados.corpo_conteudo || '')
    .split('\n')
    .map(linha => new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({
        text: linha, font: 'Arial', size: 22, color: COR_TEXTO
      })]
    }));

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        // ── Título ──
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({
            text: dados.titulo || 'Sem título',
            font: 'Arial', size: 36, bold: true, color: COR_ROXO
          })]
        }),

        // ── Tabela de metadados ──
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2600, 6426],
          rows: [
            metaRow('Módulo',          dados.modulo),
            metaRow('Etapa do funil',  dados.etapa_funil),
            metaRow('Canal',           dados.canal),
            metaRow('Tipo',            dados.tipo_conteudo),
            metaRow('GEO ou humano',   dados.geo_ou_humano),
            metaRow('Persona-alvo',    dados.persona_alvo),
            metaRow('Objetivo',        dados.objetivo),
            metaRow('Contexto/brief',  dados.contexto_brief),
            metaRow('Gerado por',      dados.usuario),
            metaRow('ID campanha',     String(dados.campanha_id)),
            metaRow('ID conteúdo',     String(dados.conteudo_id)),
            metaRow('Aprovado em',     new Date().toLocaleString('pt-BR')),
          ]
        }),

        // ── Espaço ──
        new Paragraph({ spacing: { before: 300, after: 0 }, children: [] }),

        // ── Divisor ──
        new Paragraph({
          spacing: { before: 0, after: 240 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: COR_ROXO, space: 1 }
          },
          children: []
        }),

        // ── Subtítulo do conteúdo ──
        new Paragraph({
          spacing: { before: 0, after: 160 },
          children: [new TextRun({
            text: 'Conteúdo aprovado',
            font: 'Arial', size: 28, bold: true, color: COR_TEXTO
          })]
        }),

        // ── Corpo ──
        ...linhasCorpo
      ]
    }]
  });

  return await Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────────────────
// HELPER — UPLOAD GOOGLE DRIVE
// ─────────────────────────────────────────────────────────────────────
async function uploadDrive(buffer, nomeArquivo) {
  const credenciais = process.env.GOOGLE_SERVICE_ACCOUNT;
  const folderId    = process.env.DRIVE_FOLDER_ID;

  if (!credenciais || !folderId) {
    throw new Error(
      'Variáveis GOOGLE_SERVICE_ACCOUNT ou DRIVE_FOLDER_ID não configuradas.'
    );
  }

  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(credenciais),
    scopes: ['https://www.googleapis.com/auth/drive']
  });

  const drive  = google.drive({ version: 'v3', auth });
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const { data } = await drive.files.create({
    requestBody: {
      name:     `${nomeArquivo}.docx`,
      parents:  [folderId],
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    },
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      body: stream
    },
    fields: 'id,webViewLink'
  });

  return data.webViewLink;
}

// ─────────────────────────────────────────────────────────────────────
// START
// ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ Serviço de aprovação rodando na porta ${PORT}`);
  console.log(`   DB configurado: ${process.env.DATABASE_URL ? 'sim' : 'NÃO'}`);
  console.log(`   Drive folder:   ${process.env.DRIVE_FOLDER_ID || 'NÃO configurado'}`);
  console.log(`   Service account:${process.env.GOOGLE_SERVICE_ACCOUNT ? ' sim' : ' NÃO configurado'}`);
});
