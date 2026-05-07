'use strict';

const express      = require('express');
const cors         = require('cors');
const { Pool }     = require('pg');
const { google }   = require('googleapis');
const { Readable } = require('stream');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell,
  BorderStyle, WidthType, ShadingType
} = require('docx');

// ─────────────────────────────────────────────────────────────────────
// CONFIGURAÇÃO
// ─────────────────────────────────────────────────────────────────────
const app  = express();
const PORT = process.env.PORT || 3001;

// Pool com reconexão automática
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes('railway')
    ? { rejectUnauthorized: false }
    : false,
  max: 10,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 5000
});

pool.on('error', (err) => {
  console.error('Erro inesperado no pool do banco:', err.message);
});

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ─────────────────────────────────────────────────────────────────────
// HEALTH CHECK — Railway bate aqui para confirmar que o serviço subiu
// ─────────────────────────────────────────────────────────────────────
app.get('/health', async (req, res) => {
  const status = {
    ok:        true,
    ts:        new Date().toISOString(),
    db:        'nao testado',
    drive:     process.env.DRIVE_FOLDER_ID && process.env.DRIVE_FOLDER_ID !== 'placeholder'
                 ? 'configurado' : 'pendente',
    gauth:     process.env.GOOGLE_SERVICE_ACCOUNT && process.env.GOOGLE_SERVICE_ACCOUNT !== '{}'
                 ? 'configurado' : 'pendente'
  };

  try {
    await pool.query('SELECT 1');
    status.db = 'conectado';
  } catch (e) {
    status.db  = 'erro: ' + e.message;
    status.ok  = false;
  }

  res.status(status.ok ? 200 : 500).json(status);
});

// ─────────────────────────────────────────────────────────────────────
// POST /aprovar
// Body esperado:
//   modulo, etapa_funil, canal, tipo_conteudo (obrigatórios)
//   titulo, corpo_conteudo                   (obrigatórios)
//   geo_ou_humano, persona_alvo, objetivo    (opcionais)
//   usuario, contexto_brief, formato_saida  (opcionais)
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

  // Validação dos campos obrigatórios
  const obrigatorios = { modulo, etapa_funil, canal, tipo_conteudo, titulo, corpo_conteudo };
  const faltando = Object.entries(obrigatorios)
    .filter(([, v]) => !v || !String(v).trim())
    .map(([k]) => k);

  if (faltando.length) {
    return res.status(400).json({
      ok:   false,
      erro: `Campos obrigatorios ausentes: ${faltando.join(', ')}`
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

    // 3. Inserir aprovação
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
    let docBuffer;
    try {
      docBuffer = await gerarDocx({
        titulo, modulo, etapa_funil, canal, tipo_conteudo,
        geo_ou_humano, persona_alvo, objetivo, corpo_conteudo,
        usuario, contexto_brief,
        campanha_id: campanha.id,
        conteudo_id: conteudo.id
      });
    } catch (docErr) {
      console.error('Erro ao gerar docx (nao bloqueia):', docErr.message);
    }

    // 6. Upload para o Google Drive (não bloqueia se falhar)
    let driveUrl = null;
    if (docBuffer) {
      try {
        driveUrl = await uploadDrive(docBuffer, nomeArquivo);
        await pool.query(`
          UPDATE aprovacoes
          SET url_publicado = $1, adicionado_repo = true
          WHERE conteudo_id = $2`,
          [driveUrl, conteudo.id]
        );
      } catch (driveErr) {
        console.error('Upload Drive nao realizado:', driveErr.message);
      }
    }

    return res.json({
      ok:          true,
      campanha_id: campanha.id,
      conteudo_id: conteudo.id,
      arquivo:     `${nomeArquivo}.docx`,
      drive_url:   driveUrl,
      mensagem:    driveUrl
        ? `Conteudo aprovado e salvo no Drive: ${nomeArquivo}.docx`
        : `Conteudo aprovado e salvo no banco. Drive pendente de configuracao.`
    });

  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('Erro em /aprovar:', err.message);
    return res.status(500).json({ ok: false, erro: err.message });
  } finally {
    client.release();
  }
});

// ─────────────────────────────────────────────────────────────────────
// GET /stats — resumo para o dashboard
// Query params: dias (default 90), modulo, funil, canal
// ─────────────────────────────────────────────────────────────────────
app.get('/stats', async (req, res) => {
  const dias = Math.min(Math.max(parseInt(req.query.dias || 90, 10), 1), 365);

  try {
    const { rows } = await pool.query(`
      SELECT
        c.modulo,
        c.etapa_funil,
        co.canal,
        co.tipo_conteudo,
        co.geo_ou_humano,
        COUNT(DISTINCT c.id)                          AS total_campanhas,
        COUNT(DISTINCT co.id)                         AS total_conteudos,
        COUNT(DISTINCT a.id)                          AS total_aprovados,
        COUNT(DISTINCT CASE
          WHEN a.adicionado_repo = true THEN a.id
        END)                                          AS no_drive
      FROM campanhas c
      LEFT JOIN conteudos  co ON co.campanha_id = c.id
      LEFT JOIN aprovacoes a  ON a.conteudo_id  = co.id
      WHERE c.criado_em >= NOW() - ($1 || ' days')::INTERVAL
      GROUP BY
        c.modulo, c.etapa_funil,
        co.canal, co.tipo_conteudo, co.geo_ou_humano
      ORDER BY total_conteudos DESC`,
      [dias]
    );
    res.json({ ok: true, dias, total_linhas: rows.length, dados: rows });
  } catch (err) {
    console.error('Erro em /stats:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// GET /campanhas — lista paginada para o dashboard
// Query params: dias, modulo, funil, canal, pagina (default 1)
// ─────────────────────────────────────────────────────────────────────
app.get('/campanhas', async (req, res) => {
  const dias   = Math.min(Math.max(parseInt(req.query.dias   || 90,  10), 1), 365);
  const pagina = Math.max(parseInt(req.query.pagina || 1, 10), 1);
  const limite = 50;
  const offset = (pagina - 1) * limite;

  const condicoes = [`c.criado_em >= NOW() - ($1 || ' days')::INTERVAL`];
  const params    = [String(dias)];

  if (req.query.modulo && req.query.modulo !== 'todos') {
    params.push(req.query.modulo);
    condicoes.push(`c.modulo = $${params.length}`);
  }
  if (req.query.funil && req.query.funil !== 'todos') {
    params.push(req.query.funil);
    condicoes.push(`c.etapa_funil = $${params.length}`);
  }
  if (req.query.canal && req.query.canal !== 'todos') {
    params.push(req.query.canal);
    condicoes.push(`co.canal = $${params.length}`);
  }

  const where = condicoes.join(' AND ');

  try {
    const { rows } = await pool.query(`
      SELECT
        c.id            AS campanha_id,
        c.modulo,
        c.etapa_funil,
        c.status,
        c.criado_em,
        co.id           AS conteudo_id,
        co.tipo_conteudo,
        co.canal,
        co.titulo,
        co.geo_ou_humano,
        a.aprovado_por,
        a.aprovado_em,
        a.url_publicado,
        a.adicionado_repo
      FROM campanhas c
      LEFT JOIN conteudos  co ON co.campanha_id = c.id
      LEFT JOIN aprovacoes a  ON a.conteudo_id  = co.id
      WHERE ${where}
      ORDER BY c.criado_em DESC
      LIMIT $${params.length + 1}
      OFFSET $${params.length + 2}`,
      [...params, limite, offset]
    );
    res.json({ ok: true, pagina, limite, total: rows.length, dados: rows });
  } catch (err) {
    console.error('Erro em /campanhas:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// HELPER — NOME DO ARQUIVO
// Padrão: AAAA-MM-DD_Modulo_Funil_canal_tipo_slug-titulo
// ─────────────────────────────────────────────────────────────────────
function gerarNomeArquivo(modulo, funil, canal, tipo, titulo) {
  const data = new Date().toISOString().slice(0, 10);

  const slugify = (str) =>
    (str || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9\s-]/g, '')
      .trim()
      .replace(/\s+/g, '-');

  const slug = slugify(titulo).slice(0, 50);

  return [
    data,
    slugify(modulo),
    slugify(funil),
    slugify(canal),
    slugify(tipo),
    slug || 'sem-titulo'
  ].join('_');
}

// ─────────────────────────────────────────────────────────────────────
// HELPER — GERAR DOCX
// ─────────────────────────────────────────────────────────────────────
async function gerarDocx(dados) {
  const ROXO   = '5B4FBE';
  const TEXTO  = '1A1A2E';
  const CINZA  = 'F7F7F7';
  const BRANCO = 'FFFFFF';
  const BORDA  = 'CCCCCC';

  const b = { style: BorderStyle.SINGLE, size: 1, color: BORDA };
  const bordas = { top: b, bottom: b, left: b, right: b };

  const metaRow = (label, valor) => new TableRow({
    children: [
      new TableCell({
        borders: bordas,
        width: { size: 2600, type: WidthType.DXA },
        shading: { fill: CINZA, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: label,
            font: 'Arial', size: 20, bold: true, color: ROXO
          })]
        })]
      }),
      new TableCell({
        borders: bordas,
        width: { size: 6426, type: WidthType.DXA },
        shading: { fill: BRANCO, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: String(valor || '—'),
            font: 'Arial', size: 20, color: TEXTO
          })]
        })]
      })
    ]
  });

  const linhasCorpo = (dados.corpo_conteudo || '')
    .split('\n')
    .map(linha => new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({
        text: linha,
        font: 'Arial', size: 22, color: TEXTO
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
        // Título
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({
            text: dados.titulo || 'Sem titulo',
            font: 'Arial', size: 36, bold: true, color: ROXO
          })]
        }),

        // Tabela de metadados
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2600, 6426],
          rows: [
            metaRow('Modulo',         dados.modulo),
            metaRow('Etapa do funil', dados.etapa_funil),
            metaRow('Canal',          dados.canal),
            metaRow('Tipo',           dados.tipo_conteudo),
            metaRow('GEO ou humano',  dados.geo_ou_humano),
            metaRow('Persona-alvo',   dados.persona_alvo),
            metaRow('Objetivo',       dados.objetivo),
            metaRow('Brief/contexto', dados.contexto_brief),
            metaRow('Gerado por',     dados.usuario),
            metaRow('ID campanha',    String(dados.campanha_id)),
            metaRow('ID conteudo',    String(dados.conteudo_id)),
            metaRow('Aprovado em',    new Date().toLocaleString('pt-BR'))
          ]
        }),

        // Espaço
        new Paragraph({ spacing: { before: 280, after: 0 }, children: [] }),

        // Divisor
        new Paragraph({
          spacing: { before: 0, after: 240 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: ROXO, space: 1 }
          },
          children: []
        }),

        // Subtítulo
        new Paragraph({
          spacing: { before: 0, after: 160 },
          children: [new TextRun({
            text: 'Conteudo aprovado',
            font: 'Arial', size: 28, bold: true, color: TEXTO
          })]
        }),

        // Corpo
        ...linhasCorpo
      ]
    }]
  });

  return await Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────────────────
// HELPER — UPLOAD GOOGLE DRIVE
// Só executa se as variáveis estiverem configuradas de verdade
// ─────────────────────────────────────────────────────────────────────
async function uploadDrive(buffer, nomeArquivo) {
  const credRaw    = process.env.GOOGLE_SERVICE_ACCOUNT || '';
  const folderId   = process.env.DRIVE_FOLDER_ID        || '';

  // Não tenta subir se ainda está como placeholder
  if (!credRaw || credRaw === '{}' || !folderId || folderId === 'placeholder') {
    throw new Error('Drive pendente de configuracao — verifique GOOGLE_SERVICE_ACCOUNT e DRIVE_FOLDER_ID');
  }

  let credenciais;
  try {
    credenciais = JSON.parse(credRaw);
  } catch (e) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT nao e um JSON valido');
  }

  if (!credenciais.client_email || !credenciais.private_key) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT nao tem client_email ou private_key');
  }

  const auth = new google.auth.GoogleAuth({
    credentials: credenciais,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

  const drive  = google.drive({ version: 'v3', auth });
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const mimeType =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const { data } = await drive.files.create({
    requestBody: {
      name:    `${nomeArquivo}.docx`,
      parents: [folderId],
      mimeType
    },
    media: { mimeType, body: stream },
    fields: 'id,webViewLink'
  });

  return data.webViewLink;
}

// ─────────────────────────────────────────────────────────────────────
// START
// ─────────────────────────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Servico de aprovacao rodando na porta ${PORT}`);
  console.log(`DB:     ${process.env.DATABASE_URL            ? 'configurado' : 'NAO CONFIGURADO'}`);
  console.log(`Drive:  ${process.env.DRIVE_FOLDER_ID         ? process.env.DRIVE_FOLDER_ID : 'pendente'}`);
  console.log(`GAuth:  ${process.env.GOOGLE_SERVICE_ACCOUNT  ? 'configurado' : 'pendente'}`);
});
