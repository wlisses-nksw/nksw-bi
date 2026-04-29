// api/notify-admin.js
// Painel simples para ver todas as solicitações de aviso pendentes
//
// Uso:
//   Todos os inscritos de uma variante:
//     GET /api/notify-admin?token=SEU_SECRET&variant_id=VARIANT_ID
//   Todos os produtos com inscritos (busca lenta — use com moderação):
//     GET /api/notify-admin?token=SEU_SECRET&all=1

const STORE        = 'lofty-fy.myshopify.com';
const SHOPIFY_BASE = `https://${STORE}/admin/api/2024-01`;

export default async function handler(req, res) {
  // Autenticação básica por token
  const { token, variant_id, all } = req.query;
  const SECRET = process.env.ADMIN_SECRET || 'nksw2025';

  if (token !== SECRET) {
    return res.status(401).json({ error: 'Não autorizado' });
  }

  const TOKEN = process.env.SHOPIFY_TOKEN;
  if (!TOKEN) return res.status(500).json({ error: 'SHOPIFY_TOKEN não configurado' });

  try {
    // ── Modo 1: variante específica ──────────────────────────────────
    if (variant_id) {
      const metaRes = await fetch(
        `${SHOPIFY_BASE}/variants/${variant_id}/metafields.json?namespace=avise_me&key=subscribers`,
        { headers: { 'X-Shopify-Access-Token': TOKEN } }
      );
      const metaData = await metaRes.json();

      let subscribers = [];
      if (metaData.metafields && metaData.metafields.length > 0) {
        try { subscribers = JSON.parse(metaData.metafields[0].value); } catch (_) {}
      }

      // Resposta HTML legível
      const rows = subscribers.map(s => `
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;">${s.email}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#666;">${s.variant_title || '—'}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#999;font-size:12px;">${new Date(s.registered_at).toLocaleString('pt-BR')}</td>
        </tr>`).join('');

      return res.setHeader('Content-Type', 'text/html; charset=utf-8').status(200).send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width,initial-scale=1">
          <title>Avise-me — Variant ${variant_id}</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 32px; background: #f9f9f9; color: #000; }
            h1 { font-size: 20px; margin: 0 0 4px; }
            p { color: #666; font-size: 13px; margin: 0 0 24px; }
            table { width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
            th { text-align: left; padding: 10px 12px; background: #000; color: #fff; font-size: 12px; letter-spacing: .06em; text-transform: uppercase; }
            tr:hover td { background: #f5f5f5; }
            .badge { display: inline-block; background: #000; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 12px; margin-left: 8px; vertical-align: middle; }
          </style>
        </head>
        <body>
          <h1>Avise-me — Inscritos <span class="badge">${subscribers.length}</span></h1>
          <p>Variant ID: ${variant_id}</p>
          ${subscribers.length === 0
            ? '<p style="color:#999;font-style:italic;">Nenhum inscrito nesta variante.</p>'
            : `<table>
                <thead><tr>
                  <th>Email</th>
                  <th>Variante</th>
                  <th>Data</th>
                </tr></thead>
                <tbody>${rows}</tbody>
              </table>`}
        </body>
        </html>
      `);
    }

    // ── Modo 2: JSON com dados da variante ───────────────────────────
    return res.status(400).json({
      error: 'Informe variant_id como parâmetro. Ex: /api/notify-admin?token=X&variant_id=123'
    });

  } catch (err) {
    console.error('[notify-admin]', err);
    return res.status(500).json({ error: err.message });
  }
}
