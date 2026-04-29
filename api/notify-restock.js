// api/notify-restock.js
// Recebe a solicitação de aviso do cliente via modal no produto
// Armazena em metafield da variante + notifica admin por email
//
// Variáveis de ambiente necessárias no Vercel:
//   SHOPIFY_TOKEN   → token da loja com escopos: read_products, write_products
//   RESEND_API_KEY  → chave da API Resend (resend.com — plano gratuito: 3.000 emails/mês)
//   ADMIN_EMAIL     → email que recebe notificações do admin (ex: compras@nakedsw.com.br)
//   ADMIN_SECRET    → senha para a rota /api/notify-admin (ex: nksw2025)

const STORE = 'lofty-fy.myshopify.com';
const SHOPIFY_BASE = `https://${STORE}/admin/api/2024-01`;

export default async function handler(req, res) {
  // CORS — permite chamadas do Shopify storefront
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { email, variant_id, product_id, product_title, variant_title, product_handle } = req.body || {};

  // Validações
  if (!email || !variant_id || !product_id) {
    return res.status(400).json({ error: 'Campos obrigatórios: email, variant_id, product_id' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ error: 'Email inválido' });
  }

  const TOKEN   = process.env.SHOPIFY_TOKEN;
  const RESEND  = process.env.RESEND_API_KEY;
  const ADMIN   = process.env.ADMIN_EMAIL || 'contato@nakedsw.com.br';
  const SECRET  = process.env.ADMIN_SECRET || 'nksw2025';

  if (!TOKEN) return res.status(500).json({ error: 'SHOPIFY_TOKEN não configurado' });

  try {
    // ── 1. Lê inscritos atuais do metafield da variante ──────────────
    const metaRes = await fetch(
      `${SHOPIFY_BASE}/variants/${variant_id}/metafields.json?namespace=avise_me&key=subscribers`,
      { headers: { 'X-Shopify-Access-Token': TOKEN } }
    );
    const metaData = await metaRes.json();

    let subscribers    = [];
    let existingMetaId = null;

    if (metaData.metafields && metaData.metafields.length > 0) {
      existingMetaId = metaData.metafields[0].id;
      try { subscribers = JSON.parse(metaData.metafields[0].value); } catch (_) {}
    }

    // Email já cadastrado para essa variante?
    if (subscribers.some(s => s.email === email.toLowerCase())) {
      return res.status(200).json({ success: true, already_registered: true });
    }

    // ── 2. Adiciona novo inscrito ────────────────────────────────────
    subscribers.push({
      email:         email.toLowerCase().trim(),
      product_title,
      product_handle,
      variant_title,
      registered_at: new Date().toISOString()
    });

    const metaBody = {
      metafield: {
        namespace: 'avise_me',
        key:       'subscribers',
        type:      'json',
        value:     JSON.stringify(subscribers)
      }
    };

    if (existingMetaId) {
      await fetch(`${SHOPIFY_BASE}/variants/${variant_id}/metafields/${existingMetaId}.json`, {
        method:  'PUT',
        headers: { 'X-Shopify-Access-Token': TOKEN, 'Content-Type': 'application/json' },
        body:    JSON.stringify(metaBody)
      });
    } else {
      await fetch(`${SHOPIFY_BASE}/variants/${variant_id}/metafields.json`, {
        method:  'POST',
        headers: { 'X-Shopify-Access-Token': TOKEN, 'Content-Type': 'application/json' },
        body:    JSON.stringify(metaBody)
      });
    }

    // ── 3. Email de notificação para o admin ─────────────────────────
    if (RESEND) {
      const adminUrl = `https://nksw-bi.vercel.app/api/notify-admin?token=${SECRET}&variant_id=${variant_id}`;
      const shopifyUrl = `https://${STORE}/admin/products/${product_id}`;

      await fetch('https://api.resend.com/emails', {
        method:  'POST',
        headers: { 'Authorization': `Bearer ${RESEND}`, 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          from:    'Avise-me NKSW <avisome@nakedsw.com.br>',
          to:      [ADMIN],
          subject: `🔔 Avise-me: ${product_title} — ${variant_title} (${subscribers.length} inscritos)`,
          html: `
            <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:32px 24px;color:#000;">
              <h2 style="font-size:18px;margin:0 0 24px;border-bottom:2px solid #000;padding-bottom:16px;">
                📦 Nova solicitação de aviso de reposição
              </h2>
              <table style="width:100%;border-collapse:collapse;font-size:14px;">
                <tr><td style="padding:8px 0;color:#666;width:150px;vertical-align:top;">Produto</td>
                    <td style="padding:8px 0;font-weight:700;">${product_title}</td></tr>
                <tr><td style="padding:8px 0;color:#666;vertical-align:top;">Variante</td>
                    <td style="padding:8px 0;">${variant_title}</td></tr>
                <tr><td style="padding:8px 0;color:#666;vertical-align:top;">Email cliente</td>
                    <td style="padding:8px 0;"><a href="mailto:${email}" style="color:#000;">${email}</a></td></tr>
                <tr style="border-top:2px solid #000;">
                  <td style="padding:16px 0;color:#666;vertical-align:top;">Total inscritos</td>
                  <td style="padding:16px 0;font-size:28px;font-weight:700;line-height:1;">${subscribers.length}</td>
                </tr>
              </table>
              <div style="margin-top:28px;display:flex;gap:12px;">
                <a href="${shopifyUrl}" style="background:#000;color:#fff;padding:12px 24px;text-decoration:none;font-size:12px;font-weight:700;letter-spacing:.08em;display:inline-block;margin-right:12px;">
                  VER NO SHOPIFY →
                </a>
                <a href="${adminUrl}" style="background:#f7f7f7;color:#000;padding:12px 24px;text-decoration:none;font-size:12px;font-weight:700;letter-spacing:.08em;display:inline-block;border:1px solid #ddd;">
                  VER TODOS INSCRITOS →
                </a>
              </div>
            </div>
          `
        })
      });
    }

    return res.status(200).json({ success: true, subscribers_count: subscribers.length });

  } catch (err) {
    console.error('[notify-restock]', err);
    return res.status(500).json({ error: 'Erro interno. Tente novamente.' });
  }
}
