// api/webhook-restock.js
// Recebe webhook do Shopify: inventory_levels/update
// Quando estoque voltar > 0, envia email para todos os inscritos da variante
//
// Configurar no Shopify:
//   Admin → Configurações → Notificações → Webhooks → Criar webhook
//   Evento:  inventory_levels/update
//   Formato: JSON
//   URL:     https://nksw-bi.vercel.app/api/webhook-restock

const STORE        = 'lofty-fy.myshopify.com';
const SHOPIFY_BASE = `https://${STORE}/admin/api/2024-01`;
const STORE_URL    = 'https://www.nakedsw.com.br';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const TOKEN  = process.env.SHOPIFY_TOKEN;
  const RESEND = process.env.RESEND_API_KEY;

  if (!TOKEN) return res.status(500).json({ error: 'SHOPIFY_TOKEN não configurado' });

  try {
    const { inventory_item_id, available } = req.body || {};

    // Só processa quando estoque voltar positivo
    if (!available || available <= 0) {
      return res.status(200).json({ skipped: 'sem estoque' });
    }

    // ── 1. Encontra a variante pelo inventory_item_id ────────────────
    const varRes = await fetch(
      `${SHOPIFY_BASE}/variants.json?inventory_item_ids=${inventory_item_id}&limit=1`,
      { headers: { 'X-Shopify-Access-Token': TOKEN } }
    );
    const varData = await varRes.json();

    if (!varData.variants || varData.variants.length === 0) {
      return res.status(200).json({ skipped: 'variante não encontrada' });
    }

    const variant      = varData.variants[0];
    const variant_id   = variant.id;
    const product_id   = variant.product_id;
    const variantLabel = [variant.option1, variant.option2, variant.option3]
      .filter(Boolean).join(' / ');

    // ── 2. Busca inscritos no metafield da variante ──────────────────
    const metaRes = await fetch(
      `${SHOPIFY_BASE}/variants/${variant_id}/metafields.json?namespace=avise_me&key=subscribers`,
      { headers: { 'X-Shopify-Access-Token': TOKEN } }
    );
    const metaData = await metaRes.json();

    if (!metaData.metafields || metaData.metafields.length === 0) {
      return res.status(200).json({ skipped: 'sem inscritos' });
    }

    let subscribers = [];
    try { subscribers = JSON.parse(metaData.metafields[0].value); } catch (_) {}
    const metafieldId = metaData.metafields[0].id;

    if (subscribers.length === 0) {
      return res.status(200).json({ skipped: 'lista vazia' });
    }

    // ── 3. Dados do produto ──────────────────────────────────────────
    const prodRes = await fetch(
      `${SHOPIFY_BASE}/products/${product_id}.json?fields=title,handle,images`,
      { headers: { 'X-Shopify-Access-Token': TOKEN } }
    );
    const prodData  = await prodRes.json();
    const product   = prodData.product;
    const imageUrl  = product.images?.[0]?.src || '';
    const productUrl = `${STORE_URL}/products/${product.handle}?variant=${variant_id}`;

    // ── 4. Envia email para cada inscrito ────────────────────────────
    const sent   = [];
    const failed = [];

    if (RESEND) {
      for (const sub of subscribers) {
        try {
          const emailRes = await fetch('https://api.resend.com/emails', {
            method:  'POST',
            headers: { 'Authorization': `Bearer ${RESEND}`, 'Content-Type': 'application/json' },
            body:    JSON.stringify({
              from:    'Naked SW <novidades@nakedsw.com.br>',
              to:      [sub.email],
              subject: `✨ Voltou! ${product.title}${variantLabel ? ' — ' + variantLabel : ''}`,
              html: `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Produto de volta ao estoque</title>
</head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,Helvetica,sans-serif;-webkit-text-size-adjust:100%;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;">
    <tr><td align="center" style="padding:40px 16px;">
      <table width="560" cellpadding="0" cellspacing="0" style="background:#fff;max-width:560px;width:100%;">

        <!-- Header -->
        <tr><td style="padding:28px 40px 24px;border-bottom:1px solid #eee;">
          <p style="margin:0;font-size:13px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:#000;">NAKED SW</p>
        </td></tr>

        <!-- Imagem do produto -->
        ${imageUrl ? `
        <tr><td style="padding:0;">
          <img src="${imageUrl}" alt="${product.title}" width="560"
               style="display:block;width:100%;max-width:560px;height:auto;max-height:420px;object-fit:cover;">
        </td></tr>` : ''}

        <!-- Corpo -->
        <tr><td style="padding:40px 40px 36px;">
          <p style="margin:0 0 10px;font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:#999;">
            De volta ao estoque
          </p>
          <h1 style="margin:0 0 8px;font-size:26px;font-weight:700;line-height:1.2;color:#000;letter-spacing:.01em;">
            ${product.title}
          </h1>
          ${variantLabel ? `<p style="margin:0 0 28px;font-size:14px;color:#777;">${variantLabel}</p>` : '<br>'}
          <p style="margin:0 0 32px;font-size:15px;line-height:1.7;color:#444;">
            Você pediu para ser avisada e o produto voltou ao estoque!<br>
            Corre antes de esgotar novamente.
          </p>
          <a href="${productUrl}"
             style="display:inline-block;background:#000;color:#fff;padding:16px 40px;
                    text-decoration:none;font-size:12px;font-weight:700;
                    letter-spacing:.1em;text-transform:uppercase;">
            COMPRAR AGORA →
          </a>
        </td></tr>

        <!-- Footer -->
        <tr><td style="padding:20px 40px 28px;border-top:1px solid #eee;">
          <p style="margin:0;font-size:11px;color:#bbb;line-height:1.6;">
            Você recebeu este email porque solicitou aviso de reposição em
            <a href="${STORE_URL}" style="color:#bbb;">nakedsw.com.br</a>.
          </p>
        </td></tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`
            })
          });

          if (emailRes.ok) {
            sent.push(sub.email);
          } else {
            const err = await emailRes.json();
            console.error('[webhook-restock] email erro para', sub.email, err);
            failed.push(sub.email);
          }
        } catch (e) {
          console.error('[webhook-restock] exceção para', sub.email, e);
          failed.push(sub.email);
        }
      }
    }

    // ── 5. Limpa a lista após notificar (mantém os que falharam) ─────
    const remaining = failed.length > 0
      ? subscribers.filter(s => failed.includes(s.email))
      : [];

    await fetch(`${SHOPIFY_BASE}/variants/${variant_id}/metafields/${metafieldId}.json`, {
      method:  'PUT',
      headers: { 'X-Shopify-Access-Token': TOKEN, 'Content-Type': 'application/json' },
      body:    JSON.stringify({
        metafield: { id: metafieldId, value: JSON.stringify(remaining), type: 'json' }
      })
    });

    return res.status(200).json({
      success:  true,
      product:  product.title,
      variant:  variantLabel,
      notified: sent.length,
      failed:   failed.length,
      emails:   sent
    });

  } catch (err) {
    console.error('[webhook-restock]', err);
    return res.status(500).json({ error: err.message });
  }
}
