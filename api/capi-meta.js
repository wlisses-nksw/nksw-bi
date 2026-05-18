/**
 * NKSW — Meta CAPI (Conversions API) Server-Side
 *
 * Envia eventos de Purchase para a Meta com Advanced Matching
 * para atingir Match Quality 9.3+ e melhorar o algoritmo Andromeda.
 *
 * Endpoint interno: chamado pelo webhook Shopify (order/paid)
 * GET  /api/capi-meta?action=test          → valida conexão com Meta
 * POST /api/capi-meta                      → processa pedido e envia evento
 *
 * Variáveis de ambiente necessárias no Vercel:
 *   META_PIXEL_ID         → ID do Pixel Meta (ex: 1110791822804004)
 *   META_CAPI_TOKEN       → System User Access Token (não é o token de ads)
 *   SHOPIFY_WEBHOOK_SECRET → Segredo do webhook Shopify
 */

import crypto from 'crypto';

const PIXEL_ID    = process.env.META_PIXEL_ID;
const CAPI_TOKEN  = process.env.META_CAPI_TOKEN;
const API_VERSION = 'v19.0';
const CAPI_URL    = `https://graph.facebook.com/${API_VERSION}/${PIXEL_ID}/events`;

// ── Helpers de hashing ────────────────────────────────────────────────────

const sha256 = (val) => {
  if (!val) return null;
  const clean = String(val).trim().toLowerCase();
  if (!clean) return null;
  return crypto.createHash('sha256').update(clean).digest('hex');
};

const sha256Phone = (val) => {
  if (!val) return null;
  const digits = String(val).replace(/\D/g, '');
  if (!digits) return null;
  // Garantir código do país para Brasil
  const normalized = digits.startsWith('55') ? digits : '55' + digits;
  return crypto.createHash('sha256').update(normalized).digest('hex');
};

// ── Extrair fbclid do landing_site do pedido ─────────────────────────────

function extractFbclid(landingSite) {
  if (!landingSite) return null;
  try {
    const url = new URL('https://x.com' + (landingSite.startsWith('/') ? landingSite : '/' + landingSite));
    return url.searchParams.get('fbclid') || null;
  } catch { return null; }
}

function buildFbc(fbclid, createdAt) {
  if (!fbclid) return null;
  const ts = createdAt ? Math.floor(new Date(createdAt).getTime() / 1000) : Math.floor(Date.now() / 1000);
  return `fb.1.${ts}.${fbclid}`;
}

// ── Construir payload do evento Purchase ─────────────────────────────────

function buildPurchaseEvent(order) {
  const customer  = order.customer || {};
  const billing   = order.billing_address || order.shipping_address || {};
  const fbclid    = extractFbclid(order.landing_site);
  const eventTime = Math.floor(new Date(order.created_at).getTime() / 1000);
  const eventId   = `purchase-${order.order_number}-${order.id}`;

  // User data com todos os campos disponíveis para máximo Match Quality
  const userData = {
    em:          sha256(customer.email),
    ph:          sha256Phone(customer.phone || billing.phone),
    fn:          sha256(customer.first_name || billing.first_name),
    ln:          sha256(customer.last_name  || billing.last_name),
    ct:          sha256(billing.city),
    st:          sha256(billing.province_code || billing.province),
    zp:          sha256(billing.zip),
    country:     sha256(billing.country_code || 'BR'),
    external_id: String(customer.id || order.id),
  };

  if (fbclid) {
    userData.click_id = fbclid;
    userData.fbc      = buildFbc(fbclid, order.created_at);
  }

  // Remover campos nulos
  Object.keys(userData).forEach(k => { if (!userData[k]) delete userData[k]; });

  // Produtos comprados
  const contents = (order.line_items || []).map(item => ({
    id:                String(item.variant_id || item.product_id),
    quantity:          item.quantity,
    delivery_category: 'home_delivery',
    title:             item.title,
    price:             parseFloat(item.price),
  }));

  return {
    event_name:       'Purchase',
    event_time:       eventTime,
    event_id:         eventId,
    event_source_url: `https://nakedsw.com.br/`,
    action_source:    'website',
    user_data:        userData,
    custom_data: {
      value:        parseFloat(order.total_price),
      currency:     order.currency || 'BRL',
      content_type: 'product',
      num_items:    contents.length,
      order_id:     String(order.order_number),
      contents,
    },
  };
}

// ── Enviar para a Meta CAPI ───────────────────────────────────────────────

async function sendToCAPI(events, testCode = null) {
  const body = { data: events };
  if (testCode) body.test_event_code = testCode;

  const url = `${CAPI_URL}?access_token=${CAPI_TOKEN}`;
  const res  = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(JSON.stringify(data));
  return data;
}

// ── Verificar assinatura do webhook Shopify ──────────────────────────────

function verifyShopifyWebhook(req, body) {
  const hmac   = req.headers['x-shopify-hmac-sha256'];
  const secret = process.env.SHOPIFY_WEBHOOK_SECRET;
  if (!hmac || !secret) return false;
  const expected = crypto.createHmac('sha256', secret).update(body).digest('base64');
  return crypto.timingSafeEqual(Buffer.from(hmac), Buffer.from(expected));
}

// ── Handler principal ─────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Shopify-Hmac-SHA256');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // ── GET: teste de conexão ─────────────────────────────────────────────
  if (req.method === 'GET' && req.query.action === 'test') {
    if (!PIXEL_ID || !CAPI_TOKEN) {
      return res.status(500).json({ ok: false, error: 'META_PIXEL_ID ou META_CAPI_TOKEN não configurados nas env vars do Vercel' });
    }
    // Enviar evento de teste
    try {
      const testEvent = {
        event_name:    'Purchase',
        event_time:    Math.floor(Date.now() / 1000),
        event_id:      `test-${Date.now()}`,
        action_source: 'website',
        user_data:     { em: sha256('test@nakedsw.com.br'), country: sha256('br') },
        custom_data:   { value: 100, currency: 'BRL' },
      };
      const result = await sendToCAPI([testEvent], req.query.test_code || 'TEST');
      return res.status(200).json({ ok: true, pixel_id: PIXEL_ID, result });
    } catch(e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  // ── POST: processar pedido Shopify ────────────────────────────────────
  if (req.method === 'POST') {
    // Ler body raw para verificação
    const buffers = [];
    for await (const chunk of req) buffers.push(chunk);
    const rawBody = Buffer.concat(buffers).toString('utf8');

    // Verificar assinatura (para webhook Shopify)
    const shopifyTopic = req.headers['x-shopify-topic'];
    if (shopifyTopic && !verifyShopifyWebhook(req, rawBody)) {
      return res.status(401).json({ ok: false, error: 'Assinatura inválida' });
    }

    let order;
    try { order = JSON.parse(rawBody); } catch {
      return res.status(400).json({ ok: false, error: 'JSON inválido' });
    }

    if (!PIXEL_ID || !CAPI_TOKEN) {
      console.error('[capi-meta] META_PIXEL_ID ou META_CAPI_TOKEN não configurados');
      return res.status(500).json({ ok: false, error: 'Configuração ausente' });
    }

    try {
      const event  = buildPurchaseEvent(order);
      const result = await sendToCAPI([event]);

      const matchFields = Object.keys(event.user_data).length;
      console.log(`[capi-meta] #${order.order_number} enviado · ${matchFields} campos de matching · fbclid=${!!event.user_data.click_id}`);

      return res.status(200).json({
        ok:           true,
        order:        order.order_number,
        match_fields: matchFields,
        fbclid:       !!event.user_data.click_id,
        capi_result:  result,
      });
    } catch(e) {
      console.error('[capi-meta] erro:', e.message);
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  return res.status(405).json({ error: 'Método não permitido' });
}
