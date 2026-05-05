/**
 * NKSW Shopify Sync — salva pedidos do mês no Vercel Blob
 *
 * GET /api/shopify-sync?action=full&month=2026-05&secret=nksw2025
 *   → baixa todos os pedidos do mês do Shopify e salva no Blob
 *
 * GET /api/shopify-sync?action=update&month=2026-05&secret=nksw2025
 *   → busca só pedidos atualizados desde o último sync e faz merge
 *
 * Variáveis de ambiente necessárias:
 *   SHOPIFY_STORE, SHOPIFY_CLIENT_ID, SHOPIFY_CLIENT_SECRET
 *   BLOB_READ_WRITE_TOKEN (adicionado automaticamente pelo Vercel Blob)
 *   ADMIN_SECRET (opcional, padrão: nksw2025)
 */

import { put, head } from '@vercel/blob';

const SHOPIFY_STORE  = process.env.SHOPIFY_STORE;
const CLIENT_ID      = process.env.SHOPIFY_CLIENT_ID;
const CLIENT_SECRET  = process.env.SHOPIFY_CLIENT_SECRET;
const API_VERSION    = '2024-01';

// Campos necessários para o BI — evita baixar payload completo
const FIELDS = [
  'id', 'order_number', 'created_at', 'updated_at',
  'financial_status', 'fulfillment_status', 'cancelled_at', 'cancel_reason',
  'total_price', 'subtotal_price', 'total_discounts',
  'total_shipping_price_set',
  'line_items', 'customer',
  'billing_address', 'shipping_address',
  'fulfillments', 'discount_codes',
  'payment_gateway', 'tags', 'note',
].join(',');

// Cache do token OAuth em memória (válido 23h, reseta entre cold starts)
let _token    = null;
let _tokenExp = 0;

async function getToken() {
  if (_token && Date.now() < _tokenExp) return _token;
  const res = await fetch(`https://${SHOPIFY_STORE}/admin/oauth/access_token`, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify({ grant_type: 'client_credentials', client_id: CLIENT_ID, client_secret: CLIENT_SECRET }),
  });
  if (!res.ok) throw new Error(`Shopify token error: ${res.status}`);
  const data = await res.json();
  if (!data.access_token) throw new Error('Token não retornado');
  _token    = data.access_token;
  _tokenExp = Date.now() + 23 * 60 * 60 * 1000;
  return _token;
}

async function shopifyGet(urlOrPath) {
  const token = await getToken();
  const url   = urlOrPath.startsWith('http')
    ? urlOrPath
    : `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${urlOrPath}`;
  const res = await fetch(url, { headers: { 'X-Shopify-Access-Token': token } });
  if (!res.ok) throw new Error(`Shopify ${res.status}: ${await res.text()}`);
  return { json: await res.json(), link: res.headers.get('Link') || '' };
}

// Extrai URL da próxima página do header Link
function nextPageUrl(linkHeader) {
  const m = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
  return m ? m[1] : null;
}

// Busca todos os pedidos com paginação automática
async function fetchAllOrders(initialPath) {
  const orders = [];
  let url = initialPath;
  while (url) {
    const { json, link } = await shopifyGet(url);
    orders.push(...(json.orders || []));
    url = nextPageUrl(link);
    // Pequena pausa para não estourar rate limit (40 req burst, 2/s)
    if (url) await new Promise(r => setTimeout(r, 250));
  }
  return orders;
}

// Retorna datas de início/fim do mês no fuso de Brasília (UTC-3)
function monthBounds(month) {
  const [y, m] = month.split('-').map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  return {
    start: `${month}-01T00:00:00-03:00`,
    end:   `${month}-${String(lastDay).padStart(2, '0')}T23:59:59-03:00`,
  };
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action, month, secret } = req.query;
  const SECRET = process.env.ADMIN_SECRET || 'nksw2025';

  if (secret !== SECRET) return res.status(401).json({ error: 'Não autorizado' });

  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return res.status(400).json({ error: 'Parâmetro month obrigatório (formato: YYYY-MM)' });
  }

  const blobName = `shopify-orders-${month}.json`;
  const bounds   = monthBounds(month);

  try {
    // ── FULL SYNC ──────────────────────────────────────────────────────────────
    if (action === 'full') {
      const orders = await fetchAllOrders(
        `/orders.json?status=any&created_at_min=${bounds.start}&created_at_max=${bounds.end}&limit=250&fields=${FIELDS}`
      );

      const payload = JSON.stringify({
        ok:         true,
        month,
        synced_at:  new Date().toISOString(),
        total:      orders.length,
        orders,
      });

      await put(blobName, payload, {
        access:            'public',
        contentType:       'application/json',
        addRandomSuffix:   false,
      });

      return res.status(200).json({ ok: true, action: 'full', month, total: orders.length, synced_at: new Date().toISOString() });
    }

    // ── UPDATE INCREMENTAL ────────────────────────────────────────────────────
    if (action === 'update') {
      // Lê blob existente
      let existingOrders = [];
      let lastSyncedAt   = null;

      try {
        const existing = await head(blobName);
        if (existing) {
          const r      = await fetch(existing.url);
          const parsed = await r.json();
          existingOrders = parsed.orders    || [];
          lastSyncedAt   = parsed.synced_at || null;
        }
      } catch (_) {}

      if (!lastSyncedAt) {
        return res.status(400).json({
          error: 'Full sync ainda não realizado para este mês.',
          hint:  `/api/shopify-sync?action=full&month=${month}&secret=${SECRET}`,
        });
      }

      // Busca pedidos atualizados desde o último sync
      const updatedOrders = await fetchAllOrders(
        `/orders.json?status=any&updated_at_min=${lastSyncedAt}&created_at_min=${bounds.start}&created_at_max=${bounds.end}&limit=250&fields=${FIELDS}`
      );

      // Merge: ordena existentes em mapa por ID, sobrescreve os atualizados
      const map = Object.fromEntries(existingOrders.map(o => [o.id, o]));
      updatedOrders.forEach(o => { map[o.id] = o; });
      const merged = Object.values(map).sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

      const payload = JSON.stringify({
        ok:        true,
        month,
        synced_at: new Date().toISOString(),
        total:     merged.length,
        orders:    merged,
      });

      await put(blobName, payload, {
        access:          'public',
        contentType:     'application/json',
        addRandomSuffix: false,
      });

      return res.status(200).json({
        ok:        true,
        action:    'update',
        month,
        updated:   updatedOrders.length,
        total:     merged.length,
        synced_at: new Date().toISOString(),
      });
    }

    return res.status(400).json({ error: 'action deve ser "full" ou "update"' });

  } catch (err) {
    console.error('[shopify-sync]', err);
    return res.status(500).json({ error: err.message });
  }
}
