/**
 * NKSW — Inventory Compare
 * Retorna mapa variantId → estoque Shopify para conferência.
 * GET /api/inventory-compare
 */

const SHOPIFY_STORE = process.env.SHOPIFY_STORE;
const SHOPIFY_TOKEN = process.env.SHOPIFY_ADMIN_TOKEN;
const API_VERSION   = '2024-01';
const BASE          = `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}`;

const HEADERS = {
  'X-Shopify-Access-Token': SHOPIFY_TOKEN,
  'Content-Type': 'application/json',
};

async function getAllVariants() {
  const inventory = {};
  let url = `${BASE}/products.json?limit=250&fields=variants&status=active`;

  while (url) {
    const res  = await fetch(url, { headers: HEADERS });
    if (!res.ok) throw new Error(`Shopify ${res.status}`);

    const data = await res.json();
    for (const p of (data.products || [])) {
      for (const v of (p.variants || [])) {
        inventory[String(v.id)] = v.inventory_quantity ?? 0;
      }
    }

    // Paginação via Link header
    const link = res.headers.get('link') || '';
    const next = link.match(/<([^>]+)>;\s*rel="next"/);
    url = next ? next[1] : null;
  }

  return inventory;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET')     return res.status(405).json({ error: 'Método não permitido' });

  try {
    const inventory = await getAllVariants();
    return res.status(200).json({
      ok:        true,
      total:     Object.keys(inventory).length,
      inventory,
    });
  } catch (e) {
    console.error('[inventory-compare]', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
