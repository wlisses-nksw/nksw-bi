/**
 * NKSW Products IDs API — Shopify Admin
 * Retorna todos os produtos com product_id e variant_id.
 *
 * GET /api/products-ids
 * GET /api/products-ids?format=csv  → retorna CSV para download
 *
 * Variáveis de ambiente necessárias no Vercel:
 *   SHOPIFY_STORE         → lofty-fy.myshopify.com
 *   SHOPIFY_CLIENT_ID     → client id do app Claude-voll
 *   SHOPIFY_CLIENT_SECRET → client secret do app Claude-voll
 */

const SHOPIFY_STORE         = process.env.SHOPIFY_STORE;
const SHOPIFY_CLIENT_ID     = process.env.SHOPIFY_CLIENT_ID;
const SHOPIFY_CLIENT_SECRET = process.env.SHOPIFY_CLIENT_SECRET;
const API_VERSION           = "2024-01";

let cachedToken    = null;
let tokenExpiresAt = 0;

async function getToken() {
  if (cachedToken && Date.now() < tokenExpiresAt) return cachedToken;
  const res = await fetch(`https://${SHOPIFY_STORE}/admin/oauth/access_token`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      grant_type:    "client_credentials",
      client_id:     SHOPIFY_CLIENT_ID,
      client_secret: SHOPIFY_CLIENT_SECRET,
    }),
  });
  if (!res.ok) throw new Error(`Token error: ${res.status}`);
  const data = await res.json();
  if (!data.access_token) throw new Error("Token não retornado");
  cachedToken    = data.access_token;
  tokenExpiresAt = Date.now() + 23 * 60 * 60 * 1000;
  return cachedToken;
}

async function fetchShopify(path) {
  const token = await getToken();
  const res = await fetch(`https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`, {
    headers: { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
  });
  if (!res.ok) throw new Error(`Shopify ${res.status}: ${await res.text()}`);
  return { data: await res.json(), headers: res.headers };
}

async function getAllProducts() {
  const rows = [];
  let url = `/products.json?limit=250&fields=id,title,status,variants`;

  while (url) {
    const { data, headers } = await fetchShopify(url);
    const products = data.products || [];

    for (const product of products) {
      for (const variant of product.variants || []) {
        rows.push({
          product_id:   product.id,
          variant_id:   variant.id,
          produto:      product.title,
          variante:     variant.title === "Default Title" ? "" : variant.title,
          sku:          variant.sku || "",
          preco:        variant.price || "",
          estoque:      variant.inventory_quantity ?? "",
          status:       product.status,
        });
      }
    }

    // Paginação via Link header
    const linkHeader = headers.get("link") || "";
    const nextMatch  = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
    if (nextMatch) {
      // Extrai só o path+query da URL completa
      const nextUrl = new URL(nextMatch[1]);
      url = nextUrl.pathname.replace(`/admin/api/${API_VERSION}`, "") + nextUrl.search;
    } else {
      url = null;
    }
  }

  return rows;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Método não permitido" });

  try {
    const rows   = await getAllProducts();
    const format = req.query.format;

    if (format === "csv") {
      const header = "product_id,variant_id,produto,variante,sku,preco,estoque,status";
      const lines  = rows.map(r =>
        [r.product_id, r.variant_id,
         `"${r.produto.replace(/"/g,'""')}"`,
         `"${r.variante.replace(/"/g,'""')}"`,
         `"${r.sku}"`, r.preco, r.estoque, r.status
        ].join(",")
      );
      res.setHeader("Content-Type", "text/csv; charset=utf-8");
      res.setHeader("Content-Disposition", "attachment; filename=nksw_products_ids.csv");
      return res.status(200).send([header, ...lines].join("\n"));
    }

    return res.status(200).json({
      total_produtos: [...new Set(rows.map(r => r.product_id))].length,
      total_variantes: rows.length,
      produtos: rows,
    });

  } catch (error) {
    console.error(error.message);
    return res.status(500).json({ error: error.message });
  }
}
