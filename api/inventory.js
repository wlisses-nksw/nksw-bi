/**
 * NKSW Inventory API — Shopify
 * Verifica disponibilidade de estoque por produto e/ou tamanho.
 *
 * Parâmetros:
 *   ?product=top-rina-strata   → handle do produto (slug da URL)
 *   ?size=M                    → tamanho específico (PP, P, M, G, GG)
 *   ?q=brisa                   → busca por nome parcial do produto
 */

const SHOPIFY_STORE = process.env.SHOPIFY_STORE;
const SHOPIFY_TOKEN = process.env.SHOPIFY_ADMIN_TOKEN;
const API_VERSION   = "2024-01";

async function fetchShopify(path) {
  const url = `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`;
  const res = await fetch(url, {
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_TOKEN,
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) throw new Error(`Shopify ${res.status}`);
  return res.json();
}

function formatInventory(product, sizeFilter) {
  const variantes = product.variants
    .filter((v) => !sizeFilter || v.option1?.toUpperCase() === sizeFilter.toUpperCase())
    .map((v) => ({
      tamanho: v.option1,
      cor: v.option2 || null,
      disponivel: v.inventory_quantity > 0,
      estoque: v.inventory_quantity,
      preco: `R$ ${parseFloat(v.price).toFixed(2).replace(".", ",")}`,
    }));

  const algumDisponivel = variantes.some((v) => v.disponivel);
  const tamanhosDisponiveis = variantes.filter((v) => v.disponivel).map((v) => v.tamanho);

  return {
    produto: product.title,
    handle: product.handle,
    url: `https://www.nakedsw.com.br/products/${product.handle}`,
    disponivel: algumDisponivel,
    tamanhos_disponiveis: tamanhosDisponiveis,
    variantes,
  };
}

// ───────────── Avise-me (back-in-stock) ─────────────
// Consolidado nesta função para respeitar o limite de 12 Serverless Functions do plano Hobby.
// Roteado por um rewrite em vercel.json: /api/notify-customer -> /api/inventory?fn=notify
function nowBR() {
  return new Intl.DateTimeFormat("pt-BR", {
    timeZone: "America/Sao_Paulo",
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  }).format(new Date());
}

function mergeTags(existing, additions) {
  const set = new Set((existing || "").split(",").map((t) => t.trim()).filter(Boolean));
  additions.forEach((t) => t && set.add(t));
  return Array.from(set).join(", ");
}

async function shopifyWrite(path, method, body) {
  const r = await fetch(`https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`, {
    method,
    headers: { "X-Shopify-Access-Token": SHOPIFY_TOKEN, "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await r.text();
  let json = null; try { json = text ? JSON.parse(text) : null; } catch { /* ignore */ }
  if (!r.ok) throw new Error(`Shopify ${r.status}: ${text.slice(0, 300)}`);
  return json;
}

async function notifyAviseMe(req, res) {
  if (req.method !== "POST") return res.status(405).json({ success: false, error: "Método não permitido" });
  if (!SHOPIFY_STORE || !SHOPIFY_TOKEN) return res.status(500).json({ success: false, error: "Credenciais Shopify não configuradas" });
  try {
    let body = req.body;
    if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
    body = body || {};
    const email        = String(body.email || "").trim().toLowerCase();
    const variantId    = String(body.variant_id || "").trim();
    const productTitle = String(body.product_title || "").trim();
    const variantTitle = String(body.variant_title || "").trim();
    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return res.status(400).json({ success: false, error: "E-mail inválido" });
    }
    const descr    = [productTitle, variantTitle].filter(Boolean).join(" · ");
    const noteLine = `[AVISE-ME] ${descr}${variantId ? ` | Variant ID: ${variantId}` : ""} | ${nowBR()}`;
    const newTags  = ["avise-me"];
    if (variantId) newTags.push(`avise-me:${variantId}`);

    const search   = await shopifyWrite(`/customers/search.json?query=${encodeURIComponent("email:" + email)}`, "GET");
    const existing = (search && search.customers && search.customers[0]) || null;

    if (existing) {
      const note = existing.note ? `${existing.note}\n${noteLine}` : noteLine;
      const tags = mergeTags(existing.tags, newTags);
      await shopifyWrite(`/customers/${existing.id}.json`, "PUT", { customer: { id: existing.id, tags, note } });
      return res.status(200).json({ success: true, created: false });
    }
    await shopifyWrite(`/customers.json`, "POST", { customer: { email, tags: newTags.join(", "), note: noteLine } });
    return res.status(200).json({ success: true, created: true });
  } catch (err) {
    return res.status(500).json({ success: false, error: String(err.message || err) });
  }
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();

  // Avise-me roteado para cá via rewrite (?fn=notify)
  if (req.query.fn === "notify") return notifyAviseMe(req, res);

  if (req.method !== "GET") return res.status(405).json({ error: "Método não permitido" });

  const { product: handle, size, q } = req.query;

  if (!handle && !q) {
    return res.status(400).json({
      error: "Informe o handle do produto ou uma busca",
      exemplos: [
        "/api/inventory?product=top-rina-strata",
        "/api/inventory?product=top-rina-strata&size=M",
        "/api/inventory?q=brisa&size=P",
      ],
    });
  }

  try {
    let products = [];

    if (handle) {
      // Busca direta por handle (slug exato do produto)
      const data = await fetchShopify(
        `/products.json?handle=${handle}&fields=id,title,handle,variants,status`
      );
      products = data.products.filter((p) => p.status === "active");
    } else if (q) {
      // Busca por nome parcial
      const data = await fetchShopify(
        `/products.json?title=${encodeURIComponent(q)}&limit=10&fields=id,title,handle,variants,status`
      );
      products = data.products.filter((p) => p.status === "active");
    }

    if (products.length === 0) {
      return res.status(404).json({
        encontrado: false,
        mensagem: "Produto não encontrado. Verifique o nome ou handle informado.",
      });
    }

    const resultado = products.map((p) => formatInventory(p, size));

    return res.status(200).json({
      encontrado: true,
      total: resultado.length,
      estoque: resultado,
    });
  } catch (error) {
    console.error("Erro ao verificar estoque:", error.message);
    return res.status(500).json({
      error: "Não foi possível verificar o estoque.",
      detalhes: error.message,
    });
  }
}
