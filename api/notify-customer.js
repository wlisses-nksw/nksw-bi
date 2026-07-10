/**
 * NKSW Avise-me — cria/atualiza cliente no Shopify com tag da variante desejada.
 *
 * POST /api/notify-customer
 *   body JSON: { email, variant_id, product_title, variant_title, product_handle }
 *
 * Efeito no Shopify (cliente identificado pelo e-mail):
 *   - tags:  'avise-me'  +  'avise-me:<variant_id>'
 *   - note:  linha "[AVISE-ME] <produto> · <tamanho> | Variant ID: <id> | <data BR>"
 *            (novas inscrições são anexadas em novas linhas)
 *
 * Variáveis de ambiente necessárias (já usadas por outros endpoints):
 *   SHOPIFY_STORE          ex: lofty-fy.myshopify.com
 *   SHOPIFY_ADMIN_TOKEN    token com escopo write_customers (shpat_...)
 */

const SHOPIFY_STORE = process.env.SHOPIFY_STORE;
const SHOPIFY_TOKEN = process.env.SHOPIFY_ADMIN_TOKEN;
const API_VERSION   = "2024-01";

async function shopify(path, method = "GET", body = null) {
  const url = `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`;
  const res = await fetch(url, {
    method,
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_TOKEN,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch { /* ignore */ }
  if (!res.ok) throw new Error(`Shopify ${res.status}: ${text.slice(0, 300)}`);
  return json;
}

// Data/hora no fuso de Brasília: "04/05/2026, 13:42:16"
function nowBR() {
  return new Intl.DateTimeFormat("pt-BR", {
    timeZone: "America/Sao_Paulo",
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  }).format(new Date());
}

function mergeTags(existing, additions) {
  const set = new Set(
    (existing || "")
      .split(",")
      .map((t) => t.trim())
      .filter(Boolean)
  );
  additions.forEach((t) => t && set.add(t));
  return Array.from(set).join(", ");
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ success: false, error: "Método não permitido" });

  if (!SHOPIFY_STORE || !SHOPIFY_TOKEN) {
    return res.status(500).json({ success: false, error: "Credenciais Shopify não configuradas" });
  }

  try {
    // Body pode vir já parseado (Vercel) ou como string
    let body = req.body;
    if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
    body = body || {};

    const email = String(body.email || "").trim().toLowerCase();
    const variantId    = String(body.variant_id || "").trim();
    const productTitle = String(body.product_title || "").trim();
    const variantTitle = String(body.variant_title || "").trim();

    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return res.status(400).json({ success: false, error: "E-mail inválido" });
    }

    // Linha da nota + tags a adicionar
    const descr = [productTitle, variantTitle].filter(Boolean).join(" · ");
    const noteLine = `[AVISE-ME] ${descr}${variantId ? ` | Variant ID: ${variantId}` : ""} | ${nowBR()}`;
    const newTags = ["avise-me"];
    if (variantId) newTags.push(`avise-me:${variantId}`);

    // Procura cliente pelo e-mail
    const search = await shopify(`/customers/search.json?query=${encodeURIComponent("email:" + email)}`);
    const existing = (search && search.customers && search.customers[0]) || null;

    if (existing) {
      const note = existing.note ? `${existing.note}\n${noteLine}` : noteLine;
      const tags = mergeTags(existing.tags, newTags);
      await shopify(`/customers/${existing.id}.json`, "PUT", {
        customer: { id: existing.id, tags, note },
      });
      return res.status(200).json({ success: true, created: false });
    }

    // Cliente novo
    await shopify(`/customers.json`, "POST", {
      customer: {
        email,
        tags: newTags.join(", "),
        note: noteLine,
        // não inscreve em marketing (avise-me é transacional; consentimento é o padrão "não inscrito")
      },
    });
    return res.status(200).json({ success: true, created: true });
  } catch (err) {
    return res.status(500).json({ success: false, error: String(err.message || err) });
  }
}
