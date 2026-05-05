/**
 * NKSW Products API — Shopify
 * Busca produtos em tempo real da loja Shopify.
 *
 * Parâmetros:
 *   ?q=biquini        → busca por título (nome do produto)
 *   ?collection=brisa → filtra por coleção/drop
 *   ?limit=20         → quantidade de resultados (padrão: 20, máx: 250)
 *   ?available=true   → apenas produtos com estoque disponível
 */

const SHOPIFY_STORE = process.env.SHOPIFY_STORE;       // ex: "nakedsw.myshopify.com"
const SHOPIFY_TOKEN = process.env.SHOPIFY_ADMIN_TOKEN; // Admin API access token
const API_VERSION   = "2024-01";

async function fetchShopify(path) {
  const url = `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`;
  const res = await fetch(url, {
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_TOKEN,
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Shopify API error ${res.status}: ${err}`);
  }

  return res.json();
}

const TAMANHOS_VALIDOS = new Set(["PP", "P", "M", "G", "GG", "Único", "U"]);

function detectarOpcoes(product) {
  // Detecta qual option é tamanho e qual é cor com base nos nomes das options
  const opt = product.options || [];
  let idxTamanho = 0;
  let idxCor = 1;
  opt.forEach((o, i) => {
    const nome = (o.name || "").toLowerCase();
    if (nome.includes("tamanho") || nome.includes("size")) idxTamanho = i;
    if (nome.includes("cor") || nome.includes("color") || nome.includes("estampa")) idxCor = i;
  });
  return { idxTamanho, idxCor };
}

function formatProduct(product) {
  const { idxTamanho, idxCor } = detectarOpcoes(product);
  const opcoes = ["option1", "option2", "option3"];

  return {
    id: product.id,
    nome: product.title,
    descricao: product.body_html?.replace(/<[^>]+>/g, "").trim() || "",
    tipo: product.product_type,
    tags: product.tags?.split(", ") || [],
    colecao_drop: product.vendor || "",
    url: `https://www.nakedsw.com.br/products/${product.handle}`,
    imagem: product.images?.[0]?.src || null,
    variantes: product.variants?.map((v) => ({
      id: v.id,
      tamanho: v[opcoes[idxTamanho]] || null,
      cor: v[opcoes[idxCor]] || null,
      preco: `R$ ${parseFloat(v.price).toFixed(2).replace(".", ",")}`,
      preco_comparacao: v.compare_at_price
        ? `R$ ${parseFloat(v.compare_at_price).toFixed(2).replace(".", ",")}`
        : null,
      em_promocao: v.compare_at_price && parseFloat(v.compare_at_price) > parseFloat(v.price),
      disponivel: v.inventory_quantity > 0,
      estoque: v.inventory_quantity,
    })) || [],
    disponivel: product.variants?.some((v) => v.inventory_quantity > 0) ?? false,
    status: product.status,
  };
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Método não permitido" });

  try {
    const { q, limit = "20", available } = req.query;

    // Montar query string para Shopify
    const params = new URLSearchParams({
      limit: Math.min(parseInt(limit), 250).toString(),
      status: "active",
      fields: "id,title,body_html,product_type,tags,vendor,handle,images,variants,status",
    });

    if (q) params.set("title", q);

    const data = await fetchShopify(`/products.json?${params}`);
    let products = data.products.map(formatProduct);

    // Filtrar apenas disponíveis se solicitado
    if (available === "true") {
      products = products.filter((p) => p.disponivel);
    }

    return res.status(200).json({
      total: products.length,
      produtos: products,
    });
  } catch (error) {
    console.error("Erro ao buscar produtos:", error.message);
    return res.status(500).json({
      error: "Não foi possível buscar os produtos.",
      detalhes: error.message,
    });
  }
}
