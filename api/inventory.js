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

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
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
