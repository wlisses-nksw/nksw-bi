/**
 * NKSW Order Lookup API — Shopify
 * Busca status de pedido por e-mail OU número do pedido.
 *
 * Parâmetros (ao menos um obrigatório):
 *   ?email=cliente@email.com   → busca por e-mail (retorna pedidos mais recentes)
 *   ?number=1234               → busca por número do pedido (sem #)
 *   ?email=X&number=Y          → busca combinada (mais preciso)
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

// Cache do token (válido por 23h)
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

  if (!res.ok) throw new Error(`Token error: Shopify ${res.status}`);
  const data = await res.json();
  if (!data.access_token) throw new Error("Token não retornado pelo Shopify");

  cachedToken    = data.access_token;
  tokenExpiresAt = Date.now() + 23 * 60 * 60 * 1000;
  return cachedToken;
}

const STATUS_MAP = {
  pending:            "Pagamento pendente",
  authorized:         "Pagamento autorizado",
  partially_paid:     "Pagamento parcial",
  paid:               "Pago ✅",
  partially_refunded: "Parcialmente reembolsado",
  refunded:           "Reembolsado",
  voided:             "Cancelado",
  unfulfilled:        "Aguardando envio 📦",
  fulfilled:          "Enviado ✈️",
  restocked:          "Devolvido ao estoque",
  null:               "Em processamento",
};

const CANCEL_MAP = {
  customer:  "Cancelado a pedido da cliente",
  fraud:     "Cancelado por segurança",
  inventory: "Cancelado por falta de estoque",
  declined:  "Cancelado — pagamento recusado",
  other:     "Cancelado",
};

async function fetchShopify(path) {
  const token = await getToken();
  const url   = `https://${SHOPIFY_STORE}/admin/api/${API_VERSION}${path}`;
  const res   = await fetch(url, {
    headers: {
      "X-Shopify-Access-Token": token,
      "Content-Type": "application/json",
    },
  });
  if (!res.ok) throw new Error(`Shopify ${res.status}`);
  return res.json();
}

function formatOrder(order) {
  const items = order.line_items?.map((item) => ({
    produto:    item.title,
    variante:   item.variant_title || "",
    quantidade: item.quantity,
    preco:      `R$ ${parseFloat(item.price).toFixed(2).replace(".", ",")}`,
  })) || [];

  const tracking = order.fulfillments?.flatMap((f) =>
    f.tracking_numbers?.map((num) => ({
      codigo:         num,
      transportadora: f.tracking_company || "Correios",
      url_rastreio:   f.tracking_url || `https://rastreamento.correios.com.br/app/index.php?objetos=${num}`,
    }))
  ) || [];

  return {
    numero:              `#${order.order_number}`,
    data_criacao:        new Date(order.created_at).toLocaleDateString("pt-BR"),
    status_pagamento:    STATUS_MAP[order.financial_status] || order.financial_status,
    status_envio:        STATUS_MAP[order.fulfillment_status] || "Em processamento",
    cancelado:           !!order.cancelled_at,
    motivo_cancelamento: order.cancel_reason ? CANCEL_MAP[order.cancel_reason] : null,
    itens:               items,
    total:               `R$ ${parseFloat(order.total_price).toFixed(2).replace(".", ",")}`,
    rastreamento:        tracking,
    endereco_entrega:    order.shipping_address
      ? `${order.shipping_address.city}, ${order.shipping_address.province} — ${order.shipping_address.zip}`
      : null,
    nota_interna:        order.note || null,
  };
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Método não permitido" });

  const { email, number } = req.query;

  if (!email && !number) {
    return res.status(400).json({
      error: "Informe ao menos um parâmetro: email ou number",
      exemplo: "/api/order?email=cliente@email.com  ou  /api/order?number=41099",
    });
  }

  try {
    let order = null;

    if (email && number) {
      // Busca combinada: mais precisa
      const numeroLimpo = number.replace("#", "").trim();
      const data = await fetchShopify(
        `/orders.json?email=${encodeURIComponent(email)}&status=any&limit=50&fields=id,order_number,created_at,financial_status,fulfillment_status,cancelled_at,cancel_reason,line_items,total_price,shipping_address,fulfillments,note`
      );
      order = (data.orders || []).find((o) => String(o.order_number) === numeroLimpo) || null;

    } else if (number) {
      // Busca só pelo número do pedido
      const numeroLimpo = number.replace("#", "").trim();
      const data = await fetchShopify(
        `/orders.json?name=%23${numeroLimpo}&status=any&limit=5&fields=id,order_number,created_at,financial_status,fulfillment_status,cancelled_at,cancel_reason,line_items,total_price,shipping_address,fulfillments,note`
      );
      order = (data.orders || [])[0] || null;

    } else if (email) {
      // Busca só pelo e-mail: retorna o pedido mais recente
      const data = await fetchShopify(
        `/orders.json?email=${encodeURIComponent(email)}&status=any&limit=50&fields=id,order_number,created_at,financial_status,fulfillment_status,cancelled_at,cancel_reason,line_items,total_price,shipping_address,fulfillments,note`
      );
      const orders = data.orders || [];
      // Ordena por data desc e pega o mais recente
      orders.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
      order = orders[0] || null;
    }

    if (!order) {
      return res.status(404).json({
        encontrado: false,
        mensagem: "Nenhum pedido encontrado com os dados informados.",
      });
    }

    return res.status(200).json({
      encontrado: true,
      pedido: formatOrder(order),
    });

  } catch (error) {
    console.error("Erro ao buscar pedido:", error.message);
    return res.status(500).json({
      error: "Não foi possível buscar o pedido.",
      detalhes: error.message,
    });
  }
}
