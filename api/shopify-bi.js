/**
 * NKSW Shopify BI — lê pedidos do Vercel Blob e retorna no formato do dashboard
 *
 * GET /api/shopify-bi?section=all&month=2026-05
 *
 * Retorna JSON no mesmo formato do Google Apps Script (nksw_sheets_api.js).
 * REGRA: apenas pedidos com financial_status === 'paid' entram nos cálculos de receita.
 */

import { head } from '@vercel/blob';

// ── Labels ─────────────────────────────────────────────────────────────────

const FINANCIAL_LABELS = {
  pending:             'Pendente',
  authorized:          'Autorizado',
  partially_paid:      'Parcialmente pago',
  paid:                'Pago',
  partially_refunded:  'Parcialmente reembolsado',
  refunded:            'Reembolsado',
  voided:              'Cancelado',
};

const FULFILLMENT_LABELS = {
  unfulfilled: 'Aguardando envio',
  partial:     'Envio parcial',
  fulfilled:   'Enviado',
  restocked:   'Devolvido',
};

const MONTH_LABELS = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

// ── Helpers ────────────────────────────────────────────────────────────────

const r2 = n => Math.round(n * 100) / 100;
const r1 = n => Math.round(n * 10)  / 10;

/** Retorna true para cupons de troca (não devem entrar em desconto concedido) */
const isTroca = code => typeof code === 'string' && code.toLowerCase().includes('troca');

function ptDate(isoStr) {
  return new Date(isoStr).toLocaleDateString('pt-BR');
}

function customerName(order) {
  if (order.customer) {
    const n = `${order.customer.first_name || ''} ${order.customer.last_name || ''}`.trim();
    if (n) return n;
  }
  return order.email || order.contact_email || '';
}

function shippingAmount(order) {
  return parseFloat(order.total_shipping_price_set?.shop_money?.amount ?? 0);
}

// ── Builders ───────────────────────────────────────────────────────────────

function buildVendas(orders) {
  // Apenas pedidos PAGOS para todos os cálculos financeiros
  const paid      = orders.filter(o => o.financial_status === 'paid');
  const cancelled = orders.filter(o => o.cancelled_at);

  let receita = 0, descontoTotal = 0, valorFrete = 0;
  let pedidosComFrete = 0, pedidosSemFrete = 0, comCupom = 0;

  const diarioMap   = {};
  const mesMap      = {};
  const canaisMap   = {};
  const pagMap      = {};
  const estadosMap  = {};
  const vendedorMap = {};
  const statusPagMap= {};
  const cupomMap    = {};

  for (const o of paid) {
    const total    = parseFloat(o.total_price)    || 0;
    const frete    = shippingAmount(o);

    // Desconto: exclui cupons de TROCA (troca de produto não é desconto comercial)
    const descontoTroca = (o.discount_codes || [])
      .filter(dc => isTroca(dc.code))
      .reduce((s, dc) => s + (parseFloat(dc.amount) || 0), 0);
    const desconto = Math.max(0, (parseFloat(o.total_discounts) || 0) - descontoTroca);

    receita       += total;
    descontoTotal += desconto;
    valorFrete    += frete;

    frete > 0 ? pedidosComFrete++ : pedidosSemFrete++;
    // Cupom: conta apenas se há ao menos um cupom não-troca
    const nonTrocaCodes = (o.discount_codes || []).filter(dc => !isTroca(dc.code));
    if (nonTrocaCodes.length) comCupom++;

    // ── Diário ──────────────────────────────────────────────────────────
    const dia = o.created_at.slice(0, 10);
    if (!diarioMap[dia]) diarioMap[dia] = { dia, receita: 0, pedidos: 0 };
    diarioMap[dia].receita += total;
    diarioMap[dia].pedidos++;

    // ── Por mês (para gráfico YoY) ──────────────────────────────────────
    const ym  = o.created_at.slice(0, 7);           // "2026-05"
    const mon = parseInt(o.created_at.slice(5, 7), 10); // 5
    if (!mesMap[ym]) mesMap[ym] = { mes: mon, label: MONTH_LABELS[mon - 1], receita: 0, pedidos: 0 };
    mesMap[ym].receita += total;
    mesMap[ym].pedidos++;

    // ── Canal (gateway) ─────────────────────────────────────────────────
    const canal = o.payment_gateway || 'Outros';
    if (!canaisMap[canal]) canaisMap[canal] = { canal, receita: 0, pedidos: 0 };
    canaisMap[canal].receita += total;
    canaisMap[canal].pedidos++;

    // ── Formas de pagamento ─────────────────────────────────────────────
    if (!pagMap[canal]) pagMap[canal] = { forma: canal, total: 0 };
    pagMap[canal].total++;

    // ── Por estado ──────────────────────────────────────────────────────
    const estado = o.shipping_address?.province || o.billing_address?.province || 'N/A';
    if (!estadosMap[estado]) estadosMap[estado] = { estado, receita: 0, pedidos: 0 };
    estadosMap[estado].receita += total;
    estadosMap[estado].pedidos++;

    // ── Por vendedor ────────────────────────────────────────────────────
    const vendedorTag = (o.tags || '').split(',').map(t => t.trim())
      .find(t => t.toLowerCase().startsWith('vendedor:'));
    const vendedor = vendedorTag ? vendedorTag.split(':')[1].trim() : 'Shopify';
    if (!vendedorMap[vendedor]) vendedorMap[vendedor] = { vendedor, receita: 0, pedidos: 0 };
    vendedorMap[vendedor].receita += total;
    vendedorMap[vendedor].pedidos++;

    // ── Status de pagamento ─────────────────────────────────────────────
    const spLabel = FINANCIAL_LABELS[o.financial_status] || o.financial_status || 'Desconhecido';
    if (!statusPagMap[spLabel]) statusPagMap[spLabel] = { status: spLabel, pedidos: 0, receita: 0 };
    statusPagMap[spLabel].pedidos++;
    statusPagMap[spLabel].receita += total;

    // ── Cupons (exclui cupons de TROCA) ──────────────────────────────────
    for (const dc of (o.discount_codes || [])) {
      if (isTroca(dc.code)) continue; // troca de produto — não exibir como desconto
      const key = dc.code;
      if (!cupomMap[key]) cupomMap[key] = { cupom: key, pedidos: 0, valor: 0 };
      cupomMap[key].pedidos++;
      cupomMap[key].valor += parseFloat(dc.amount) || 0;
    }
  }

  const pedidos        = paid.length;
  const ticket         = pedidos > 0 ? receita / pedidos : 0;
  const receitaSemFrete = r2(receita - valorFrete);

  const diario = Object.values(diarioMap)
    .sort((a, b) => a.dia.localeCompare(b.dia))
    .map(d => ({ ...d, receita: r2(d.receita) }));

  const porMes = Object.values(mesMap)
    .sort((a, b) => a.mes - b.mes)
    .map(m => ({ ...m, receita: r2(m.receita) }));

  return {
    kpis: {
      receita:         r2(receita),
      pedidos,
      ticket:          r2(ticket),
      lucro:           0,
      custo:           0,
      taxaCancel:      r1(orders.length > 0 ? (cancelled.length / orders.length) * 100 : 0),
      descontoTotal:   r2(descontoTotal),
      pctCupom:        r1(pedidos > 0 ? (comCupom / pedidos) * 100 : 0),
      receitaSemFrete,
      valorFrete:      r2(valorFrete),
      pctFrete:        r1(receita > 0 ? (valorFrete / receita) * 100 : 0),
      pedidosComFrete,
      pedidosSemFrete,
      pctComFrete:     r1(pedidos > 0 ? (pedidosComFrete / pedidos) * 100 : 0),
      pctSemFrete:     r1(pedidos > 0 ? (pedidosSemFrete / pedidos) * 100 : 0),
    },
    diario,
    porMes,
    canais:          Object.values(canaisMap).sort((a, b) => b.receita - a.receita),
    pagamentos:      Object.values(pagMap).sort((a, b) => b.total - a.total),
    porEstado:       Object.values(estadosMap).sort((a, b) => b.receita - a.receita),
    porVendedor:     Object.values(vendedorMap).sort((a, b) => b.receita - a.receita),
    porCupom:        Object.values(cupomMap).sort((a, b) => b.valor - a.valor).map(c => ({
                       ...c, valor: r2(c.valor)
                     })),
    statusPagamento: Object.values(statusPagMap),
  };
}

function buildPedidos(orders) {
  const contadores = { aprovados: 0, pendentes: 0, em_transito: 0, cancelados: 0, entregues: 0 };

  for (const o of orders) {
    if (o.cancelled_at)                            contadores.cancelados++;
    else if (o.fulfillment_status === 'fulfilled') contadores.entregues++;
    else if (o.fulfillment_status === 'partial')   contadores.em_transito++;
    else if (o.financial_status   === 'paid')      contadores.aprovados++;
    else                                           contadores.pendentes++;
  }

  const lista = orders.map(o => ({
    id:      String(o.order_number),
    produto: o.line_items?.[0]?.title || '',
    status:  o.cancelled_at
               ? 'Cancelado'
               : (FULFILLMENT_LABELS[o.fulfillment_status] || 'Aguardando envio'),
    valor:   parseFloat(o.total_price) || 0,
    data:    ptDate(o.created_at),
    cliente: customerName(o),
    email:   o.customer?.email || o.contact_email || '',
    rastreio: o.fulfillments?.flatMap(f => f.tracking_numbers || [])[0] || null,
  }));

  return { contadores, lista };
}

function buildLogistica(orders) {
  const statusCount = {};

  const pedidos = orders
    .filter(o => !o.cancelled_at)
    .map(o => {
      let status;
      if      (o.fulfillment_status === 'fulfilled') status = 'Entregue';
      else if (o.fulfillment_status === 'partial')   status = 'Envio Parcial';
      else if (o.fulfillments?.length > 0)           status = 'Em Trânsito';
      else                                           status = 'Aguardando Envio';

      statusCount[status] = (statusCount[status] || 0) + 1;

      const tracking = o.fulfillments?.flatMap(f => f.tracking_numbers || []) || [];
      return {
        pedido:         String(o.order_number),
        data:           ptDate(o.created_at),
        cliente:        customerName(o),
        status,
        rastreio:       tracking[0] || null,
        transportadora: o.fulfillments?.[0]?.tracking_company || null,
        url_rastreio:   o.fulfillments?.[0]?.tracking_url     || null,
      };
    });

  const statusList = Object.entries(statusCount)
    .map(([status, total]) => ({ status, total }))
    .sort((a, b) => b.total - a.total);

  return { statusList, pedidos };
}

function buildClientes(orders) {
  // Apenas pedidos pagos para análise de clientes
  const paid = orders.filter(o => o.financial_status === 'paid');

  const clienteMap = {};

  for (const o of paid) {
    const cid = o.customer?.id || o.email || o.contact_email || 'anonimo';
    const total = parseFloat(o.total_price) || 0;
    if (!clienteMap[cid]) {
      clienteMap[cid] = {
        nome:    customerName(o),
        email:   o.customer?.email || o.contact_email || '',
        pedidos: 0,
        receita: 0,
      };
    }
    clienteMap[cid].pedidos++;
    clienteMap[cid].receita += total;
  }

  const clientes   = Object.values(clienteMap);
  const total      = clientes.length;
  const recompras  = clientes.filter(c => c.pedidos > 1).length;
  const totalRec   = paid.reduce((s, o) => s + (parseFloat(o.total_price) || 0), 0);
  const ltv        = total > 0 ? totalRec / total : 0;

  // Top clientes por receita
  const topClientes = clientes
    .sort((a, b) => b.receita - a.receita)
    .slice(0, 20)
    .map(c => ({ ...c, receita: r2(c.receita) }));

  return {
    kpis: {
      total,
      recompras,
      pctRecomp:      r1(total > 0 ? recompras / total * 100 : 0),
      ltv:            r2(ltv),
      totalReceita:   r2(totalRec),
      avgDaysBetween: 0, // histórico não disponível por mês único
    },
    topClientes,
    // rfm e abc requerem histórico multi-mês — não disponível via sync mensal
    rfm: [],
    abc: [],
  };
}

// ── Handler ────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { section = 'all', month } = req.query;

  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return res.status(400).json({ error: 'Parâmetro month obrigatório (formato: YYYY-MM)' });
  }

  const blobName = `shopify-orders-${month}.json`;

  try {
    const blobMeta = await head(blobName);

    if (!blobMeta) {
      return res.status(404).json({
        ok:    false,
        error: `Dados de ${month} ainda não sincronizados.`,
        hint:  `/api/shopify-sync?action=full&month=${month}&secret=nksw2025`,
      });
    }

    const blobRes  = await fetch(blobMeta.url);
    const { orders = [], synced_at } = await blobRes.json();

    // ── Verificação de totais (log para debug) ──────────────────────────
    const paid = orders.filter(o => o.financial_status === 'paid');
    console.log(`[shopify-bi] ${month}: ${orders.length} pedidos total | ${paid.length} pagos`);
    if (paid.length > 0) {
      const rec  = paid.reduce((s, o) => s + (parseFloat(o.total_price)    || 0), 0);
      const fret = paid.reduce((s, o) => s + shippingAmount(o), 0);
      const desc = paid.reduce((s, o) => s + (parseFloat(o.total_discounts) || 0), 0);
      const last = paid.sort((a,b) => b.order_number - a.order_number)[0]?.order_number;
      console.log(`[shopify-bi] receita=${rec.toFixed(2)} frete=${fret.toFixed(2)} desconto=${desc.toFixed(2)} último=#${last}`);
    }

    const out = { ok: true, month, synced_at, source: 'shopify' };

    if (section === 'vendas'   || section === 'all') out.vendas    = buildVendas(orders);
    if (section === 'pedidos'  || section === 'all') out.pedidos   = buildPedidos(orders);
    if (section === 'logistica'|| section === 'all') out.logistica = buildLogistica(orders);
    if (section === 'clientes' || section === 'all') out.clientes  = buildClientes(orders);

    const isCurrentMonth = month === new Date().toISOString().slice(0, 7);
    res.setHeader('Cache-Control', isCurrentMonth ? 's-maxage=300' : 's-maxage=3600');

    return res.status(200).json(out);

  } catch (err) {
    console.error('[shopify-bi]', err);
    return res.status(500).json({ ok: false, error: err.message });
  }
}
