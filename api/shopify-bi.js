/**
 * NKSW Shopify BI — lê pedidos do Vercel Blob e retorna no formato do dashboard
 *
 * GET /api/shopify-bi?section=all&month=2026-05
 *
 * REGRAS:
 *   - Apenas pedidos com financial_status === 'paid' entram nos KPIs de receita/ticket/pedidos
 *   - Datas sempre em fuso de Brasília (America/Sao_Paulo)
 *   - D-1: para o mês corrente, exclui pedidos do dia atual (Brasília)
 */

import { head } from '@vercel/blob';

const TZ = 'America/Sao_Paulo';

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

/** Retorna YYYY-MM-DD no fuso de Brasília */
function brISO(isoStr) {
  return new Date(isoStr).toLocaleDateString('en-CA', { timeZone: TZ });
}

/** Retorna data formatada pt-BR no fuso de Brasília */
function brDate(isoStr) {
  return new Date(isoStr).toLocaleDateString('pt-BR', { timeZone: TZ });
}

/** Retorna YYYY-MM no fuso de Brasília */
function brYM(isoStr) {
  return brISO(isoStr).slice(0, 7);
}

/** Retorna true para cupons de troca */
const isTroca = code => typeof code === 'string' && code.toLowerCase().includes('troca');

/** Calcula dias úteis entre duas datas (exclui sáb e dom) */
function diasUteis(dataInicio, dataFim) {
  const start = new Date(dataInicio + 'T00:00:00-03:00');
  const end   = new Date(dataFim   + 'T23:59:59-03:00');
  let count = 0;
  const cur = new Date(start);
  while (cur <= end) {
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

/** Status do pedido baseado em fulfillment e dias úteis */
function calcStatusPedido(order) {
  if (order.cancelled_at) return 'Cancelado';
  const fulfillments = order.fulfillments || [];
  const hasTracking  = fulfillments.some(f => f.tracking_number && f.tracking_number.trim());
  // Enviado = tem rastreamento (fulfilled ou partial com tracking)
  if (hasTracking || order.fulfillment_status === 'fulfilled') return 'Enviado';
  // Sem envio → verifica prazo de 7 dias úteis desde a data do pedido
  const orderDateBR    = brISO(order.created_at);
  const todayBR        = new Date().toLocaleDateString('en-CA', { timeZone: TZ });
  const diasDecorridos = diasUteis(orderDateBR, todayBR);
  if (diasDecorridos >= 7) return 'Atrasado';
  return 'No Prazo';
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

/**
 * Retorna data D-1 em Brasília (YYYY-MM-DD).
 * Para o mês corrente, usamos como corte o dia anterior (dados do dia atual
 * ainda não estão completos). Para meses históricos, incluímos tudo.
 */
function getD1BR() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return d.toLocaleDateString('en-CA', { timeZone: TZ });
}

// ── Builders ───────────────────────────────────────────────────────────────

function buildVendas(orders, isCurrent, filterFn) {
  // Filtro de data: período personalizado tem prioridade sobre D-1
  let base;
  if (filterFn) {
    base = orders.filter(filterFn);
  } else {
    const cutoff = isCurrent ? getD1BR() : null;
    base = cutoff ? orders.filter(o => brISO(o.created_at) <= cutoff) : orders;
  }

  // Apenas pedidos PAGOS para KPIs financeiros
  const paid      = base.filter(o => o.financial_status === 'paid');
  const cancelled = base.filter(o => o.cancelled_at);

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
    const total  = parseFloat(o.total_price) || 0;
    const frete  = shippingAmount(o);

    const descontoTroca = (o.discount_codes || [])
      .filter(dc => isTroca(dc.code))
      .reduce((s, dc) => s + (parseFloat(dc.amount) || 0), 0);
    const desconto = Math.max(0, (parseFloat(o.total_discounts) || 0) - descontoTroca);

    receita       += total;
    descontoTotal += desconto;
    valorFrete    += frete;

    frete > 0 ? pedidosComFrete++ : pedidosSemFrete++;

    const nonTrocaCodes = (o.discount_codes || []).filter(dc => !isTroca(dc.code));
    if (nonTrocaCodes.length) comCupom++;

    // ── Diário (data em Brasília) ────────────────────────────────────────
    const dia = brISO(o.created_at);
    if (!diarioMap[dia]) diarioMap[dia] = { dia, receita: 0, pedidos: 0 };
    diarioMap[dia].receita += total;
    diarioMap[dia].pedidos++;

    // ── Por mês ──────────────────────────────────────────────────────────
    const ym  = brYM(o.created_at);
    const mon = parseInt(ym.slice(5, 7), 10);
    if (!mesMap[ym]) mesMap[ym] = { mes: mon, label: MONTH_LABELS[mon - 1], receita: 0, pedidos: 0 };
    mesMap[ym].receita += total;
    mesMap[ym].pedidos++;

    // ── Canal ────────────────────────────────────────────────────────────
    const canal = o.payment_gateway || 'Outros';
    if (!canaisMap[canal]) canaisMap[canal] = { canal, receita: 0, pedidos: 0 };
    canaisMap[canal].receita += total;
    canaisMap[canal].pedidos++;

    if (!pagMap[canal]) pagMap[canal] = { forma: canal, total: 0 };
    pagMap[canal].total++;

    // ── Por estado ───────────────────────────────────────────────────────
    const estado = o.shipping_address?.province || o.billing_address?.province || 'N/A';
    if (!estadosMap[estado]) estadosMap[estado] = { estado, receita: 0, pedidos: 0 };
    estadosMap[estado].receita += total;
    estadosMap[estado].pedidos++;

    // ── Por vendedor ─────────────────────────────────────────────────────
    const vendedorTag = (o.tags || '').split(',').map(t => t.trim())
      .find(t => t.toLowerCase().startsWith('vendedor:'));
    const vendedor = vendedorTag ? vendedorTag.split(':')[1].trim() : 'Shopify';
    if (!vendedorMap[vendedor]) vendedorMap[vendedor] = { vendedor, receita: 0, pedidos: 0 };
    vendedorMap[vendedor].receita += total;
    vendedorMap[vendedor].pedidos++;

    // ── Status pagamento ─────────────────────────────────────────────────
    const spLabel = FINANCIAL_LABELS[o.financial_status] || o.financial_status || 'Desconhecido';
    if (!statusPagMap[spLabel]) statusPagMap[spLabel] = { status: spLabel, pedidos: 0, receita: 0 };
    statusPagMap[spLabel].pedidos++;
    statusPagMap[spLabel].receita += total;

    // ── Cupons (valor = receita sem frete do pedido, não o desconto) ────
    for (const dc of (o.discount_codes || [])) {
      if (isTroca(dc.code)) continue;
      const key = dc.code;
      if (!cupomMap[key]) cupomMap[key] = { cupom: key, pedidos: 0, valor: 0 };
      cupomMap[key].pedidos++;
      cupomMap[key].valor += Math.max(0, total - frete); // receita s/ frete
    }
  }

  const pedidos         = paid.length;
  const ticket          = pedidos > 0 ? receita / pedidos : 0;
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
      taxaCancel:      r1(base.length > 0 ? (cancelled.length / base.length) * 100 : 0),
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
    porCupom:        Object.values(cupomMap).sort((a, b) => b.valor - a.valor)
                       .map(c => ({ ...c, valor: r2(c.valor) })),
    statusPagamento: Object.values(statusPagMap),
  };
}

function buildPedidos(orders, isCurrent, filterFn) {
  let filtered;
  if (filterFn) {
    filtered = orders.filter(filterFn);
  } else {
    const cutoff = isCurrent ? getD1BR() : null;
    filtered = cutoff ? orders.filter(o => brISO(o.created_at) <= cutoff) : orders;
  }

  // Contadores por status de pagamento
  const contadores = { pagos: 0, pendentes: 0, parciais: 0, cancelados: 0, entregues: 0 };
  for (const o of filtered) {
    const fs = o.financial_status;
    if (o.cancelled_at || fs === 'voided' || fs === 'refunded' || fs === 'partially_refunded') {
      contadores.cancelados++;
    } else if (fs === 'paid') {
      contadores.pagos++;
      if (o.fulfillment_status === 'fulfilled') contadores.entregues++;
    } else if (fs === 'partially_paid') {
      contadores.parciais++;
    } else {
      contadores.pendentes++;
    }
  }

  // Lista completa — TODOS os pedidos (não só pagos)
  const lista = filtered.map(o => {
    const fulfillments = o.fulfillments || [];
    const tracking     = fulfillments.flatMap(f => f.tracking_numbers || [])[0] || null;
    const hasTracking  = fulfillments.some(f => f.tracking_number && f.tracking_number.trim());
    // Status de entrega: espelha exatamente o que o Shopify Admin mostra
    // "fulfilled" no Shopify = enviado/rastreado, NÃO entregue ao cliente
    let statusEntrega = '';
    if (o.cancelled_at) {
      statusEntrega = 'Cancelado';
    } else if (o.fulfillment_status === 'partial') {
      statusEntrega = hasTracking ? 'Rastreamento adicionado' : 'Envio parcial';
    } else if (o.fulfillment_status === 'fulfilled' || hasTracking) {
      statusEntrega = 'Rastreamento adicionado';
    }
    // sem fulfillment_status e sem tracking → string vazia (como o Shopify mostra)
    return {
      id:             String(o.order_number),
      produto:        o.line_items?.[0]?.title || '',
      cliente:        customerName(o),
      email:          o.customer?.email || o.contact_email || '',
      pagamento:      FINANCIAL_LABELS[o.financial_status] || o.financial_status || 'Desconhecido',
      metodo_envio:   o.shipping_lines?.[0]?.title || '—',
      status_entrega: statusEntrega,
      status_pedido:  calcStatusPedido(o),
      tags:           o.tags || '',
      valor:          parseFloat(o.total_price) || 0,
      data:           brDate(o.created_at),
      rastreio:       tracking,
    };
  });

  return { contadores, lista };
}

function buildLogistica(orders, isCurrent, filterFn) {
  let filtered;
  if (filterFn) {
    filtered = orders.filter(filterFn);
  } else {
    const cutoff = isCurrent ? getD1BR() : null;
    filtered = cutoff ? orders.filter(o => brISO(o.created_at) <= cutoff) : orders;
  }

  const statusCount = {};
  const pedidos = filtered
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
        data:           brDate(o.created_at),
        cliente:        customerName(o),
        status,
        rastreio:       tracking[0] || null,
        transportadora: o.fulfillments?.[0]?.tracking_company || null,
        url_rastreio:   o.fulfillments?.[0]?.tracking_url     || null,
      };
    });

  return {
    statusList: Object.entries(statusCount)
      .map(([status, total]) => ({ status, total }))
      .sort((a, b) => b.total - a.total),
    pedidos,
  };
}

function buildClientes(orders, isCurrent, filterFn) {
  let base;
  if (filterFn) {
    base = orders.filter(filterFn);
  } else {
    const cutoff = isCurrent ? getD1BR() : null;
    base = cutoff ? orders.filter(o => brISO(o.created_at) <= cutoff) : orders;
  }
  const paid     = base.filter(o => o.financial_status === 'paid');
  const clienteMap = {};

  for (const o of paid) {
    const cid   = o.customer?.id || o.email || o.contact_email || 'anonimo';
    const total = parseFloat(o.total_price) || 0;
    if (!clienteMap[cid]) {
      clienteMap[cid] = { nome: customerName(o), email: o.customer?.email || o.contact_email || '', pedidos: 0, receita: 0 };
    }
    clienteMap[cid].pedidos++;
    clienteMap[cid].receita += total;
  }

  const clientes  = Object.values(clienteMap);
  const total     = clientes.length;
  const recompras = clientes.filter(c => c.pedidos > 1).length;
  const totalRec  = paid.reduce((s, o) => s + (parseFloat(o.total_price) || 0), 0);
  const ltv       = total > 0 ? totalRec / total : 0;

  return {
    kpis: {
      total, recompras,
      pctRecomp:      r1(total > 0 ? recompras / total * 100 : 0),
      ltv:            r2(ltv),
      totalReceita:   r2(totalRec),
      avgDaysBetween: 0,
    },
    topClientes: clientes.sort((a, b) => b.receita - a.receita).slice(0, 20).map(c => ({ ...c, receita: r2(c.receita) })),
    rfm: [], abc: [],
  };
}

// ── Handler ────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { section = 'all', month, startDate, endDate } = req.query;

  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return res.status(400).json({ error: 'Parâmetro month obrigatório (formato: YYYY-MM)' });
  }

  const blobName = `shopify-orders-${month}.json`;

  try {
    const blobMeta = await head(blobName);
    if (!blobMeta) {
      return res.status(404).json({
        ok: false,
        error: `Dados de ${month} ainda não sincronizados.`,
        hint: `/api/shopify-sync?action=full&month=${month}&secret=nksw2025`,
      });
    }

    const blobRes  = await fetch(blobMeta.url);
    const { orders = [], synced_at } = await blobRes.json();

    // Verifica se é o mês corrente (em Brasília) para aplicar D-1
    const nowBR      = new Date().toLocaleDateString('en-CA', { timeZone: TZ });
    const isCurrent  = month === nowBR.slice(0, 7);
    // Se startDate/endDate fornecidos, filtra por período personalizado (sobrescreve D-1)
    const hasCustomRange = startDate && endDate && /^\d{4}-\d{2}-\d{2}$/.test(startDate);

    const paid = orders.filter(o => o.financial_status === 'paid');
    const cutoff = isCurrent ? getD1BR() : null;
    console.log(`[shopify-bi] ${month} | total=${orders.length} pagos=${paid.length} cutoff=${cutoff || 'none'} TZ=Brasília`);

    const out = { ok: true, month, synced_at, source: 'shopify', startDate, endDate };

    // Se período personalizado: filtra pelos dias selecionados
    // Caso contrário: aplica D-1 (mês corrente) ou sem filtro (histórico)
    const filterFn = hasCustomRange
      ? (o) => { const d = brISO(o.created_at); return d >= startDate && d <= endDate; }
      : null;

    if (section === 'vendas'    || section === 'all') out.vendas    = buildVendas(orders,    isCurrent, filterFn);
    if (section === 'pedidos'   || section === 'all') out.pedidos   = buildPedidos(orders,   isCurrent, filterFn);
    if (section === 'logistica' || section === 'all') out.logistica = buildLogistica(orders, isCurrent, filterFn);
    if (section === 'clientes'  || section === 'all') out.clientes  = buildClientes(orders,  isCurrent, filterFn);

    res.setHeader('Cache-Control', isCurrent ? 's-maxage=60' : 's-maxage=3600');
    return res.status(200).json(out);

  } catch (err) {
    console.error('[shopify-bi]', err);
    return res.status(500).json({ ok: false, error: err.message });
  }
}
