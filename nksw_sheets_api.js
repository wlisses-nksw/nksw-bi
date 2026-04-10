// ================================================================
// NKSW Dashboard — API de Planilhas (Google Apps Script)
// Versão 2.1
// ================================================================
//
// COMO USAR:
// 1. Acesse https://script.google.com → Novo projeto
// 2. Cole TODO este código no editor
// 3. Preencha CONFIG abaixo com os dados da sua planilha
// 4. Clique em: Implantar → Nova implantação → Tipo: App da Web
//      Execute como: Sua conta Google
//      Quem pode acessar: Qualquer pessoa (anônimo)
// 5. Copie a URL gerada (termina em /exec)
// 6. Cole no dashboard em "Conectar Planilhas"
//
// ESTRUTURA DA ABA PRINCIPAL (Base Completa — Pedidos NKSW):
//
//   Número do Pedido | E-mail | Data | Mês | Ano | Status do Pagamento
//   Status do Fulfillment | Status de Entrega | Moeda | Subtotal | Desconto
//   Valor do Frete | Total | Venda sem Frete | Nome do comprador | ...
//   Cidade | Estado | Forma de Pagamento | Cupom de Desconto | Data de Entrega
//   Quantidade | SKU | Produto Final | Pessoa que registrou | Vendedor
//   Motivo do cancelamento | ...
//
// Os nomes das colunas são flexíveis — veja a seção COL_MAP abaixo.
// ================================================================

/* ===== CONFIGURAÇÃO (preencha aqui) ===== */
var CONFIG = {
  // ID da planilha principal (encontre na URL do Google Sheets:
  // docs.google.com/spreadsheets/d/[ESTE_ID]/edit )
  MASTER_ID: '1sA3lJ0DIE7lHYtov-UYOwduZ6zaMgiK9g9jOcEELF-o',

  // Nome exato das abas na sua planilha
  ABA_VENDAS:   'Pedidos',          // aba principal com todos os pedidos
  ABA_PRODUTOS: 'Produtos',
  ABA_PEDIDOS:  'Pedidos',
  // Clientes agora é calculado dinamicamente da ABA_VENDAS (Pedidos)

  // Se quiser usar planilhas SEPARADAS, cole os IDs aqui.
  // Deixe '' para usar somente MASTER_ID acima.
  ID_VENDAS:   '',
  ID_PRODUTOS: '',
  ID_PEDIDOS:  '',
};

/* ===== MAPEAMENTO DE COLUNAS =====
   Adicione aqui os nomes das colunas que você usa na sua planilha.
   A busca é case-insensitive e ignora espaços extras. */
var COL_MAP = {
  vendas: {
    email:          ['e-mail','email','Email','E-mail','email_cliente','cliente_email','email_comprador'],
    data:           ['data','date','Data','Data Pedido','Created at','Paid at','data_pedido'],
    pedido:         ['pedido','ordem','Número do Pedido','N° Pedido','Pedido','Name','order','id_pedido','#'],
    canal:          ['canal','Canal','origem','channel','Gateway','Fonte','fonte'],
    produto:        ['Produto Final','produto final','produto','item','Produto','items','Lineitem name','descrição','SKU'],
    receita:        ['receita','total','Total','Valor','valor','Receita','Total Pedido','Total paid','Subtotal'],
    custo:          ['custo','Custo','cost','custo_produto','Custo Produto','custo_total'],
    status_pag:     ['Status do Pagamento','status do pagamento','Financial Status','status_pagamento','pagamento','Pagamento','status','Status'],
    status_entrega: ['Status de Entrega','status de entrega','status_entrega','Entrega'],
    desconto:       ['Desconto','desconto','discount','Desconto Total'],
    venda_sem_frete:['Venda sem Frete','venda sem frete','venda_sem_frete','receita_sem_frete'],
    forma_pag:      ['Forma de Pagamento','forma de pagamento','Forma de','forma_pagamento','payment_method'],
    cupom:          ['Cupom de Desconto','cupom de desconto','Cupom','cupom','coupon','discount_code'],
    qtd:            ['Quantidade','quantidade','Qtd','qtd','quantity','Lineitem quantity'],
    estado:         ['Estado','estado','state','province','UF','uf'],
    cidade:         ['Cidade','cidade','city'],
    vendedor:       ['Pessoa que registrou a venda','pessoa que registrou a venda','Pessoa que registrou','pessoa que registrou','Pessoa qu','Vendedor','vendedor','seller','Atendente','atendente'],
    frete:          ['Valor do Frete','valor do frete','Frete','frete','shipping','freight','Shipping'],
    motivo_cancel:  ['Motivo do cancelamento','motivo do cancelamento','Motivo do','motivo_cancelamento','cancel_reason'],
    mes:            ['Mês','mes','month','Mês (número)'],
    ano:            ['Ano','ano','year'],
    sku:            ['SKU','sku','Código do produto','cod_produto'],
    nome:           ['Nome do comprador','nome do comprador','Nome','Cliente','Customer','nome_cliente','comprador'],
  },
  produtos: {
    nome:     ['nome','name','produto','Produto','item','Nome','descrição','Descrição'],
    categoria:['categoria','category','Categoria','tipo','Tipo','linha','Linha'],
    custo:    ['custo','Custo','cost','custo_unitario','Custo Unitário','custo_medio'],
    preco:    ['preco','preco_venda','Preço','Preço de Venda','price','pvp','valor'],
    qtd:      ['qtd','quantidade','Qtd Vendida','qtd_vendida','vendas','Vendas','sold','Quantidade Vendida'],
    lucro:    ['lucro','Lucro','lucro_total','profit','Lucro Total'],
    receita:  ['receita','Receita','Receita Total','revenue','faturamento','Faturamento'],
    margem:   ['margem','Margem','margem_%','margem_percent','margin','Margem %'],
  },
  pedidos: {
    id:      ['pedido','id','N° Pedido','Pedido','Name','order','ordem','#','numero'],
    data:    ['data','Data','date','Data Pedido','Created at','Paid at','data_pedido'],
    produto: ['produto','item','Produto','produtos','items','Lineitem name','descricao'],
    status:  ['status','Status','situacao','Situação','Financial Status','pagamento'],
    valor:   ['valor','total','Total','Valor','receita','Total paid','Subtotal','Total Pedido'],
    cliente: ['cliente','Cliente','nome','customer','Email','email','nome_cliente'],
  },
};

/* ===== ENTRADA PRINCIPAL ===== */
function doGet(e) {
  try {
    var p       = e.parameter || {};
    var section = p.section   || 'all';
    var period  = p.period    || 'last_30d';
    var start   = p.startDate || '';
    var end     = p.endDate   || '';

    var dates  = getDateRange(period, start, end);
    var result = { ok: true };

    if (section === 'debug_headers') { return jsonOut(debugHeaders()); }
    if (section === 'debug_mari')    { return jsonOut(debugMari(p.startDate, p.endDate)); }
    if (section === 'vendas'    || section === 'all') result.vendas    = getVendas(dates);
    if (section === 'produtos'  || section === 'all') result.produtos  = getProdutos();
    if (section === 'pedidos'   || section === 'all') result.pedidos   = getPedidos(dates);
    if (section === 'clientes'  || section === 'all') result.clientes  = getClientes(p);
    if (section === 'logistica' || section === 'all') result.logistica = getLogistica(dates);

    result.periodo = {
      inicio: fmtDate(dates.start),
      fim:    fmtDate(dates.end),
      label:  period,
    };

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

/* ===== VENDAS ===== */
function getVendas(dates, statusFilter) {
  var rows = readSheet(CONFIG.ABA_VENDAS, CONFIG.ID_VENDAS);
  if (rows.length < 2) return { kpis: {}, diario: [], canais: [] };

  // Normaliza filtro: array de strings lowercase, ou null se sem filtro
  var statusAllowed = (statusFilter && statusFilter.trim())
    ? statusFilter.split(',').map(function(s){ return s.trim().toLowerCase(); })
    : null;

  var h  = rows[0];
  var cm = COL_MAP.vendas;

  // Índices de colunas
  var iData          = findCol(h, cm.data);
  var iCanal         = findCol(h, cm.canal);
  var iReceita       = findCol(h, cm.receita);
  var iCusto         = findCol(h, cm.custo);
  var iStatusPag     = findCol(h, cm.status_pag);
  var iStatusEntr    = findCol(h, cm.status_entrega);
  var iDesconto      = findCol(h, cm.desconto);
  var iVendaSemFrete = findCol(h, cm.venda_sem_frete);
  var iFreteCol      = findCol(h, cm.frete);
  var iFormaPag      = findCol(h, cm.forma_pag);
  var iCupom         = findCol(h, cm.cupom);
  var iQtd           = findCol(h, cm.qtd);
  var iEstado        = findCol(h, cm.estado);
  var iVendedor      = findCol(h, ['Pessoa que registrou a venda','pessoa que registrou a venda','Pessoa que registrou','pessoa que registrou']); // col AS
  // Fallback: usa "Vendedor" (col AU) quando col AS está vazia na linha
  var iVendedor2     = findCol(h, ['Vendedor','vendedor','seller','Atendente','atendente']);
  var iMotivo        = findCol(h, cm.motivo_cancel);
  var iProduto       = findCol(h, cm.produto);

  // Acumuladores
  var receita = 0, custo = 0, pedidos = 0;
  var totalDesconto = 0, pedidosComCupom = 0;
  var receitaSemFrete = 0, valorFrete = 0;
  var cancelados = 0, pedidosComFrete = 0, pedidosSemFrete = 0;

  var byDia = {}, byCanal = {}, byMes = {};
  var byFormaPag = {}, byEstado = {}, byVendedor = {};
  var byStatusPag = {}, byMotivo = {};
  var byProduto = {}, byCupom = {};

  rows.slice(1).forEach(function (r) {
    var d = parseDate(r[iData]);
    if (!d || d < dates.start || d > dates.end) return;

    var val         = parseNum(r[iReceita]);
    var cst         = parseNum(r[iCusto]);
    var desc        = iDesconto      >= 0 ? parseNum(r[iDesconto])      : 0;
    var semFrete    = iVendaSemFrete >= 0 ? parseNum(r[iVendaSemFrete]) : val;
    var freVal      = iFreteCol      >= 0 ? parseNum(r[iFreteCol])      : Math.max(0, val - semFrete);
    var canal       = iCanal         >= 0 ? trim(r[iCanal])  || 'Direto' : 'Direto';
    var statusPag   = iStatusPag     >= 0 ? trim(r[iStatusPag])  || 'Outros' : 'Outros';
    var statusEntr  = iStatusEntr    >= 0 ? trim(r[iStatusEntr]) || 'Outros' : 'Outros';
    var formaPag    = iFormaPag      >= 0 ? trim(r[iFormaPag])   || 'Outros' : 'Outros';
    var cupom       = iCupom         >= 0 ? trim(r[iCupom]) : '';
    var estado      = iEstado        >= 0 ? trim(r[iEstado])     || 'N/D'    : 'N/D';
    var _v1      = iVendedor  >= 0 ? trim(r[iVendedor])  : '';
    var _v2      = iVendedor2 >= 0 ? trim(r[iVendedor2]) : '';
    var vendedor = _v1 || _v2 || 'Outros';
    var motivo      = iMotivo        >= 0 ? trim(r[iMotivo]) : '';
    var produto     = iProduto       >= 0 ? trim(r[iProduto])    || 'Outros' : 'Outros';
    var qtd         = iQtd           >= 0 ? parseNum(r[iQtd]) || 1 : 1;
    var dia         = fmtDate(d);
    var mesKey      = d.getFullYear() + '-' + pad(d.getMonth() + 1);

    // ── Totais globais (sem filtro) ──
    receita         += val;
    custo           += cst;
    pedidos         += 1;
    totalDesconto   += desc;
    receitaSemFrete += semFrete;
    valorFrete      += freVal;
    if (freVal > 0) pedidosComFrete++; else pedidosSemFrete++;
    if (cupom) {
      pedidosComCupom++;
      byCupom[cupom] = byCupom[cupom] || { pedidos: 0, valor: 0 };
      byCupom[cupom].pedidos += 1;
      byCupom[cupom].valor   += semFrete; // valor líquido (col N / Venda sem Frete)
    }

    var stEntrLow = statusEntr.toLowerCase();
    if (stEntrLow === 'cancelado' || stEntrLow.includes('cancel')) cancelados++;

    byDia[dia] = byDia[dia] || { receita: 0, pedidos: 0 };
    byDia[dia].receita  += val;
    byDia[dia].pedidos  += 1;

    byCanal[canal] = byCanal[canal] || { receita: 0, pedidos: 0 };
    byCanal[canal].receita += val;
    byCanal[canal].pedidos += 1;

    byMes[mesKey] = byMes[mesKey] || { receita: 0, pedidos: 0, label: '' };
    byMes[mesKey].receita  += val;
    byMes[mesKey].pedidos  += 1;
    byMes[mesKey].label     = mesKey;

    byFormaPag[formaPag] = byFormaPag[formaPag] || { receita: 0, pedidos: 0 };
    byFormaPag[formaPag].receita += val;
    byFormaPag[formaPag].pedidos += 1;

    if (estado && estado !== 'N/D') {
      byEstado[estado] = byEstado[estado] || { receita: 0, pedidos: 0 };
      byEstado[estado].receita += val;
      byEstado[estado].pedidos += 1;
    }

    byVendedor[vendedor] = byVendedor[vendedor] || { receita: 0, pedidos: 0 };
    byVendedor[vendedor].receita += semFrete;
    byVendedor[vendedor].pedidos += 1;

    if (produto && produto !== 'Outros') {
      byProduto[produto] = byProduto[produto] || { receita: 0, qtd: 0 };
      byProduto[produto].receita += val;
      byProduto[produto].qtd     += qtd;
    }

    if (motivo) {
      byMotivo[motivo] = byMotivo[motivo] || 0;
      byMotivo[motivo] += 1;
    }

    // ── Acumulação por status (para filtro client-side) ──
    byStatusPag[statusPag] = byStatusPag[statusPag] || {
      receita: 0, custo: 0, pedidos: 0, desconto: 0, semFrete: 0, frete: 0, cupom: 0, cancelados: 0,
      comFrete: 0, semFreteCount: 0,
      byMes: {}, byVendedor: {}, byEstado: {}
    };
    var sp = byStatusPag[statusPag];
    sp.receita   += val;
    sp.custo     += cst;
    sp.pedidos   += 1;
    sp.desconto  += desc;
    sp.semFrete  += semFrete;
    sp.frete     += freVal;
    if (freVal > 0) sp.comFrete++; else sp.semFreteCount++;
    if (cupom) sp.cupom++;
    if (stEntrLow === 'cancelado' || stEntrLow.includes('cancel')) sp.cancelados++;

    sp.byMes[mesKey] = sp.byMes[mesKey] || { receita: 0, pedidos: 0 };
    sp.byMes[mesKey].receita += val;
    sp.byMes[mesKey].pedidos += 1;

    sp.byVendedor[vendedor] = sp.byVendedor[vendedor] || { receita: 0, pedidos: 0 };
    sp.byVendedor[vendedor].receita += semFrete;
    sp.byVendedor[vendedor].pedidos += 1;

    if (estado && estado !== 'N/D') {
      sp.byEstado[estado] = sp.byEstado[estado] || { receita: 0, pedidos: 0 };
      sp.byEstado[estado].receita += val;
      sp.byEstado[estado].pedidos += 1;
    }
  });

  var ticket        = pedidos > 0 ? receita / pedidos : 0;
  var taxaCancel    = pedidos > 0 ? cancelados / pedidos * 100 : 0;
  var pctCupom      = pedidos > 0 ? pedidosComCupom / pedidos * 100 : 0;

  var diario = Object.keys(byDia).sort().map(function (d) {
    return { dia: d, receita: round2(byDia[d].receita), pedidos: byDia[d].pedidos };
  });

  var canais = Object.keys(byCanal).map(function (c) {
    return { canal: c, receita: round2(byCanal[c].receita), pedidos: byCanal[c].pedidos };
  }).sort(function (a, b) { return b.receita - a.receita; });

  var porMes = Object.keys(byMes).sort().map(function (k) {
    var parts = k.split('-');
    var nomeMes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][parseInt(parts[1]) - 1];
    return { mes: k, label: nomeMes + '/' + parts[0].slice(2), receita: round2(byMes[k].receita), pedidos: byMes[k].pedidos };
  });

  var formaPag = Object.keys(byFormaPag).map(function (f) {
    return { forma: f, receita: round2(byFormaPag[f].receita), pedidos: byFormaPag[f].pedidos };
  }).sort(function (a, b) { return b.receita - a.receita; });

  var porEstado = Object.keys(byEstado).map(function (e) {
    return { estado: e, receita: round2(byEstado[e].receita), pedidos: byEstado[e].pedidos };
  }).sort(function (a, b) { return b.receita - a.receita; }).slice(0, 15);

  var porVendedor = Object.keys(byVendedor).map(function (v) {
    return { vendedor: v, receita: round2(byVendedor[v].receita), pedidos: byVendedor[v].pedidos };
  }).sort(function (a, b) { return b.receita - a.receita; });

  var MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var statusPagamento = Object.keys(byStatusPag).map(function (s) {
    var sp = byStatusPag[s];
    var spTicket = sp.pedidos > 0 ? sp.receita / sp.pedidos : 0;
    var spPorMes = Object.keys(sp.byMes).sort().map(function (k) {
      var pts = k.split('-');
      return { mes: k, label: MESES[parseInt(pts[1])-1]+'/'+pts[0].slice(2), receita: round2(sp.byMes[k].receita), pedidos: sp.byMes[k].pedidos };
    });
    var spPorVendedor = Object.keys(sp.byVendedor).map(function (v) {
      return { vendedor: v, receita: round2(sp.byVendedor[v].receita), pedidos: sp.byVendedor[v].pedidos };
    }).sort(function (a,b){ return b.receita - a.receita; });
    var spPorEstado = Object.keys(sp.byEstado).map(function (e) {
      return { estado: e, receita: round2(sp.byEstado[e].receita), pedidos: sp.byEstado[e].pedidos };
    }).sort(function (a,b){ return b.receita - a.receita; }).slice(0,10);
    return {
      status:      s,
      receita:     round2(sp.receita),
      pedidos:     sp.pedidos,
      kpis: {
        receita:         round2(sp.receita),
        pedidos:         sp.pedidos,
        ticket:          round2(spTicket),
        lucro:           round2(sp.receita - sp.custo),
        custo:           round2(sp.custo),
        taxaCancel:      round1(sp.pedidos > 0 ? sp.cancelados / sp.pedidos * 100 : 0),
        descontoTotal:   round2(sp.desconto),
        pctCupom:        round1(sp.pedidos > 0 ? sp.cupom / sp.pedidos * 100 : 0),
        receitaSemFrete:  round2(sp.semFrete),
        valorFrete:       round2(sp.frete),
        pctFrete:         round1(sp.receita > 0 ? sp.frete / sp.receita * 100 : 0),
        pedidosComFrete:  sp.comFrete,
        pedidosSemFrete:  sp.semFreteCount,
        pctComFrete:      round1(sp.pedidos > 0 ? sp.comFrete       / sp.pedidos * 100 : 0),
        pctSemFrete:      round1(sp.pedidos > 0 ? sp.semFreteCount  / sp.pedidos * 100 : 0),
      },
      porMes:      spPorMes,
      porVendedor: spPorVendedor,
      porEstado:   spPorEstado,
    };
  }).sort(function (a, b) { return b.receita - a.receita; });

  var topProdutos = Object.keys(byProduto).map(function (p) {
    return { produto: p, receita: round2(byProduto[p].receita), qtd: byProduto[p].qtd };
  }).sort(function (a, b) { return b.receita - a.receita; }).slice(0, 10);

  var motivosCancel = Object.keys(byMotivo).map(function (m) {
    return { motivo: m, qtd: byMotivo[m] };
  }).sort(function (a, b) { return b.qtd - a.qtd; }).slice(0, 8);

  var porCupom = Object.keys(byCupom).map(function (c) {
    return { cupom: c, pedidos: byCupom[c].pedidos, valor: round2(byCupom[c].valor) };
  }).sort(function (a, b) { return b.valor - a.valor; });

  return {
    kpis: {
      receita:         round2(receita),
      pedidos:         pedidos,
      ticket:          round2(ticket),
      lucro:           round2(receita - custo),
      custo:           round2(custo),
      taxaCancel:      round1(taxaCancel),
      descontoTotal:   round2(totalDesconto),
      pctCupom:        round1(pctCupom),
      receitaSemFrete:  round2(receitaSemFrete),
      valorFrete:       round2(valorFrete),
      pctFrete:         round1(receita > 0 ? valorFrete / receita * 100 : 0),
      pedidosComFrete:  pedidosComFrete,
      pedidosSemFrete:  pedidosSemFrete,
      pctComFrete:      round1(pedidos > 0 ? pedidosComFrete / pedidos * 100 : 0),
      pctSemFrete:      round1(pedidos > 0 ? pedidosSemFrete / pedidos * 100 : 0),
    },
    diario:         diario,
    canais:         canais,
    porMes:         porMes,
    formaPag:       formaPag,
    porEstado:      porEstado,
    porVendedor:    porVendedor,
    statusPagamento:statusPagamento,
    topProdutos:    topProdutos,
    motivosCancel:  motivosCancel,
    porCupom:       porCupom,
  };
}

/* ===== PRODUTOS ===== */
function getProdutos() {
  var rows;
  try { rows = readSheet(CONFIG.ABA_PRODUTOS, CONFIG.ID_PRODUTOS); } catch(e) { return { lista: [], categorias: [] }; }
  if (rows.length < 2) return { lista: [], categorias: [] };

  var h  = rows[0];
  var cm = COL_MAP.produtos;
  var iNome  = findCol(h, cm.nome);
  var iCat   = findCol(h, cm.categoria);
  var iCusto = findCol(h, cm.custo);
  var iPreco = findCol(h, cm.preco);
  var iQtd   = findCol(h, cm.qtd);
  var iLucro = findCol(h, cm.lucro);
  var iRec   = findCol(h, cm.receita);
  var iMarg  = findCol(h, cm.margem);

  var lista = rows.slice(1).filter(function (r) { return r[iNome]; }).map(function (r) {
    var custo = parseNum(r[iCusto]);
    var preco = parseNum(r[iPreco]);
    var qtd   = parseNum(r[iQtd]);
    var rec   = iRec   >= 0 ? parseNum(r[iRec])   : preco * qtd;
    var luc   = iLucro >= 0 ? parseNum(r[iLucro]) : rec - custo * qtd;
    var marg  = iMarg  >= 0 ? parseNum(r[iMarg])  : (preco > 0 ? (preco - custo) / preco * 100 : 0);
    return {
      nome:      trim(r[iNome]),
      categoria: trim(r[iCat]) || 'Geral',
      receita:   round2(rec),
      lucro:     round2(luc),
      margem:    round1(marg),
      qtd:       Math.round(qtd),
    };
  }).sort(function (a, b) { return b.receita - a.receita; });

  // Agrupar por categoria
  var cats = {};
  lista.forEach(function (p) {
    cats[p.categoria] = cats[p.categoria] || { receita: 0, lucro: 0, qtd: 0, produtos: 0 };
    cats[p.categoria].receita  += p.receita;
    cats[p.categoria].lucro    += p.lucro;
    cats[p.categoria].qtd      += p.qtd;
    cats[p.categoria].produtos += 1;
  });

  var categorias = Object.keys(cats).map(function (c) {
    var ct = cats[c];
    return {
      categoria:   c,
      receita:     round2(ct.receita),
      margem:      ct.receita > 0 ? round1(ct.lucro / ct.receita * 100) : 0,
      qtdProdutos: ct.produtos,
    };
  }).sort(function (a, b) { return b.receita - a.receita; });

  return { lista: lista, categorias: categorias };
}

/* ===== PEDIDOS ===== */
function getPedidos(dates) {
  var rows = readSheet(CONFIG.ABA_PEDIDOS, CONFIG.ID_PEDIDOS);
  if (rows.length < 2) return { contadores: {}, lista: [] };

  var h  = rows[0];
  var cm = COL_MAP.pedidos;
  var iId     = findCol(h, cm.id);
  var iData   = findCol(h, cm.data);
  var iProd   = findCol(h, cm.produto);
  var iStatus = findCol(h, cm.status);
  var iValor  = findCol(h, cm.valor);
  var iClien  = findCol(h, cm.cliente);

  var cont = { aprovados: 0, pendentes: 0, em_transito: 0, cancelados: 0, entregues: 0 };
  var lista = [];

  rows.slice(1).forEach(function (r) {
    var d = parseDate(r[iData]);
    if (!d || d < dates.start || d > dates.end) return;

    var st  = trim(r[iStatus]).toLowerCase();
    var sNorm = normalizeStatus(st);

    if (sNorm === 'Aprovado')     cont.aprovados++;
    else if (sNorm === 'Pendente')    cont.pendentes++;
    else if (sNorm === 'Em Trânsito') cont.em_transito++;
    else if (sNorm === 'Cancelado')   cont.cancelados++;
    else if (sNorm === 'Entregue')    cont.entregues++;

    lista.push({
      id:      trim(r[iId]),
      produto: trim(r[iProd]),
      status:  sNorm,
      valor:   parseNum(r[iValor]),
      data:    fmtDateBR(d),
      cliente: trim(r[iClien]),
    });
  });

  // Ordenar do mais recente para o mais antigo (pelo índice original)
  lista.reverse();

  return { contadores: cont, lista: lista.slice(0, 100) };
}

/* ===== CLIENTES — RFM + Cohort + Curva ABC ===== */
function getClientes(params) {
  try {
    params = params || {};
    var cohortMonths = parseInt(params.cohortMonths) || 24;
    var rows = readSheet(CONFIG.ABA_VENDAS, CONFIG.ID_VENDAS);
    if (rows.length < 2) return { kpis: {}, rfm: [], cohort: [], abc: [] };

    var h = rows[0];

    var iEmail   = findCol(h, ['email', 'e-mail', 'email_cliente', 'cliente_email', 'email_comprador']);
    var iData    = findCol(h, ['data', 'data_pedido', 'data_criacao', 'dt_pedido', 'date', 'created_at']);
    var iReceita = findCol(h, ['venda_s/_frete', 'venda sem frete', 'venda_sem_frete', 'receita', 'valor', 'total', 'receita_aprovada', 'valor_aprovado']);
    var iStatus  = findCol(h, ['status', 'status_pagamento', 'situacao', 'payment_status', 'status_do_pedido']);
    var iNome    = findCol(h, ['nome', 'name', 'nome_cliente', 'cliente', 'comprador', 'nome_comprador', 'nome do comprador']);

    if (iEmail < 0) {
      return { kpis: {}, rfm: [], cohort: [], abc: [], erro: 'Coluna de email não encontrada. Cabeçalhos: ' + h.join(', ') };
    }

    var today = new Date(); today.setHours(0, 0, 0, 0);
    var clientesMap = {};

    rows.slice(1).forEach(function(r) {
      var email = trim(r[iEmail]).toLowerCase();
      if (!email || email.indexOf('@') < 0) return;

      if (iStatus >= 0) {
        var st = trim(r[iStatus]).toLowerCase();
        if (st.indexOf('cancel') >= 0 || st.indexOf('estorn') >= 0 || st.indexOf('reembol') >= 0) return;
      }

      var d = parseDate(r[iData]);
      if (!d) return;
      d.setHours(0, 0, 0, 0);

      var receita = iReceita >= 0 ? parseNum(r[iReceita]) : 0;
      var nome    = iNome    >= 0 ? trim(r[iNome])        : '';

      if (!clientesMap[email]) clientesMap[email] = { email: email, nome: nome, orders: [] };
      clientesMap[email].orders.push({ data: d, receita: receita });
    });

    var clientList = Object.keys(clientesMap).map(function(k) { return clientesMap[k]; });
    if (!clientList.length) return { kpis: {}, rfm: [], cohort: [], abc: [] };

    // ── Métricas por cliente ──
    clientList.forEach(function(c) {
      c.orders.sort(function(a, b) { return a.data - b.data; });
      c.firstOrder = c.orders[0].data;
      c.lastOrder  = c.orders[c.orders.length - 1].data;
      c.numOrders  = c.orders.length;
      c.totalSpent = c.orders.reduce(function(s, o) { return s + o.receita; }, 0);
      c.recency    = Math.floor((today - c.lastOrder) / 86400000);
      if (c.numOrders > 1) {
        var totalGap = 0;
        for (var i = 1; i < c.orders.length; i++) {
          totalGap += (c.orders[i].data - c.orders[i - 1].data) / 86400000;
        }
        c.avgGap = totalGap / (c.numOrders - 1);
      }
    });

    // ── KPIs globais ──
    var total        = clientList.length;
    var recompras    = clientList.filter(function(c) { return c.numOrders >= 2; }).length;
    var totalReceita = clientList.reduce(function(s, c) { return s + c.totalSpent; }, 0);
    var ltv          = total > 0 ? totalReceita / total : 0;
    var gapList      = clientList.filter(function(c) { return c.avgGap != null; });
    var avgDays      = gapList.length > 0
      ? gapList.reduce(function(s, c) { return s + c.avgGap; }, 0) / gapList.length : 0;

    // ── RFM Scoring ──
    var sortedSpent = clientList.map(function(c) { return c.totalSpent; }).sort(function(a, b) { return a - b; });
    var sn = sortedSpent.length;
    var p20 = sortedSpent[Math.floor(sn * 0.2)] || 0;
    var p40 = sortedSpent[Math.floor(sn * 0.4)] || 0;
    var p60 = sortedSpent[Math.floor(sn * 0.6)] || 0;
    var p80 = sortedSpent[Math.floor(sn * 0.8)] || 0;

    function rScore(days) { return days <= 30 ? 5 : days <= 60 ? 4 : days <= 90 ? 3 : days <= 180 ? 2 : 1; }
    function fScore(num)  { return num  >= 9  ? 5 : num  >= 5  ? 4 : num  >= 3  ? 3 : num  >= 2  ? 2 : 1; }
    function mScore(val)  { return val  > p80 ? 5 : val  > p60 ? 4 : val  > p40 ? 3 : val  > p20 ? 2 : 1; }

    var rfmDefs = {
      'Champions':  { icon: '🏆', racional: 'Compram muito, frequentemente e de forma recente — núcleo do negócio',              acao: 'Programa fidelidade & embaixadores' },
      'VIPs':       { icon: '💎', racional: 'Alto valor e alta fidelidade — potencial de Champions com o estímulo certo',          acao: 'Tratamento exclusivo & upsell premium' },
      'Fiéis':      { icon: '🔁', racional: 'Compram com regularidade — base sólida de receita recorrente',                       acao: 'Programa de pontos & recompensa' },
      'Potenciais': { icon: '🌱', racional: 'Fizeram 2ª compra recentemente — janela de oportunidade para fidelizar',              acao: 'Incentivar 2ª compra com oferta especial' },
      'Novos':      { icon: '✨', racional: 'Primeira compra recente — vão decidir se viram fiéis ou somem',                       acao: 'Boas-vindas & onboarding por email' },
      'Em Risco':   { icon: '⚠️', racional: 'Foram bons clientes mas estão sumindo — janela de reativação fechando',               acao: 'Campanha de reativação urgente' },
      'Dormentes':  { icon: '💤', racional: 'Compraram algumas vezes e pararam — difíceis de reativar sem estímulo forte',         acao: 'Desconto exclusivo de reativação' },
      'Perdidos':   { icon: '❌', racional: 'Muito tempo sem comprar e baixo engajamento — avaliar custo x benefício de reativar', acao: 'Win-back agressivo ou limpeza da base' },
    };
    var rfmGroups = {};
    Object.keys(rfmDefs).forEach(function(k) { rfmGroups[k] = { base: 0, pedidos: 0, receita: 0 }; });

    clientList.forEach(function(c) {
      var r = rScore(c.recency), f = fScore(c.numOrders), m = mScore(c.totalSpent);
      var seg;
      if      (r >= 4 && f >= 4 && m >= 4)   seg = 'Champions';
      else if (m >= 5 || (f >= 4 && m >= 4)) seg = 'VIPs';
      else if (r >= 3 && f >= 3)              seg = 'Fiéis';
      else if (r >= 4 && f === 1)             seg = 'Novos';
      else if (r >= 3 && f === 2)             seg = 'Potenciais';
      else if (r <= 2 && f >= 3)              seg = 'Em Risco';
      else if (r >= 2 && r <= 3)              seg = 'Dormentes';
      else                                    seg = 'Perdidos';
      rfmGroups[seg].base++;
      rfmGroups[seg].pedidos += c.numOrders;
      rfmGroups[seg].receita += c.totalSpent;
    });

    var rfm = Object.keys(rfmDefs)
      .filter(function(k) { return rfmGroups[k].base > 0; })
      .map(function(k) {
        var g = rfmGroups[k];
        return { segmento: rfmDefs[k].icon + ' ' + k, base: g.base, pedidos: g.pedidos, receita: Math.round(g.receita), racional: rfmDefs[k].racional, acao: rfmDefs[k].acao };
      })
      .sort(function(a, b) { return b.receita - a.receita; });

    // ── Curva ABC (regras de negócio) ──
    // A: > 3 compras E >= R$3.000 · B: <= 2 compras E >= R$2.000 · C: restante
    var abcStats = {
      A: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
      B: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
      C: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
    };
    clientList.forEach(function(c) {
      var cls;
      if      (c.numOrders >= 10 && c.totalSpent >= 10000) cls = 'A';
      else if (c.numOrders >= 5  && c.totalSpent >= 5000)  cls = 'B';
      else                                                cls = 'C';
      c.abcClass = cls;
      var st = abcStats[cls];
      st.count++;
      st.revenue     += c.totalSpent;
      st.totalOrders += c.numOrders;
      if (c.totalSpent > 0) {
        st.minTicket = Math.min(st.minTicket, c.totalSpent);
        st.maxTicket = Math.max(st.maxTicket, c.totalSpent);
      }
      if (c.avgGap != null) { st.gapSum += c.avgGap; st.gapCount++; }
    });
    var abcTotal = clientList.reduce(function(s, c) { return s + c.totalSpent; }, 0);

    // Subgrupos de VIP (Classe A) por recência
    var vipA = clientList.filter(function(c) { return c.abcClass === 'A'; });
    var vipAtivos   = vipA.filter(function(c) { return c.recency <= 180; });
    var vipTrabalho = vipA.filter(function(c) { return c.recency > 180 && c.recency <= 540; });
    var vipPerdidos = vipA.filter(function(c) { return c.recency > 540; });
    function subGroup(list) {
      var rev    = list.reduce(function(s, c) { return s + c.totalSpent; }, 0);
      var orders = list.reduce(function(s, c) { return s + c.numOrders; }, 0);
      var avgRec = list.length > 0 ? Math.round(list.reduce(function(s, c) { return s + c.recency; }, 0) / list.length) : 0;
      return { count: list.length, revenue: Math.round(rev), avgRecency: avgRec, ticketMedio: orders > 0 ? Math.round(rev / orders) : 0 };
    }
    var abc = ['A', 'B', 'C'].map(function(cls) {
      var st = abcStats[cls];
      var row = {
        classe:         cls,
        count:          st.count,
        revenue:        Math.round(st.revenue),
        pctCount:       total > 0 ? Math.round(st.count / total * 100) : 0,
        pctRevenue:     abcTotal > 0 ? Math.round(st.revenue / abcTotal * 100) : 0,
        ticketMedio:    st.totalOrders > 0 ? Math.round(st.revenue / st.totalOrders) : 0,
        avgDaysBetween: st.gapCount > 0 ? Math.round(st.gapSum / st.gapCount) : 0,
        minTicket:      st.minTicket === Infinity ? 0 : Math.round(st.minTicket),
        maxTicket:      Math.round(st.maxTicket),
      };
      if (cls === 'A') {
        row.vipAtivos   = subGroup(vipAtivos);
        row.vipTrabalho = subGroup(vipTrabalho);
        row.vipPerdidos = subGroup(vipPerdidos);
      }
      return row;
    });

    // ── Cohort Analysis ──
    var MONTHS_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    var cohortMap = {};
    clientList.forEach(function(c) {
      var yr  = c.firstOrder.getFullYear();
      var mo  = c.firstOrder.getMonth();
      var key = yr + '-' + pad(mo + 1);
      if (!cohortMap[key]) cohortMap[key] = { yr: yr, mo: mo, clients: [] };
      cohortMap[key].clients.push(c);
    });

    var todayYr  = today.getFullYear();
    var todayMo  = today.getMonth();
    var todayKey = todayYr + '-' + pad(todayMo + 1);

    var cohort = Object.keys(cohortMap).sort().slice(-13).map(function(key) {
      var g        = cohortMap[key];
      var cn       = g.clients.length;
      var totalRec = g.clients.reduce(function(s, c) { return s + c.totalSpent; }, 0);
      var m        = [100];
      for (var off = 1; off <= cohortMonths; off++) {
        var tYr  = g.yr + Math.floor((g.mo + off) / 12);
        var tMo  = (g.mo + off) % 12;
        var tKey = tYr + '-' + pad(tMo + 1);
        if (tKey > todayKey) { m.push(null); continue; }
        var retained = g.clients.filter(function(c) {
          return c.orders.some(function(o) {
            return o.data.getFullYear() === tYr && o.data.getMonth() === tMo;
          });
        }).length;
        m.push(cn > 0 ? Math.round(retained / cn * 100) : 0);
      }
      return { safra: MONTHS_PT[g.mo] + '. ' + g.yr, clientes: cn, receita: Math.round(totalRec), cac: 0, gasto: 0, m: m };
    });

    return {
      kpis: {
        total:          total,
        recompras:      recompras,
        pctRecomp:      total > 0 ? Math.round(recompras / total * 100) : 0,
        ltv:            Math.round(ltv),
        avgDaysBetween: Math.round(avgDays),
        totalReceita:   Math.round(totalReceita),
      },
      rfm:    rfm,
      cohort: cohort,
      abc:    abc,
    };

  } catch (err) {
    return { kpis: {}, rfm: [], cohort: [], abc: [], erro: 'Erro em getClientes: ' + err.message };
  }
}

/* ===== LOGISTICA ===== */
function getLogistica(dates) {
  var rows = readSheet(CONFIG.ABA_VENDAS, CONFIG.ID_VENDAS);
  if (rows.length < 2) return { statusList: [], pedidos: [], totalPedidos: 0 };

  var h = rows[0];
  var iData   = findCol(h, ['data','date','Data','Data Pedido','Created at','Paid at','data_pedido']);
  var iPedido = findCol(h, ['Número do Pedido','N° Pedido','Numero do Pedido','pedido','ordem','Name','#']);
  var iNome   = findCol(h, ['Nome do comprador','Nome do Comprador','nome do comprador','Nome','Cliente','Customer','nome_cliente']);
  var iStatus = findCol(h, ['Status de Entrega','status de entrega','Status Entrega','status_entrega','Entrega']);

  if (iStatus < 0) return { statusList: [], pedidos: [], totalPedidos: 0, erro: 'Coluna "Status de Entrega" não encontrada' };

  var contadores = {};
  var pedidos = [];
  var totalPedidos = 0;

  rows.slice(1).forEach(function(r) {
    var d = parseDate(r[iData]);
    if (!d || d < dates.start || d > dates.end) return;

    var status = trim(r[iStatus]);
    if (!status) return;

    totalPedidos++;
    contadores[status] = (contadores[status] || 0) + 1;

    pedidos.push({
      pedido:  iPedido >= 0 ? trim(r[iPedido]) : '',
      data:    fmtDateBR(d),
      cliente: iNome   >= 0 ? trim(r[iNome])   : '',
      status:  status,
    });
  });

  var statusList = Object.keys(contadores).map(function(s) {
    return { status: s, total: contadores[s] };
  }).sort(function(a, b) { return b.total - a.total; });

  pedidos.reverse();

  return {
    statusList:   statusList,
    pedidos:      pedidos,
    totalPedidos: totalPedidos,
  };
}

/* ===== DEBUG ===== */

function debugHeaders() {
  var rows = readSheet(CONFIG.ABA_VENDAS, CONFIG.ID_VENDAS);
  if (rows.length < 1) return { error: 'Aba vazia' };
  var h = rows[0];
  var colLetter = function(i) {
    var s = '';
    var n = i + 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  var headers = h.map(function(v, i) {
    return { index: i, col: colLetter(i), header: String(v) };
  }).filter(function(x){ return x.header.trim() !== ''; });

  var cm = COL_MAP.vendas;
  var iVendedor  = findCol(h, ['Pessoa que registrou a venda','pessoa que registrou a venda','Pessoa que registrou','pessoa que registrou']);
  var iVendedor2 = findCol(h, ['Vendedor','vendedor','seller','Atendente','atendente']);
  var iSemFrete  = findCol(h, cm.venda_sem_frete);
  var iReceita   = findCol(h, cm.receita);
  var iStatusPag = findCol(h, cm.status_pag);
  var iData      = findCol(h, cm.data);
  var iEmail     = findCol(h, cm.email);
  return {
    totalCols:    h.length,
    headers:      headers,
    colsDetected: {
      email:        { index: iEmail,     col: iEmail     >= 0 ? colLetter(iEmail)     : 'NÃO ENCONTRADO', header: iEmail     >= 0 ? h[iEmail]     : null },
      vendedor:     { index: iVendedor,  col: iVendedor  >= 0 ? colLetter(iVendedor)  : 'NÃO ENCONTRADO', header: iVendedor  >= 0 ? h[iVendedor]  : null },
      vendedor2:    { index: iVendedor2, col: iVendedor2 >= 0 ? colLetter(iVendedor2) : 'NÃO ENCONTRADO', header: iVendedor2 >= 0 ? h[iVendedor2] : null },
      semFrete:     { index: iSemFrete,  col: iSemFrete  >= 0 ? colLetter(iSemFrete)  : 'NÃO ENCONTRADO', header: iSemFrete  >= 0 ? h[iSemFrete]  : null },
      receita:      { index: iReceita,   col: iReceita   >= 0 ? colLetter(iReceita)   : 'NÃO ENCONTRADO', header: iReceita   >= 0 ? h[iReceita]   : null },
      statusPag:    { index: iStatusPag, col: iStatusPag >= 0 ? colLetter(iStatusPag) : 'NÃO ENCONTRADO', header: iStatusPag >= 0 ? h[iStatusPag] : null },
      data:         { index: iData,      col: iData      >= 0 ? colLetter(iData)      : 'NÃO ENCONTRADO', header: iData      >= 0 ? h[iData]      : null },
    }
  };
}

function debugMari(startStr, endStr) {
  var rows = readSheet(CONFIG.ABA_VENDAS, CONFIG.ID_VENDAS);
  if (rows.length < 2) return { error: 'Sem dados' };
  var h = rows[0];
  var cm = COL_MAP.vendas;

  var iData      = findCol(h, cm.data);
  var iReceita   = findCol(h, cm.receita);
  var iSemFrete  = findCol(h, cm.venda_sem_frete);
  var iStatusPag = findCol(h, cm.status_pag);
  var iVendedor  = findCol(h, ['Pessoa que registrou a venda','pessoa que registrou a venda','Pessoa que registrou','pessoa que registrou']);
  var iVendedor2 = findCol(h, ['Vendedor','vendedor','seller','Atendente','atendente']);

  var colLetter = function(i) {
    var s = ''; var n = i + 1;
    while (n > 0) { var r = (n-1)%26; s = String.fromCharCode(65+r)+s; n = Math.floor((n-1)/26); }
    return s;
  };

  var start = startStr ? parseDate(startStr) : new Date(2026,2,1);
  var end   = endStr   ? parseDate(endStr)   : new Date(2026,2,31);
  if (start) start.setHours(0,0,0,0);
  if (end)   end.setHours(23,59,59,999);

  var mariRows = [];
  var mariTotal = 0, mariConfirmadoTotal = 0;
  rows.slice(1).forEach(function(r) {
    var d  = parseDate(r[iData]);
    if (!d || d < start || d > end) return;
    var v1 = iVendedor  >= 0 ? trim(r[iVendedor])  : '';
    var v2 = iVendedor2 >= 0 ? trim(r[iVendedor2]) : '';
    var vend = v1 || v2 || 'Outros';
    if (vend !== 'Mari') return;
    var st  = iStatusPag >= 0 ? trim(r[iStatusPag]) : '';
    var val = iReceita   >= 0 ? parseNum(r[iReceita]) : 0;
    var sf  = iSemFrete  >= 0 ? parseNum(r[iSemFrete])  : val;
    mariTotal += sf;
    if (st === 'Confirmado') mariConfirmadoTotal += sf;
    if (mariRows.length < 15) {
      mariRows.push({ colV1: colLetter(iVendedor), v1: v1, colV2: colLetter(iVendedor2), v2: v2, status: st, receita: val, semFrete: sf });
    }
  });
  return {
    iVendedorPrimario:      { index: iVendedor,  col: iVendedor  >= 0 ? colLetter(iVendedor)  : 'N/A', header: iVendedor  >= 0 ? String(h[iVendedor])  : null },
    iVendedorFallback:      { index: iVendedor2, col: iVendedor2 >= 0 ? colLetter(iVendedor2) : 'N/A', header: iVendedor2 >= 0 ? String(h[iVendedor2]) : null },
    iSemFrete:              { index: iSemFrete,  col: iSemFrete  >= 0 ? colLetter(iSemFrete)  : 'N/A', header: iSemFrete  >= 0 ? String(h[iSemFrete])  : null },
    mariTotalSemFrete:       round2(mariTotal),
    mariConfirmadoSemFrete:  round2(mariConfirmadoTotal),
    primeiraLinhasMari:      mariRows,
  };
}

/* ===== UTILITÁRIOS ===== */

function readSheet(tabName, altId) {
  var id = (altId && altId.length > 5) ? altId : CONFIG.MASTER_ID;
  if (!id || id === 'COLE_O_ID_DA_PLANILHA_AQUI') {
    throw new Error('ID da planilha não configurado. Preencha CONFIG.MASTER_ID no script.');
  }
  var ss    = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Aba "' + tabName + '" não encontrada em ' + id);
  return sheet.getDataRange().getValues();
}

function findCol(headers, options) {
  var lc = headers.map(function (h) { return String(h).toLowerCase().trim(); });
  for (var i = 0; i < options.length; i++) {
    var idx = lc.indexOf(options[i].toLowerCase().trim());
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/R\$\s?/g, '').replace(/\./g, '').replace(',', '.').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim();
  var m;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    var yr = parseInt(m[3]); if (yr < 100) yr += 2000;
    return new Date(yr, parseInt(m[2]) - 1, parseInt(m[1]));
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function normalizeStatus(st) {
  if (st.includes('pago') || st.includes('paid') || st.includes('aprovado') || st.includes('approved') || st.includes('concluído'))
    return 'Aprovado';
  if (st.includes('transit') || st.includes('trânsito') || st.includes('enviado') || st.includes('shipped') || st.includes('saiu'))
    return 'Em Trânsito';
  if (st.includes('cancel') || st.includes('recusado') || st.includes('refused') || st.includes('rejeitado'))
    return 'Cancelado';
  if (st.includes('entreg') || st.includes('delivered') || st.includes('finaliz'))
    return 'Entregue';
  return 'Pendente';
}

function getDateRange(period, startStr, endStr) {
  var end   = new Date(); end.setHours(23, 59, 59, 999);
  var start = new Date(end);

  if (period === 'custom' && startStr && endStr) {
    start = parseDate(startStr); start.setHours(0, 0, 0, 0);
    end   = parseDate(endStr);   end.setHours(23, 59, 59, 999);
  } else if (period === 'last_7d')  {
    start.setDate(end.getDate() - 6);  start.setHours(0, 0, 0, 0);
  } else if (period === 'last_14d') {
    start.setDate(end.getDate() - 13); start.setHours(0, 0, 0, 0);
  } else if (period === 'last_30d') {
    start.setDate(end.getDate() - 29); start.setHours(0, 0, 0, 0);
  } else if (period === 'this_month') {
    start = new Date(end.getFullYear(), end.getMonth(), 1);
  } else if (period === 'last_month') {
    start = new Date(end.getFullYear(), end.getMonth() - 1, 1);
    end   = new Date(end.getFullYear(), end.getMonth(), 0);
    end.setHours(23, 59, 59, 999);
  } else {
    start.setDate(end.getDate() - 29); start.setHours(0, 0, 0, 0);
  }
  return { start: start, end: end };
}

function fmtDate(d) {
  if (!d) return '';
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}
function fmtDateBR(d) {
  if (!d) return '';
  return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear();
}
function pad(n)    { return n < 10 ? '0' + n : '' + n; }
function trim(v)   { return String(v || '').trim(); }
function round2(n) { return Math.round(n * 100) / 100; }
function round1(n) { return Math.round(n * 10)  / 10;  }

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
