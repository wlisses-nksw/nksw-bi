/**
 * NKSW BI — Módulo Clientes
 * Adicione este código ao seu Google Apps Script existente.
 *
 * No seu doGet(e), inclua o case abaixo:
 *
 *   if (section === 'clientes') {
 *     return respond(getClientesData(e.parameter));
 *   }
 *
 * onde `respond` é sua função que retorna ContentService com JSON.
 */

function getClientesData(params) {
  try {
    params = params || {};
    var cohortMonths = parseInt(params.cohortMonths) || 24;
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Pedidos');
    if (!sheet) return { ok: false, error: 'Aba "Pedidos" não encontrada' };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: false, error: 'Sem dados na aba "Pedidos"' };

    // ── Auto-detecta colunas pelo cabeçalho ──────────────────────────────
    var rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var headers = rawHeaders.map(function(h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, '_');
    });

    function findCol(names) {
      for (var i = 0; i < names.length; i++) {
        var idx = headers.indexOf(names[i]);
        if (idx >= 0) return idx;
      }
      return -1;
    }

    var iEmail   = findCol(['email', 'e-mail', 'email_cliente', 'cliente_email', 'email_comprador']);
    var iData    = findCol(['data', 'data_pedido', 'data_criacao', 'dt_pedido', 'date', 'created_at']);
    var iReceita = findCol(['receita', 'valor', 'total', 'venda_sem_frete', 'receita_aprovada', 'valor_aprovado', 'venda_s/_frete']);
    var iStatus  = findCol(['status', 'status_pagamento', 'situacao', 'payment_status', 'status_do_pedido']);
    var iNome    = findCol(['nome', 'name', 'nome_cliente', 'cliente', 'comprador', 'nome_comprador']);

    if (iEmail < 0) {
      return { ok: false, error: 'Coluna de email não encontrada. Cabeçalhos detectados: ' + rawHeaders.join(', ') };
    }

    // ── Lê todos os dados ─────────────────────────────────────────────────
    var rows  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var clientes = {};

    rows.forEach(function(row) {
      var email = String(row[iEmail] || '').toLowerCase().trim();
      if (!email || email.indexOf('@') < 0) return;

      // Ignora Pedidos cancelados/estornados
      if (iStatus >= 0) {
        var st = String(row[iStatus] || '').toLowerCase();
        if (st.indexOf('cancel') >= 0 || st.indexOf('estorn') >= 0 || st.indexOf('reembol') >= 0) return;
      }

      var dataVal = iData >= 0 ? row[iData] : null;
      if (!dataVal) return;
      var dt = dataVal instanceof Date ? new Date(dataVal) : new Date(dataVal);
      if (isNaN(dt.getTime())) return;
      dt.setHours(0, 0, 0, 0);

      var receita = iReceita >= 0
        ? (parseFloat(String(row[iReceita]).replace(',', '.')) || 0) : 0;
      var nome = iNome >= 0 ? String(row[iNome] || '') : '';

      if (!clientes[email]) {
        clientes[email] = { email: email, nome: nome, orders: [] };
      }
      clientes[email].orders.push({ data: dt, receita: receita });
    });

    var clientList = Object.values(clientes);
    if (!clientList.length) return { ok: false, error: 'Nenhum cliente encontrado em "Pedidos"' };

    // ── Métricas por cliente ──────────────────────────────────────────────
    clientList.forEach(function(c) {
      c.orders.sort(function(a, b) { return a.data - b.data; });
      c.firstOrder = c.orders[0].data;
      c.lastOrder  = c.orders[c.orders.length - 1].data;
      c.numOrders  = c.orders.length;
      c.totalSpent = c.orders.reduce(function(s, o) { return s + o.receita; }, 0);
      c.recency    = Math.floor((today - c.lastOrder) / 86400000); // dias desde última compra

      if (c.numOrders > 1) {
        var totalGap = 0;
        for (var i = 1; i < c.orders.length; i++) {
          totalGap += (c.orders[i].data - c.orders[i - 1].data) / 86400000;
        }
        c.avgGap = totalGap / (c.numOrders - 1);
      }
    });

    // ── KPIs globais ──────────────────────────────────────────────────────
    var total        = clientList.length;
    var recompras    = clientList.filter(function(c) { return c.numOrders >= 2; }).length;
    var totalReceita = clientList.reduce(function(s, c) { return s + c.totalSpent; }, 0);
    var ltv          = total > 0 ? totalReceita / total : 0;
    var gapList      = clientList.filter(function(c) { return c.avgGap != null; });
    var avgDays      = gapList.length > 0
      ? gapList.reduce(function(s, c) { return s + c.avgGap; }, 0) / gapList.length : 0;

    // ── RFM Scoring (R percentil da base, F fixo, M percentil monetario) ─
    var sortedSpent = clientList.map(function(c) { return c.totalSpent; }).sort(function(a, b) { return a - b; });
    var n = sortedSpent.length;
    var p20 = sortedSpent[Math.floor(n * 0.2)] || 0;
    var p40 = sortedSpent[Math.floor(n * 0.4)] || 0;
    var p60 = sortedSpent[Math.floor(n * 0.6)] || 0;
    var p80 = sortedSpent[Math.floor(n * 0.8)] || 0;

    // R: percentil de recência (menor = mais recente = score maior)
    var sortedRec = clientList.map(function(c) { return c.recency; }).sort(function(a, b) { return a - b; });
    var rp20 = sortedRec[Math.floor(n * 0.2)] || 0;
    var rp40 = sortedRec[Math.floor(n * 0.4)] || 0;
    var rp60 = sortedRec[Math.floor(n * 0.6)] || 0;
    var rp80 = sortedRec[Math.floor(n * 0.8)] || 0;

    function rScore(days) { return days <= rp20 ? 5 : days <= rp40 ? 4 : days <= rp60 ? 3 : days <= rp80 ? 2 : 1; }
    function fScore(num)  { return num  >= 9  ? 5 : num  >= 5  ? 4 : num  >= 3  ? 3 : num  >= 2  ? 2 : 1; }
    function mScore(val)  { return val  > p80 ? 5 : val  > p60 ? 4 : val  > p40 ? 3 : val  > p20 ? 2 : 1; }

    // ── 13 Segmentos RFM (metodologia planilha Curva ABC RFM) ────────────
    var rfmDefs = {
      'VIP Ativo':            { icon: '💎', racional: 'Top clientes com score RFM maximo (14-15) comprando nos ultimos 6 meses — nucleo e embaixadores da marca.',           acao: 'Programa VIP exclusivo, acesso antecipado a lancamentos, presentes personalizados' },
      'VIP Reaquecer':        { icon: '🔔', racional: 'Top clientes esfriando (6-12 meses sem compra) — janela de reconquista antes da perda definitiva.',                   acao: 'Contato personalizado pelo atendente, oferta exclusiva de reengajamento' },
      'VIP Em Risco':         { icon: '⚠️', racional: 'Top clientes inativos ha mais de 1 ano — risco real de perda permanente de alto valor.',                              acao: 'Win-back urgente com beneficio exclusivo, ligacao direta do gerente de contas' },
      'Alto Valor Ativo':     { icon: '🌟', racional: 'Gasto historico no top 20% da base, comprando nos ultimos 6 meses — candidatos naturais ao grupo VIP.',               acao: 'Incentivar frequencia, convidar para clube exclusivo, apresentar produtos premium' },
      'Alto Valor Morno':     { icon: '🌡️', racional: 'Alto gasto historico (top 20%) mas comprando menos — risco de migrar para concorrente.',                              acao: 'Reativacao personalizada, oferta de produto complementar ao historico de compras' },
      'Alto Valor Em Risco':  { icon: '🚨', racional: 'Muito gastaram, mas inativos ha 1 a 2 anos — recuperacao dificil, impacto alto se perdidos.',                         acao: 'Campanha win-back agressiva, desconto significativo ou produto exclusivo' },
      'Alto Valor Hibernado': { icon: '❄️', racional: 'Ex-clientes de alto valor sem compra ha mais de 2 anos — avaliar custo x beneficio de reativacao.',                   acao: 'Campanha em data especial (aniversario), avaliar ROI antes de investir na reativacao' },
      'Potencial Ativo':      { icon: '🚀', racional: 'Bom score RFM (12-13) comprando regularmente — caminho claro para Alto Valor se a frequencia aumentar.',              acao: 'Cross-sell e upsell, apresentar linha premium, recomendacao personalizada' },
      'Potencial Morno':      { icon: '💡', racional: 'Bom historico mas comprando menos — precisam de estimulo para voltar ao ritmo anterior.',                              acao: 'E-mail de reengajamento com curadoria baseada no historico de compras' },
      'Potencial Inativo':    { icon: '⏸️', racional: 'Tinham bom comportamento de compra mas sumiu — reativacao pode gerar retorno relevante.',                             acao: 'Fluxo de reativacao automatizado com incentivo progressivo (oferta escalonada)' },
      'Novo/Eventual':        { icon: '✨', racional: 'Poucos pedidos e baixo score, mas compra recente — ainda avaliando a marca, momento de criar vinculo.',                acao: 'Sequencia de boas-vindas, pesquisa de satisfacao, oferta exclusiva de 2a compra' },
      'Baixo Engajamento':    { icon: '💤', racional: 'Compras raras ou unicas ha muito tempo — base de baixo potencial imediato, custo alto de reativacao.',                 acao: 'Automacao de baixo custo, acionar apenas em campanhas sazonais de alto volume' },
      'Emergente':            { icon: '🌱', racional: 'Primeira compra muito recente — ainda formando opiniao sobre a marca, momento decisivo da experiencia.',               acao: 'Impressionar na pos-compra (unboxing, agradecimento), capturar feedback imediato' },
    };
    var rfmGroups = {};
    Object.keys(rfmDefs).forEach(function(k) { rfmGroups[k] = { base: 0, Pedidos: 0, receita: 0 }; });

    clientList.forEach(function(c) {
      var r = rScore(c.recency), f = fScore(c.numOrders), m = mScore(c.totalSpent);
      var score = r + f + m;
      c.rScore = r; c.fScore = f; c.mScore = m; c.rfmScore = score;

      // ABC por score RFM: A=14-15, B=12-13 ou M=5, C=restante
      if      (score >= 14)            c.abcClass = 'A';
      else if (score >= 12 || m === 5) c.abcClass = 'B';
      else                             c.abcClass = 'C';

      // Segmento RFM (13 segmentos com recencia)
      var rec = c.recency;
      var seg;
      if (score >= 14) {
        seg = rec <= 180 ? 'VIP Ativo' : rec <= 365 ? 'VIP Reaquecer' : 'VIP Em Risco';
      } else if (m === 5) {
        seg = rec <= 180 ? 'Alto Valor Ativo' : rec <= 365 ? 'Alto Valor Morno' : rec <= 730 ? 'Alto Valor Em Risco' : 'Alto Valor Hibernado';
      } else if (score >= 12) {
        seg = rec <= 180 ? 'Potencial Ativo' : rec <= 365 ? 'Potencial Morno' : 'Potencial Inativo';
      } else {
        seg = (rec <= 60 && c.numOrders === 1) ? 'Emergente' : rec <= 365 ? 'Novo/Eventual' : 'Baixo Engajamento';
      }
      rfmGroups[seg].base++;
      rfmGroups[seg].Pedidos += c.numOrders;
      rfmGroups[seg].receita += c.totalSpent;
    });

    var rfm = Object.keys(rfmDefs)
      .filter(function(k) { return rfmGroups[k].base > 0; })
      .map(function(k) {
        var g = rfmGroups[k];
        return {
          segmento: rfmDefs[k].icon + ' ' + k,
          base:     g.base,
          Pedidos:  g.Pedidos,
          receita:  Math.round(g.receita),
          racional: rfmDefs[k].racional,
          acao:     rfmDefs[k].acao,
        };
      })
      .sort(function(a, b) { return b.receita - a.receita; });

    // ── Curva ABC (baseada em score RFM — classificacao ja feita acima) ──
    // A: Score 14-15  ·  B: Score 12-13 ou M=5  ·  C: restante
    var abcStats = {
      A: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
      B: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
      C: { count: 0, revenue: 0, totalOrders: 0, minTicket: Infinity, maxTicket: 0, gapSum: 0, gapCount: 0 },
    };
    clientList.forEach(function(c) {
      var st = abcStats[c.abcClass];
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

    // Subgrupos de VIP (Classe A) por recencia — alinhado com segmentos RFM
    var vipA       = clientList.filter(function(c) { return c.abcClass === 'A'; });
    var vipAtivos   = vipA.filter(function(c) { return c.recency <= 180; });
    var vipTrabalho = vipA.filter(function(c) { return c.recency > 180 && c.recency <= 365; });
    var vipPerdidos = vipA.filter(function(c) { return c.recency > 365; });
    function subGroup(list) {
      var rev     = list.reduce(function(s, c) { return s + c.totalSpent; }, 0);
      var orders  = list.reduce(function(s, c) { return s + c.numOrders; }, 0);
      var avgRec  = list.length > 0 ? Math.round(list.reduce(function(s, c) { return s + c.recency; }, 0) / list.length) : 0;
      return { count: list.length, revenue: Math.round(rev), avgRecency: avgRec, ticketMedio: orders > 0 ? Math.round(rev / orders) : 0 };
    }
    var abc = ['A', 'B', 'C'].map(function(cls) {
      var st = abcStats[cls];
      var row = {
        classe:          cls,
        count:           st.count,
        revenue:         Math.round(st.revenue),
        pctCount:        total > 0 ? Math.round(st.count / total * 100) : 0,
        pctRevenue:      abcTotal > 0 ? Math.round(st.revenue / abcTotal * 100) : 0,
        ticketMedio:     st.totalOrders > 0 ? Math.round(st.revenue / st.totalOrders) : 0,
        avgDaysBetween:  st.gapCount > 0 ? Math.round(st.gapSum / st.gapCount) : 0,
        minTicket:       st.minTicket === Infinity ? 0 : Math.round(st.minTicket),
        maxTicket:       Math.round(st.maxTicket),
      };
      if (cls === 'A') {
        row.vipAtivos   = subGroup(vipAtivos);
        row.vipTrabalho = subGroup(vipTrabalho);
        row.vipPerdidos = subGroup(vipPerdidos);
      }
      return row;
    });

    // ── Cohort Analysis ───────────────────────────────────────────────────
    var MONTHS_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    var cohortMap = {};

    clientList.forEach(function(c) {
      var yr  = c.firstOrder.getFullYear();
      var mo  = c.firstOrder.getMonth();
      var key = yr + '-' + (mo < 9 ? '0' : '') + (mo + 1);
      if (!cohortMap[key]) cohortMap[key] = { yr: yr, mo: mo, clients: [] };
      cohortMap[key].clients.push(c);
    });

    var todayYr = today.getFullYear();
    var todayMo = today.getMonth();
    var todayKey = todayYr + '-' + (todayMo < 9 ? '0' : '') + (todayMo + 1);

    var cohort = Object.keys(cohortMap).sort().slice(-13).map(function(key) {
      var g    = cohortMap[key];
      var cn   = g.clients.length;
      var totalRec = g.clients.reduce(function(s, c) { return s + c.totalSpent; }, 0);
      var m    = [100]; // Mês 0 sempre 100%

      for (var off = 1; off <= cohortMonths; off++) {
        var tYr  = g.yr + Math.floor((g.mo + off) / 12);
        var tMo  = (g.mo + off) % 12;
        var tKey = tYr + '-' + (tMo < 9 ? '0' : '') + (tMo + 1);
        if (tKey > todayKey) { m.push(null); continue; }

        var retained = g.clients.filter(function(c) {
          return c.orders.some(function(o) {
            return o.data.getFullYear() === tYr && o.data.getMonth() === tMo;
          });
        }).length;
        m.push(cn > 0 ? Math.round(retained / cn * 100) : 0);
      }

      return {
        safra:    MONTHS_PT[g.mo] + '. ' + g.yr,
        clientes: cn,
        receita:  Math.round(totalRec),
        cac:      0,
        gasto:    0,
        m:        m,
      };
    });

    return {
      ok: true,
      clientes: {
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
      },
    };

  } catch (err) {
    return { ok: false, error: 'Erro em getClientesData: ' + err.message };
  }
}
