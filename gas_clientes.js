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

    // ── RFM Scoring ───────────────────────────────────────────────────────
    var sortedSpent = clientList.map(function(c) { return c.totalSpent; }).sort(function(a, b) { return a - b; });
    var n = sortedSpent.length;
    var p20 = sortedSpent[Math.floor(n * 0.2)] || 0;
    var p40 = sortedSpent[Math.floor(n * 0.4)] || 0;
    var p60 = sortedSpent[Math.floor(n * 0.6)] || 0;
    var p80 = sortedSpent[Math.floor(n * 0.8)] || 0;

    function rScore(days) { return days <= 30 ? 5 : days <= 60 ? 4 : days <= 90 ? 3 : days <= 180 ? 2 : 1; }
    function fScore(num)  { return num  >= 9  ? 5 : num  >= 5  ? 4 : num  >= 3  ? 3 : num  >= 2  ? 2 : 1; }
    function mScore(val)  { return val  > p80 ? 5 : val  > p60 ? 4 : val  > p40 ? 3 : val  > p20 ? 2 : 1; }

    var rfmDefs = {
      'Champions':  { icon: '🏆', acao: 'Programa fidelidade & embaixadores' },
      'VIPs':       { icon: '💎', acao: 'Tratamento exclusivo & upsell premium' },
      'Fiéis':      { icon: '🔁', acao: 'Programa de pontos & recompensa' },
      'Potenciais': { icon: '🌱', acao: 'Incentivar 2ª compra com oferta especial' },
      'Novos':      { icon: '✨', acao: 'Boas-vindas & onboarding por email' },
      'Em Risco':   { icon: '⚠️', acao: 'Campanha de reativação urgente' },
      'Dormentes':  { icon: '💤', acao: 'Desconto exclusivo de reativação' },
      'Perdidos':   { icon: '❌', acao: 'Win-back agressivo ou limpeza da base' },
    };
    var rfmGroups = {};
    Object.keys(rfmDefs).forEach(function(k) {
      rfmGroups[k] = { base: 0, Pedidos: 0, receita: 0 };
    });

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
      else                                     seg = 'Perdidos';

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
          acao:     rfmDefs[k].acao,
        };
      })
      .sort(function(a, b) { return b.receita - a.receita; });

    // ── Curva ABC ─────────────────────────────────────────────────────────
    var sortedByValue = clientList.slice().sort(function(a, b) { return b.totalSpent - a.totalSpent; });
    var abcTotal = sortedByValue.reduce(function(s, c) { return s + c.totalSpent; }, 0);
    var abcStats = {
      A: { count: 0, revenue: 0, minTicket: Infinity, maxTicket: 0 },
      B: { count: 0, revenue: 0, minTicket: Infinity, maxTicket: 0 },
      C: { count: 0, revenue: 0, minTicket: Infinity, maxTicket: 0 },
    };
    var cumRev = 0;
    sortedByValue.forEach(function(c) {
      cumRev += c.totalSpent;
      var cumPct = abcTotal > 0 ? cumRev / abcTotal : 1;
      c.abcClass = cumPct <= 0.80 ? 'A' : cumPct <= 0.95 ? 'B' : 'C';
      var st = abcStats[c.abcClass];
      st.count++;
      st.revenue += c.totalSpent;
      if (c.totalSpent > 0) {
        st.minTicket = Math.min(st.minTicket, c.totalSpent);
        st.maxTicket = Math.max(st.maxTicket, c.totalSpent);
      }
    });
    // Subgrupos de VIP (Classe A) por recência — base da decisão
    var vipA = sortedByValue.filter(function(c) { return c.abcClass === 'A'; });
    var vipAtivos   = vipA.filter(function(c) { return c.recency <= 180; });
    var vipTrabalho = vipA.filter(function(c) { return c.recency > 180 && c.recency <= 540; });
    var vipPerdidos = vipA.filter(function(c) { return c.recency > 540; });
    function subGroup(list) {
      var rev = list.reduce(function(s, c) { return s + c.totalSpent; }, 0);
      var avgRec = list.length > 0 ? Math.round(list.reduce(function(s, c) { return s + c.recency; }, 0) / list.length) : 0;
      var avgTicket = list.length > 0 ? Math.round(rev / list.length) : 0;
      return { count: list.length, revenue: Math.round(rev), avgRecency: avgRec, ticketMedio: avgTicket };
    }
    var abc = ['A', 'B', 'C'].map(function(cls) {
      var st = abcStats[cls];
      var row = {
        classe:      cls,
        count:       st.count,
        revenue:     Math.round(st.revenue),
        pctCount:    total > 0 ? Math.round(st.count / total * 100) : 0,
        pctRevenue:  abcTotal > 0 ? Math.round(st.revenue / abcTotal * 100) : 0,
        ticketMedio: st.count > 0 ? Math.round(st.revenue / st.count) : 0,
        minTicket:   st.minTicket === Infinity ? 0 : Math.round(st.minTicket),
        maxTicket:   Math.round(st.maxTicket),
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
