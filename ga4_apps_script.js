// ═══════════════════════════════════════════════════════════════════════════
// GA4 Dashboard API — Google Apps Script
// Naked SW / Ananda Fróes
//
// INSTRUÇÕES DE DEPLOY:
// 1. Acesse script.google.com → Novo projeto
// 2. Apague o código padrão e cole TODO este arquivo
// 3. Clique em: Implantar → Nova implantação
//    → Tipo: App da Web
//    → Executar como: Eu mesmo
//    → Quem tem acesso: Qualquer pessoa
// 4. Autorize as permissões quando solicitado
// 5. Copie a URL gerada e cole no dashboard
// ═══════════════════════════════════════════════════════════════════════════

// ─── CONFIGURAÇÃO ──────────────────────────────────────────────────────────
const GA4_PROPERTY_ID = '374872616'; // Sua property ID

// ─── ENTRY POINT ───────────────────────────────────────────────────────────
function doGet(e) {
  try {
    // Instala o gatilho diário automaticamente na primeira execução
    autoInstallTrigger();

    const period    = e?.parameter?.period    || 'last_7d';
    const section   = e?.parameter?.section   || 'all';
    const startDate = e?.parameter?.startDate || null;
    const endDate   = e?.parameter?.endDate   || null;

    const data = fetchAllData(period, section, startDate, endDate);

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders ? addCors(data) : buildResponse(data);
  } catch (err) {
    return buildErrorResponse(err.message);
  }
}

// Cria o gatilho diário às 8h uma única vez — detecta automaticamente se já existe
function autoInstallTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('triggerInstalled') === 'true') return; // já instalado

  // Remove duplicatas por segurança
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  props.setProperty('triggerInstalled', 'true');
  Logger.log('✅ Gatilho diário instalado automaticamente às 8h.');
}

function buildResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function buildErrorResponse(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addCors(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─── FETCH PRINCIPAL ───────────────────────────────────────────────────────
function fetchAllData(period, section, startDate, endDate) {
  const range = getDateRange(period, startDate, endDate);

  const result = {
    period: period,
    dateRange: range,
    generatedAt: new Date().toISOString(),
    propertyId: GA4_PROPERTY_ID
  };

  if (section === 'all' || section === 'overview') {
    result.overview = fetchOverview(range);
  }
  if (section === 'all' || section === 'channels') {
    result.channels = fetchByChannel(range);
  }
  if (section === 'all' || section === 'daily') {
    result.daily = fetchByDate(range);
  }
  if (section === 'all' || section === 'landing') {
    result.landingPages = fetchLandingPages(range);
  }
  if (section === 'all' || section === 'sources') {
    result.sources = fetchBySources(range);
  }
  if (section === 'all' || section === 'funnel') {
    result.funnel = fetchFunnel(range);
  }
  if (section === 'all' || section === 'device') {
    result.device = fetchByDevice(range);
  }
  if (section === 'all' || section === 'geo') {
    result.geo = fetchByGeo(range);
  }
  if (section === 'all' || section === 'demographics') {
    result.demographics = fetchDemographics(range);
  }
  if (section === 'all' || section === 'productFunnel') {
    result.productFunnel = fetchProductFunnel(range);
  }
  if (section === 'all' || section === 'allSources') {
    result.allSources = fetchAllSources(range);
  }

  return result;
}

// ─── MÉTRICAS GERAIS ───────────────────────────────────────────────────────
// GA4 limita 10 métricas por requisição — dividimos em 2 chamadas e mesclamos
function fetchOverview(range) {
  // Requisição 1: métricas de sessão + e-commerce base (10 métricas)
  const req1 = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    metrics: [
      { name: 'sessions' },
      { name: 'activeUsers' },
      { name: 'newUsers' },
      { name: 'bounceRate' },
      { name: 'averageSessionDuration' },
      { name: 'screenPageViews' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'sessionConversionRate' },
      { name: 'addToCarts' }
    ]
  };

  // Requisição 2: funil restante (2 métricas)
  const req2 = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    metrics: [
      { name: 'checkouts' },
      { name: 'ecommercePurchases' }
    ]
  };

  const res1 = AnalyticsData.Properties.runReport(req1, `properties/${GA4_PROPERTY_ID}`);
  const res2 = AnalyticsData.Properties.runReport(req2, `properties/${GA4_PROPERTY_ID}`);

  const row1 = res1.rows?.[0]?.metricValues || [];
  const row2 = res2.rows?.[0]?.metricValues || [];

  const sessions     = safeInt(row1[0]);
  const users        = safeInt(row1[1]);
  const newUsers     = safeInt(row1[2]);
  const bounce       = safeFloat(row1[3]);
  const duration     = safeFloat(row1[4]);
  const screenPageViews = safeInt(row1[5]);
  const transactions = safeInt(row1[6]);
  const revenue      = safeFloat(row1[7]);
  const convRate     = safeFloat(row1[8]);
  const addToCarts   = safeInt(row1[9]);
  const checkouts    = safeInt(row2[0]);
  const purchases    = safeInt(row2[1]);
  const aov          = transactions > 0 ? revenue / transactions : 0;

  return {
    sessions,
    users,
    newUsers,
    newUserRate: sessions > 0 ? (newUsers / sessions) : 0,
    bounceRate: bounce,
    avgSessionDuration: duration,
    screenPageViews,
    transactions,
    purchaseRevenue: revenue,
    sessionConversionRate: convRate,
    addToCarts,
    checkouts,
    purchases,
    averageOrderValue: aov,
    // Funil resumido
    funnelSummary: {
      sessions,
      addToCarts,
      checkouts,
      purchases,
      sessionToCartRate:     sessions > 0   ? addToCarts  / sessions   : 0,
      cartToCheckoutRate:    addToCarts > 0  ? checkouts   / addToCarts : 0,
      checkoutToPurchaseRate: checkouts > 0  ? purchases   / checkouts  : 0,
      overallConvRate:       sessions > 0   ? purchases   / sessions   : 0
    }
  };
}

// ─── SESSÕES POR CANAL ─────────────────────────────────────────────────────
function fetchByChannel(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [{ name: 'sessionDefaultChannelGroup' }],
    metrics: [
      { name: 'sessions' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'bounceRate' },
      { name: 'averageSessionDuration' },
      { name: 'sessionConversionRate' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 10
  };

  const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);

  return (response.rows || []).map(row => ({
    channel:     row.dimensionValues[0].value,
    sessions:    safeInt(row.metricValues[0]),
    transactions: safeInt(row.metricValues[1]),
    revenue:     safeFloat(row.metricValues[2]),
    bounceRate:  safeFloat(row.metricValues[3]),
    avgDuration: safeFloat(row.metricValues[4]),
    convRate:    safeFloat(row.metricValues[5])
  }));
}

// ─── RECEITA POR DIA ───────────────────────────────────────────────────────
function fetchByDate(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [{ name: 'date' }],
    metrics: [
      { name: 'sessions' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'addToCarts' }
    ],
    orderBys: [{ dimension: { dimensionName: 'date' } }]
  };

  const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);

  return (response.rows || []).map(row => {
    const raw  = row.dimensionValues[0].value; // YYYYMMDD
    const date = `${raw.slice(6,8)}/${raw.slice(4,6)}/${raw.slice(0,4)}`;
    return {
      date,
      dateRaw:      raw,
      sessions:     safeInt(row.metricValues[0]),
      transactions: safeInt(row.metricValues[1]),
      revenue:      safeFloat(row.metricValues[2]),
      addToCarts:   safeInt(row.metricValues[3])
    };
  });
}

// ─── LANDING PAGES (tráfego pago) ─────────────────────────────────────────
function fetchLandingPages(range) {
  // Sem filtro de canal — puxa todas, deixa o dashboard filtrar
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [
      { name: 'landingPage' },
      { name: 'sessionDefaultChannelGroup' }
    ],
    metrics: [
      { name: 'sessions' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'bounceRate' },
      { name: 'averageSessionDuration' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 50
  };

  const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);

  return (response.rows || []).map(row => ({
    page:         row.dimensionValues[0].value,
    channel:      row.dimensionValues[1].value,
    sessions:     safeInt(row.metricValues[0]),
    transactions: safeInt(row.metricValues[1]),
    revenue:      safeFloat(row.metricValues[2]),
    bounceRate:   safeFloat(row.metricValues[3]),
    avgDuration:  safeFloat(row.metricValues[4]),
    convRate:     safeInt(row.metricValues[0]) > 0
                    ? safeInt(row.metricValues[1]) / safeInt(row.metricValues[0])
                    : 0
  }));
}

// ─── FONTES / UTM ─────────────────────────────────────────────────────────
function fetchBySources(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [
      { name: 'sessionSource' },
      { name: 'sessionMedium' },
      { name: 'sessionCampaignName' }
    ],
    metrics: [
      { name: 'sessions' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'bounceRate' },
      { name: 'sessionConversionRate' }
    ],
    dimensionFilter: {
      orGroup: {
        expressions: [
          { filter: { fieldName: 'sessionMedium', stringFilter: { matchType: 'EXACT', value: 'paid_social' } } },
          { filter: { fieldName: 'sessionMedium', stringFilter: { matchType: 'EXACT', value: 'cpc' } } },
          { filter: { fieldName: 'sessionSource', stringFilter: { matchType: 'CONTAINS', value: 'facebook' } } },
          { filter: { fieldName: 'sessionSource', stringFilter: { matchType: 'CONTAINS', value: 'instagram' } } }
        ]
      }
    },
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 20
  };

  try {
    const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);
    return (response.rows || []).map(row => ({
      source:       row.dimensionValues[0].value,
      medium:       row.dimensionValues[1].value,
      campaign:     row.dimensionValues[2].value,
      sessions:     safeInt(row.metricValues[0]),
      transactions: safeInt(row.metricValues[1]),
      revenue:      safeFloat(row.metricValues[2]),
      bounceRate:   safeFloat(row.metricValues[3]),
      convRate:     safeFloat(row.metricValues[4])
    }));
  } catch(e) {
    // Se não tiver UTMs configurados, retorna vazio com aviso
    return [{ _warning: 'Sem dados UTM — configure utm_source/utm_medium nos anúncios', sessions: 0 }];
  }
}

// ─── FUNIL DE E-COMMERCE ───────────────────────────────────────────────────
function fetchFunnel(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    metrics: [
      { name: 'itemsViewed' },
      { name: 'addToCarts' },
      { name: 'checkouts' },
      { name: 'ecommercePurchases' },
      { name: 'purchaseRevenue' },
      { name: 'itemsPurchased' }
    ]
  };

  try {
    const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);
    const row = response.rows?.[0]?.metricValues || [];
    const itemsViewed = safeInt(row[0]);
    const addToCarts  = safeInt(row[1]);
    const checkouts   = safeInt(row[2]);
    const purchases   = safeInt(row[3]);
    const revenue     = safeFloat(row[4]);
    const itemsBought = safeInt(row[5]);

    return {
      itemsViewed,
      addToCarts,
      checkouts,
      purchases,
      revenue,
      itemsPurchased: itemsBought,
      rates: {
        viewToCart:     itemsViewed > 0 ? (addToCarts / itemsViewed)  : 0,
        cartToCheckout: addToCarts > 0  ? (checkouts  / addToCarts)   : 0,
        checkoutToPay:  checkouts > 0   ? (purchases  / checkouts)    : 0,
        overall:        itemsViewed > 0 ? (purchases  / itemsViewed)  : 0
      }
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ─── DISPOSITIVOS / TECNOLOGIA ─────────────────────────────────────────────
function fetchByDevice(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [{ name: 'deviceCategory' }],
    metrics: [
      { name: 'sessions' },
      { name: 'activeUsers' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'bounceRate' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }]
  };
  const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);
  return (response.rows || []).map(row => ({
    device:       row.dimensionValues[0].value,
    sessions:     safeInt(row.metricValues[0]),
    users:        safeInt(row.metricValues[1]),
    transactions: safeInt(row.metricValues[2]),
    revenue:      safeFloat(row.metricValues[3]),
    bounceRate:   safeFloat(row.metricValues[4])
  }));
}

// ─── GEOGRÁFICO (estados + cidades) ───────────────────────────────────────
function fetchByGeo(range) {
  const stateReq = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [{ name: 'region' }],
    metrics: [
      { name: 'sessions' },
      { name: 'activeUsers' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 10
  };
  const cityReq = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [{ name: 'city' }],
    metrics: [
      { name: 'sessions' },
      { name: 'activeUsers' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 10
  };
  const stateRes = AnalyticsData.Properties.runReport(stateReq, `properties/${GA4_PROPERTY_ID}`);
  const cityRes  = AnalyticsData.Properties.runReport(cityReq,  `properties/${GA4_PROPERTY_ID}`);
  return {
    states: (stateRes.rows || []).map(row => ({
      state:        row.dimensionValues[0].value,
      sessions:     safeInt(row.metricValues[0]),
      users:        safeInt(row.metricValues[1]),
      transactions: safeInt(row.metricValues[2]),
      revenue:      safeFloat(row.metricValues[3])
    })),
    cities: (cityRes.rows || []).map(row => ({
      city:         row.dimensionValues[0].value,
      sessions:     safeInt(row.metricValues[0]),
      users:        safeInt(row.metricValues[1]),
      transactions: safeInt(row.metricValues[2]),
      revenue:      safeFloat(row.metricValues[3])
    }))
  };
}

// ─── DEMOGRÁFICO (idade + gênero — requer Google Signals) ──────────────────
function fetchDemographics(range) {
  try {
    const ageReq = {
      property: `properties/${GA4_PROPERTY_ID}`,
      dateRanges: [range],
      dimensions: [{ name: 'userAgeBracket' }],
      metrics: [{ name: 'activeUsers' }, { name: 'sessions' }],
      orderBys: [{ metric: { metricName: 'activeUsers' }, desc: true }]
    };
    const genderReq = {
      property: `properties/${GA4_PROPERTY_ID}`,
      dateRanges: [range],
      dimensions: [{ name: 'userGender' }],
      metrics: [{ name: 'activeUsers' }, { name: 'sessions' }]
    };
    const ageRes    = AnalyticsData.Properties.runReport(ageReq,    `properties/${GA4_PROPERTY_ID}`);
    const genderRes = AnalyticsData.Properties.runReport(genderReq, `properties/${GA4_PROPERTY_ID}`);
    return {
      age: (ageRes.rows || []).map(row => ({
        bracket:  row.dimensionValues[0].value,
        users:    safeInt(row.metricValues[0]),
        sessions: safeInt(row.metricValues[1])
      })),
      gender: (genderRes.rows || []).map(row => ({
        gender:   row.dimensionValues[0].value,
        users:    safeInt(row.metricValues[0]),
        sessions: safeInt(row.metricValues[1])
      }))
    };
  } catch(e) {
    return { error: e.message, age: [], gender: [] };
  }
}

// ─── TODAS AS ORIGENS / MÍDIAS ─────────────────────────────────────────────
function fetchAllSources(range) {
  const request = {
    property: `properties/${GA4_PROPERTY_ID}`,
    dateRanges: [range],
    dimensions: [
      { name: 'sessionSource' },
      { name: 'sessionMedium' }
    ],
    metrics: [
      { name: 'sessions' },
      { name: 'transactions' },
      { name: 'purchaseRevenue' },
      { name: 'sessionConversionRate' },
      { name: 'bounceRate' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 10
  };
  try {
    const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);
    return (response.rows || []).map(row => ({
      source:       row.dimensionValues[0].value,
      medium:       row.dimensionValues[1].value,
      sessions:     safeInt(row.metricValues[0]),
      transactions: safeInt(row.metricValues[1]),
      revenue:      safeFloat(row.metricValues[2]),
      convRate:     safeFloat(row.metricValues[3]),
      bounceRate:   safeFloat(row.metricValues[4])
    }));
  } catch(e) {
    return [];
  }
}

// ─── FUNIL POR PRODUTO ─────────────────────────────────────────────────────
function fetchProductFunnel(range) {
  try {
    const request = {
      property: `properties/${GA4_PROPERTY_ID}`,
      dateRanges: [range],
      dimensions: [{ name: 'itemName' }],
      metrics: [
        { name: 'itemsViewed' },
        { name: 'addToCarts' },
        { name: 'checkouts' },
        { name: 'ecommercePurchases' },
        { name: 'purchaseRevenue' }
      ],
      orderBys: [{ metric: { metricName: 'itemsViewed' }, desc: true }],
      limit: 20
    };
    const response = AnalyticsData.Properties.runReport(request, `properties/${GA4_PROPERTY_ID}`);
    return (response.rows || [])
      .filter(row => row.dimensionValues[0].value !== '(not set)')
      .map(row => ({
        product:     row.dimensionValues[0].value,
        viewed:      safeInt(row.metricValues[0]),
        addedToCart: safeInt(row.metricValues[1]),
        checkouts:   safeInt(row.metricValues[2]),
        purchases:   safeInt(row.metricValues[3]),
        revenue:     safeFloat(row.metricValues[4])
      }));
  } catch(e) {
    return [];
  }
}

// ─── HELPERS ───────────────────────────────────────────────────────────────
function getDateRange(period, startDate, endDate) {
  // Custom date range from dashboard (YYYY-MM-DD)
  if (startDate && endDate) {
    return { startDate: startDate, endDate: endDate };
  }
  const map = {
    today:      { startDate: 'today',       endDate: 'today' },
    yesterday:  { startDate: 'yesterday',   endDate: 'yesterday' },
    last_7d:    { startDate: '7daysAgo',    endDate: 'yesterday' },
    last_14d:   { startDate: '14daysAgo',   endDate: 'yesterday' },
    last_30d:   { startDate: '30daysAgo',   endDate: 'yesterday' },
    this_month: { startDate: 'firstDayOfMonth', endDate: 'today' },
    last_month: { startDate: '32daysAgo',   endDate: 'firstDayOfMonth' }
  };
  return map[period] || map['last_7d'];
}

function safeInt(val)   { return val?.value ? parseInt(val.value)   || 0 : 0; }
function safeFloat(val) { return val?.value ? parseFloat(val.value) || 0 : 0; }

// ─── TESTE LOCAL ───────────────────────────────────────────────────────────
// Para testar sem deploy, rode esta função no editor do Apps Script:
function testRun() {
  const result = fetchAllData('last_7d', 'overview');
  Logger.log(JSON.stringify(result, null, 2));
}

// ═══════════════════════════════════════════════════════════════════════════
// SLACK — RESUMO DIÁRIO DE CAMPANHAS
// ═══════════════════════════════════════════════════════════════════════════
//
// CONFIGURAÇÃO:
// 1. Preencha SLACK_WEBHOOK_URL com a URL do Incoming Webhook do canal #patrocinados
//    Como criar: slack.com/apps → "Incoming WebHooks" → Add → selecione #patrocinados
// 2. Preencha META_TOKEN_SLACK com o mesmo token usado no dashboard
// 3. Para envio automático diário: Gatilhos (ícone relógio) → + Adicionar gatilho
//    → sendDailyReport → Acionado por tempo → Temporizador diário → 08:00–09:00
// ═══════════════════════════════════════════════════════════════════════════
const SLACK_WEBHOOK_URL   = 'SUA_SLACK_WEBHOOK_URL'; // cole aqui a URL do Incoming Webhook do Slack
const META_TOKEN_SLACK    = 'EAAVdI3PzZADYBRHGsqLGaN2hlE9ZCm6WhcLqInkq6Igg23F6Gu4ZAnuC5OvKYeoUflfR42HoGgUzcNEZBRG0ZA8on59HwkEVhk1vQvu1CMYhOODjqRZCZAJ99eROWZAMZC63I5eLgtgZCOPytJVx5obijlagUOPAWQsHVnFxnrLnYFZBimCVPmrVpYEBJys9wfD';
const META_ACCOUNT_SLACK  = 'act_282317327';
const GRAPH_BASE          = 'https://graph.facebook.com/v19.0';
const ROAS_META           = 5.0; // Meta de ROAS

function sendDailyReport() {
  try {
    const token   = META_TOKEN_SLACK;
    const account = META_ACCOUNT_SLACK;

    // 1. Métricas gerais da conta (ontem)
    const acctResp = UrlFetchApp.fetch(
      `${GRAPH_BASE}/${account}/insights?fields=impressions,clicks,spend,reach,ctr,cpc,cpm,frequency,actions,purchase_roas,action_values&date_preset=yesterday&access_token=${token}`
    );
    const acctJson = JSON.parse(acctResp.getContentText());
    const d = acctJson.data?.[0];

    // 2. Busca campanhas ativas
    const campResp = UrlFetchApp.fetch(
      `${GRAPH_BASE}/${account}/campaigns?fields=id,name,status&effective_status=%5B%22ACTIVE%22%5D&limit=50&access_token=${token}`
    );
    const camps = JSON.parse(campResp.getContentText()).data || [];

    // 3. Sessões GA4 por nome de campanha para calcular CPS
    const ga4Sessions = fetchGA4SessionsByCampaign();

    // 4. Insights por campanha + anúncios melhor/pior
    const campInsights = [];
    for (const c of camps) {
      try {
        // Métricas da campanha
        const r = UrlFetchApp.fetch(
          `${GRAPH_BASE}/${c.id}/insights?fields=campaign_name,spend,actions,purchase_roas,ctr,cpm,frequency,action_values&date_preset=yesterday&access_token=${token}`
        );
        const ci = JSON.parse(r.getContentText());
        if (!ci.data?.[0]) continue;

        // Anúncios desta campanha (nível ad)
        const adResp = UrlFetchApp.fetch(
          `${GRAPH_BASE}/${c.id}/insights?fields=ad_name,spend,actions,purchase_roas,ctr&date_preset=yesterday&level=ad&limit=20&access_token=${token}`
        );
        const adRows = (JSON.parse(adResp.getContentText()).data || [])
          .map(ad => ({
            name:      (ad.ad_name || 'Anúncio').slice(0, 40),
            spend:     parseFloat(ad.spend || 0),
            roas:      ad.purchase_roas ? parseFloat(ad.purchase_roas[0]?.value || 0) : 0,
            purchases: slackAction(ad.actions, 'purchase'),
            ctr:       parseFloat(ad.ctr || 0)
          }))
          .filter(ad => ad.spend > 0);

        const adsWithRoas = adRows.filter(a => a.roas > 0);
        const bestAd  = adsWithRoas.length > 0
          ? adsWithRoas.reduce((a, b) => a.roas >= b.roas ? a : b)
          : null;
        const worstAd = adsWithRoas.length > 1
          ? adsWithRoas.reduce((a, b) => a.roas <= b.roas ? a : b)
          : null;

        const campName = ci.data[0].campaign_name || c.name;
        campInsights.push({
          id:          c.id,
          name:        c.name,
          ...ci.data[0],
          bestAd,
          worstAd:     worstAd?.name !== bestAd?.name ? worstAd : null,
          ga4Sessions: ga4Sessions[campName] || ga4Sessions[c.name] || 0
        });
      } catch(e) { /* campanha sem dados ontem */ }
    }

    // Ordena por ROAS desc
    campInsights.sort((a, b) => {
      const ra = a.purchase_roas ? parseFloat(a.purchase_roas[0]?.value || 0) : 0;
      const rb = b.purchase_roas ? parseFloat(b.purchase_roas[0]?.value || 0) : 0;
      return rb - ra;
    });

    // Métricas gerais da conta
    const spend     = parseFloat(d?.spend || 0);
    const purchases = slackAction(d?.actions, 'purchase');
    const purchaseVal = slackActionValue(d?.action_values, 'purchase');
    const roas      = d?.purchase_roas ? parseFloat(d.purchase_roas[0]?.value || 0) : 0;
    const ctr       = parseFloat(d?.ctr || 0);
    const cpm       = parseFloat(d?.cpm || 0);
    const cpa       = purchases > 0 ? spend / purchases : 0;
    const roasEmoji = roas >= ROAS_META ? '🟢' : roas >= 3 ? '🟡' : roas > 0 ? '🔴' : '⚫';

    const today = new Date().toLocaleDateString('pt-BR');

    // ── Monta blocos Slack ──────────────────────────────────────────────────
    const blocks = [
      {
        type: 'header',
        text: { type: 'plain_text', text: `📊 Meta Ads — Resumo de ${today}` }
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*💰 Investimento*\nR$ ${spend.toFixed(2)}` },
          { type: 'mrkdwn', text: `*🛒 Compras*\n${purchases}` },
          { type: 'mrkdwn', text: `*${roasEmoji} ROAS*\n${roas.toFixed(2)}x  _(meta: ${ROAS_META}x)_` },
          { type: 'mrkdwn', text: `*👆 CTR*\n${ctr.toFixed(2)}%` },
          { type: 'mrkdwn', text: `*📺 CPM*\nR$ ${cpm.toFixed(2)}` },
          { type: 'mrkdwn', text: `*🎯 CPA*\n${cpa > 0 ? 'R$ ' + cpa.toFixed(2) : '—'}` }
        ]
      },
      { type: 'divider' },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: `*📌 Detalhes por Campanha (ranking ROAS):*` }
      }
    ];

    // Um bloco por campanha
    for (const c of campInsights) {
      const cr   = c.purchase_roas ? parseFloat(c.purchase_roas[0]?.value || 0) : 0;
      const cs   = parseFloat(c.spend || 0);
      const cp   = slackAction(c.actions, 'purchase');
      const cv   = slackActionValue(c.action_values, 'purchase');
      const cctr = parseFloat(c.ctr || 0);
      const ccpm = parseFloat(c.cpm || 0);
      const ccpa = cp > 0 ? cs / cp : 0;
      const sess = c.ga4Sessions;
      const ccps = sess > 0 ? cs / sess : 0;
      const em   = cr >= ROAS_META ? '🟢' : cr >= 3 ? '🟡' : cr > 0 ? '🔴' : '⚫';
      const freq = parseFloat(c.frequency || 0);

      const analysis = getCampaignAnalysis(cr, cctr, ccpa, ccpm, freq, cs);

      let text = `${em} *${c.name.slice(0, 45)}*\n`;
      text += `   CTR *${cctr.toFixed(2)}%*  ·  ROAS *${cr.toFixed(1)}x*  ·  CPM *R$ ${ccpm.toFixed(2)}*\n`;
      text += `   CPA *${ccpa > 0 ? 'R$ ' + ccpa.toFixed(2) : '—'}*  ·  CPS *${ccps > 0 ? 'R$ ' + ccps.toFixed(2) : '—'}*  ·  Compras *R$ ${cv.toFixed(2)}*\n`;

      if (c.bestAd) {
        text += `   🥇 *Melhor:* ${c.bestAd.name} _(ROAS ${c.bestAd.roas.toFixed(1)}x · CTR ${c.bestAd.ctr.toFixed(2)}%)_\n`;
      }
      if (c.worstAd) {
        text += `   📉 *Pior:* ${c.worstAd.name} _(ROAS ${c.worstAd.roas.toFixed(1)}x · CTR ${c.worstAd.ctr.toFixed(2)}%)_\n`;
      }

      text += `   💡 _${analysis}_`;

      blocks.push({
        type: 'section',
        text: { type: 'mrkdwn', text }
      });
    }

    if (campInsights.length === 0) {
      blocks.push({
        type: 'section',
        text: { type: 'mrkdwn', text: '_Nenhuma campanha com dados ontem._' }
      });
    }

    // Observações gerais
    blocks.push({ type: 'divider' });

    const obs = [];
    if (roas >= ROAS_META)        obs.push(`🚀 ROAS geral atingiu a meta de ${ROAS_META}x — considere escalar orçamento.`);
    if (roas < 1.5 && spend > 0)  obs.push('🔴 ROAS abaixo do breakeven — revisar campanhas hoje.');
    const highFreq = campInsights.filter(c => parseFloat(c.frequency || 0) >= 4);
    if (highFreq.length > 0)      obs.push(`⚠️ ${highFreq.length} campanha(s) com frequência ≥4x — renovar criativos.`);
    const belowMeta = campInsights.filter(c => {
      const r = c.purchase_roas ? parseFloat(c.purchase_roas[0]?.value || 0) : 0;
      return r > 0 && r < ROAS_META;
    });
    if (belowMeta.length > 0)     obs.push(`📊 ${belowMeta.length} campanha(s) abaixo da meta de ROAS ${ROAS_META}x — revisar.`);
    if (obs.length === 0)         obs.push('✅ Tudo dentro do esperado. Monitore ao longo do dia.');

    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: `*Observações:*\n${obs.map(o => `• ${o}`).join('\n')}` }
    });
    blocks.push({
      type: 'context',
      elements: [{ type: 'mrkdwn', text: `_Gerado automaticamente · Dashboard BI Naked SW · Ananda Fróes_` }]
    });

    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ blocks })
    });

    Logger.log('✅ Relatório Slack enviado com sucesso!');
  } catch(e) {
    Logger.log('❌ Erro ao enviar Slack: ' + e.message);
  }
}

// Busca sessões GA4 por nome de campanha (ontem) — usado para calcular CPS
function fetchGA4SessionsByCampaign() {
  try {
    const req = {
      property: `properties/${GA4_PROPERTY_ID}`,
      dateRanges: [{ startDate: 'yesterday', endDate: 'yesterday' }],
      dimensions: [{ name: 'sessionCampaignName' }],
      metrics: [{ name: 'sessions' }],
      orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
      limit: 50
    };
    const response = AnalyticsData.Properties.runReport(req, `properties/${GA4_PROPERTY_ID}`);
    const map = {};
    (response.rows || []).forEach(row => {
      const name = row.dimensionValues[0].value;
      map[name] = safeInt(row.metricValues[0]);
    });
    return map;
  } catch(e) {
    Logger.log('Aviso GA4 sessões por campanha: ' + e.message);
    return {};
  }
}

// Análise e sugestão de melhoria por campanha
function getCampaignAnalysis(roas, ctr, cpa, cpm, freq, spend) {
  if (freq >= 5)                          return `Frequência muito alta (${freq.toFixed(1)}x) — público esgotado, troque criativos urgente`;
  if (roas >= ROAS_META && ctr >= 1.8)    return `ROAS e CTR excelentes — escale o orçamento em 20-30% para maximizar resultado`;
  if (roas >= ROAS_META)                  return `ROAS acima da meta — monitore CTR (${ctr.toFixed(2)}%) e amplie públicos semelhantes`;
  if (roas >= 4 && ctr >= 1.2)            return `ROAS próximo da meta — teste novos criativos para elevar CTR e alcançar ${ROAS_META}x`;
  if (roas >= 3 && ctr < 1.0)             return `CTR baixo (${ctr.toFixed(2)}%) — renove imagens/vídeos; criativos estão cansados`;
  if (roas >= 3)                           return `ROAS médio — revise segmentação de público e teste novos ângulos de oferta`;
  if (freq >= 4)                           return `Frequência alta (${freq.toFixed(1)}x) — troca de criativos necessária para recuperar performance`;
  if (cpm > 80)                            return `CPM elevado (R$${cpm.toFixed(2)}) — revise segmentação ou experimente públicos mais amplos`;
  if (roas >= 1.5)                         return `ROAS abaixo do breakeven — revise landing page, oferta e criativos com urgência`;
  if (roas > 0)                            return `ROAS crítico — considere pausar e reestruturar campanha do zero`;
  return `Sem conversões registradas — verifique pixel, rastreamento e página de destino`;
}

// Helper: extrai quantidade de uma action
function slackAction(actions, type) {
  if (!actions) return 0;
  const a = actions.find(x => x.action_type === type);
  return a ? parseInt(a.value) : 0;
}

// Helper: extrai valor monetário de uma action_value
function slackActionValue(actionValues, type) {
  if (!actionValues) return 0;
  const a = actionValues.find(x => x.action_type === type);
  return a ? parseFloat(a.value) : 0;
}

// Função para testar o envio manualmente no editor:
function testSlack() {
  sendDailyReport();
}

// ─── AGENDAMENTO AUTOMÁTICO ────────────────────────────────────────────────
// Execute esta função UMA VEZ no editor para criar o gatilho diário às 8h.
// Após rodar, o sendDailyReport() disparará automaticamente todo dia — mesmo
// com o computador desligado, pois roda nos servidores do Google.
function criarGatilhoDiario() {
  // Remove gatilhos antigos do sendDailyReport para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Cria novo gatilho: todo dia entre 08:00 e 09:00
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('✅ Gatilho criado! sendDailyReport vai rodar todo dia entre 08:00 e 09:00.');
}

// Para cancelar o envio automático:
function removerGatilho() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑 Gatilho removido.');
    }
  });
}
