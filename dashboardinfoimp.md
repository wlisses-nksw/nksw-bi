# NKSW BI Dashboard — Informações Importantes
Última atualização: 2026-05-05
---

## 🌐 URLs de Produção
| Serviço | URL |
|---------|-----|
| Dashboard + API (principal) | https://nksw-api.vercel.app |
| Alias (redireciona para cima) | https://nksw-bi.vercel.app |

> Dashboard: `https://nksw-api.vercel.app/` → index.html  
> API: `https://nksw-api.vercel.app/api/shopify-bi`, `/api/shopify-sync`, etc.

---

## 📁 Repositório e Pasta Local
| Projeto | GitHub | Pasta Local |
|---------|--------|-------------|
| Dashboard + API (tudo) | github.com/wlisses-nksw/nksw-bi (branch `main`) | `C:\Users\Wlisses Vicente\OneDrive\Documentos\GitHub\nksw-api` |

---

## ☁️ Vercel
- **Conta:** `wlisses-4983` (Hobby)
- **Projeto:** `nksw-api` → https://nksw-api.vercel.app
- **Auto-deploy:** ativo — cada push no `main` deploya automaticamente

### Deploy (único fluxo)
```powershell
cd "C:\Users\Wlisses Vicente\OneDrive\Documentos\GitHub\nksw-api"
git add -A
git commit -m "descrição"
git pull origin main --rebase
git push origin main
```

> ⚠️ NÃO usar `vercel --prod` — usar sempre `git push`  
> ⚠️ Sempre `git pull --rebase` antes do push  
> ⚠️ Se aparecer erro `index.lock`: `Remove-Item .git\index.lock` e tente novamente

---

## 🔑 Credenciais e Tokens

### Shopify
- **Loja:** `lofty-fy.myshopify.com` / https://nakedsw.com.br
- **Admin Token:** `shpat_XXXXXXXX (ver Shopify Admin ? Apps ? API)`
- **Sync secret:** `SYNC_SECRET (ver Vercel env vars)`
- **Escopos:** `read_orders, read_products`

### Meta Ads (Facebook)
- **Token:** constante `META_TOKEN` no `index.html` (linha ~1466)
- **Mesmo token** também em `ga4_apps_script.js` → constante `META_TOKEN_SLACK`
- ⚠️ **Expira em ~60 dias.** Quando erro 400 "session invalidated":
  1. Novo token em: https://developers.facebook.com/tools/explorer/
  2. Atualizar `META_TOKEN` em `index.html`
  3. Atualizar `META_TOKEN_SLACK` em `ga4_apps_script.js`
  4. git commit + push

### Google Apps Script — Auth
- **URL:** `https://script.google.com/macros/s/AKfycbxMN8FUeTiwCVREZvO1GD4ruaYjsv8s2GwQRI1QV7tglZrj8Tb4-vjw7Vru7NlbYec1DA/exec`
- **Arquivo referência:** `auth-script.gs`
- Constante `DEFAULT_AUTH_URL` no `index.html`

### Google Apps Script — GA4 / Sheets
- **URL:** `https://script.google.com/macros/s/AKfycbzcRQgtMkSY2sDpO4I8cf60e9lmqII1AeqZqg1mJyc_AjWn1FBVOeIWhwHLCzhrplbe/exec`
- **Constante:** `GA4_URL` (index.html linha ~1220)
- **Arquivo referência:** `ga4_apps_script.js`
- Parâmetros: `period=custom&startDate=YYYY-MM-DD&endDate=YYYY-MM-DD`

### Vercel Blob
- **Nome:** `nksw-api-blob`
- Chave por mês: `shopify-orders-YYYY-MM.json`
- Env var no Vercel: `BLOB_READ_WRITE_TOKEN`

---

## 🏗️ Arquitetura de Dados

### Fontes por período
| Período | Fonte |
|---------|-------|
| Até abr/2026 | Google Sheets via GAS — arquivos estáticos em `/data/` |
| Mai/2026 em diante | Shopify API via `/api/shopify-bi` + Vercel Blob |

```javascript
const SHOPIFY_CUTOFF_MONTH = 4; // 0-indexed = maio
const SHOPIFY_CUTOFF_YEAR  = 2026;
// isShopifyPeriod() faz o roteamento automático
```

### Endpoints da API
| Endpoint | Função |
|----------|--------|
| `/api/shopify-sync?action=full&month=2026-05&secret=SYNC_SECRET (ver Vercel env vars)` | Baixa todos os pedidos do mês → Blob |
| `/api/shopify-sync?action=update&month=2026-05&secret=SYNC_SECRET (ver Vercel env vars)` | Atualiza pedidos novos/modificados |
| `/api/shopify-bi?section=all&month=2026-05` | Lê Blob → retorna JSON completo |

---

## 📦 Estrutura dos Arquivos

```
GitHub/nksw-api/  →  nksw-api.vercel.app
├── index.html                    ← Dashboard completo (HTML + JS + CSS inline)
├── vercel.json                   ← Config Vercel (CORS + cache headers)
├── package.json                  ← Deps (@vercel/blob)
├── dashboardinfoimp.md           ← Este arquivo
├── api/
│   ├── shopify-bi.js             ← BI principal (lê Blob → JSON dashboard)
│   ├── shopify-sync.js           ← Sync Shopify → Vercel Blob
│   ├── inventory.js              ← Estoque
│   ├── order.js                  ← Pedido individual
│   ├── products.js               ← Produtos
│   ├── products-ids.js           ← IDs de produtos
│   ├── knowledge.js              ← Base de conhecimento
│   └── auth/callback.js          ← OAuth Shopify (eventual)
├── data/                         ← JSONs históricos pré-mai/2026 (somente leitura)
│   ├── pedidos_YYYY-MM.json
│   ├── vendas_YYYY-MM.json
│   └── ...
├── auth-script.gs                ← Referência GAS (autenticação)
├── ga4_apps_script.js            ← Referência GAS (GA4 + relatório Slack)
├── gas_clientes.js               ← Referência GAS (clientes)
├── nksw_sheets_api.js            ← Referência GAS (Sheets API)
└── .github/workflows/
    └── update-data.yml           ← Só workflow_dispatch manual (sem cron)
```

---

## 🔐 Autenticação do Dashboard
- Sistema próprio via Google Apps Script (`auth-script.gs`)
- URL hardcoded em `DEFAULT_AUTH_URL` no `index.html`
- Sessão em `sessionStorage` (chave: `nksw_auth`)
- Perfis: `admin` (acesso total + editar metas + aba Usuários) / `viewer` (somente leitura)
- Login: email + senha hasheada SHA-256 client-side

---

## 🛒 Aba Pedidos (mai/2026)

### Cards de status (lado a lado, clicáveis)
| Card | Critério | Cor |
|------|---------|-----|
| Pagos | `financial_status === 'paid'` | Verde |
| Entregues | `fulfillment_status === 'fulfilled'` (dentro dos pagos) | Verde |
| Pag. Parcial | `financial_status === 'partially_paid'` | Azul |
| Pendentes | `pending` ou `authorized` | Amarelo |
| Cancelados | `cancelled_at` ou `voided/refunded` | Vermelho |

> Clicar em um card aplica o filtro correspondente automaticamente.

### Tabela de pedidos
Colunas: **# Pedido · Cliente · E-mail · Pagamento · Entrega · Método de Envio · Valor · Data**

### Filtros
- Dropdown **Pagamento** e dropdown **Entrega**
- Botão **✕ Limpar** para resetar filtros
- Contador de pedidos visíveis

### Download CSV
- Botão **⬇ Baixar CSV** — exporta os pedidos visíveis (respeita filtros)
- Nome do arquivo: `pedidos-YYYY-MM.csv`
- Encoding: UTF-8 com BOM (abre corretamente no Excel)

### D-1
- Dados sempre até ontem (server-side no Shopify + client-side no GAS)
- Sem botão "Atualizar Shopify"

---

## 🧠 Variáveis Globais Principais (index.html)
```javascript
_indRealizadoReceita    // receita total do mês
_indVendasPedidos       // total de pedidos
_indVendasTicket        // ticket médio
_indMktSpend            // gasto Meta Ads
_indMktCpp              // CAC
_indRoasAtual           // ROAS Meta Ads
_indGA4Sessions         // sessões GA4
_pedidosLista           // cache global da lista de pedidos (para filtros e CSV)
_metasDiarioCache       // { "YYYY_MM": [{dia, receita, pedidos, ticket, ...}] }
navMonth / navYear      // mês/ano navegado (navMonth = 0-indexed)
goalMonth / goalYear    // mês/ano da aba Metas
```

## 🔧 Funções-Chave (index.html)
| Função | O que faz |
|--------|-----------|
| `loadShopifyData()` | Carrega dados Shopify do mês |
| `isShopifyPeriod(m,y)` | Retorna true se mês >= mai/2026 |
| `applyPedidosData(d)` | Aplica dados de pedidos + popula `_pedidosLista` |
| `renderPedidosTable()` | Renderiza tabela respeitando filtros ativos |
| `filterPedidos(campo, valor)` | Aplica filtro e re-renderiza tabela |
| `downloadPedidosCSV()` | Gera e baixa CSV dos pedidos filtrados |
| `getMonthRealized(key,m,y)` | Retorna KPI realizado do mês (cache diário) |
| `renderGoals()` | Renderiza aba Metas |
| `isAdmin()` | Retorna true se usuário é admin |

---

## 🔀 Regras de Negócio
- Cupons com `"troca"` (case-insensitive) excluídos de desconto e % cupom
- Apenas `financial_status === 'paid'` nos KPIs de receita/ticket/pedidos
- Realizado sempre D-1 (dia anterior)
- Comparação YoY mai/2026: usa abr/2026 como base (sem dados de mai/2025)

---

## 🐛 Problemas Conhecidos e Soluções
| Problema | Causa | Solução |
|----------|-------|---------|
| Login não funciona (clique sem resposta) | Erro de sintaxe JS no arquivo | Verificar com `node --check` e corrigir; nunca usar caracteres especiais literais em strings JS |
| Login não funciona (`doLogin is not defined`) | `index.html` truncado | Verificar se termina com `</script></body></html>` |
| Push rejeitado (non-fast-forward) | GitHub Actions commitou dados automaticamente | `git pull origin main --rebase` antes do push |
| Erro `index.lock` no git | Processo do sandbox deixou lock | `Remove-Item .git\index.lock` |
| Token Meta expirado (erro 400) | Expira em ~60 dias | Novo token em developers.facebook.com → atualizar `index.html` e `ga4_apps_script.js` |
| Edições não aparecem no site | Cache do browser | `Ctrl+Shift+R` ou abrir em aba anônima |

---

## 📊 Fluxo de Dados — Mai/2026 (Shopify)
```
1. isShopifyPeriod() = true → loadShopifyData()
2. GET /api/shopify-bi?section=all&month=2026-05
3. applyVendasData() + applyPedidosData() + applyLogisticaData() + applyClientesData()
4. _populateShopifyDiarioCache() → cache diário D-1
5. _enrichShopifyCacheAsync() → enriquece com Meta Ads + GA4 em background
6. Aba Metas: getMonthRealized() / getDayRealized() leem do cache
```

## 📊 Fluxo de Dados — Até Abr/2026 (Sheets/estático)
```
1. isShopifyPeriod() = false → loadAllSheetsData()
2. fetchStaticJSON() → lê /data/SECAO_YYYY-MM.json
3. Fallback: fetchSheets() → GAS API
```
