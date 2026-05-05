/**
 * NKSW — Shopify OAuth Callback
 *
 * Endpoint temporário para capturar o access token do Shopify via OAuth.
 * Após obter o token, você pode copiá-lo e usar como SHOPIFY_ADMIN_TOKEN nas
 * variáveis de ambiente do Vercel. Depois disso, este endpoint não é mais necessário.
 *
 * Fluxo:
 *   1. Shopify redireciona para este endpoint com ?code=XXXX&shop=XXXX
 *   2. Este endpoint troca o code pelo access token
 *   3. Exibe o access token na tela para você copiar
 *
 * URL para instalar o app (substitua CLIENT_ID pelo seu):
 *   https://lofty-fy.myshopify.com/admin/oauth/authorize
 *     ?client_id=b9101169f3e4a3d91271a982e079179f
 *     &scope=read_products,read_inventory,read_orders
 *     &redirect_uri=https://nksw-api.vercel.app/api/auth/callback
 */

const CLIENT_ID     = process.env.SHOPIFY_CLIENT_ID;
const CLIENT_SECRET = process.env.SHOPIFY_CLIENT_SECRET;

export default async function handler(req, res) {
  const { code, shop, error } = req.query;

  if (error) {
    return res.status(400).send(`
      <h2>❌ Erro na autenticação</h2>
      <p>Shopify retornou erro: <strong>${error}</strong></p>
    `);
  }

  if (!code || !shop) {
    return res.status(400).send(`
      <h2>❌ Parâmetros ausentes</h2>
      <p>Este endpoint deve ser chamado pelo Shopify após a autorização OAuth.</p>
      <hr/>
      <h3>Para iniciar o fluxo OAuth, acesse esta URL:</h3>
      <a href="https://${process.env.SHOPIFY_STORE || 'lofty-fy.myshopify.com'}/admin/oauth/authorize?client_id=${CLIENT_ID}&scope=read_products,read_inventory,read_orders&redirect_uri=https://nksw-api.vercel.app/api/auth/callback">
        👉 Clique aqui para autorizar o app
      </a>
    `);
  }

  try {
    // Troca o authorization code pelo access token
    const tokenRes = await fetch(`https://${shop}/admin/oauth/access_token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        code,
      }),
    });

    if (!tokenRes.ok) {
      const errBody = await tokenRes.text();
      return res.status(500).send(`
        <h2>❌ Erro ao trocar o código pelo token</h2>
        <p>Status: ${tokenRes.status}</p>
        <pre>${errBody}</pre>
      `);
    }

    const data = await tokenRes.json();
    const accessToken = data.access_token;
    const scope = data.scope;

    // Exibe o token para copiar — após isso, coloque no Vercel como SHOPIFY_ADMIN_TOKEN
    return res.status(200).send(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>NKSW API — Token obtido com sucesso</title>
        <style>
          body { font-family: sans-serif; max-width: 700px; margin: 40px auto; padding: 0 20px; }
          .token { background: #f0f0f0; padding: 16px; border-radius: 8px; word-break: break-all; font-family: monospace; font-size: 14px; }
          .step { background: #e8f5e9; border-left: 4px solid #4caf50; padding: 12px 16px; margin: 16px 0; border-radius: 4px; }
          .warning { background: #fff3e0; border-left: 4px solid #ff9800; padding: 12px 16px; margin: 16px 0; border-radius: 4px; }
          button { background: #333; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; }
        </style>
      </head>
      <body>
        <h1>✅ Token obtido com sucesso!</h1>
        <p><strong>Loja:</strong> ${shop}</p>
        <p><strong>Escopos autorizados:</strong> ${scope}</p>

        <div class="warning">
          ⚠️ <strong>Atenção:</strong> Este token aparece apenas uma vez. Copie agora e salve em local seguro.
        </div>

        <h2>Seu SHOPIFY_ADMIN_TOKEN:</h2>
        <div class="token" id="token">${accessToken}</div>
        <br>
        <button onclick="navigator.clipboard.writeText('${accessToken}').then(() => alert('Copiado!'))">
          📋 Copiar token
        </button>

        <h2>Próximos passos:</h2>
        <div class="step">
          <strong>1.</strong> Acesse <a href="https://vercel.com" target="_blank">vercel.com</a> → seu projeto <strong>nksw-api</strong> → Settings → Environment Variables
        </div>
        <div class="step">
          <strong>2.</strong> Adicione a variável:<br>
          <code>SHOPIFY_ADMIN_TOKEN</code> = o token acima
        </div>
        <div class="step">
          <strong>3.</strong> Adicione também:<br>
          <code>SHOPIFY_STORE</code> = <strong>${shop}</strong>
        </div>
        <div class="step">
          <strong>4.</strong> Faça redeploy do projeto (ou aguarde o próximo deploy automático)
        </div>
        <div class="step">
          <strong>5.</strong> Teste: <a href="https://nksw-api.vercel.app/api/products" target="_blank">https://nksw-api.vercel.app/api/products</a>
        </div>
      </body>
      </html>
    `);
  } catch (err) {
    return res.status(500).send(`
      <h2>❌ Erro interno</h2>
      <pre>${err.message}</pre>
    `);
  }
}
