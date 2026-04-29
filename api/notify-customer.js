// api/notify-customer.js
// Cria ou atualiza cliente no Shopify com tag avise-me:VARIANT_ID
// Token necessário: read_customers + write_customers (já configurado em SHOPIFY_TOKEN)
//
// Resultado no Admin:
//   Clientes → filtrar por tag "avise-me" ou "avise-me:VARIANT_ID"
//   Nota do cliente registra produto/variante/data
//   Quando reposto: filtrar tag → enviar email → remover tag

const STORE        = 'lofty-fy.myshopify.com';
const SHOPIFY_BASE = `https://${STORE}/admin/api/2024-01`;

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).json({ error: 'Method not allowed' });

  const { email, variant_id, product_title, variant_title, product_handle } = req.body || {};

  if (!email || !variant_id) {
    return res.status(400).json({ error: 'Campos obrigatórios: email, variant_id' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ error: 'Email inválido' });
  }

  const TOKEN = process.env.SHOPIFY_TOKEN;
  if (!TOKEN) return res.status(500).json({ error: 'SHOPIFY_TOKEN não configurado' });

  // Tags: avise-me (geral) + avise-me:VARIANT_ID (específico para filtro de reposição)
  const tagVariant  = `avise-me:${variant_id}`;
  const tagGeneral  = 'avise-me';
  const now         = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  const noteEntry   = `[AVISE-ME] ${product_title || ''}${variant_title ? ' · ' + variant_title : ''} | Variant ID: ${variant_id} | ${now}`;

  try {
    // ── 1. Busca cliente por email ───────────────────────────────────
    const searchRes = await fetch(
      `${SHOPIFY_BASE}/customers/search.json?query=email:${encodeURIComponent(email)}&limit=1&fields=id,email,tags,note`,
      { headers: { 'X-Shopify-Access-Token': TOKEN } }
    );
    const searchData = await searchRes.json();
    const existing   = searchData.customers?.[0];

    let customer;

    if (existing) {
      // ── 2a. Atualiza cliente existente ───────────────────────────
      const currentTags = (existing.tags || '')
        .split(',')
        .map(t => t.trim())
        .filter(Boolean);

      // Adiciona tags apenas se ainda não existirem
      if (!currentTags.includes(tagGeneral))  currentTags.push(tagGeneral);
      if (!currentTags.includes(tagVariant))  currentTags.push(tagVariant);

      // Acrescenta linha na nota (preserva histórico)
      const updatedNote = existing.note
        ? existing.note + '\n' + noteEntry
        : noteEntry;

      const updateRes = await fetch(
        `${SHOPIFY_BASE}/customers/${existing.id}.json`,
        {
          method:  'PUT',
          headers: { 'X-Shopify-Access-Token': TOKEN, 'Content-Type': 'application/json' },
          body:    JSON.stringify({
            customer: {
              id:   existing.id,
              tags: currentTags.join(', '),
              note: updatedNote
            }
          })
        }
      );
      customer = (await updateRes.json()).customer;

    } else {
      // ── 2b. Cria novo cliente ────────────────────────────────────
      const createRes = await fetch(
        `${SHOPIFY_BASE}/customers.json`,
        {
          method:  'POST',
          headers: { 'X-Shopify-Access-Token': TOKEN, 'Content-Type': 'application/json' },
          body:    JSON.stringify({
            customer: {
              email:                 email.toLowerCase().trim(),
              tags:                  `${tagGeneral}, ${tagVariant}`,
              note:                  noteEntry,
              accepts_marketing:     false,
              send_email_welcome:    false  // não dispara email de boas-vindas
            }
          })
        }
      );
      const createData = await createRes.json();

      if (createData.errors) {
        // Email pode já existir em edge-case de race-condition — trata como sucesso
        if (JSON.stringify(createData.errors).includes('taken')) {
          return res.status(200).json({ success: true, note: 'email already exists' });
        }
        throw new Error(JSON.stringify(createData.errors));
      }
      customer = createData.customer;
    }

    return res.status(200).json({
      success:     true,
      customer_id: customer.id,
      tags:        customer.tags
    });

  } catch (err) {
    console.error('[notify-customer]', err);
    return res.status(500).json({ error: 'Erro interno: ' + err.message });
  }
}
