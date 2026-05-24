/**
 * Cron: Shopify Sync diário — 03:00 BRT (06:00 UTC)
 *
 * Chamado automaticamente pelo Vercel Cron todos os dias às 06:00 UTC.
 * Aciona action=update para o mês corrente em horário de Brasília (UTC-3).
 * No dia 1 de cada mês, também sincroniza o mês anterior para fechamento.
 *
 * Configurado em vercel.json:
 *   { "path": "/api/cron/shopify-sync", "schedule": "0 6 * * *" }
 *
 * Vercel envia o header: Authorization: Bearer {CRON_SECRET}
 */

const BASE_URL = 'https://nksw-api.vercel.app';

function brtMonth(date) {
  // UTC-3 (Brasília)
  const brt = new Date(date.getTime() - 3 * 60 * 60 * 1000);
  const y   = brt.getUTCFullYear();
  const m   = String(brt.getUTCMonth() + 1).padStart(2, '0');
  const d   = brt.getUTCDate();
  return { month: `${y}-${m}`, day: d, year: y, monthNum: brt.getUTCMonth() + 1 };
}

export default async function handler(req, res) {
  /* ── Autorização: Vercel Cron envia Bearer {CRON_SECRET} ── */
  const cronSecret = process.env.CRON_SECRET;
  if (cronSecret) {
    const auth = req.headers['authorization'] ?? '';
    if (auth !== `Bearer ${cronSecret}`) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
  }

  const SECRET = process.env.ADMIN_SECRET || 'nksw2025';
  const now    = new Date();
  const { month, day, year, monthNum } = brtMonth(now);

  /* ── Meses a sincronizar ── */
  const months = [month];

  /* No dia 1 de cada mês, fecha também o mês anterior */
  if (day === 1) {
    const prevDate   = new Date(Date.UTC(year, monthNum - 2, 1)); // mês anterior
    const py = prevDate.getUTCFullYear();
    const pm = String(prevDate.getUTCMonth() + 1).padStart(2, '0');
    months.push(`${py}-${pm}`);
  }

  /* ── Executa sync para cada mês ── */
  const results = [];
  for (const m of months) {
    try {
      const url = `${BASE_URL}/api/shopify-sync?action=update&month=${m}&secret=${SECRET}`;
      const r   = await fetch(url, { headers: { 'User-Agent': 'nksw-cron/1.0' } });
      const data = await r.json().catch(() => ({}));
      results.push({ month: m, ok: r.ok, status: r.status, data });
    } catch (err) {
      results.push({ month: m, ok: false, error: err.message });
    }
  }

  const allOk = results.every(r => r.ok);

  return res.status(allOk ? 200 : 207).json({
    ok:        allOk,
    synced:    months,
    timestamp: now.toISOString(),
    brtMonth:  month,
    brtDay:    day,
    results,
  });
}
