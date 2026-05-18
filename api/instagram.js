/**
 * NKSW — Instagram Insights API
 * GET /api/instagram?period=30  → insights dos últimos N dias
 * GET /api/instagram?period=7   → insights dos últimos 7 dias
 */

const IG_TOKEN = process.env.META_IG_TOKEN;
const IG_ID    = process.env.META_IG_ID || '36179889998275967';
const IG_BASE  = 'https://graph.instagram.com/v19.0';

async function igFetch(path) {
  const sep = path.includes('?') ? '&' : '?';
  const res = await fetch(`${IG_BASE}${path}${sep}access_token=${IG_TOKEN}`);
  if (!res.ok) throw new Error(`IG ${res.status}: ${await res.text()}`);
  return res.json();
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (!IG_TOKEN) return res.status(500).json({ ok: false, error: 'META_IG_TOKEN não configurado' });

  try {
    const days  = parseInt(req.query.period || '30');
    const until = Math.floor(Date.now() / 1000);
    const since = until - days * 86400;

    // Perfil básico
    const profile = await igFetch(`/${IG_ID}?fields=username,followers_count,media_count,biography,website,profile_picture_url`);

    // Insights diários — separar métricas por tipo de período
    const metricsDay  = 'reach,follower_count,profile_views,accounts_engaged,total_interactions,likes,comments,shares,saves,follows_and_unfollows';
    const metricsLife = 'views'; // views às vezes só funciona com lifetime
    const insights = await igFetch(`/${IG_ID}/insights?metric=${metricsDay}&period=day&since=${since}&until=${until}`);
    // Tentar views separadamente
    let viewsInsight = null;
    try {
      const vr = await igFetch(`/${IG_ID}/insights?metric=views&period=day&since=${since}&until=${until}`);
      if (vr.data && vr.data.length) viewsInsight = vr.data[0];
    } catch(_) {}

    // Últimas 12 mídias
    const media = await igFetch(`/${IG_ID}/media?fields=id,caption,media_type,timestamp,like_count,comments_count,thumbnail_url,media_url,permalink&limit=12`);

    // Processar insights em objeto por nome
    const insightMap = {};
    for (const item of (insights.data || [])) {
      insightMap[item.name] = item.values || [];
    }

    // Calcular totais e médias do período
    const sum   = (arr) => arr.reduce((s, v) => s + (v.value || 0), 0);
    const avg   = (arr) => arr.length ? Math.round(sum(arr) / arr.length) : 0;
    const last  = (arr) => arr.length ? arr[arr.length - 1].value : 0;
    const prev  = (arr, n) => arr.length >= n ? arr[arr.length - n].value : 0;

    const reach     = insightMap['reach']               || [];
    const followers = insightMap['follower_count']       || [];
    const engaged   = insightMap['accounts_engaged']     || [];
    const interacts = insightMap['total_interactions']   || [];
    const likes     = insightMap['likes']                || [];
    const comments  = insightMap['comments']             || [];
    const shares    = insightMap['shares']               || [];
    const saves     = insightMap['saves']                || [];
    const profViews = insightMap['profile_views']        || [];
    // Views do endpoint separado
    const views     = viewsInsight ? (viewsInsight.values || []) : [];

    const totalNewFollowers = sum(followers);
    const totalReach        = sum(reach);
    const avgDailyReach     = avg(reach);
    const avgEngaged        = avg(engaged);
    // Usar soma de likes+comments+shares+saves se total_interactions vier zerado
    const totalInteractsSum = sum(interacts) || (sum(likes) + sum(comments) + sum(shares) + sum(saves));
    const avgInteractions   = reach.length > 0 ? totalInteractsSum / reach.length : 0;
    const totalViews        = sum(views);
    const totalProfViews    = sum(profViews);
    const totalLikes        = sum(likes);
    const totalComments     = sum(comments);
    const totalShares       = sum(shares);
    const totalSaves        = sum(saves);

    // Taxa de engajamento — usa média de interações por alcance (mais preciso)
    const engRate = totalReach > 0
      ? ((totalInteractsSum / totalReach) * 100).toFixed(3)
      : (profile.followers_count > 0
          ? ((avgInteractions / profile.followers_count) * 100).toFixed(3)
          : '0.000');

    // Tendência: últimos 7 dias vs 7 anteriores
    const reach7  = reach.slice(-7);
    const reach7p = reach.slice(-14, -7);
    const reachTrend = reach7p.length
      ? Math.round(((sum(reach7) - sum(reach7p)) / Math.max(sum(reach7p), 1)) * 100)
      : 0;

    const fol7  = sum(followers.slice(-7));
    const fol7p = sum(followers.slice(-14, -7));
    const folTrend = fol7p > 0 ? Math.round(((fol7 - fol7p) / fol7p) * 100) : 0;

    return res.status(200).json({
      ok: true,
      period: days,
      profile: {
        username:       profile.username,
        followers:      profile.followers_count,
        media_count:    profile.media_count,
        biography:      profile.biography,
        website:        profile.website,
        picture:        profile.profile_picture_url,
      },
      kpis: {
        new_followers:      totalNewFollowers,
        total_reach:        totalReach,
        avg_daily_reach:    avgDailyReach,
        avg_engaged:        avgEngaged,
        eng_rate:           parseFloat(engRate),
        total_views:        totalViews,
        profile_views:      totalProfViews,
        reach_trend_pct:    reachTrend,
        follower_trend_pct: folTrend,
        total_likes:        totalLikes,
        total_comments:     totalComments,
        total_shares:       totalShares,
        total_saves:        totalSaves,
        total_interactions: totalInteractsSum,
      },
      daily: {
        reach:     reach.map(v => ({ date: v.end_time?.slice(0,10), value: v.value })),
        followers: followers.map(v => ({ date: v.end_time?.slice(0,10), value: v.value })),
        engaged:   engaged.map(v => ({ date: v.end_time?.slice(0,10), value: v.value })),
        views:     views.map(v => ({ date: v.end_time?.slice(0,10), value: v.value })),
      },
      media: (media.data || []).map(m => ({
        id:        m.id,
        type:      m.media_type,
        caption:   (m.caption || '').slice(0, 120),
        date:      m.timestamp?.slice(0,10),
        likes:     m.like_count || 0,
        comments:  m.comments_count || 0,
        url:       m.permalink,
        thumb:     m.thumbnail_url || m.media_url,
      })),
    });

  } catch (e) {
    console.error('[instagram]', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
