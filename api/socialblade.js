/**
 * NKSW — SocialBlade Proxy
 * Busca dados públicos do SocialBlade para @nakedswimwear
 * GET /api/socialblade
 */
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const url = 'https://socialblade.com/instagram/user/nakedswimwear';
    const r   = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
      }
    });

    if (!r.ok) throw new Error(`SocialBlade HTTP ${r.status}`);
    const html = await r.text();

    // Extrair dados com regex
    const extract = (pattern, def = null) => {
      const m = html.match(pattern);
      return m ? m[1].trim() : def;
    };

    const engRate   = extract(/Engagement Rate[\s\S]*?<\/span>([\d.]+%)/);
    const avgLikes  = extract(/Average Likes[\s\S]*?<\/span>([\d.,]+)/);
    const avgComm   = extract(/Average Comments[\s\S]*?<\/span>([\d.,]+)/);
    const grade     = extract(/Grade[\s\S]*?([A-F][+\-]?)\n/);
    const followers = extract(/Followers[\s\S]*?<\/span>([\d,]+)/);
    const sbRank    = extract(/SB Rank[\s\S]*?(\d[\d.,]+th)/);

    // Variação 14 dias — extrair da tabela
    const last14Match = html.match(/Last 14 Days[\s\S]*?(-?[\d,]+)\s+(-?[\d,]+)/);
    const fol14d = last14Match ? last14Match[1].replace(/,/g, '') : null;
    const med14d = last14Match ? last14Match[2] : null;

    // Daily average
    const dailyAvgMatch = html.match(/Daily Average[\s\S]*?(-?[\d,]+)/);
    const dailyAvg = dailyAvgMatch ? dailyAvgMatch[1].replace(/,/g, '') : null;

    return res.status(200).json({
      ok: true,
      source: 'socialblade',
      cached_at: new Date().toISOString(),
      data: {
        followers:      followers ? parseInt(followers.replace(/,/g, '')) : null,
        eng_rate:       engRate   || null,
        avg_likes:      avgLikes  ? parseFloat(avgLikes.replace(/,/g, '')) : null,
        avg_comments:   avgComm   ? parseFloat(avgComm.replace(/,/g, '')) : null,
        grade:          grade     || null,
        sb_rank:        sbRank    || null,
        fol_14d:        fol14d    ? parseInt(fol14d) : null,
        media_14d:      med14d    ? parseInt(med14d) : null,
        daily_avg_fol:  dailyAvg  ? parseInt(dailyAvg) : null,
      }
    });

  } catch(e) {
    // Retorna dados estáticos como fallback se SocialBlade bloquear
    return res.status(200).json({
      ok: true,
      source: 'static',
      cached_at: new Date().toISOString(),
      data: {
        followers: 259682,
        eng_rate: '0.01%',
        avg_likes: 23.19,
        avg_comments: 2.38,
        grade: 'B+',
        sb_rank: '43.296th',
        fol_14d: -1489,
        media_14d: 13,
        daily_avg_fol: -39,
      }
    });
  }
}
