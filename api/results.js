// Vercel serverless function: trae resultados del Mundial desde football-data.org.
// El token va en la env var FOOTBALL_DATA_TOKEN (NUNCA en el cliente).
// Cachea en el CDN de Vercel para respetar el rate limit (10 req/min del plan free).
module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  const token = process.env.FOOTBALL_DATA_TOKEN || '';
  if (!token) {
    res.status(200).json({ ok: false, error: 'Falta FOOTBALL_DATA_TOKEN' });
    return;
  }
  try {
    const r = await fetch('https://api.football-data.org/v4/competitions/WC/matches', {
      headers: { 'X-Auth-Token': token, 'Accept': 'application/json' }
    });
    if (!r.ok) {
      res.status(200).json({ ok: false, error: 'football-data HTTP ' + r.status, available: r.headers.get('X-Requests-Available-Minute') });
      return;
    }
    const data = await r.json();
    const LIVE_STATUSES = ['IN_PLAY', 'PAUSED'];
    const INCLUDE_STATUSES = ['FINISHED', 'IN_PLAY', 'PAUSED'];
    const matches = (data.matches || [])
      .filter(m => m && INCLUDE_STATUSES.includes(m.status))
      .map(m => {
        const live = LIVE_STATUSES.includes(m.status);
        const ft = m.score && m.score.fullTime;
        const ht = m.score && m.score.halfTime;
        const scoreSource = (ft && ft.home != null) ? ft : ht;
        return {
          stage: m.stage,
          group: m.group,
          utcDate: m.utcDate,
          home: m.homeTeam && m.homeTeam.tla,
          away: m.awayTeam && m.awayTeam.tla,
          hg: scoreSource ? scoreSource.home : null,
          ag: scoreSource ? scoreSource.away : null,
          winner: m.score ? m.score.winner : null,
          status: live ? (m.status === 'PAUSED' ? 'halftime' : 'live') : 'finished'
        };
      });
    const hasLive = matches.some(m => m.status === 'live' || m.status === 'halftime');
    // Cache dinámico: 60s cuando hay partidos en vivo, 120s si no
    res.setHeader('Cache-Control', hasLive
      ? 's-maxage=60, stale-while-revalidate=120'
      : 's-maxage=120, stale-while-revalidate=600');
    res.status(200).json({
      ok: true,
      updatedAt: new Date().toISOString(),
      season: (data.competition && data.competition.name) || 'FIFA World Cup',
      played: (data.resultSet && data.resultSet.played) || matches.length,
      count: matches.length,
      hasLive,
      matches
    });
  } catch (e) {
    res.status(200).json({ ok: false, error: String((e && e.message) || e) });
  }
};
