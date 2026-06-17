// Vercel serverless function: trae resultados del Mundial desde football-data.org.
// El token va en la env var FOOTBALL_DATA_TOKEN (NUNCA en el cliente).
// Cachea en el CDN de Vercel para respetar el rate limit (10 req/min del plan free).
module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  // Cache en el edge: 1 llamada real ~cada 2 min sin importar cuántos usuarios entren.
  res.setHeader('Cache-Control', 's-maxage=120, stale-while-revalidate=600');

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
    const matches = (data.matches || [])
      .filter(m => m && m.status === 'FINISHED')
      .map(m => ({
        stage: m.stage,
        group: m.group,
        utcDate: m.utcDate,
        home: m.homeTeam && m.homeTeam.tla,
        away: m.awayTeam && m.awayTeam.tla,
        hg: m.score && m.score.fullTime ? m.score.fullTime.home : null,
        ag: m.score && m.score.fullTime ? m.score.fullTime.away : null,
        winner: m.score ? m.score.winner : null
      }));
    res.status(200).json({
      ok: true,
      updatedAt: new Date().toISOString(),
      season: (data.competition && data.competition.name) || 'FIFA World Cup',
      played: (data.resultSet && data.resultSet.played) || matches.length,
      count: matches.length,
      matches
    });
  } catch (e) {
    res.status(200).json({ ok: false, error: String((e && e.message) || e) });
  }
};
