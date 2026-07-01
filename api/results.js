// Vercel serverless function: trae resultados del Mundial desde football-data.org.
// El token va en la env var FOOTBALL_DATA_TOKEN (NUNCA en el cliente).
// Cachea en el CDN de Vercel para respetar el rate limit (10 req/min del plan free).
function hasScore(score) {
  return score && score.home != null && score.away != null;
}

function subtractScoreParts(base, parts) {
  if (!hasScore(base)) return null;
  let home = Number(base.home);
  let away = Number(base.away);
  for (const part of parts) {
    if (!hasScore(part)) continue;
    home -= Number(part.home);
    away -= Number(part.away);
  }
  if (!Number.isFinite(home) || !Number.isFinite(away) || home < 0 || away < 0) return null;
  return { home, away };
}

function regularScoreFromParts(score) {
  if (!score) return null;
  const duration = score.duration;
  if (duration === 'REGULAR') return hasScore(score.fullTime) ? score.fullTime : null;
  if (duration === 'EXTRA_TIME') return subtractScoreParts(score.fullTime, [score.extraTime]);
  if (duration === 'PENALTY_SHOOTOUT') return subtractScoreParts(score.fullTime, [score.extraTime, score.penalties]);
  return null;
}

function regularScoreFromGoals(match) {
  const goals = Array.isArray(match && match.goals) ? match.goals : [];
  let score = null;
  for (const goal of goals) {
    const minute = Number(goal && goal.minute);
    if (!Number.isFinite(minute) || minute > 90 || !hasScore(goal.score)) continue;
    score = { home: goal.score.home, away: goal.score.away };
  }
  return score;
}

function firstScore() {
  for (const score of arguments) {
    if (hasScore(score)) return score;
  }
  return null;
}

// Correcciones manuales a errores puntuales del feed oficial (football-data.org).
// El feed reportó Bélgica-Senegal (16avos) como 3-2 en tiempo reglamentario, pero
// el partido terminó 2-2 en los 90' y Bélgica avanzó en el alargue. Para la penca
// importa el marcador de los 90' + el equipo que pasa, así que forzamos el 2-2/ET.
// Clave: `${stage}|${homeTLA}|${awayTLA}`.
const FEED_CORRECTIONS = {
  'LAST_32|BEL|SEN': { hg: 2, ag: 2, duration: 'EXTRA_TIME' }
};

module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  const token = process.env.FOOTBALL_DATA_TOKEN || '';
  if (!token) {
    res.status(200).json({ ok: false, error: 'Falta FOOTBALL_DATA_TOKEN' });
    return;
  }
  try {
    const r = await fetch('https://api.football-data.org/v4/competitions/WC/matches', {
      headers: { 'X-Auth-Token': token, 'Accept': 'application/json', 'X-Unfold-Goals': 'true' }
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
        let duration = m.score ? m.score.duration : null;
        const rt = m.score && (m.score.regularTime || m.score.regulationTime || m.score.normalTime);
        const ft = m.score && m.score.fullTime;
        const ht = m.score && m.score.halfTime;
        const scoreSource = firstScore(rt, regularScoreFromGoals(m), regularScoreFromParts(m.score), duration === 'REGULAR' ? ft : null, ht);
        const homeTla = m.homeTeam && m.homeTeam.tla;
        const awayTla = m.awayTeam && m.awayTeam.tla;
        let hg = scoreSource ? scoreSource.home : null;
        let ag = scoreSource ? scoreSource.away : null;
        let winner = m.score ? m.score.winner : null;
        const corr = FEED_CORRECTIONS[m.stage + '|' + homeTla + '|' + awayTla];
        if (corr) {
          if (corr.hg != null) hg = corr.hg;
          if (corr.ag != null) ag = corr.ag;
          if (corr.duration) duration = corr.duration;
          if (corr.winner) winner = corr.winner;
        }
        return {
          stage: m.stage,
          group: m.group,
          utcDate: m.utcDate,
          duration,
          home: homeTla,
          away: awayTla,
          hg: hg,
          ag: ag,
          winner: winner,
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
