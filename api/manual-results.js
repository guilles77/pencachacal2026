const KEY = process.env.MANUAL_RESULTS_KV_KEY || 'pencachacal2026:manual-results';
const FALLBACK_ADMIN_PIN = 'PencaChacal2026!';

function readBody(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', chunk => { raw += chunk; });
    req.on('end', () => {
      if (!raw) return resolve({});
      try { resolve(JSON.parse(raw)); } catch (err) { reject(new Error('JSON invalido')); }
    });
    req.on('error', reject);
  });
}

function storageConfig() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL || '';
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN || '';
  return { url, token, configured: !!(url && token) };
}

async function kvCommand(command) {
  const cfg = storageConfig();
  if (!cfg.configured) {
    const err = new Error('Falta configurar KV_REST_API_URL/KV_REST_API_TOKEN en Vercel.');
    err.code = 'NO_STORAGE';
    throw err;
  }
  const r = await fetch(cfg.url, {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + cfg.token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(command)
  });
  const text = await r.text();
  let data = {};
  try { data = text ? JSON.parse(text) : {}; } catch (e) { data = { raw: text }; }
  if (!r.ok || data.error) throw new Error(data.error || ('KV HTTP ' + r.status));
  return data.result;
}

function hasScore(score) {
  return score && score.homeGoals !== '' && score.awayGoals !== '' && score.homeGoals != null && score.awayGoals != null
    && Number.isFinite(Number(score.homeGoals)) && Number.isFinite(Number(score.awayGoals));
}

function cleanActual(input) {
  const src = input && typeof input === 'object' ? input : {};
  const out = { groupScores: {}, knockoutScores: {}, knockoutWinners: {}, r32Overrides: {}, awards: {}, matchMeta: {} };
  Object.keys(src.groupScores || {}).forEach(id => {
    const s = src.groupScores[id] || {};
    if (hasScore(s)) out.groupScores[id] = { homeGoals: Number(s.homeGoals), awayGoals: Number(s.awayGoals) };
  });
  Object.keys(src.knockoutScores || {}).forEach(id => {
    const s = src.knockoutScores[id] || {};
    if (hasScore(s)) out.knockoutScores[id] = { homeGoals: Number(s.homeGoals), awayGoals: Number(s.awayGoals) };
  });
  Object.keys(src.knockoutWinners || {}).forEach(id => {
    const winner = String(src.knockoutWinners[id] || '').trim();
    if (winner) out.knockoutWinners[id] = winner;
  });
  Object.keys(src.r32Overrides || {}).forEach(id => {
    const over = src.r32Overrides[id] || {};
    const home = String(over.home || '').trim();
    const away = String(over.away || '').trim();
    if (home || away) out.r32Overrides[id] = { home, away };
  });
  Object.keys(src.awards || {}).forEach(key => {
    const value = String(src.awards[key] || '').trim();
    if (value) out.awards[key] = value;
  });
  Object.keys(src.matchMeta || {}).forEach(id => {
    const meta = src.matchMeta[id] || {};
    const duration = String(meta.duration || meta.decision || '').trim();
    const note = String(meta.note || '').trim();
    if (duration || note) out.matchMeta[id] = { duration, note, source: 'manual' };
  });
  return out;
}

function checkPin(req, body) {
  const expected = String(process.env.ADMIN_PIN || FALLBACK_ADMIN_PIN).trim();
  const given = String((req.headers && (req.headers['x-admin-pin'] || req.headers['X-Admin-Pin'])) || body.pin || '').trim();
  if (!given || given !== expected) {
    const err = new Error('PIN de admin invalido.');
    err.status = 401;
    throw err;
  }
}

module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  try {
    const cfg = storageConfig();
    if (req.method === 'GET') {
      let stored = null;
      if (cfg.configured) {
        const raw = await kvCommand(['GET', KEY]);
        stored = raw ? JSON.parse(raw) : null;
      }
      res.status(200).json({ ok: true, configured: cfg.configured, updatedAt: stored && stored.updatedAt, actual: stored && stored.actual ? stored.actual : {} });
      return;
    }
    if (req.method !== 'POST' && req.method !== 'DELETE') {
      res.status(405).json({ ok: false, error: 'Metodo no soportado' });
      return;
    }
    const body = await readBody(req);
    checkPin(req, body || {});
    if ((body && body.action) === 'auth') {
      res.status(200).json({ ok: true, configured: cfg.configured });
      return;
    }
    if (req.method === 'DELETE' || (body && body.action === 'clear')) {
      await kvCommand(['DEL', KEY]);
      res.status(200).json({ ok: true, cleared: true });
      return;
    }
    const actual = cleanActual((body && body.actual) || {});
    const stored = { actual, updatedAt: new Date().toISOString() };
    await kvCommand(['SET', KEY, JSON.stringify(stored)]);
    res.status(200).json({ ok: true, updatedAt: stored.updatedAt, actual });
  } catch (err) {
    res.status(err.status || (err.code === 'NO_STORAGE' ? 503 : 500)).json({ ok: false, error: String((err && err.message) || err) });
  }
};
