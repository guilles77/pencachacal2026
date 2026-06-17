/**
 * Penca Los Chacales - backend Google Apps Script - v19 Official Match Center + Live Data Ready
 *
 * Uso:
 * 1) Crear un Google Sheet.
 * 2) Extensions / Extensiones -> Apps Script.
 * 3) Pegar este archivo como Code.gs y crear index.html con el otro archivo.
 * 4) Ejecutar setup() una vez desde Apps Script y autorizar.
 * 5) Deploy -> New deployment -> Web app -> Execute as: Me; Access: Anyone with the link.
 */

const LC_SHEETS = {
  CONFIG: 'Config',
  PLAYERS: 'Players',
  PREDICTIONS: 'Predictions',
  ACTUAL: 'Actual',
  SQUADS: 'Squads',
  AUDIT: 'Audit'
};

const LC_HEADERS = {
  Config: ['Key', 'Value', 'Notes'],
  Players: ['PlayerId', 'Name', 'Token', 'SubmittedAt', 'CreatedAt', 'LastSavedAt', 'Email', 'PasswordHash', 'PasswordSalt', 'AuthUpdatedAt', 'TempPassword'],
  Predictions: ['PlayerId', 'DraftJson', 'SubmittedJson', 'UpdatedAt'],
  Actual: ['Key', 'Json', 'UpdatedAt'],
  Squads: ['Team', 'Player', 'Position', 'Number'],
  Audit: ['Timestamp', 'Action', 'PlayerId', 'Info']
};

const LC_APP_VERSION = 'v20.5-my-prediction-cleanup-2026-06';

const LC_DEFAULT_CONFIG = {
  ADMIN_PIN: 'cambiar-este-pin',
  PREDICTION_DEADLINE: '2026-06-12T00:00:00-03:00',
  REVEAL_PREDICTIONS: 'TRUE',
  OFFICIAL_RESULTS_URL: 'https://www.fifa.com/en/tournaments/mens/worldcup/canadamexicousa2026/scores-fixtures',
  POINT_RESULT_EXACT: '3',
  POINT_RESULT_OUTCOME: '1',
  POINT_GROUP_EXACT: '3',
  POINT_GROUP_OUTCOME: '1',
  POINT_R32: '1',
  POINT_R16: '2',
  POINT_QF: '4',
  POINT_SF: '6',
  POINT_THIRD_PLACE: '8',
  POINT_FINALIST: '10',
  POINT_CHAMPION: '20',
  POINT_AWARD: '5',
  ACTIVITY_MAX_ITEMS: '12',
  OFFICIAL_RESULTS_SYNC_URL: '',
  OFFICIAL_RESULTS_API_KEY: '',
  OFFICIAL_SQUADS_SYNC_URL: '',
  OFFICIAL_SQUADS_API_KEY: '',
  OFFICIAL_SYNC_AUTO: 'FALSE',
  OFFICIAL_SYNC_MINUTES: '15',
  OFFICIAL_LAST_SYNC_AT: '',
  OFFICIAL_LAST_SYNC_STATUS: '',
  OFFICIAL_SQUADS_LAST_SYNC_AT: '',
  OFFICIAL_SQUADS_LAST_SYNC_STATUS: ''
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Los Chacales')
    .addItem('Inicializar / reparar hojas', 'setup')
    .addItem('Abrir URL de la web app', 'showWebAppUrl')
    .addToUi();
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl() || 'Todavia no hay deployment. Usar Deploy > New deployment > Web app.';
  SpreadsheetApp.getUi().alert(url);
}

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const serviceUrl = ScriptApp.getService().getUrl() || '';
  if (String(params.asset || '') === 'manifest') {
    return ContentService
      .createTextOutput(JSON.stringify(buildPwaManifest_(serviceUrl), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (String(params.asset || '') === 'sw') {
    const out = ContentService.createTextOutput(buildPwaServiceWorker_(serviceUrl));
    try {
      out.setMimeType(ContentService.MimeType.JAVASCRIPT);
    } catch (err) {
      out.setMimeType(ContentService.MimeType.TEXT);
    }
    return out;
  }
  if (String(params.diag || '') === 'server') {
    return ContentService
      .createTextOutput(JSON.stringify(diagnostics(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Apps Script sirve las web apps dentro de un iframe con origen googleusercontent.com.
  // Por eso el frontend no siempre puede leer la URL original /exec?... desde location.search.
  // Pasamos los parametros recibidos por doGet al HTML como datos iniciales.
  const template = HtmlService.createTemplateFromFile('index');
  template.initialParamsJson = safeJsonForHtml_(params);
  template.serviceUrlJson = safeJsonForHtml_(serviceUrl);
  template.pwaManifestUrlJson = safeJsonForHtml_(buildAssetUrl_(serviceUrl, 'manifest'));
  template.pwaServiceWorkerUrlJson = safeJsonForHtml_(buildAssetUrl_(serviceUrl, 'sw'));
  template.pwaIconDataUrlJson = safeJsonForHtml_(buildPwaIconDataUrl_());
  return template.evaluate()
    .setTitle('Penca Mundial 2026 - Los Chacales')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API JSON para el frontend hospedado fuera de Apps Script (ej. Vercel).
 * Recibe { fn: 'getBootstrap', args: [...] } como texto plano (sin preflight CORS)
 * y devuelve { ok:true, result } o { ok:false, error }.
 */
function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    }
    const fn = String(body.fn || '');
    const args = Array.isArray(body.args) ? body.args : [];
    const api = {
      getBootstrap: getBootstrap,
      registerPlayer: registerPlayer,
      loginPlayer: loginPlayer,
      recoverPlayerAccess: recoverPlayerAccess,
      updateMyAuth: updateMyAuth,
      saveDraft: saveDraft,
      submitPrediction: submitPrediction,
      getAdminState: getAdminState,
      updateActual: updateActual,
      saveSettings: saveSettings,
      unlockPrediction: unlockPrediction,
      resetPlayerAccess: resetPlayerAccess,
      deletePlayer: deletePlayer,
      importSquads: importSquads,
      recalculateRanking: recalculateRanking,
      diagnostics: diagnostics,
      previewOfficialResults: previewOfficialResults,
      syncOfficialResults: syncOfficialResults,
      syncOfficialSquads: syncOfficialSquads,
      installOfficialResultsSync: installOfficialResultsSync,
      removeOfficialResultsSync: removeOfficialResultsSync
    };
    if (!Object.prototype.hasOwnProperty.call(api, fn)) {
      throw new Error('Accion no permitida: ' + fn);
    }
    const result = api[fn].apply(null, args);
    out.setContent(JSON.stringify({ ok: true, result: (result === undefined ? null : result) }));
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: (err && err.message) ? err.message : String(err) }));
  }
  return out;
}

function buildAssetUrl_(serviceUrl, asset) {
  const base = String(serviceUrl || '').trim();
  if (!base) return '';
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'asset=' + encodeURIComponent(asset);
}

function buildPwaIconDataUrl_() {
  const svg = [
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">',
    '<defs>',
    '<linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#3a1426"/><stop offset="55%" stop-color="#1a1018"/><stop offset="100%" stop-color="#0c0810"/></linearGradient>',
    '<radialGradient id="gl" cx="50%" cy="6%" r="78%"><stop offset="0%" stop-color="#ff669b" stop-opacity="0.55"/><stop offset="55%" stop-color="#ff669b" stop-opacity="0"/></radialGradient>',
    '<linearGradient id="pk" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#ffd7e6"/><stop offset="46%" stop-color="#ff90b6"/><stop offset="100%" stop-color="#ff5a94"/></linearGradient>',
    '</defs>',
    '<rect width="512" height="512" rx="116" fill="url(#bg)"/>',
    '<rect width="512" height="512" rx="116" fill="url(#gl)"/>',
    '<rect x="14" y="14" width="484" height="484" rx="104" fill="none" stroke="#ff669b" stroke-opacity="0.42" stroke-width="8"/>',
    '<polygon points="256,70 268,104 304,104 275,126 286,160 256,139 226,160 237,126 208,104 244,104" fill="#ffd86b"/>',
    '<polygon points="256,148 300,154 374,86 332,196 360,236 372,266 372,300 344,300 338,330 300,394 284,432 256,452 228,432 212,394 174,330 168,300 140,300 140,266 152,236 180,196 138,86 212,154" fill="url(#pk)"/>',
    '<polygon points="316,166 360,112 332,194" fill="#190e16"/><polygon points="196,166 152,112 180,194" fill="#190e16"/>',
    '<polygon points="296,254 334,242 328,272 300,280" fill="#150b13"/><polygon points="216,254 178,242 184,272 212,280" fill="#150b13"/>',
    '<circle cx="320" cy="252" r="5" fill="#ffe3ee"/><circle cx="192" cy="252" r="5" fill="#ffe3ee"/>',
    '<polygon points="256,306 282,320 272,410 256,432 240,410 230,320" fill="#ffffff" fill-opacity="0.17"/>',
    '<polygon points="236,352 276,352 256,380" fill="#150b13"/>',
    '</svg>'
  ].join('');
  return 'data:image/svg+xml;base64,' + Utilities.base64Encode(svg);
}

function buildPwaManifest_(serviceUrl) {
  const startUrl = String(serviceUrl || '').trim();
  const icon = buildPwaIconDataUrl_();
  return {
    id: startUrl || 'penca-chacal',
    name: 'Penca Chacal',
    short_name: 'Chacales',
    description: 'Pronósticos, resultados y ranking de la Penca Chacal.',
    start_url: startUrl,
    scope: startUrl,
    display: 'standalone',
    orientation: 'portrait',
    background_color: '#060407',
    theme_color: '#100b12',
    icons: [
      { src: icon, sizes: '192x192', type: 'image/svg+xml', purpose: 'any maskable' },
      { src: icon, sizes: '512x512', type: 'image/svg+xml', purpose: 'any maskable' }
    ]
  };
}

function buildPwaServiceWorker_(serviceUrl) {
  const startUrl = String(serviceUrl || '').trim();
  const manifestUrl = buildAssetUrl_(serviceUrl, 'manifest');
  const cacheName = 'penca-chacal-' + String(LC_APP_VERSION || 'v1').replace(/[^a-z0-9_-]/gi, '-').toLowerCase();
  return [
    "const CACHE_NAME = " + JSON.stringify(cacheName) + ";",
    "const START_URL = " + JSON.stringify(startUrl) + ";",
    "const ASSETS = [" + JSON.stringify(startUrl) + ", " + JSON.stringify(manifestUrl) + "].filter(Boolean);",
    "self.addEventListener('install', event => {",
    "  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting()));",
    "});",
    "self.addEventListener('activate', event => {",
    "  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))).then(() => self.clients.claim()));",
    "});",
    "self.addEventListener('fetch', event => {",
    "  if (event.request.method !== 'GET') return;",
    "  const url = new URL(event.request.url);",
    "  if (event.request.mode === 'navigate') {",
    "    event.respondWith(fetch(event.request).catch(() => caches.match(START_URL)));",
    "    return;",
    "  }",
    "  if (url.origin !== self.location.origin) return;",
    "  event.respondWith(caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {",
    "    const copy = response.clone();",
    "    caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));",
    "    return response;",
    "  })));",
    "});"
  ].join('\\n');
}

function safeJsonForHtml_(value) {
  return JSON.stringify(value || {})
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}



function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(LC_HEADERS).forEach(name => ensureSheet_(ss, name, LC_HEADERS[name]));
  seedConfig_();
  seedActual_();
  seedSquadsExample_();
  migratePlayerAuth_();
  log_('SETUP', '', 'Hojas inicializadas');
  return { ok: true, message: 'Penca Los Chacales inicializada.' };
}

function getBootstrap(params) {
  setupLight_();
  const autoSubmit = autoSubmitSavedDraftsAfterDeadline_();
  params = params || {};
  const config = getConfig_();
  const reveal = String(config.REVEAL_PREDICTIONS || '').toUpperCase() === 'TRUE';
  const players = getPlayers_().map(p => ({
    playerId: p.PlayerId,
    name: p.Name,
    submittedAt: p.SubmittedAt,
    createdAt: p.CreatedAt,
    lastSavedAt: p.LastSavedAt
  }));
  const actual = getActual_();
  // Público: los pronósticos ajenos se revelan recién cuando pasó la fecha límite.
  // Antes de esa fecha, solo cada usuario ve su propio pronóstico desde requester.
  // Admin sigue viendo todo desde getAdminState().
  const revealOpen = reveal && predictionsRevealOpen_(config);
  const predictions = revealOpen ? getSubmittedPredictions_() : {};
  const requester = getRequester_(params.playerId, params.token);
  return {
    ok: true,
    version: LC_APP_VERSION,
    appUrl: ScriptApp.getService().getUrl(),
    config: publicConfig_(config),
    players,
    predictions,
    actual,
    squads: getSquads_(),
    activity: getPublicActivity_(parseInt(config.ACTIVITY_MAX_ITEMS || '24', 10) || 24),
    privacy: { predictionsRevealOpen: revealOpen, deadline: config.PREDICTION_DEADLINE || '', revealPredictions: reveal, autoSubmittedDrafts: autoSubmit.submitted || 0 },
    requester
  };
}

function registerPlayer(name, email, password) {
  setupLight_();
  name = String(name || '').trim();
  email = normalizeEmail_(email || '');
  password = String(password || '').trim();
  if (!name) throw new Error('Falta el nombre del jugador.');
  if (name.length > 60) throw new Error('El nombre es demasiado largo.');
  if (email && !isValidEmail_(email)) throw new Error('El mail no parece valido.');
  if (!password || password.length < 4) throw new Error('La clave debe tener al menos 4 caracteres.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const players = getPlayers_();
    const normalized = normalizeName_(name);
    const existing = players.find(p => normalizeName_(p.Name) === normalized);
    if (existing) throw new Error('Ese nombre ya esta cargado. Entra con usuario y clave, o pedile al admin que resetee tu acceso.');
    if (email) {
      const sameEmail = players.find(p => normalizeEmail_(p.Email || '') === email);
      if (sameEmail) throw new Error('Ese mail ya esta asociado a otro jugador.');
    }

    const now = nowIso_();
    const playerId = Utilities.getUuid();
    const token = makeToken_();
    const auth = makePasswordAuth_(password);
    getSheet_(LC_SHEETS.PLAYERS).appendRow([playerId, name, token, '', now, now, email, auth.hash, auth.salt, now, '']);
    getSheet_(LC_SHEETS.PREDICTIONS).appendRow([playerId, '{}', '', now]);
    log_('REGISTER_PLAYER', playerId, name + (email ? ' <' + email + '>' : ''));
    return { ok: true, playerId, token, name, email, hasPassword: true };
  } finally {
    lock.releaseLock();
  }
}

function loginPlayer(identifier, password) {
  setupLight_();
  identifier = String(identifier || '').trim();
  password = String(password || '').trim();
  if (!identifier) throw new Error('Ingresa tu nombre, apodo o mail.');
  if (!password) throw new Error('Ingresa tu clave.');
  const player = findPlayerByLogin_(identifier);
  if (!player) throw new Error('No encontre ese usuario. Revisa el nombre/mail o pedi al admin que resetee tu acceso.');
  const hasPassword = !!String(player.PasswordHash || '').trim();
  let ok = false;
  if (hasPassword) ok = verifyPassword_(password, player.PasswordSalt || '', player.PasswordHash || '');
  else ok = String(player.TempPassword || '').trim().toUpperCase() === password.toUpperCase() || String(player.Token || '') === password;
  if (!ok) throw new Error('Clave incorrecta.');
  updatePlayerAuthLogin_(player.PlayerId);
  log_('LOGIN_PLAYER', player.PlayerId, 'Login correcto');
  return { ok: true, playerId: player.PlayerId, token: player.Token, name: player.Name, email: player.Email || '', needsPasswordSetup: !hasPassword };
}


function recoverPlayerAccess(identifier) {
  setupLight_();
  identifier = String(identifier || '').trim();
  if (!identifier) throw new Error('Ingresa tu nombre, apodo o mail.');
  const generic = { ok: true, message: 'Si el usuario existe y tiene mail cargado, enviaremos una clave temporal.' };
  const cache = CacheService.getScriptCache();
  const cacheKey = 'LC_RECOVERY_' + hashPassword_(normalizeName_(identifier), 'recover').slice(0, 32);
  if (cache.get(cacheKey)) return generic;
  cache.put(cacheKey, '1', 600);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const player = findPlayerByLogin_(identifier);
    if (!player || !isValidEmail_(player.Email || '')) {
      log_('RECOVER_ACCESS_REQUEST', player ? player.PlayerId : '', 'Sin mail valido o usuario no encontrado');
      return generic;
    }
    const temp = makeTempPassword_();
    const sheet = getSheet_(LC_SHEETS.PLAYERS);
    const data = sheet.getDataRange().getValues();
    const h = headerMap_(data[0]);
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][h.PlayerId]) === String(player.PlayerId)) {
        const row = r + 1;
        if (h.PasswordHash != null) sheet.getRange(row, h.PasswordHash + 1).setValue('');
        if (h.PasswordSalt != null) sheet.getRange(row, h.PasswordSalt + 1).setValue('');
        if (h.TempPassword != null) sheet.getRange(row, h.TempPassword + 1).setValue(temp);
        if (h.AuthUpdatedAt != null) sheet.getRange(row, h.AuthUpdatedAt + 1).setValue(nowIso_());
        MailApp.sendEmail({
          to: player.Email,
          subject: 'Penca Chacal - clave temporal',
          body: 'Hola ' + player.Name + ',\n\nTu clave temporal para entrar a la Penca Chacal es: ' + temp + '\n\nEntrá con tu nombre/mail y esta clave. Luego cambiá la clave desde Mi pronóstico > Usuario y clave.\n\nSi no pediste esto, avisale al admin.'
        });
        log_('RECOVER_ACCESS_SENT', player.PlayerId, 'Clave temporal enviada por mail');
        return generic;
      }
    }
    return generic;
  } finally {
    lock.releaseLock();
  }
}

function updateMyAuth(playerId, token, email, password) {
  setupLight_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const player = requirePlayer_(playerId, token);
    email = normalizeEmail_(email || '');
    password = String(password || '').trim();
    if (email && !isValidEmail_(email)) throw new Error('El mail no parece valido.');
    if (!password && !String(player.PasswordHash || '').trim()) throw new Error('Para activar el acceso desde otros dispositivos, guarda una clave.');
    if (password && password.length < 4) throw new Error('La clave debe tener al menos 4 caracteres.');
    if (email) {
      const duplicate = getPlayers_().find(p => String(p.PlayerId) !== String(playerId) && normalizeEmail_(p.Email || '') === email);
      if (duplicate) throw new Error('Ese mail ya esta asociado a otro jugador.');
    }
    const sheet = getSheet_(LC_SHEETS.PLAYERS);
    const data = sheet.getDataRange().getValues();
    const h = headerMap_(data[0]);
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][h.PlayerId]) === String(playerId)) {
        const row = r + 1;
        if (h.Email != null) sheet.getRange(row, h.Email + 1).setValue(email);
        if (password) {
          const auth = makePasswordAuth_(password);
          if (h.PasswordHash != null) sheet.getRange(row, h.PasswordHash + 1).setValue(auth.hash);
          if (h.PasswordSalt != null) sheet.getRange(row, h.PasswordSalt + 1).setValue(auth.salt);
          if (h.TempPassword != null) sheet.getRange(row, h.TempPassword + 1).setValue('');
        }
        if (h.AuthUpdatedAt != null) sheet.getRange(row, h.AuthUpdatedAt + 1).setValue(nowIso_());
        log_('UPDATE_AUTH', playerId, 'Jugador actualizo mail/clave');
        return { ok: true, email: email, hasPassword: true };
      }
    }
    throw new Error('Jugador no encontrado.');
  } finally {
    lock.releaseLock();
  }
}

function saveDraft(playerId, token, prediction) {
  setupLight_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const player = requirePlayer_(playerId, token);
    assertCanEditPlayer_(player);
    upsertPrediction_(playerId, prediction, false);
    updatePlayerLastSaved_(playerId, false);
    clearComputedCaches_();
    log_('SAVE_DRAFT', playerId, 'Borrador guardado');
    return { ok: true, savedAt: nowIso_() };
  } finally {
    lock.releaseLock();
  }
}

function submitPrediction(playerId, token, prediction) {
  setupLight_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const player = requirePlayer_(playerId, token);
    assertCanEditPlayer_(player);
    const submittedAt = nowIso_();
    upsertPrediction_(playerId, prediction, true);
    updatePlayerLastSaved_(playerId, true, submittedAt);
    clearComputedCaches_();
    const receiptCode = makeReceiptCode_(playerId, submittedAt, safeStringify_(prediction || {}));
    log_('SUBMIT_PREDICTION', playerId, 'Pronostico bloqueado; comprobante ' + receiptCode);
    return { ok: true, submittedAt, receiptCode };
  } finally {
    lock.releaseLock();
  }
}

function getAdminState(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  const autoSubmit = autoSubmitSavedDraftsAfterDeadline_();
  const config = getConfig_();
  const appUrl = ScriptApp.getService().getUrl();
  const players = getPlayers_().map(p => {
    const pred = getPredictionRow_(p.PlayerId);
    const receiptCode = pred && pred.SubmittedJson && p.SubmittedAt ? makeReceiptCode_(p.PlayerId, p.SubmittedAt, String(pred.SubmittedJson)) : '';
    return {
      playerId: p.PlayerId,
      name: p.Name,
      token: p.Token,
      email: p.Email || '',
      hasPassword: !!String(p.PasswordHash || '').trim(),
      needsPasswordSetup: !String(p.PasswordHash || '').trim(),
      tempPassword: p.TempPassword || '',
      submittedAt: p.SubmittedAt,
      createdAt: p.CreatedAt,
      lastSavedAt: p.LastSavedAt,
      receiptCode: receiptCode,
      link: appUrl ? appUrl + '?p=' + encodeURIComponent(p.PlayerId) + '&t=' + encodeURIComponent(p.Token) : ''
    };
  });
  return {
    ok: true,
    appUrl,
    config: publicConfig_(config),
    privateConfig: privateConfigForAdmin_(config),
    officialSync: officialSyncStatus_(config),
    players,
    predictions: getAllPredictions_(),
    actual: getActual_(),
    squads: getSquads_(),
    activity: getPublicActivity_(40),
    autoSubmittedDrafts: autoSubmit.submitted || 0
  };
}

function updateActual(adminPin, actual) {
  setupLight_();
  requireAdmin_(adminPin);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_(LC_SHEETS.ACTUAL);
    const rows = sheet.getDataRange().getValues();
    const json = safeStringify_(actual || {});
    const now = nowIso_();
    let updated = false;
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) === 'actual') {
        sheet.getRange(r + 1, 2, 1, 2).setValues([[json, now]]);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow(['actual', json, now]);
    clearComputedCaches_();
    log_('UPDATE_ACTUAL', '', 'Resultados reales actualizados');
    return { ok: true, updatedAt: now };
  } finally {
    lock.releaseLock();
  }
}

function saveSettings(adminPin, settings) {
  setupLight_();
  requireAdmin_(adminPin);
  settings = settings || {};
  const allowed = [
    'PREDICTION_DEADLINE', 'REVEAL_PREDICTIONS', 'OFFICIAL_RESULTS_URL',
    'POINT_RESULT_EXACT', 'POINT_RESULT_OUTCOME', 'POINT_GROUP_EXACT', 'POINT_GROUP_OUTCOME', 'POINT_R32', 'POINT_R16',
    'POINT_QF', 'POINT_SF', 'POINT_THIRD_PLACE', 'POINT_FINALIST', 'POINT_CHAMPION', 'POINT_AWARD',
    'ACTIVITY_MAX_ITEMS',
    'OFFICIAL_RESULTS_SYNC_URL', 'OFFICIAL_RESULTS_API_KEY', 'OFFICIAL_SQUADS_SYNC_URL', 'OFFICIAL_SQUADS_API_KEY',
    'OFFICIAL_SYNC_AUTO', 'OFFICIAL_SYNC_MINUTES'
  ];
  allowed.forEach(k => {
    if (Object.prototype.hasOwnProperty.call(settings, k)) {
      const rawValue = String(settings[k] == null ? '' : settings[k]);
      setConfigValue_(k, k === 'PREDICTION_DEADLINE' ? normalizeDeadlineInput_(rawValue) : rawValue);
    }
  });
  clearComputedCaches_();
  log_('SAVE_SETTINGS', '', JSON.stringify(settings));
  return { ok: true, config: publicConfig_(getConfig_()) };
}

function unlockPrediction(adminPin, playerId) {
  setupLight_();
  requireAdmin_(adminPin);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_(LC_SHEETS.PLAYERS);
    const data = sheet.getDataRange().getValues();
    const h = headerMap_(data[0]);
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][h.PlayerId]) === String(playerId)) {
        sheet.getRange(r + 1, h.SubmittedAt + 1).setValue('');
        clearSubmittedPredictionJson_(playerId);
        log_('UNLOCK_PREDICTION', playerId, 'Admin desbloqueo pronostico');
        return { ok: true };
      }
    }
    throw new Error('Jugador no encontrado.');
  } finally {
    lock.releaseLock();
  }
}

function resetPlayerAccess(adminPin, playerId) {
  setupLight_();
  requireAdmin_(adminPin);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const temp = makeTempPassword_();
    const sheet = getSheet_(LC_SHEETS.PLAYERS);
    const data = sheet.getDataRange().getValues();
    const h = headerMap_(data[0]);
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][h.PlayerId]) === String(playerId)) {
        const row = r + 1;
        if (h.PasswordHash != null) sheet.getRange(row, h.PasswordHash + 1).setValue('');
        if (h.PasswordSalt != null) sheet.getRange(row, h.PasswordSalt + 1).setValue('');
        if (h.TempPassword != null) sheet.getRange(row, h.TempPassword + 1).setValue(temp);
        if (h.AuthUpdatedAt != null) sheet.getRange(row, h.AuthUpdatedAt + 1).setValue(nowIso_());
        log_('RESET_ACCESS', playerId, 'Admin reseteo acceso');
        return { ok: true, tempPassword: temp };
      }
    }
    throw new Error('Jugador no encontrado.');
  } finally {
    lock.releaseLock();
  }
}

function deletePlayer(adminPin, playerId) {
  setupLight_();
  requireAdmin_(adminPin);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    deleteRowById_(LC_SHEETS.PLAYERS, 'PlayerId', playerId);
    deleteRowById_(LC_SHEETS.PREDICTIONS, 'PlayerId', playerId);
    log_('DELETE_PLAYER', playerId, 'Admin borro jugador');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function importSquads(adminPin, squads) {
  setupLight_();
  requireAdmin_(adminPin);
  const rows = [];
  if (Array.isArray(squads)) {
    squads.forEach(item => {
      const team = String(item.team || item.Team || '').trim();
      const player = String(item.player || item.Player || item.name || item.Name || '').trim();
      const position = String(item.position || item.Position || item.posicion || item.Posicion || item.role || item.Role || '').trim();
      const number = item.number || item.Number || item.numero || item.Numero || '';
      if (team && player) rows.push([team, player, position, number]);
    });
  } else if (squads && typeof squads === 'object') {
    Object.keys(squads).forEach(team => {
      const list = squads[team] || [];
      if (Array.isArray(list)) {
        list.forEach(item => {
          if (typeof item === 'string') rows.push([team, item, '', '']);
          else if (item && typeof item === 'object') {
            const player = String(item.name || item.Name || item.player || item.Player || '').trim();
            const position = String(item.position || item.Position || item.posicion || item.Posicion || item.role || item.Role || '').trim();
            const number = item.number || item.Number || item.numero || item.Numero || '';
            if (player) rows.push([team, player, position, number]);
          }
        });
      }
    });
  }
  if (!rows.length) throw new Error('No encontre jugadores para importar.');

  const sheet = getSheet_(LC_SHEETS.SQUADS);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, LC_HEADERS.Squads.length).setValues([LC_HEADERS.Squads]);
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  clearComputedCaches_();
  log_('IMPORT_SQUADS', '', rows.length + ' jugadores');
  return { ok: true, count: rows.length, squads: getSquads_() };
}

function recalculateRanking(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  clearComputedCaches_();
  log_('RECALCULATE', '', 'Admin recalculo ranking/cache');
  return { ok: true, updatedAt: nowIso_(), version: LC_APP_VERSION };
}

function recalculateEverything(adminPin) {
  return recalculateRanking(adminPin);
}



function warmupNews() {
  return { ok: false, disabled: true, version: LC_APP_VERSION, message: 'Modulo externo desactivado en v20.' };
}

function refreshNewsFeed() {
  return getNewsFeed();
}

function getNewsFeed() {
  return { ok: false, disabled: true, version: LC_APP_VERSION, updatedAt: nowIso_(), sources: [], message: 'Modulo externo desactivado en v20.' };
}


/**
 * v19 - Sincronización opcional de resultados y planteles desde una fuente JSON/API.
 * La app NO raspa la página visual de FIFA. El admin puede configurar un endpoint JSON
 * propio/licenciado y la app normaliza el payload hacia el formato interno.
 */
function previewOfficialResults(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  const result = fetchAndNormalizeOfficialResults_();
  return Object.assign({ ok: true, preview: true, updatedAt: nowIso_() }, result.summary);
}

function syncOfficialResults(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  return syncOfficialResultsInternal_('manual-admin');
}

function scheduledOfficialResultsSync() {
  setupLight_();
  const cfg = getConfig_();
  if (String(cfg.OFFICIAL_SYNC_AUTO || '').toUpperCase() !== 'TRUE') {
    return { ok: false, skipped: true, message: 'Sincronizacion automatica desactivada.' };
  }
  return syncOfficialResultsInternal_('scheduled');
}

function syncOfficialResultsInternal_(source) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const result = fetchAndNormalizeOfficialResults_();
    const sheet = getSheet_(LC_SHEETS.ACTUAL);
    const rows = sheet.getDataRange().getValues();
    const now = nowIso_();
    const json = safeStringify_(result.actual);
    let updated = false;
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) === 'actual') {
        sheet.getRange(r + 1, 2, 1, 2).setValues([[json, now]]);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow(['actual', json, now]);
    setConfigValue_('OFFICIAL_LAST_SYNC_AT', now);
    setConfigValue_('OFFICIAL_LAST_SYNC_STATUS', 'OK: ' + result.summary.importedMatches + ' partidos importados');
    clearComputedCaches_();
    log_('SYNC_OFFICIAL_RESULTS', '', source + ' · ' + result.summary.importedMatches + ' partidos');
    return Object.assign({ ok: true, updatedAt: now, source: source || 'manual' }, result.summary);
  } catch (err) {
    setConfigValue_('OFFICIAL_LAST_SYNC_AT', nowIso_());
    setConfigValue_('OFFICIAL_LAST_SYNC_STATUS', 'ERROR: ' + err.message);
    log_('SYNC_OFFICIAL_RESULTS_ERROR', '', err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function fetchAndNormalizeOfficialResults_() {
  const cfg = getConfig_();
  const url = String(cfg.OFFICIAL_RESULTS_SYNC_URL || '').trim();
  if (!url) throw new Error('Falta OFFICIAL_RESULTS_SYNC_URL. Configura una URL JSON/API en Admin.');
  const payload = fetchJson_(url, cfg.OFFICIAL_RESULTS_API_KEY || '');
  return normalizeExternalResultsToActual_(payload, getActual_());
}

function syncOfficialSquads(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  const cfg = getConfig_();
  const url = String(cfg.OFFICIAL_SQUADS_SYNC_URL || '').trim();
  if (!url) throw new Error('Falta OFFICIAL_SQUADS_SYNC_URL. Configura una URL JSON/API de planteles en Admin.');
  const payload = fetchJson_(url, cfg.OFFICIAL_SQUADS_API_KEY || '');
  const rows = normalizeExternalSquadRows_(payload);
  if (!rows.length) throw new Error('No encontre jugadores en el JSON recibido. Usa el importador manual o revisa el mapeo del proveedor.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_(LC_SHEETS.SQUADS);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, LC_HEADERS.Squads.length).setValues([LC_HEADERS.Squads]);
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    const now = nowIso_();
    setConfigValue_('OFFICIAL_SQUADS_LAST_SYNC_AT', now);
    setConfigValue_('OFFICIAL_SQUADS_LAST_SYNC_STATUS', 'OK: ' + rows.length + ' jugadores importados');
    clearComputedCaches_();
    log_('SYNC_OFFICIAL_SQUADS', '', rows.length + ' jugadores');
    return { ok: true, count: rows.length, updatedAt: now, squads: getSquads_() };
  } finally {
    lock.releaseLock();
  }
}

function installOfficialResultsSync(adminPin, minutes) {
  setupLight_();
  requireAdmin_(adminPin);
  minutes = Number(minutes || getConfig_().OFFICIAL_SYNC_MINUTES || 15);
  const allowed = [1, 5, 10, 15, 30];
  if (allowed.indexOf(minutes) < 0) minutes = 15;
  removeOfficialResultsSyncTriggers_();
  ScriptApp.newTrigger('scheduledOfficialResultsSync').timeBased().everyMinutes(minutes).create();
  setConfigValue_('OFFICIAL_SYNC_AUTO', 'TRUE');
  setConfigValue_('OFFICIAL_SYNC_MINUTES', String(minutes));
  log_('INSTALL_OFFICIAL_SYNC_TRIGGER', '', 'Cada ' + minutes + ' minutos');
  return { ok: true, minutes: minutes };
}

function removeOfficialResultsSync(adminPin) {
  setupLight_();
  requireAdmin_(adminPin);
  const count = removeOfficialResultsSyncTriggers_();
  setConfigValue_('OFFICIAL_SYNC_AUTO', 'FALSE');
  log_('REMOVE_OFFICIAL_SYNC_TRIGGER', '', count + ' triggers removidos');
  return { ok: true, removed: count };
}

function removeOfficialResultsSyncTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'scheduledOfficialResultsSync') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  return count;
}

function fetchJson_(url, apiKey) {
  url = fillApiKey_(String(url || '').trim(), String(apiKey || '').trim());
  if (!/^https:\/\//i.test(url)) throw new Error('La URL de sincronizacion debe ser https:// y devolver JSON.');
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'Accept': 'application/json', 'User-Agent': 'PencaLosChacales/' + LC_APP_VERSION }
  });
  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  if (code < 200 || code >= 300) throw new Error('La fuente oficial/API respondio HTTP ' + code + ': ' + text.slice(0, 180));
  try {
    return JSON.parse(text);
  } catch (err) {
    throw new Error('La URL respondio, pero no devolvio JSON parseable. Usa un endpoint API, no una pagina HTML visual. ' + err.message);
  }
}

function fillApiKey_(url, apiKey) {
  if (!apiKey) return url;
  let out = url.replace(/\{\{API_KEY\}\}|\{API_KEY\}|\{key\}/g, encodeURIComponent(apiKey));
  if (out === url && !/[?&](key|api_key|apikey|apiKey)=/i.test(out)) {
    out += (out.indexOf('?') >= 0 ? '&' : '?') + 'key=' + encodeURIComponent(apiKey);
  }
  return out;
}

function normalizeExternalResultsToActual_(payload, current) {
  const actual = normalizeServerEntry_(current || {});
  if (payload && payload.actual) payload = payload.actual;
  if (payload && (payload.groupScores || payload.knockoutScores || payload.knockoutWinners || payload.awards)) {
    const direct = normalizeServerEntry_(payload);
    actual.groupScores = Object.assign(actual.groupScores, direct.groupScores || {});
    actual.knockoutScores = Object.assign(actual.knockoutScores, direct.knockoutScores || {});
    actual.knockoutWinners = Object.assign(actual.knockoutWinners, direct.knockoutWinners || {});
    actual.awards = Object.assign(actual.awards, direct.awards || {});
    return { actual, summary: { importedMatches: Object.keys(direct.groupScores || {}).length + Object.keys(direct.knockoutScores || {}).length, groupImported: Object.keys(direct.groupScores || {}).length, koImported: Object.keys(direct.knockoutScores || {}).length, winnersImported: Object.keys(direct.knockoutWinners || {}).length, directFormat: true } };
  }

  const matches = [];
  collectMatchObjects_(payload, matches, 0);
  let groupImported = 0, koImported = 0, winnersImported = 0;
  matches.forEach(m => {
    const no = extractMatchNo_(m);
    const id = officialMatchIdFromNo_(no) || String(firstPath_(m, ['internalId', 'internal_id', 'matchKey', 'key']) || '').trim();
    if (!id) return;
    const score = extractScore_(m);
    if (score && score.homeGoals !== '' && score.awayGoals !== '') {
      if (/^K\d+/i.test(id) || Number(no) >= 73) { actual.knockoutScores[id] = score; koImported++; }
      else { actual.groupScores[id] = score; groupImported++; }
    }
    const winner = normalizeTeamName_(firstPath_(m, ['winner', 'winnerTeam', 'winner_team', 'qualified', 'qualifiedTeam', 'winner.name', 'winner.code', 'winnerTeam.name', 'team_winner', 'winningTeam']));
    if (winner && (/^K\d+/i.test(id) || Number(no) >= 73)) { actual.knockoutWinners[id] = winner; winnersImported++; }
  });
  return { actual, summary: { importedMatches: groupImported + koImported, groupImported, koImported, winnersImported, scannedMatches: matches.length, directFormat: false } };
}

function normalizeServerEntry_(entry) {
  entry = entry || {};
  return {
    groupScores: Object.assign({}, entry.groupScores || {}),
    knockoutScores: Object.assign({}, entry.knockoutScores || {}),
    knockoutWinners: Object.assign({}, entry.knockoutWinners || {}),
    r32Overrides: Object.assign({}, entry.r32Overrides || {}),
    awards: Object.assign({}, entry.awards || {})
  };
}

function collectMatchObjects_(node, out, depth) {
  if (!node || depth > 8) return;
  if (Array.isArray(node)) { node.forEach(x => collectMatchObjects_(x, out, depth + 1)); return; }
  if (typeof node !== 'object') return;
  if (looksLikeMatchObject_(node)) out.push(node);
  Object.keys(node).forEach(k => collectMatchObjects_(node[k], out, depth + 1));
}

function looksLikeMatchObject_(m) {
  const no = extractMatchNo_(m);
  const home = firstPath_(m, ['home', 'homeTeam', 'home_team', 'team1', 'localTeam', 'home.name', 'homeTeam.name']);
  const away = firstPath_(m, ['away', 'awayTeam', 'away_team', 'team2', 'visitorTeam', 'away.name', 'awayTeam.name']);
  const score = extractScore_(m);
  return !!(no || (home && away && score));
}

function extractMatchNo_(m) {
  const raw = firstPath_(m, ['matchNo', 'match_no', 'matchNumber', 'match_number', 'number', 'gameNumber', 'game_number', 'fixtureNumber', 'fixture_number', 'id', 'match_id']);
  const text = String(raw == null ? '' : raw);
  const direct = parseInt(text, 10);
  if (isFinite(direct) && direct >= 1 && direct <= 104) return direct;
  const found = text.match(/(?:match|game|fixture|#|K)?\s*(\d{1,3})/i);
  const n = found ? parseInt(found[1], 10) : NaN;
  return isFinite(n) && n >= 1 && n <= 104 ? n : null;
}

function extractScore_(m) {
  const h = firstPath_(m, ['homeGoals', 'home_goals', 'homeScore', 'home_score', 'score.home', 'score.homeGoals', 'score.fullTime.home', 'score.ft.home', 'goals.home', 'goalsHomeTeam', 'home.goals', 'home.score', 'homeTeam.score', 'team1_score']);
  const a = firstPath_(m, ['awayGoals', 'away_goals', 'awayScore', 'away_score', 'score.away', 'score.awayGoals', 'score.fullTime.away', 'score.ft.away', 'goals.away', 'goalsAwayTeam', 'away.goals', 'away.score', 'awayTeam.score', 'team2_score']);
  if (h === '' || h == null || a === '' || a == null) return null;
  const hn = Number(h), an = Number(a);
  if (!isFinite(hn) || !isFinite(an)) return null;
  return { homeGoals: Math.max(0, Math.floor(hn)), awayGoals: Math.max(0, Math.floor(an)) };
}

function officialMatchIdFromNo_(no) {
  no = Number(no);
  if (!isFinite(no)) return '';
  const groupMap = officialGroupMatchMap_();
  if (groupMap[no]) return groupMap[no];
  if (no >= 73 && no <= 104) return 'K' + no;
  return '';
}

function officialGroupMatchMap_() {
  const numbers = {
    A: [1, 2, 28, 25, 53, 54], B: [3, 8, 27, 26, 51, 52], C: [7, 5, 29, 30, 49, 50], D: [4, 6, 32, 31, 59, 60],
    E: [10, 9, 33, 34, 55, 56], F: [11, 12, 35, 36, 57, 58], G: [16, 15, 39, 40, 64, 63], H: [14, 13, 38, 37, 66, 65],
    I: [17, 18, 42, 41, 61, 62], J: [19, 20, 43, 44, 70, 69], K: [23, 24, 47, 48, 71, 72], L: [22, 21, 45, 46, 67, 68]
  };
  const map = {};
  Object.keys(numbers).forEach(g => numbers[g].forEach((n, i) => map[n] = g + (i + 1)));
  return map;
}

function firstPath_(obj, paths) {
  for (let i = 0; i < paths.length; i++) {
    const v = readPath_(obj, paths[i]);
    if (v !== undefined && v !== null && String(v) !== '') return v;
  }
  return '';
}

function readPath_(obj, path) {
  const parts = String(path || '').split('.');
  let cur = obj;
  for (let i = 0; i < parts.length; i++) {
    if (cur == null) return undefined;
    cur = cur[parts[i]];
  }
  return cur;
}

function normalizeTeamName_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const map = officialTeamAliasMap_();
  const key = normalizeAlias_(raw);
  return map[key] || raw;
}

function normalizeAlias_(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, ' ').trim();
}

function officialTeamAliasMap_() {
  const aliases = {
    'México': ['Mexico','MEX'], 'Sudáfrica': ['South Africa','RSA','SAF'], 'Corea del Sur': ['Korea Republic','South Korea','KOR'], 'República Checa': ['Czechia','Czech Republic','CZE'],
    'Canadá': ['Canada','CAN'], 'Bosnia y Herzegovina': ['Bosnia and Herzegovina','Bosnia-Herzegovina','BIH'], 'Qatar': ['QAT'], 'Suiza': ['Switzerland','SUI'],
    'Brasil': ['Brazil','BRA'], 'Marruecos': ['Morocco','MAR'], 'Haití': ['Haiti','HAI'], 'Escocia': ['Scotland','SCO'],
    'Estados Unidos': ['United States','USA','US','United States of America'], 'Paraguay': ['PAR'], 'Australia': ['AUS'], 'Turquía': ['Turkey','Turkiye','Türkiye','TUR'],
    'Alemania': ['Germany','GER','DEU'], 'Curazao': ['Curacao','Curaçao','CUW'], 'Costa de Marfil': ["Cote d Ivoire", "Côte d'Ivoire", 'Ivory Coast','CIV'], 'Ecuador': ['ECU'],
    'Países Bajos': ['Netherlands','Holland','NED'], 'Japón': ['Japan','JPN'], 'Suecia': ['Sweden','SWE'], 'Túnez': ['Tunisia','TUN'],
    'Bélgica': ['Belgium','BEL'], 'Egipto': ['Egypt','EGY'], 'Irán': ['Iran','IR Iran','IRN'], 'Nueva Zelanda': ['New Zealand','NZL'],
    'España': ['Spain','ESP'], 'Cabo Verde': ['Cape Verde','Cabo Verde','CPV'], 'Arabia Saudita': ['Saudi Arabia','KSA','Saudi'], 'Uruguay': ['URU'],
    'Francia': ['France','FRA'], 'Senegal': ['SEN'], 'Irak': ['Iraq','IRQ'], 'Noruega': ['Norway','NOR'],
    'Argentina': ['ARG'], 'Argelia': ['Algeria','ALG'], 'Austria': ['AUT'], 'Jordania': ['Jordan','JOR'],
    'Portugal': ['POR'], 'RD Congo': ['DR Congo','Congo DR','Democratic Republic of Congo','COD'], 'Uzbekistán': ['Uzbekistan','UZB'], 'Colombia': ['COL'],
    'Inglaterra': ['England','ENG'], 'Croacia': ['Croatia','CRO'], 'Ghana': ['GHA'], 'Panamá': ['Panama','PAN']
  };
  const map = {};
  Object.keys(aliases).forEach(team => {
    map[normalizeAlias_(team)] = team;
    aliases[team].forEach(a => map[normalizeAlias_(a)] = team);
  });
  return map;
}

function normalizeExternalSquadRows_(payload) {
  const rows = [];
  if (payload && payload.data) payload = payload.data;
  if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
    Object.keys(payload).forEach(teamKey => {
      const list = payload[teamKey];
      if (Array.isArray(list)) addSquadListRows_(rows, normalizeTeamName_(teamKey), list);
    });
  }
  collectSquadRows_(payload, rows, 0);
  const seen = {};
  return rows.filter(r => {
    const key = normalizeAlias_(r[0] + '|' + r[1]);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).sort((a,b) => String(a[0]).localeCompare(String(b[0])) || String(a[3] || '').localeCompare(String(b[3] || ''), undefined, {numeric:true}) || String(a[1]).localeCompare(String(b[1])));
}

function collectSquadRows_(node, rows, depth) {
  if (!node || depth > 8) return;
  if (Array.isArray(node)) { node.forEach(x => collectSquadRows_(x, rows, depth + 1)); return; }
  if (typeof node !== 'object') return;
  const team = normalizeTeamName_(firstPath_(node, ['team', 'teamName', 'team_name', 'country', 'nation', 'team.name', 'country.name']));
  const list = firstPath_(node, ['players', 'squad', 'roster', 'members']);
  if (team && Array.isArray(list)) addSquadListRows_(rows, team, list);
  const player = firstPath_(node, ['player', 'playerName', 'player_name', 'fullName', 'full_name', 'name', 'person.name']);
  if (team && player && !Array.isArray(list) && !extractMatchNo_(node)) {
    rows.push([team, String(player).trim(), String(firstPath_(node, ['position', 'pos', 'role', 'positionName', 'position_name']) || '').trim(), firstPath_(node, ['number', 'shirtNumber', 'shirt_number', 'jersey', 'jerseyNumber']) || '']);
  }
  Object.keys(node).forEach(k => collectSquadRows_(node[k], rows, depth + 1));
}

function addSquadListRows_(rows, team, list) {
  if (!team || !Array.isArray(list)) return;
  list.forEach(p => {
    if (typeof p === 'string') rows.push([team, p, '', '']);
    else if (p && typeof p === 'object') {
      const player = firstPath_(p, ['name', 'fullName', 'full_name', 'playerName', 'player_name', 'knownName', 'shortName', 'person.name']);
      if (player) rows.push([team, String(player).trim(), String(firstPath_(p, ['position', 'pos', 'role', 'positionName', 'position_name']) || '').trim(), firstPath_(p, ['number', 'shirtNumber', 'shirt_number', 'jersey', 'jerseyNumber']) || '']);
    }
  });
}

function diagnostics() {
  const startedAt = nowIso_();
  const out = {
    ok: true,
    version: LC_APP_VERSION,
    startedAt: startedAt,
    scriptTimeZone: Session.getScriptTimeZone(),
    serviceUrl: '',
    spreadsheet: null,
    sheets: {},
    warnings: []
  };
  try {
    out.serviceUrl = ScriptApp.getService().getUrl() || '';
  } catch (err) {
    out.warnings.push('No se pudo leer URL del servicio: ' + err.message);
  }
  try {
    setupLight_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    out.spreadsheet = { name: ss.getName(), id: ss.getId() };
    Object.keys(LC_HEADERS).forEach(name => {
      const sh = ss.getSheetByName(name);
      out.sheets[name] = sh ? { exists: true, rows: sh.getLastRow(), columns: sh.getLastColumn() } : { exists: false, rows: 0, columns: 0 };
    });
    out.players = getPlayers_().length;
    out.submitted = getPlayers_().filter(p => p.SubmittedAt).length;
    out.playersWithPassword = getPlayers_().filter(p => String(p.PasswordHash || '').trim()).length;
    out.squads = Object.values(getSquads_()).reduce((acc, list) => acc + list.length, 0);
    out.actualHasScores = Object.keys((getActual_().groupScores || {})).length;
    const cfg = getConfig_();
    out.config = publicConfig_(cfg);
    out.adminPinConfigured = !!String(cfg.ADMIN_PIN || '').trim() && String(cfg.ADMIN_PIN || '').trim() !== 'cambiar-este-pin';
    out.predictionsRevealOpen = predictionsRevealOpen_(cfg);
  } catch (err) {
    out.ok = false;
    out.error = err.message;
    out.stack = err.stack || '';
  }
  return out;
}

function getRequester_(playerId, token) {
  if (!playerId || !token) return null;
  try {
    const player = requirePlayer_(playerId, token);
    const pred = getPredictionRow_(playerId);
    return {
      playerId: player.PlayerId,
      name: player.Name,
      email: player.Email || '',
      needsPasswordSetup: !String(player.PasswordHash || '').trim(),
      submittedAt: player.SubmittedAt,
      lastSavedAt: player.LastSavedAt,
      draft: pred ? parseJson_(pred.DraftJson, {}) : {},
      submitted: pred && pred.SubmittedJson && player.SubmittedAt ? parseJson_(pred.SubmittedJson, {}) : null,
      receiptCode: pred && pred.SubmittedJson && player.SubmittedAt ? makeReceiptCode_(player.PlayerId, player.SubmittedAt, String(pred.SubmittedJson)) : '',
      canEdit: !player.SubmittedAt && !deadlinePassed_()
    };
  } catch (err) {
    return { error: err.message };
  }
}

function setupLight_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(LC_HEADERS).forEach(name => ensureSheet_(ss, name, LC_HEADERS[name]));
  seedConfig_();
  seedActual_();
  migratePlayerAuth_();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const first = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const missing = headers.some((h, i) => String(first[i] || '') !== h);
  if (sheet.getLastRow() === 0 || missing) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedConfig_() {
  const sheet = getSheet_(LC_SHEETS.CONFIG);
  const values = sheet.getDataRange().getValues();
  const keys = new Set(values.slice(1).map(r => String(r[0] || '')));
  Object.keys(LC_DEFAULT_CONFIG).forEach(key => {
    if (!keys.has(key)) sheet.appendRow([key, LC_DEFAULT_CONFIG[key], configNote_(key)]);
  });
}

function seedActual_() {
  const sheet = getSheet_(LC_SHEETS.ACTUAL);
  if (sheet.getLastRow() < 2) sheet.appendRow(['actual', '{}', nowIso_()]);
}

function seedSquadsExample_() {
  const sheet = getSheet_(LC_SHEETS.SQUADS);
  if (sheet.getLastRow() > 1) return;
  sheet.appendRow(['Argentina', 'Ejemplo: importar planteles oficiales cuando FIFA los publique', '', '']);
}

function configNote_(key) {
  const notes = {
    ADMIN_PIN: 'Cambiar antes de compartir. Permite editar resultados reales, settings y planteles.',
    PREDICTION_DEADLINE: 'Fecha/hora límite de carga. Los pronósticos ajenos se revelan recién después de esta fecha. Ejemplo: 2026-06-12T00:00:00-03:00.',
    REVEAL_PREDICTIONS: 'TRUE permite liberar pronósticos después del deadline. FALSE los mantiene ocultos para usuarios comunes.',
    OFFICIAL_RESULTS_URL: 'Link que abre el boton Fuente oficial FIFA.',
    OFFICIAL_RESULTS_SYNC_URL: 'URL JSON/API para sincronizar resultados. No usar la URL visual si no devuelve JSON.',
    OFFICIAL_RESULTS_API_KEY: 'API key opcional. Se oculta en getBootstrap para usuarios comunes.',
    OFFICIAL_SQUADS_SYNC_URL: 'URL JSON/API para importar planteles oficiales cuando esten publicados.',
    OFFICIAL_SQUADS_API_KEY: 'API key opcional para planteles. Se oculta en getBootstrap para usuarios comunes.',
    OFFICIAL_SYNC_AUTO: 'TRUE activa trigger programado de resultados si el admin lo instala.',
    OFFICIAL_SYNC_MINUTES: 'Intervalo del trigger: 1, 5, 10, 15 o 30 minutos.',
    POINT_RESULT_EXACT: 'Puntos por acertar resultado exacto en cualquier partido.',
    POINT_RESULT_OUTCOME: 'Puntos por acertar ganador correcto en cualquier partido.',
    POINT_THIRD_PLACE: 'Puntos por cada equipo pronosticado en el partido de tercer y cuarto puesto.'
  };
  return notes[key] || '';
}


function publicConfig_(config) {
  const cfg = Object.assign({}, config || {});
  delete cfg.ADMIN_PIN;
  delete cfg.OFFICIAL_RESULTS_SYNC_URL;
  delete cfg.OFFICIAL_RESULTS_API_KEY;
  delete cfg.OFFICIAL_SQUADS_SYNC_URL;
  delete cfg.OFFICIAL_SQUADS_API_KEY;
  return cfg;
}

function privateConfigForAdmin_(config) {
  const cfg = Object.assign({}, config || getConfig_());
  delete cfg.ADMIN_PIN;
  return cfg;
}

function officialSyncStatus_(config) {
  config = config || getConfig_();
  return {
    resultsUrlConfigured: !!String(config.OFFICIAL_RESULTS_SYNC_URL || '').trim(),
    squadsUrlConfigured: !!String(config.OFFICIAL_SQUADS_SYNC_URL || '').trim(),
    auto: String(config.OFFICIAL_SYNC_AUTO || '').toUpperCase() === 'TRUE',
    minutes: String(config.OFFICIAL_SYNC_MINUTES || '15'),
    lastResultsSyncAt: config.OFFICIAL_LAST_SYNC_AT || '',
    lastResultsSyncStatus: config.OFFICIAL_LAST_SYNC_STATUS || '',
    lastSquadsSyncAt: config.OFFICIAL_SQUADS_LAST_SYNC_AT || '',
    lastSquadsSyncStatus: config.OFFICIAL_SQUADS_LAST_SYNC_STATUS || ''
  };
}

function getConfig_() {
  const sheet = getSheet_(LC_SHEETS.CONFIG);
  const rows = sheet.getDataRange().getValues();
  const cfg = Object.assign({}, LC_DEFAULT_CONFIG);
  for (let r = 1; r < rows.length; r++) {
    const key = String(rows[r][0] || '').trim();
    if (key) cfg[key] = String(rows[r][1] == null ? '' : rows[r][1]);
  }
  if (!String(cfg.PREDICTION_DEADLINE || '').trim()) cfg.PREDICTION_DEADLINE = LC_DEFAULT_CONFIG.PREDICTION_DEADLINE;
  return cfg;
}

function setConfigValue_(key, value) {
  const sheet = getSheet_(LC_SHEETS.CONFIG);
  const rows = sheet.getDataRange().getValues();
  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][0]) === key) {
      sheet.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, configNote_(key)]);
}

function getPlayers_() {
  return objectsFromSheet_(LC_SHEETS.PLAYERS);
}

function getPredictionRow_(playerId) {
  const rows = objectsFromSheet_(LC_SHEETS.PREDICTIONS);
  return rows.find(r => String(r.PlayerId) === String(playerId)) || null;
}

function getSubmittedPredictions_() {
  const out = {};
  const submittedByPlayer = {};
  getPlayers_().forEach(p => {
    if (p.SubmittedAt) submittedByPlayer[String(p.PlayerId)] = true;
  });
  const rows = objectsFromSheet_(LC_SHEETS.PREDICTIONS);
  rows.forEach(r => {
    if (r.SubmittedJson && submittedByPlayer[String(r.PlayerId)]) out[r.PlayerId] = parseJson_(r.SubmittedJson, {});
  });
  return out;
}

function autoSubmitSavedDraftsAfterDeadline_() {
  const cfg = getConfig_();
  if (!deadlinePassed_(cfg)) return { submitted: 0, skippedEmpty: 0 };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const playersById = {};
    getPlayers_().forEach(p => playersById[String(p.PlayerId)] = p);
    const submittedAt = nowIso_();
    let submitted = 0;
    let skippedEmpty = 0;

    objectsFromSheet_(LC_SHEETS.PREDICTIONS).forEach(row => {
      const playerId = String(row.PlayerId || '');
      const player = playersById[playerId];
      if (!player || player.SubmittedAt) return;

      const draft = parseJson_(row.DraftJson, {});
      if (!predictionHasSavedContent_(draft)) {
        skippedEmpty++;
        return;
      }

      upsertPrediction_(playerId, draft, true);
      updatePlayerLastSaved_(playerId, true, submittedAt);
      log_('AUTO_SUBMIT_DEADLINE', playerId, 'Borrador guardado enviado automaticamente al cierre');
      submitted++;
    });

    if (submitted) clearComputedCaches_();
    return { submitted, skippedEmpty };
  } finally {
    lock.releaseLock();
  }
}

function predictionHasSavedContent_(entry) {
  entry = entry || {};
  const groupScores = entry.groupScores || {};
  if (Object.keys(groupScores).some(k => scoreHasAnyValue_(groupScores[k]))) return true;

  const knockoutScores = entry.knockoutScores || {};
  if (Object.keys(knockoutScores).some(k => scoreHasAnyValue_(knockoutScores[k]))) return true;

  const winners = entry.knockoutWinners || {};
  if (Object.keys(winners).some(k => String(winners[k] || '').trim())) return true;

  const overrides = entry.r32Overrides || {};
  if (Object.keys(overrides).some(k => {
    const row = overrides[k] || {};
    return String(row.home || '').trim() || String(row.away || '').trim();
  })) return true;

  const awards = entry.awards || {};
  return Object.keys(awards).some(k => String(awards[k] || '').trim());
}

function scoreHasAnyValue_(score) {
  if (!score) return false;
  return score.homeGoals !== '' && score.homeGoals != null || score.awayGoals !== '' && score.awayGoals != null;
}

function getPublicSubmittedPredictions_(actual, config) {
  // Compatibilidad con versiones anteriores: ahora se revela todo después del deadline.
  // Se mantiene la función por si algún código viejo la llama.
  return predictionsRevealOpen_(config || getConfig_()) ? getSubmittedPredictions_() : {};
}

function normalizeDeadlineInput_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]+)(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(Z|[+-]\d{2}:?\d{2})?$/);
  if (!m) return raw;
  let tz = m[7] || '-03:00';
  if (/^[+-]\d{4}$/.test(tz)) tz = tz.slice(0, 3) + ':' + tz.slice(3);
  return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0') + 'T' + String(m[4]).padStart(2, '0') + ':' + m[5] + ':' + (m[6] || '00') + tz;
}

function parseFlexibleDate_(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  let d = new Date(raw);
  if (!isNaN(d.getTime())) return d;
  const normalized = normalizeDeadlineInput_(raw);
  if (normalized && normalized !== raw) {
    d = new Date(normalized);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function predictionsRevealOpen_(config) {
  config = config || getConfig_();
  const reveal = String(config.REVEAL_PREDICTIONS || '').toUpperCase() === 'TRUE';
  if (!reveal) return false;
  const raw = String(config.PREDICTION_DEADLINE || '').trim();
  if (!raw) return true;
  const d = parseFlexibleDate_(raw);
  if (!d) return false;
  return new Date() >= d;
}

function maskPredictionForPublic_(entry, actual) {
  // Ya no se usa para privacidad pública. Se deja como no-op para compatibilidad.
  return entry || {};
}

function scoreCompleteServer_(score) {
  if (!score) return false;
  const hg = score.homeGoals;
  const ag = score.awayGoals;
  return hg !== '' && ag !== '' && hg != null && ag != null && isFinite(Number(hg)) && isFinite(Number(ag));
}

function allGroupActualComplete_(actual) {
  actual = actual || {};
  const gs = actual.groupScores || {};
  // 48 equipos en 12 grupos de 4: 6 partidos por grupo = 72 partidos.
  return Object.keys(gs).filter(k => scoreCompleteServer_(gs[k])).length >= 72;
}

function getAllPredictions_() {
  const out = {};
  const submittedByPlayer = {};
  getPlayers_().forEach(p => {
    if (p.SubmittedAt) submittedByPlayer[String(p.PlayerId)] = true;
  });
  const rows = objectsFromSheet_(LC_SHEETS.PREDICTIONS);
  rows.forEach(r => {
    const submitted = r.SubmittedJson && submittedByPlayer[String(r.PlayerId)] ? parseJson_(r.SubmittedJson, {}) : null;
    out[r.PlayerId] = {
      draft: parseJson_(r.DraftJson, {}),
      submitted: submitted,
      updatedAt: r.UpdatedAt
    };
  });
  return out;
}

function getActual_() {
  const rows = objectsFromSheet_(LC_SHEETS.ACTUAL);
  const row = rows.find(r => String(r.Key) === 'actual');
  return row ? parseJson_(row.Json, {}) : {};
}

function getSquads_() {
  const rows = objectsFromSheet_(LC_SHEETS.SQUADS);
  const out = {};
  rows.forEach(r => {
    const team = String(r.Team || '').trim();
    const player = String(r.Player || '').trim();
    if (!team || !player || player.indexOf('Ejemplo:') === 0) return;
    if (!out[team]) out[team] = [];
    out[team].push({ name: player, position: r.Position || '', number: r.Number || '' });
  });
  Object.keys(out).forEach(team => out[team].sort((a, b) => String(a.name).localeCompare(String(b.name))));
  return out;
}


function migratePlayerAuth_() {
  const sheet = getSheet_(LC_SHEETS.PLAYERS);
  if (sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  if (h.TempPassword == null || h.PasswordHash == null) return;
  for (let r = 1; r < data.length; r++) {
    const hasHash = String(data[r][h.PasswordHash] || '').trim();
    const hasTemp = String(data[r][h.TempPassword] || '').trim();
    if (!hasHash && !hasTemp) {
      sheet.getRange(r + 1, h.TempPassword + 1).setValue(makeTempPassword_());
      if (h.AuthUpdatedAt != null) sheet.getRange(r + 1, h.AuthUpdatedAt + 1).setValue(nowIso_());
    }
  }
}

function findPlayerByLogin_(identifier) {
  const value = String(identifier || '').trim();
  const email = normalizeEmail_(value);
  const name = normalizeName_(value);
  const players = getPlayers_();
  if (email && value.indexOf('@') >= 0) {
    const byEmail = players.find(p => normalizeEmail_(p.Email || '') === email);
    if (byEmail) return byEmail;
  }
  return players.find(p => normalizeName_(p.Name || '') === name) || null;
}

function makePasswordAuth_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  return { salt: salt, hash: hashPassword_(password, salt) };
}

function hashPassword_(password, salt) {
  const raw = String(salt || '') + '|' + String(password || '');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
}

function verifyPassword_(password, salt, hash) {
  return !!hash && hashPassword_(password, salt) === String(hash || '');
}

function makeTempPassword_() {
  return 'LC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function updatePlayerAuthLogin_(playerId) {
  const sheet = getSheet_(LC_SHEETS.PLAYERS);
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  if (h.AuthUpdatedAt == null) return;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][h.PlayerId]) === String(playerId)) {
      sheet.getRange(r + 1, h.AuthUpdatedAt + 1).setValue(nowIso_());
      return;
    }
  }
}

function requirePlayer_(playerId, token) {
  const players = getPlayers_();
  const player = players.find(p => String(p.PlayerId) === String(playerId));
  if (!player) throw new Error('Link de jugador no encontrado.');
  if (String(player.Token) !== String(token)) throw new Error('Token invalido para este jugador.');
  return player;
}

function assertCanEditPlayer_(player) {
  if (player.SubmittedAt) throw new Error('Este pronostico ya fue enviado y esta bloqueado.');
  const cfg = getConfig_();
  const deadline = String(cfg.PREDICTION_DEADLINE || '').trim();
  if (deadline) {
    const deadlineDate = parseFlexibleDate_(deadline);
    if (deadlineDate && new Date().getTime() > deadlineDate.getTime()) {
      throw new Error('La fecha limite ya paso. El pronostico quedo cerrado.');
    }
  }
}

function deadlinePassed_(cfg) {
  cfg = cfg || getConfig_();
  const deadline = String(cfg.PREDICTION_DEADLINE || '').trim();
  if (!deadline) return false;
  const deadlineDate = parseFlexibleDate_(deadline);
  return !!(deadlineDate && new Date().getTime() > deadlineDate.getTime());
}

function requireAdmin_(adminPin) {
  const cfg = getConfig_();
  const expected = String(cfg.ADMIN_PIN || '').trim();
  if (!expected || expected === 'cambiar-este-pin') throw new Error('Primero cambia ADMIN_PIN en la hoja Config.');
  if (String(adminPin || '').trim() !== expected) throw new Error('PIN de admin incorrecto.');
}

function upsertPrediction_(playerId, prediction, submit) {
  const sheet = getSheet_(LC_SHEETS.PREDICTIONS);
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const json = safeStringify_(prediction || {});
  const now = nowIso_();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][h.PlayerId]) === String(playerId)) {
      sheet.getRange(r + 1, h.DraftJson + 1).setValue(json);
      if (submit) sheet.getRange(r + 1, h.SubmittedJson + 1).setValue(json);
      sheet.getRange(r + 1, h.UpdatedAt + 1).setValue(now);
      return;
    }
  }
  sheet.appendRow([playerId, json, submit ? json : '', now]);
}

function clearSubmittedPredictionJson_(playerId) {
  const sheet = getSheet_(LC_SHEETS.PREDICTIONS);
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  if (h.SubmittedJson == null) return;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][h.PlayerId]) === String(playerId)) {
      sheet.getRange(r + 1, h.SubmittedJson + 1).setValue('');
      if (h.UpdatedAt != null) sheet.getRange(r + 1, h.UpdatedAt + 1).setValue(nowIso_());
      return;
    }
  }
}

function updatePlayerLastSaved_(playerId, submitted, submittedAt) {
  const sheet = getSheet_(LC_SHEETS.PLAYERS);
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  const now = nowIso_();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][h.PlayerId]) === String(playerId)) {
      sheet.getRange(r + 1, h.LastSavedAt + 1).setValue(now);
      if (submitted) sheet.getRange(r + 1, h.SubmittedAt + 1).setValue(submittedAt || now);
      return;
    }
  }
}

function deleteRowById_(sheetName, idHeader, id) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  const h = headerMap_(data[0]);
  if (h[idHeader] == null) return;
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][h[idHeader]]) === String(id)) sheet.deleteRow(r + 1);
  }
}

function objectsFromSheet_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] == null ? '' : row[i]);
    return obj;
  });
}

function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h)] = i);
  return map;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Falta la hoja ' + name + '. Ejecuta setup().');
  return sheet;
}

function parseJson_(text, fallback) {
  try {
    if (!text) return fallback;
    return JSON.parse(String(text));
  } catch (err) {
    return fallback;
  }
}

function safeStringify_(obj) {
  const text = JSON.stringify(obj || {});
  if (text.length > 45000) throw new Error('El pronostico es demasiado grande para una celda de Google Sheets.');
  return text;
}

function makeToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function normalizeName_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}


function makeReceiptCode_(playerId, submittedAt, submittedJson) {
  const raw = String(playerId || '') + '|' + String(submittedAt || '') + '|' + String(submittedJson || '');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  const hex = bytes.map(function(b) {
    const n = b < 0 ? b + 256 : b;
    return ('0' + n.toString(16)).slice(-2);
  }).join('').slice(0, 8).toUpperCase();
  return 'LC-' + hex;
}

function getPublicActivity_(limit) {
  limit = Math.max(3, Math.min(30, parseInt(limit || LC_DEFAULT_CONFIG.ACTIVITY_MAX_ITEMS || '12', 10) || 12));
  try {
    const rows = objectsFromSheet_(LC_SHEETS.AUDIT).slice(-limit).reverse();
    const names = {};
    getPlayers_().forEach(function(p) { names[String(p.PlayerId)] = p.Name; });
    return rows.map(function(r) {
      return {
        timestamp: r.Timestamp || '',
        action: r.Action || '',
        playerId: r.PlayerId || '',
        playerName: names[String(r.PlayerId || '')] || '',
        text: publicActivityText_(r.Action || '', names[String(r.PlayerId || '')] || '', r.Info || '')
      };
    }).filter(function(item) { return item.text; });
  } catch (err) {
    return [];
  }
}

function publicActivityText_(action, playerName, info) {
  const name = playerName || 'Un chacal';
  switch (String(action || '')) {
    case 'REGISTER_PLAYER': return name + ' se sumo a la penca.';
    case 'SAVE_DRAFT': return name + ' guardo un borrador.';
    case 'SUBMIT_PREDICTION': return name + ' envio su pronostico definitivo.';
    case 'AUTO_SUBMIT_DEADLINE': return name + ' quedo enviado automaticamente al cierre.';
    case 'UNLOCK_PREDICTION': return 'El admin desbloqueo el pronostico de ' + name + '.';
    case 'UPDATE_ACTUAL': return 'El admin cargo resultados reales.';
    case 'IMPORT_SQUADS': return 'El admin actualizo planteles.';
    case 'SAVE_SETTINGS': return 'El admin ajusto la configuracion.';
    case 'RECALCULATE': return 'El admin recalculo la penca.';
    default: return '';
  }
}

function clearComputedCaches_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('LC_BOOTSTRAP_V18');
  } catch (err) {
    // Best effort.
  }
}

function nowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function log_(action, playerId, info) {
  try {
    getSheet_(LC_SHEETS.AUDIT).appendRow([nowIso_(), action, playerId || '', info || '']);
  } catch (err) {
    // No interrumpir la app por problemas de auditoria.
  }
}
