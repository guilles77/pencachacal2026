/* Penca Chacal · service worker (Vercel) */
const CACHE = 'penca-chacal-v26';
const SHELL = [
  '/',
  '/index.html',
  '/manifest.webmanifest',
  '/icons/icon-192-any.png',
  '/icons/icon-512-any.png',
  '/icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(SHELL))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;                 // nunca cachear POST (API)
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;  // no tocar pedidos cross-origin (backend, banderas, etc.)
  if (url.pathname.startsWith('/api/')) return;      // nunca cachear la API

  // Network-first para HTML y data.js (siempre tomar la última versión; offline cae al cache)
  if (req.mode === 'navigate' || url.pathname === '/data.js') {
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((cache) => cache.put(req, copy));
        return res;
      }).catch(() => caches.match(req.mode === 'navigate' ? '/index.html' : req).then((r) => r || caches.match('/')))
    );
    return;
  }

  // Cache-first para estáticos (íconos, manifest, etc.).
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((cache) => cache.put(req, copy));
      return res;
    }))
  );
});
