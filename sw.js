/* 108 Stitches service worker — network-first for the page, cache-first for static assets.
   Bump CACHE to force old caches out on the next visit. */
const CACHE = '108-stitches-v1';
const STATIC = [
  './',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // The board HTML: always try the network first so updates land without clearing data.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put('./', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./').then(r => r || caches.match(req)))
    );
    return;
  }

  const sameOrigin = url.origin === location.origin;
  const cacheableCDN = /gstatic|googleapis|cdnjs|sheetjs|unpkg|jsdelivr/i.test(url.href);
  // Never intercept live data (ESPN relay, Firebase, etc.) — let it hit the network directly.
  if (!sameOrigin && !cacheableCDN) return;

  // Static assets: serve from cache, fall back to network and cache the result.
  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res && res.status === 200) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      }
      return res;
    }))
  );
});
