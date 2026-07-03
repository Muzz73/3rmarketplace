/* 3R Oilfield Equipment — service worker
   Makes the site installable (PWA) and fast/offline for the app shell.
   Strategy: network-first for page loads (so updates always show),
   cache-first for static assets, and NEVER touch the login API or forms. */
const CACHE = '3r-v1';
const CORE = [
  '/', '/index.html', '/manifest.webmanifest',
  '/icon-192.png', '/icon-512.png', '/icon-maskable-512.png',
  '/apple-touch-icon.png', '/logo.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(CORE)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;                     // never cache POSTs (forms, login)
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;           // let fonts, chat, Power BI, etc. pass through
  if (url.pathname.startsWith('/.netlify/')) return;    // never cache Identity/login or functions

  // Page navigations: try network first, fall back to cached shell when offline
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((r) => { const cp = r.clone(); caches.open(CACHE).then((c) => c.put('/index.html', cp)); return r; })
        .catch(() => caches.match('/index.html'))
    );
    return;
  }

  // Same-origin static assets: cache first, then network (and cache it)
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((r) => {
      if (r && r.ok) { const cp = r.clone(); caches.open(CACHE).then((c) => c.put(req, cp)); }
      return r;
    }))
  );
});
