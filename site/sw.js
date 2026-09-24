// QCMS site mode service worker: keeps the page (one HTML file with everything in it) and its icons so the
// app opens with no signal. A new build has a new cache name; the old one is dropped when it takes over.
const CACHE = 'qcms-site-20260924202516444';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k.startsWith('qcms-site-') && k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// Photo to text (Tesseract.js, its engine and the English data) from jsDelivr: versioned files, so the kept
// copy is used first and kept across builds (the reader then works with no signal after its first use).
const CDN = 'qcms-cdn';
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).host !== 'cdn.jsdelivr.net') return;
  e.respondWith(
    caches.open(CDN).then((c) => c.match(req).then((hit) => hit || fetch(req).then((res) => {
      if (res.ok || res.type === 'opaque') c.put(req, res.clone());
      return res;
    }))),
  );
});

// The page: network first (so an update shows when there is signal), the kept copy when there is none.
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(
    fetch(req).then((res) => {
      if (res.ok) caches.open(CACHE).then((c) => c.put(req, res.clone()));
      return res.clone();
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || caches.match('./index.html'))),
  );
});
