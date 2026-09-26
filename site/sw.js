// QCMS site mode service worker: keeps the page (one HTML file with everything in it) and its icons so the
// app opens with no signal. A new build has a new cache name; the old one is dropped when it takes over.
const CACHE = 'qcms-site-20260926054559309';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => (k.startsWith('qcms-site-') && k !== CACHE) || (k.startsWith('qcms-ocr-') && k !== OCR) || k === 'qcms-cdn').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// Photo to text (tesseract/: Tesseract.js, its engine builds and the English data, about 14 MB): kept in
// their own cache named by the Tesseract version, so a new page build doesn't fetch them again. Kept copy
// first. All fetched once in the background after the page opens with a signal (the three engine builds
// too: which one the phone uses is only known inside Tesseract's worker).
const OCR = 'qcms-ocr-7.0.0';
const OCR_FILES = ['./tesseract/tesseract.min.js', './tesseract/worker.min.js', './tesseract/lang/eng.traineddata.gz',
  './tesseract/tesseract-core-lstm.wasm.js', './tesseract/tesseract-core-simd-lstm.wasm.js', './tesseract/tesseract-core-relaxedsimd-lstm.wasm.js'];
const isOcr = (url) => url.origin === location.origin && url.pathname.includes('/tesseract/');

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || !isOcr(new URL(req.url))) return;
  e.respondWith(
    caches.open(OCR).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req).then((res) => {
      if (res.ok) c.put(req, res.clone());
      return res;
    }))),
  );
});

/** Fetches the photo-to-text files not kept yet (in the background; a failure just means next time). */
async function warmOcr() {
  const c = await caches.open(OCR);
  for (const f of OCR_FILES) {
    const url = new URL(f, self.registration.scope).href;
    if (await c.match(url)) continue;
    try {
      const res = await fetch(url);
      if (res.ok) await c.put(url, res);
    } catch {
      return;
    }
  }
}
self.addEventListener('message', (e) => {
  if (e.data === 'warm-ocr') e.waitUntil(warmOcr());
});

// The page: network first (so an update shows when there is signal), the kept copy when there is none.
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin || isOcr(new URL(req.url))) return;
  e.respondWith(
    fetch(req).then((res) => {
      if (res.ok) caches.open(CACHE).then((c) => c.put(req, res.clone()));
      return res.clone();
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || caches.match('./index.html'))),
  );
});
