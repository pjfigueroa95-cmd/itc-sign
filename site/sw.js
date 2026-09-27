// QCMS site mode service worker: keeps the page (one HTML file with everything in it) and its icons so the
// app opens with no signal. A new build has a new cache name; the old one is dropped when it takes over.
const CACHE = 'qcms-site-20260927071117169';
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

// The photo reader (ocr/: onnxruntime-web, PaddleOCR's models, zxing; about 25 MB; site/src/ocrfiles.ts, the
// build writes the name and the list in): kept in their own cache named by their versions, so a new page
// build doesn't fetch them again. Kept copy first. Fetched once in the background when the app opens.
const OCR = "qcms-ocr-ppocrv5m-en-ort1.23.0-zx3.1.4";
const OCR_FILES = ["./ocr/ort.wasm.min.mjs","./ocr/ort-wasm-simd-threaded.mjs","./ocr/ort-wasm-simd-threaded.wasm","./ocr/det.onnx","./ocr/rec-en.onnx","./ocr/rec-en-dict.json","./ocr/zxing_reader.wasm"];
const isOcr = (url) => url.origin === location.origin && url.pathname.includes('/ocr/');

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => (k.startsWith('qcms-site-') && k !== CACHE) || (k.startsWith('qcms-ocr-') && k !== OCR) || k === 'qcms-cdn').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

/**
 * The page served cross-origin isolated (Cross-Origin-Opener-Policy and -Embedder-Policy), so the photo
 * reader can run on several threads (SharedArrayBuffer); GitHub Pages can't send the headers itself. The page
 * loads nothing from other sites, so nothing is blocked. Takes effect from the first reload the worker serves.
 */
function isolated(res) {
  if (!res || res.status === 0 || res.type === 'opaque') return res;
  const headers = new Headers(res.headers);
  headers.set('Cross-Origin-Opener-Policy', 'same-origin');
  headers.set('Cross-Origin-Embedder-Policy', 'require-corp');
  headers.set('Cross-Origin-Resource-Policy', 'same-origin');
  return new Response(res.body, { status: res.status, statusText: res.statusText, headers });
}

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || !isOcr(new URL(req.url))) return;
  e.respondWith(
    caches.open(OCR).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req).then((res) => {
      if (res.ok) c.put(req, res.clone());
      return res;
    }))).then(isolated),
  );
});

/** Fetches the photo reader's files not kept yet (in the background; a failure just means next time). */
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
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || caches.match('./index.html'))).then(isolated),
  );
});
