// QCMS site mode service worker: keeps the page (one HTML file with everything in it) and its icons so the
// app opens with no signal, and makes updates behave like a published app's (site/src/update.ts): the app
// always opens from the copy kept here; a new version (a new sw.js, which the browser checks for in the
// background) is fetched and kept, then waits until the user taps "Update ready" (the page sends 'update-now').
// It never switches by itself. An update never touches the app's data (IndexedDB): only its own caches.
//
// Two copies of the app run on the same site (site/src/channel.ts): the live one (…/site/) and the test one
// (…/site-test/), from the same build. Cache Storage is shared by the whole site, so each copy names its caches
// apart and only ever deletes its own: the live copy's are qcms-site-<build> and qcms-ocr-<version> (the names
// it always had), the test copy's qcms-site-test-<build> and qcms-ocr-test-<version>.
const BUILD = '20261001012620009';
const TEST = /\/site-test\/$/.test(new URL(self.registration.scope).pathname);
const CACHE = `qcms-site-${TEST ? 'test-' : ''}${BUILD}`;
const FILES = ['./', './index.html', './manifest.webmanifest', './manifest-test.webmanifest', './icon-192.png', './icon-512.png'];

// The photo reader (ocr/: onnxruntime-web, PaddleOCR's models, zxing; about 25 MB; site/src/ocrfiles.ts, the
// build writes the version and the list in): kept in their own cache named by their version, so a new page
// build doesn't fetch them again. Kept copy first. Fetched once in the background when the app opens.
const OCR_VERSION = "ppocrv5m-en-ort1.23.0-zx3.1.4";
const OCR = `qcms-ocr-${TEST ? 'test-' : ''}${OCR_VERSION}`;
const OCR_FILES = ["./ocr/ort.wasm.min.mjs","./ocr/ort-wasm-simd-threaded.mjs","./ocr/ort-wasm-simd-threaded.wasm","./ocr/det.onnx","./ocr/rec-en.onnx","./ocr/rec-en-dict.json","./ocr/zxing_reader.wasm"];
const isOcr = (url) => url.origin === location.origin && url.pathname.includes('/ocr/');

/** This copy's own page caches, and its own photo reader caches: the only ones it ever deletes. */
const ownPage = (k) => (TEST ? /^qcms-site-test-\d+$/ : /^qcms-site-\d+$/).test(k);
const ownOcr = (k) => (TEST ? k.startsWith('qcms-ocr-test-') : k.startsWith('qcms-ocr-') && !k.startsWith('qcms-ocr-test-'));

self.addEventListener('install', (e) => {
  // fetched fresh (not from the browser's HTTP cache), kept, then waiting: no skipWaiting here
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))));
});

self.addEventListener('activate', (e) => {
  // this copy's older page caches (and the old CDN cache the live copy once had); the photo reader's older
  // cache goes only once the new one is complete (warmOcr), so the reader keeps working with no signal
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => (ownPage(k) && k !== CACHE) || (!TEST && k === 'qcms-cdn')).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('message', (e) => {
  // the tap on "Update ready": this waiting version takes over (the page opens again from it)
  if (e.data === 'update-now') self.skipWaiting();
  else if (e.data === 'warm-ocr') e.waitUntil(warmOcr());
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

/** Fetches the photo reader's files not kept yet (in the background; a failure just means next time); once all are kept, this copy's older reader caches go. */
async function warmOcr() {
  const c = await caches.open(OCR);
  for (const f of OCR_FILES) {
    const url = new URL(f, self.registration.scope).href;
    if (await c.match(url)) continue;
    try {
      const res = await fetch(url);
      if (!res.ok) return;
      await c.put(url, res);
    } catch {
      return;
    }
  }
  const keys = await caches.keys();
  await Promise.all(keys.filter((k) => ownOcr(k) && k !== OCR).map((k) => caches.delete(k)));
}

// The page: the copy kept with this version, always (so the app never changes under the user); the network
// only for something not kept (then kept), and the kept page when there is no signal.
self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin || isOcr(url) || !req.url.startsWith(self.registration.scope)) return;
  e.respondWith(
    caches.open(CACHE).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req).then((res) => {
      if (res.ok) c.put(req, res.clone());
      return res;
    }).catch(() => c.match('./index.html')))).then(isolated),
  );
});
