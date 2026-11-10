/* Акты АЗС, веб-версия (iPhone): все файлы приложения хранятся в кэше, работает без сети.
   Новая версия (новый sw.js) загружается заранее и включается после согласия пользователя (postMessage "skipWaiting")
   или при следующем запуске. Данные приложения (localStorage) при этом не трогаются. */
const CACHE = "akty-3.24";
const FILES = ["./", "app.js", "changelog.js", "data.js", "fonts/Carlito-400-cyrillic-ext.woff2", "fonts/Carlito-400-cyrillic.woff2", "fonts/Carlito-400-latin-ext.woff2", "fonts/Carlito-400-latin.woff2", "fonts/Carlito-400i-cyrillic-ext.woff2", "fonts/Carlito-400i-cyrillic.woff2", "fonts/Carlito-400i-latin-ext.woff2", "fonts/Carlito-400i-latin.woff2", "fonts/Carlito-700-cyrillic-ext.woff2", "fonts/Carlito-700-cyrillic.woff2", "fonts/Carlito-700-latin-ext.woff2", "fonts/Carlito-700-latin.woff2", "fonts/Carlito-700i-cyrillic-ext.woff2", "fonts/Carlito-700i-cyrillic.woff2", "fonts/Carlito-700i-latin-ext.woff2", "fonts/Carlito-700i-latin.woff2", "fonts/Tinos-400-cyrillic-ext.woff2", "fonts/Tinos-400-cyrillic.woff2", "fonts/Tinos-400-latin-ext.woff2", "fonts/Tinos-400-latin.woff2", "fonts/Tinos-400i-cyrillic-ext.woff2", "fonts/Tinos-400i-cyrillic.woff2", "fonts/Tinos-400i-latin-ext.woff2", "fonts/Tinos-400i-latin.woff2", "fonts/Tinos-700-cyrillic-ext.woff2", "fonts/Tinos-700-cyrillic.woff2", "fonts/Tinos-700-latin-ext.woff2", "fonts/Tinos-700-latin.woff2", "fonts/fonts.css", "icon-180.png", "icon-192.png", "icon-512.png", "index.html", "layouts.js", "manifest.webmanifest", "vendor/html2canvas.min.js", "vendor/jspdf.umd.min.js", "vendor/jszip.min.js", "vendor/tesseract/lang/eng.traineddata.gz", "vendor/tesseract/tesseract-core-lstm.wasm.js", "vendor/tesseract/tesseract-core-simd-lstm.wasm.js", "vendor/tesseract/tesseract.min.js", "vendor/tesseract/worker.min.js", "vendor/zxing.min.js", "version.js", "xlsxwriter.js"];
self.addEventListener("install", e => e.waitUntil(
  caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, {cache: "reload"}))))));
self.addEventListener("activate", e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("akty-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())));
self.addEventListener("message", e => { if (e.data === "skipWaiting") self.skipWaiting(); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin || u.search) return;   // version.js?t=… — всегда с сервера
  e.respondWith(caches.open(CACHE).then(c => c.match(e.request)).then(r => r || fetch(e.request)));
});
