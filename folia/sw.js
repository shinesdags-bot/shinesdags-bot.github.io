// Folia service worker: network-first, falls back to cache offline. CACHE stamped on deploy.
const CACHE = "folia-20261126223934";
const SHELL = ["./", "index.html", "app.css", "app.js", "sync.js", "word.js", "excel.js", "ppt.js", "ocr.js", "inline.js", "build.js", "vendor/blank.pptx", "vendor/exceljs.min.js", "vendor/hyperformula.full.min.js", "vendor/mammoth.browser.min.js", "vendor/fflate.min.js", "vendor/docx.iife.js", "vendor/docx-math.iife.js", "vendor/docx-html.iife.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "vendor/pdf.min.mjs", "vendor/pdf.worker.min.mjs", "vendor/pdf-lib.min.js"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => Promise.allSettled(SHELL.map(u => c.add(new Request(u, { cache: "reload" }))))));  self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("folia-") && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET") return;
  const same = u.origin === location.origin && u.pathname.startsWith("/folia/");
  const font = /fonts\.(googleapis|gstatic)\.com$/.test(u.hostname);
  if (!same && !font) return;
  const put = res => { if (res.ok || res.type === "opaque") { const c = res.clone(); caches.open(CACHE).then(x => x.put(r, c)); } return res; };
  // big vendor libraries: cache first (the cache name changes every deploy, so they stay fresh)
  if (same && u.pathname.includes("/vendor/")) { e.respondWith(caches.match(r, { ignoreSearch: true }).then(m => m || fetch(r).then(put))); return; }
  e.respondWith(fetch(r, same ? { cache: "no-cache" } : {}).then(put)
    .catch(() => caches.match(r, { ignoreSearch: true }).then(m => m || (same && r.mode === "navigate" ? caches.match("index.html") : Response.error()))));
});
