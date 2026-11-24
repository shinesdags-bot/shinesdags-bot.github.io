// Folia service worker: network-first, falls back to cache offline. CACHE stamped on deploy.
const CACHE = "folia-20261124155322";
const SHELL = ["./", "index.html", "app.css", "app.js", "sync.js", "word.js", "vendor/mammoth.browser.min.js", "vendor/fflate.min.js", "vendor/docx.iife.js", "vendor/docx-math.iife.js", "vendor/docx-html.iife.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "vendor/pdf.min.mjs", "vendor/pdf.worker.min.mjs", "vendor/pdf-lib.min.js"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" })))).catch(() => {})); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("folia-") && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET") return;
  const same = u.origin === location.origin && u.pathname.startsWith("/folia/");
  const font = /fonts\.(googleapis|gstatic)\.com$/.test(u.hostname);
  if (!same && !font) return;
  e.respondWith(fetch(r, same ? { cache: "no-cache" } : {}).then(res => { if (res.ok || res.type === "opaque") { const c = res.clone(); caches.open(CACHE).then(x => x.put(r, c)); } return res; })
    .catch(() => caches.match(r, { ignoreSearch: true }).then(m => m || (same ? caches.match("index.html") : undefined))));
});
