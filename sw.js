// Neska Music service worker. Network-first for the app shell, revalidating with
// the server every time so a new deploy shows up on the next open.
// Never caches songs, Supabase, fonts or lyrics. CACHE is stamped by deploy.sh.
const CACHE = "neska-shell-20261124075214";
const SHELL = ["./", "index.html", "hark.css", "hark-hero.css", "hark.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "neska-emblem.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" })))).catch(() => {})); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(fetch(r, { cache: "no-cache" }).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(r, c)); } return res; })
    .catch(() => caches.match(r, { ignoreSearch: true }).then(m => m || caches.match("index.html"))));
});
