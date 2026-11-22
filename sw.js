// Neska Music: minimal service worker so browsers can install the app.
// Network-first for the app shell; never caches songs or Supabase calls.
const CACHE = "neska-shell-v2";
const SHELL = ["./", "index.html", "hark.css", "hark-hero.css", "hark.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "neska-emblem.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {})); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin) return; // songs, Supabase, fonts, lyrics go straight to the network
  e.respondWith(fetch(r).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(r, c)); } return res; }).catch(() => caches.match(r).then(m => m || caches.match("index.html"))));
});
