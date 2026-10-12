/* Dollar Store Oman – makes the catalogue installable and open fast / offline */
const CACHE = "ds-v26";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())
));
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin) return;           // orders, WhatsApp, fonts: untouched
  const put = res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(r, c)); } return res; };
  if (u.pathname.includes("/img/")) {                                         // photos: cache first
    e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(put)));
  } else {                                                                    // page: newest first, cached copy if offline
    e.respondWith(fetch(r).then(put).catch(() => caches.match(r).then(hit => hit || caches.match("./"))));
  }
});
