const CACHE = "evidencija-shell-v2";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Mrežno-prvo za sve isto-porijeklo GET zahtjeve (aplikacija, JS, CSS, ikone) — ono što uspješno stigne
// se sprema u cache, pa ako sljedeći put nema interneta app ipak može otvoriti (podaci s Supabasea neće doći,
// ali sama aplikacija se učita). Supabase i drugi vanjski pozivi se ne diraju.
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;

  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        return res;
      })
      .catch(() =>
        caches.match(e.request).then((cached) => cached || caches.match("/index.html"))
      )
  );
});
