// Minimal cache-first service worker for shell assets.
const CACHE = "sorghum-pwa-v2";
const ASSETS = [
  "/app",
  "/static/app.js",
  "/static/manifest.webmanifest",
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  // Network-first for API; cache-first for the shell.
  if (url.pathname.startsWith("/sync") || url.pathname.startsWith("/nursery")) {
    return; // let it go to network; failures are handled in app.js via outbox
  }
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request).then(resp => {
      const clone = resp.clone();
      caches.open(CACHE).then(c => c.put(e.request, clone));
      return resp;
    }).catch(() => caches.match("/app")))
  );
});
