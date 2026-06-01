// Service worker INTENTIONALLY EMPTY.
//
// Earlier versions cached app.js + the shell. That caching kept serving stale
// app.js after updates, which led to confusing TDZ errors on inlined constants
// that had been removed from the source. The launcher runs a local FastAPI on
// the same machine, so there is no offline use case — every request hits the
// in-process server anyway.
//
// This stub:
//   1. Unregisters itself on install (so existing installations clean up).
//   2. Deletes any previously cached responses.
//   3. Never intercepts a fetch.

self.addEventListener("install", (event) => {
  // Take over immediately
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    // Wipe every cache this SW (or earlier versions) ever created.
    const keys = await caches.keys();
    await Promise.all(keys.map((k) => caches.delete(k)));
    // Tell controlled pages to reload so they fetch the fresh code.
    const clients = await self.clients.matchAll({ type: "window" });
    for (const c of clients) {
      try { c.navigate(c.url); } catch (_) {}
    }
    // Then unregister ourselves so future loads bypass SW entirely.
    self.registration.unregister();
  })());
});

// Pass every fetch through to the network — never serve from cache.
self.addEventListener("fetch", () => { /* no-op; default browser fetch wins */ });
