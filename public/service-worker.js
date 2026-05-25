/* Nourish PWA – Service Worker
   Strategy: Cache-first for app shell, network-first for API calls */

const CACHE_NAME = "nourish-v1";
const STATIC_ASSETS = [
  "/",
  "/index.html",
  "/static/js/main.chunk.js",
  "/static/js/bundle.js",
  "/manifest.json"
];

// Install: cache app shell
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(STATIC_ASSETS).catch(() => {
        // Non-fatal if individual assets miss during install
      });
    })
  );
  self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: cache-first for same-origin, network-only for Anthropic API
self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);

  // Never cache Anthropic API calls
  if (url.hostname === "api.anthropic.com") {
    event.respondWith(fetch(event.request));
    return;
  }

  // Cache-first for everything else
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // Only cache valid same-origin GET responses
        if (
          response.ok &&
          event.request.method === "GET" &&
          url.origin === self.location.origin
        ) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => caches.match("/index.html")); // Offline fallback
    })
  );
});
