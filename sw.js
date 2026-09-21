// Bump this on every deploy that changes cached files — it's what makes the
// browser notice this file changed and install a new service worker.
const CACHE_NAME = "cellgrind-v4";
const ASSETS = [
  "./",
  "index.html",
  "style.css",
  "game.js",
  "manifest.webmanifest",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "icons/apple-touch-icon.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

// Network-first, and *actually* bypassing the browser's ordinary HTTP
// cache (not just our own Cache Storage) via { cache: "no-store" } — a
// plain fetch() still honors GitHub Pages' Cache-Control headers and can
// silently return a stale response from the browser's disk cache even
// though this looks "network-first". no-store forces a real round trip.
// Our own cache is only ever read as an offline fallback.
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    fetch(event.request, { cache: "no-store" })
      .then((response) => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
