const CACHE_NAME = "tagaddod-cache-settlement-v49";
const APP_SHELL = ["./", "./index.html", "./styles.css?v=settlement-v9", "./app.js?v=settlement-v9", "./analysis.js?v=share-only-v11", "./manifest.json", "./icon.png", "./icon.svg"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin) return;
  if (url.pathname.includes("/exec")) return;
  const networkPromise = fetch(request, { cache: "no-cache" }).then((response) => {
    if (response && response.ok) {
      const copy = response.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => {});
    }
    return response;
  });
  event.waitUntil(networkPromise.then(() => {}, () => {}));
  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return networkPromise.catch((error) => {
        if (request.mode === "navigate") return caches.match("./index.html");
        throw error;
      });
    })
  );
});