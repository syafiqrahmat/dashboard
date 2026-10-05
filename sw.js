const CACHE_NAME = "sw-dashboard-__SW_VERSION__";
const APP_SHELL = [
  "/manifest.json",
  "/sw.svg",
  "/static/icon-192.png",
  "/static/icon-512.png",
  "/static/icon-maskable-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // App shell assets rarely change: cache-first.
  if (APP_SHELL.includes(url.pathname)) {
    event.respondWith(
      caches.match(req).then((cached) => cached || fetch(req))
    );
    return;
  }

  // Everything else (the dashboard itself, filtered/uploaded data) is
  // dynamic and deliberately NOT intercepted: the old cache-on-success /
  // serve-stale-on-failure pair here masked a dead backend completely --
  // the page and its lazy-loaded tabs still rendered from cache while
  // every save/upload failed with a bare "Failed to fetch", which is
  // near-impossible to diagnose from the UI. Letting the browser handle
  // these requests means a down server fails loudly (browser error page)
  // instead of pretending to work.
});
