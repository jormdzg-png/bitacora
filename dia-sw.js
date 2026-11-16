const CACHE_NAME = 'dia-v3';
const ASSETS = ['./dia.html'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});
// Stale-while-revalidate: abre al instante desde caché y se actualiza en segundo plano.
// Así las nuevas versiones llegan solas, sin tener que cambiar el nombre del caché.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.open(CACHE_NAME).then((cache) =>
      cache.match(event.request).then((cached) => {
        const network = fetch(event.request).then((res) => {
          if (res && res.ok && new URL(event.request.url).origin === self.location.origin) cache.put(event.request, res.clone());
          return res;
        }).catch(() => cached);
        return cached || network;
      })
    )
  );
});
