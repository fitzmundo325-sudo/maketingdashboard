const CACHE_NAME = 'hyg-marketing-pwa-v1';
const STATIC_ASSETS = [
  '/static/icons/icon-192.png',
  '/static/icons/icon-512.png',
  '/static/icons/maskable-512.png',
  '/static/manifest.webmanifest'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  // Navigation requests: network-first with offline cache fallback
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request).then((response) => {
        return response;
      }).catch(() => {
        return caches.match(request).then((cached) => {
          return cached || new Response('HYG Marketing is offline. Please reconnect and reload.', {
            headers: { 'Content-Type': 'text/plain; charset=utf-8' },
            status: 503
          });
        });
      })
    );
    return;
  }

  // API requests: network-first with cache fallback
  if (url.pathname.startsWith('/apis/')) {
    event.respondWith(
      fetch(request).then((response) => {
        if (response && response.status === 200) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(request, copy);
          });
        }
        return response;
      }).catch(() => {
        return caches.match(request).then((cached) => {
          return cached || new Response('{"error": "offline"}', {
            headers: { 'Content-Type': 'application/json' },
            status: 503
          });
        });
      })
    );
    return;
  }

  // Static assets: cache-first
  if (url.pathname.startsWith('/static/')) {
    event.respondWith(
      caches.match(request).then((cached) => cached || fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      }))
    );
  }
});
