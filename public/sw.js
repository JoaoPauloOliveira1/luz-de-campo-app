const SHELL_CACHE_NAME = 'luz-de-campo-shell-v4';
const TILE_CACHE_NAME = 'luz-de-campo-tiles-v1';
const RUNTIME_CACHE_NAME = 'luz-de-campo-runtime-v1';
const OFFLINE_TILE_URL = '/offline-tile.svg';

const SHELL_ASSETS = [
  '/',
  '/index.html',
  '/manifest.webmanifest',
  OFFLINE_TILE_URL,
];

function isCartoTileRequest(url) {
  return url.hostname.endsWith('basemaps.cartocdn.com') && url.pathname.includes('/rastertiles/');
}

function isRuntimeCacheableRequest(url) {
  return url.hostname === 'nominatim.openstreetmap.org';
}

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  const validCaches = [SHELL_CACHE_NAME, TILE_CACHE_NAME, RUNTIME_CACHE_NAME];
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys
        .filter((key) => !validCaches.includes(key))
        .map((key) => caches.delete(key))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  if (url.origin === self.location.origin) {
    const isNavigationRequest = request.mode === 'navigate'
      || (request.headers.get('accept') || '').includes('text/html');

    if (isNavigationRequest || url.pathname.startsWith('/assets/')) {
      event.respondWith(
        fetch(request)
          .then((response) => {
            if (!response || response.status !== 200) return response;
            const cloned = response.clone();
            caches.open(SHELL_CACHE_NAME).then((cache) => cache.put(request, cloned));
            return response;
          })
          .catch(async () => {
            const cached = await caches.match(request);
            if (cached) return cached;
            return caches.match('/index.html');
          })
      );
      return;
    }

    event.respondWith(
      caches.match(request).then((cached) => {
        if (cached) return cached;
        return fetch(request).then((response) => {
          if (!response || response.status !== 200) return response;
          const cloned = response.clone();
          caches.open(SHELL_CACHE_NAME).then((cache) => cache.put(request, cloned));
          return response;
        });
      })
    );
    return;
  }

  if (isCartoTileRequest(url)) {
    event.respondWith(
      caches.open(TILE_CACHE_NAME).then(async (cache) => {
        const cached = await cache.match(request);
        if (cached) return cached;

        try {
          const response = await fetch(request);
          if (response && (response.status === 200 || response.type === 'opaque')) {
            cache.put(request, response.clone());
          }
          return response;
        } catch {
          return caches.match(OFFLINE_TILE_URL);
        }
      })
    );
    return;
  }

  if (isRuntimeCacheableRequest(url)) {
    event.respondWith(
      caches.open(RUNTIME_CACHE_NAME).then(async (cache) => {
        try {
          const response = await fetch(request);
          if (response && response.status === 200) {
            cache.put(request, response.clone());
          }
          return response;
        } catch {
          const cached = await cache.match(request);
          if (cached) return cached;
          throw new Error('runtime request failed');
        }
      })
    );
  }
});
