/**
 * VENTO MSL - Offline-First Service Worker
 * Caches all application shell files, Excel parser library, icons, and fonts
 * so the app runs 100% offline on mobile phones after first visit.
 */

const CACHE_NAME = 'vento-msl-offline-v1';

const CORE_ASSETS = [
  './',
  './index.html',
  './index.css',
  './app.js',
  './manifest.json',
  './libs/xlsx.full.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon.svg'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(CORE_ASSETS);
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // Cache-first with background update (Stale-While-Revalidate) for same-origin and Google Fonts
  const isSameOrigin = url.origin === self.location.origin;
  const isGoogleFonts = url.hostname.includes('googleapis.com') || url.hostname.includes('gstatic.com');

  if (isSameOrigin || isGoogleFonts) {
    event.respondWith(
      caches.open(CACHE_NAME).then((cache) => {
        return cache.match(event.request, { ignoreSearch: isSameOrigin }).then((cachedResponse) => {
          const networkFetch = fetch(event.request)
            .then((networkResponse) => {
              if (networkResponse && (networkResponse.status === 200 || networkResponse.type === 'opaque')) {
                cache.put(event.request, networkResponse.clone());
              }
              return networkResponse;
            })
            .catch(() => cachedResponse);

          // Return cached copy immediately for instant offline loading, fallback to network
          return cachedResponse || networkFetch;
        });
      })
    );
  }
});

