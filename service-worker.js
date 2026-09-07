const CACHE_NAME = 'ceo-checkin-pwa-v4';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './offline.html',
  './icons/apple-touch-icon.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(APP_SHELL);
    })
  );

  self.skipWaiting();
});


self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(
          keys.map(function (key) {
            if (key !== CACHE_NAME) {
              return caches.delete(key);
            }
          })
        );
      })
      .then(function () {
        return self.clients.claim();
      })
  );
});


self.addEventListener('fetch', function (event) {
  const request = event.request;

  if (request.method !== 'GET') {
    return;
  }

  const url = new URL(request.url);

  // Apps Script 與 Google Sheets 不由 PWA Service Worker 攔截。
  if (url.origin !== self.location.origin) {
    return;
  }

  // HTML 導航採 Network First，
  // 可降低手機一直看到舊版 GOOGLE_SHEET_URL 的情況。
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(
        request,
        {
          cache: 'no-store'
        }
      )
        .then(function (response) {
          if (
            response &&
            response.ok
          ) {
            const copy =
              response.clone();

            caches
              .open(CACHE_NAME)
              .then(function (cache) {
                cache.put(
                  request,
                  copy
                );
              });
          }

          return response;
        })
        .catch(function () {
          return caches
            .match(request)
            .then(function (cached) {
              return (
                cached ||
                caches.match(
                  './offline.html'
                )
              );
            });
        })
    );

    return;
  }

  // 靜態資源採 Cache First。
  event.respondWith(
    caches
      .match(request)
      .then(function (cached) {
        if (cached) {
          return cached;
        }

        return fetch(request)
          .then(function (response) {
            if (
              !response ||
              !response.ok
            ) {
              return response;
            }

            const copy =
              response.clone();

            caches
              .open(CACHE_NAME)
              .then(function (cache) {
                cache.put(
                  request,
                  copy
                );
              });

            return response;
          });
      })
  );
});
