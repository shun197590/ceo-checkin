const CACHE_NAME = 'ceo-checkin-pwa-v6';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './offline.html',
  './apple-touch-icon.png',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png'
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

  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  // 外部 Apps Script / Google Sheets 不攔截
  if (url.origin !== self.location.origin) return;

  // Scanner 永遠優先取得最新版
  if (url.pathname.includes('/ceo-checkin/scanner/')) {
    event.respondWith(
      fetch(request, { cache: 'no-store' })
        .catch(function () {
          return new Response(
            '掃碼頁目前無法連線，請確認網路後重新啟動。',
            {
              status: 503,
              headers: {
                'Content-Type': 'text/plain; charset=utf-8'
              }
            }
          );
        })
    );
    return;
  }

  // PWA 導航頁：Network First
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request, { cache: 'no-store' })
        .then(function (response) {
          if (response && response.ok) {
            const copy = response.clone();

            caches.open(CACHE_NAME)
              .then(function (cache) {
                cache.put(request, copy);
              });
          }

          return response;
        })
        .catch(function () {
          return caches.match(request)
            .then(function (cached) {
              return cached || caches.match('./offline.html');
            });
        })
    );

    return;
  }

  // 其他靜態檔：Cache First
  event.respondWith(
    caches.match(request)
      .then(function (cached) {
        if (cached) return cached;

        return fetch(request)
          .then(function (response) {
            if (!response || !response.ok) return response;

            const copy = response.clone();

            caches.open(CACHE_NAME)
              .then(function (cache) {
                cache.put(request, copy);
              });

            return response;
          });
      })
  );
});
