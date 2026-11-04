const CACHE_NAME = 'shallow-well-survey-offline-v15';
// Everything the app needs to run with no internet, including the offline Iloilo City map.
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './iloilo-map.js',
  './dashboard.html'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      // Cache files one by one so a single missing file does not stop the rest from caching.
      .then(cache => Promise.all(APP_SHELL.map(url => cache.add(new Request(url, {cache: 'reload'})).catch(() => {}))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  const requestUrl = new URL(event.request.url);
  const isSameOrigin = requestUrl.origin === self.location.origin;

  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put('./index.html', copy)).catch(() => {});
          return response;
        })
        .catch(() => caches.match('./index.html').then(cached => cached || caches.match('./')))
    );
    return;
  }

  if (isSameOrigin) {
    event.respondWith(
      caches.match(event.request, {ignoreSearch: true}).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(response => {
          if (response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy)).catch(() => {});
          }
          return response;
        }).catch(() => new Response('', { status: 504, statusText: 'Offline' }));
      })
    );
    return;
  }

  // Online map tiles and Google Sheets sync go straight to the network.
  // The app falls back to the built-in offline Iloilo map when tiles cannot load.
  event.respondWith(
    fetch(event.request).catch(() => new Response('', { status: 504, statusText: 'Offline' }))
  );
});
