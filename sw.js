/* Органайзер — service worker: работа без интернета.
   При обновлении index.html увеличьте номер версии ниже. */
const VERSION = 'organizer-v3';
const SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/apple-touch-icon.png',
  'lib/qrcode.js'
];
const CACHE_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Страница: сначала сеть (чтобы получать обновления), без сети — из кэша
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(VERSION).then(c => c.put('index.html', copy));
          return res;
        })
        .catch(() => caches.match('index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // Свои файлы, шрифты и библиотека QR: из кэша сразу, в фоне обновляем
  if (url.origin === self.location.origin || CACHE_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.open(VERSION).then(cache =>
        cache.match(req).then(cached => {
          const network = fetch(req)
            .then(res => {
              if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
              return res;
            })
            .catch(() => cached);
          return cached || network;
        })
      )
    );
  }
  // Всё остальное (в том числе запросы к Google Apps Script) идёт напрямую в сеть
});
