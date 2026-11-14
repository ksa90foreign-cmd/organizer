/* Органайзер — service worker: работа без интернета.
   При обновлении index.html увеличьте номер версии ниже. */
const VERSION = 'organizer-v6';
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
const NAV_TIMEOUT = 4000; // сколько ждать сеть, прежде чем открыть сохранённую версию

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

function withTimeout(promise, ms) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('timeout')), ms);
    promise.then(v => { clearTimeout(t); resolve(v); }, e => { clearTimeout(t); reject(e); });
  });
}

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Страница: сначала сеть (чтобы получать обновления), но не дольше NAV_TIMEOUT —
  // иначе открываем сохранённую версию, а свежая подтянется в фоне
  if (req.mode === 'navigate') {
    const network = fetch(req).then(res => {
      if (res && res.ok) {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put('index.html', copy));
      }
      return res;
    });
    event.respondWith(
      withTimeout(network, NAV_TIMEOUT)
        .catch(() => caches.match('index.html')
          .then(r => r || caches.match('./'))
          .then(r => r || network))
    );
    return;
  }

  // Свои файлы и шрифты: из кэша сразу, в фоне обновляем
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
  // Всё остальное (Google Apps Script, курсы валют) идёт напрямую в сеть
});
