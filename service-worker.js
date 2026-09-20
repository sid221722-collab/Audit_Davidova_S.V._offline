const CACHE_NAME = 'audit-pwa-emergency-v4';
const APP_SHELL = ['./', './index.html', './app-config.js', './manifest.webmanifest', './service-worker.js'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  try {
    if (new URL(req.url).origin !== self.location.origin) return;
  } catch (e) {
    return;
  }

  // HTML и локальные файлы сначала берём из сети, чтобы обновления
  // аварийной версии подхватывались сразу. При отсутствии сети
  // используем локальный кэш.
  event.respondWith(
    fetch(req).then(resp => {
      if (resp && resp.ok) {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(req, copy)).catch(() => {});
      }
      return resp;
    }).catch(() => caches.match(req).then(cached => cached || caches.match('./index.html')))
  );
});
