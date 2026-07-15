const CACHE_NAME = 'meowney-v1';
self.addEventListener('install', (e) => {
  self.skipWaiting();
});
self.addEventListener('activate', (e) => {
  self.clients.claim();
});
self.addEventListener('fetch', (e) => {
  // Network-first, falls back to cache if offline. Keeps sheet sync always fresh when online.
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
