// KIMMIZU Service Worker - Offline Support
const CACHE_NAME = 'kimmizu-v1';
const CACHE_URLS = [
  '/nhanvien.html',
  '/manifest.json',
];

// Install: cache files
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(CACHE_URLS))
  );
  self.skipWaiting();
});

// Activate: clear old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: network first, fallback to cache
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // API calls: network only (không cache)
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ success: false, error: 'offline' }), {
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }

  // App files: network first, fallback cache
  event.respondWith(
    fetch(event.request)
      .then(res => {
        const clone = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});

// Background sync (khi có mạng trở lại)
self.addEventListener('sync', event => {
  if (event.tag === 'sync-offline') {
    event.waitUntil(
      self.clients.matchAll().then(clients => {
        clients.forEach(client => client.postMessage({ type: 'SYNC_OFFLINE' }));
      })
    );
  }
});
