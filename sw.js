// KIMMIZU Service Worker v1.0
// Cache các file tĩnh để app load nhanh hơn khi có mạng yếu
const CACHE_NAME = 'kimmizu-nv-v1';
const STATIC_FILES = [
  './nhanvien.html',
  './icon-192.png',
  './icon-512.png',
  './manifest.json',
];

// Cài đặt: cache file tĩnh
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(STATIC_FILES))
      .then(() => self.skipWaiting())
  );
});

// Activate: xoá cache cũ
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Fetch strategy: Network first, cache fallback
// API calls luôn dùng network (không cache)
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // API calls → không cache, luôn dùng network
  if (url.hostname.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(() =>
      new Response(JSON.stringify({success:false,error:'Offline'}),
        {headers:{'Content-Type':'application/json'}})
    ));
    return;
  }

  // File tĩnh → network first, fallback cache
  e.respondWith(
    fetch(e.request)
      .then(res => {
        // Cập nhật cache khi có network
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
