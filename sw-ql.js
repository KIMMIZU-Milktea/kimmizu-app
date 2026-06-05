const CACHE = 'kimmizu-ql-v1';
const FILES = ['/quanly.html', '/manifest-ql.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (new URL(e.request.url).hostname.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(() =>
      new Response(JSON.stringify({success:false,error:'offline'}), {headers:{'Content-Type':'application/json'}})
    )); return;
  }
  e.respondWith(fetch(e.request).then(r => {
    caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r;
  }).catch(() => caches.match(e.request)));
});
