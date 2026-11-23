const CACHE = 'proj-v9';
const ARQS = ['./', 'index.html', 'manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(k => Promise.all(k.filter(n => n !== CACHE).map(n => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;                      // não intercepta POST
  if (!e.request.url.startsWith(self.location.origin)) return; // não intercepta API nem CDN
  e.respondWith(
    fetch(e.request)
      .then(r => {
        if (r.ok) { const c = r.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); }
        return r;
      })
      .catch(() => caches.match(e.request))
  );
});
