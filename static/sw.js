// Yve.01 Service Worker — offline support + cache
const CACHE = 'yve01-v1';
const STATIC = [
  '/static/manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js',
];

// Install: cache static assets
self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC).catch(() => {})));
});

// Activate: clean old caches
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

// Fetch: network-first for API, cache-first for static
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  
  // API calls: network only
  if (url.pathname.startsWith('/api/') || url.pathname.startsWith('/fb/') || 
      url.pathname.startsWith('/admin/') || url.pathname === '/') {
    return; // Let browser handle normally
  }
  
  // Static assets: cache-first
  if (url.pathname.startsWith('/static/') || url.hostname.includes('cdn') || 
      url.hostname.includes('fonts')) {
    e.respondWith(
      caches.match(e.request).then(cached => {
        if (cached) return cached;
        return fetch(e.request).then(resp => {
          if (resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE).then(c => c.put(e.request, clone));
          }
          return resp;
        }).catch(() => cached);
      })
    );
  }
});
