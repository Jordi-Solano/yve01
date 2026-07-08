/*  Yve.01 — Service Worker
 *  · Cache-first para estáticos (CSS, JS, iconos, fuentes)
 *  · Network-first para navegaciones, con página offline propia
 *  · Network-only para APIs y datos financieros (nunca se cachean)
 *  · Auto-actualización con banner (skipWaiting bajo demanda del cliente)
 *  · Notificaciones push + click → abre el tab correcto
 */
const SW_VERSION    = 'yve01-v2.0.0';
const STATIC_CACHE  = SW_VERSION + '-static';
const RUNTIME_CACHE = SW_VERSION + '-runtime';
const OFFLINE_URL   = '/static/offline.html';

const PRECACHE = [
  OFFLINE_URL,
  '/static/yve.css',
  '/static/manifest.json',
  '/static/icons/favicon.svg',
  '/static/icons/yve-logo-192.png',
  '/static/icons/yve-logo-512.png',
  '/static/icons/yve-logo-maskable-192.png',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// ── Install: precache del app shell (tolerante a fallos individuales) ──
self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(STATIC_CACHE);
    await Promise.all(PRECACHE.map(async url => {
      try {
        const resp = await fetch(new Request(url, { cache: 'reload' }));
        if (resp && (resp.ok || resp.type === 'opaque')) await cache.put(url, resp);
      } catch (e) { /* nunca bloquea la instalación */ }
    }));
    // Sin skipWaiting automático: el cliente decide cuándo activar (banner "Actualización disponible")
  })());
});

// ── Activate: limpia caches viejas + toma control ──
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => !k.startsWith(SW_VERSION)).map(k => caches.delete(k)));
    if (self.registration.navigationPreload) {
      try { await self.registration.navigationPreload.enable(); } catch (e) {}
    }
    await self.clients.claim();
  })());
});

// ── Mensajes desde la página (activar nueva versión) ──
self.addEventListener('message', event => {
  const d = event.data;
  if (d === 'SKIP_WAITING' || (d && d.type === 'SKIP_WAITING')) self.skipWaiting();
});

const isStatic = u => u.pathname.startsWith('/static/');
const isCdn = u =>
  u.hostname.includes('cdn.jsdelivr.net') ||
  u.hostname.includes('fonts.googleapis.com') ||
  u.hostname.includes('fonts.gstatic.com') ||
  u.hostname.includes('cdnjs.cloudflare.com');
const isApi = u =>
  u.pathname.startsWith('/api/') || u.pathname.startsWith('/fb/') ||
  u.pathname.startsWith('/admin/') || u.pathname.startsWith('/oracle/') ||
  u.pathname.startsWith('/aprobaciones');

// ── Fetch ──
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return;

  // 1) Navegaciones → network-first, fallback offline
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const preload = await event.preloadResponse;
        if (preload) return preload;
        return await fetch(req);
      } catch (e) {
        const cache = await caches.open(STATIC_CACHE);
        return (await cache.match(OFFLINE_URL)) ||
          new Response('Sin conexión', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
      }
    })());
    return;
  }

  // 2) APIs / datos financieros → network-only (jamás cachear), error JSON controlado
  if (url.origin === self.location.origin && isApi(url)) {
    event.respondWith((async () => {
      try { return await fetch(req); }
      catch (e) {
        return new Response(JSON.stringify({ ok: false, offline: true, error: 'Sin conexión' }),
          { status: 503, headers: { 'Content-Type': 'application/json' } });
      }
    })());
    return;
  }

  // 3) Estáticos propios + CDN/fuentes → cache-first + revalidación en segundo plano
  if ((url.origin === self.location.origin && isStatic(url)) || isCdn(url)) {
    event.respondWith((async () => {
      const cache = await caches.open(isCdn(url) ? RUNTIME_CACHE : STATIC_CACHE);
      const cached = await cache.match(req);
      const network = fetch(req).then(resp => {
        if (resp && (resp.ok || resp.type === 'opaque')) cache.put(req, resp.clone());
        return resp;
      }).catch(() => null);
      return cached || (await network) || new Response('', { status: 504 });
    })());
    return;
  }

  // 4) Resto mismo origen → network-first con fallback a cache
  if (url.origin === self.location.origin) {
    event.respondWith((async () => {
      try { return await fetch(req); }
      catch (e) { return (await caches.match(req)) || new Response('', { status: 504 }); }
    })());
  }
});

// ── Push: muestra la notificación (funciona con la app cerrada) ──
self.addEventListener('push', event => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; }
  catch (e) { data = { title: 'Yve.01', body: (event.data && event.data.text()) || 'Nueva alerta' }; }
  const title = data.title || 'Yve.01';
  const options = {
    body: data.body || '',
    icon: data.icon || '/static/icons/yve-logo-192.png',
    badge: '/static/icons/yve-logo-96.png',
    tag: data.tag || 'yve-alert',
    renotify: !!data.renotify,
    requireInteraction: data.requireInteraction !== false,
    data: Object.assign({ url: data.url || '/app' }, data.data || {}),
    vibrate: [80, 40, 80],
    timestamp: Date.now()
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

// ── Click en la notificación → enfoca o abre Yve en el tab correcto ──
self.addEventListener('notificationclick', event => {
  event.notification.close();
  const target = (event.notification.data && event.notification.data.url) || '/app';
  event.waitUntil((async () => {
    const all = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of all) {
      try {
        const cu = new URL(c.url);
        if (cu.origin === self.location.origin && 'focus' in c) {
          await c.focus();
          if ('navigate' in c && target) { try { await c.navigate(target); } catch (e) {} }
          return;
        }
      } catch (e) {}
    }
    if (clients.openWindow) return clients.openWindow(target);
  })());
});
