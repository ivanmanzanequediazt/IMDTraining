// Service worker de IMDTraining
// 1) Recibe los archivos que se comparten a la app desde otras apps (WhatsApp, Archivos…) y los deja listos para importar.
// 2) Guarda la app para que se abra sin conexión (primero intenta la red, así siempre carga la última versión).
const APP_CACHE = 'imdt-app-v1';
const SHARE_CACHE = 'imdt-share';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);

  // Archivos compartidos a la app: POST a ./share-target
  if (req.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      let n = 0;
      try {
        const form = await req.formData();
        const cache = await caches.open(SHARE_CACHE);
        for (const f of form.getAll('files')) {
          if (!(f instanceof File)) continue;
          const key = new URL(`shared/${Date.now()}-${n++}`, self.registration.scope).href;
          await cache.put(key, new Response(f, { headers: { 'x-name': encodeURIComponent(f.name || 'datos.json') } }));
        }
      } catch (err) { /* si falla, la app avisará de que no ha llegado ningún archivo */ }
      return Response.redirect(new URL(`./?shared=${n}`, self.registration.scope).href, 303);
    })());
    return;
  }

  // App sin conexión: solo peticiones GET del propio sitio
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith((async () => {
    try {
      const res = await fetch(req);
      if (res.ok && (req.mode === 'navigate' || /\.(png|json|js)$/.test(url.pathname))) {
        const cache = await caches.open(APP_CACHE);
        cache.put(req.mode === 'navigate' ? new URL('./', self.registration.scope).href : req, res.clone());
      }
      return res;
    } catch (err) {
      const cache = await caches.open(APP_CACHE);
      return (await cache.match(req.mode === 'navigate' ? new URL('./', self.registration.scope).href : req)) || Response.error();
    }
  })());
});
