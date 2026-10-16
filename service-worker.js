// Hallazgos de Planta · service worker v3
// Primero la red (así siempre ves la versión más nueva) y, sin internet, la copia guardada.
// Nunca guarda llamadas a Google (datos, fotos, fuentes): esas van siempre directo a la red.
const CACHE = 'hallazgos-v3';
const FILES = ['./Hallazgos_Planta.html', './manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== self.location.origin) return;
  e.respondWith(
    fetch(r).then(res => {
      if (res && res.ok) { const copia = res.clone(); caches.open(CACHE).then(c => c.put(r, copia)); }
      return res;
    }).catch(() => caches.match(r).then(m => m || caches.match('./Hallazgos_Planta.html')))
  );
});
