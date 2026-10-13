/* Offline-Speicher der App.
   Mit Internet: immer die neueste Fassung von GitHub holen (und für offline merken).
   Ohne Internet oder bei schwachem Netz (nach 4 Sekunden): gespeicherte Fassung nehmen. */
const VERSION = 'ws-feld-1.2.2';
const FILES = ['./', 'index.html', 'app.js', 'xlsx.js', 'jszip.min.js', 'pdf-lib.min.js', 'pdf.js', 'map.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return; // Kartenkacheln usw. direkt
  e.respondWith((async () => {
    const cache = await caches.open(VERSION);
    const net = fetch(req, { cache: 'no-cache' }).then(r => { if (r.ok) cache.put(req, r.clone()); return r; });
    const timeout = new Promise(res => setTimeout(() => res(null), 4000));
    try {
      const r = await Promise.race([net, timeout]);
      if (r && r.ok) return r;
    } catch (err) { /* offline */ }
    const hit = await cache.match(req, { ignoreSearch: true }) || (req.mode === 'navigate' && await cache.match('index.html'));
    if (hit) return hit;
    try { return await net; } catch (err) { return new Response('Offline', { status: 503 }); }
  })());
});
