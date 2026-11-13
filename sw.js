/* Mi espacio · funcionamiento sin conexión.
 * Guarda una copia de la propia app (la página y el icono) en el dispositivo.
 * Con internet siempre se pide la versión más nueva; sin internet (o si tarda mucho) se abre la copia.
 * Aquí no se guarda ningún dato tuyo: las notas viajan aparte, entre la app y tu hoja de Google. */
const CACHE = 'miespacio-app-v1';
const BASICO = ['./', 'manifest.webmanifest', 'icon-180.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASICO)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

function conLimite(promesa, ms) {
  return new Promise((ok, mal) => { const t = setTimeout(() => mal(new Error('lento')), ms); promesa.then(v => { clearTimeout(t); ok(v); }, e => { clearTimeout(t); mal(e); }); });
}
async function pagina(req) {
  const cache = await caches.open(CACHE);
  const red = fetch(req, { cache: 'no-cache' }).then(res => { if (res && res.ok) cache.put('./', res.clone()); return res; });
  try {
    const res = await conLimite(red, 4000);
    if (res && res.ok) return res;
    throw new Error('respuesta');
  } catch (err) {
    red.catch(() => {});
    return (await cache.match('./')) || red;
  }
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // lo que va a tu hoja de Google no se toca
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (req.mode === 'navigate') { e.respondWith(pagina(req)); return; }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
    if (res && res.ok) { const copia = res.clone(); caches.open(CACHE).then(c => c.put(req, copia)); }
    return res;
  })));
});
