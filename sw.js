// Service worker: la app abre sin conexión y se actualiza sola cuando hay internet.
const C = 'bodega-v1';
const BASE = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './favicon-48.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(BASE))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== C).map(x => caches.delete(x)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET') return;
  const mine = u.origin === location.origin, cdn = u.hostname === 'cdnjs.cloudflare.com';
  if (!mine && !cdn) return; // Apps Script y WhatsApp siempre van directo a internet
  e.respondWith(fetch(r).then(res => {
    const copy = res.clone(); caches.open(C).then(c => c.put(r, copy)); return res;
  }).catch(() => caches.match(r).then(m => m || caches.match('./index.html'))));
});
