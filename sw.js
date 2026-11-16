// House Finances: lets the app install on phones. The data itself always comes live
// from the Google Apps Script app, so only the wrapper page and icons are cached.
const CACHE_NAME = 'finances-20261116-162359';
const ASSETS = ['/house-finances/', '/house-finances/manifest.json', '/house-finances/icon-192.png', '/house-finances/icon-512.png', '/house-finances/icon-full-192.png', '/house-finances/icon-full-512.png', '/house-finances/icon-mono-192.png', '/house-finances/icon-mono-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith('finances-') && k !== CACHE_NAME).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).catch(() => caches.match(e.request)));
});
