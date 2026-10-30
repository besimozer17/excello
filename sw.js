// Excello çevrimdışı önbellek
const V = 'excello-v7';
const FILES = ['./', './index.html', './manifest.webmanifest', './css/app.css',
  './js/formula.js', './js/model.js', './js/ui.js', './js/chart.js', './js/io.js', './js/wizard.js', './css/modern.css', './js/modern.js',
  './vendor/exceljs.min.js', './vendor/xlsx.full.min.js', './icons/icon-180.png', './icons/icon-192.png', './icons/icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(r => r || fetch(e.request).then(res => { const cp = res.clone(); caches.open(V).then(c => c.put(e.request, cp)); return res; }).catch(() => caches.match('./index.html'))));
});
