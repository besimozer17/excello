// Excello çevrimdışı önbellek: internet varsa en yeni sürüm, yoksa cihazdaki kopya
const V = 'excello-v9';
const FILES = ['./', './index.html', './manifest.webmanifest', './css/app.css',
  './js/formula.js', './js/model.js', './js/ui.js', './js/chart.js', './js/io.js', './js/wizard.js', './css/modern.css', './js/modern.js',
  './vendor/exceljs.min.js', './vendor/xlsx.full.min.js', './icons/icon-180.png', './icons/icon-192.png', './icons/icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  const cached = () => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('./index.html'));
  // büyük kütüphaneler değişmez: önce cihazdaki kopya
  if (/\/vendor\//.test(e.request.url)) { e.respondWith(caches.match(e.request).then(r => r || fetch(e.request))); return; }
  // diğerleri: önce internet (en fazla 4 sn), olmazsa cihazdaki kopya
  e.respondWith(new Promise(resolve => {
    let done = false; const fallback = () => { if (!done) { done = true; cached().then(resolve); } };
    const t = setTimeout(fallback, 4000);
    fetch(e.request, { cache: 'no-cache' }).then(res => {
      clearTimeout(t); if (res && res.ok) { const cp = res.clone(); caches.open(V).then(c => c.put(e.request, cp)); }
      if (!done) { done = true; resolve(res && res.ok ? res : cached()); }
    }).catch(() => { clearTimeout(t); fallback(); });
  }));
});
