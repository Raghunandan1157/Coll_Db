var CACHE_NAME = 'coll-report-v135';
var NAV_FALLBACKS = ['./employee.html', './index.html'];
var ASSETS = [
  './',
  './index.html',
  './employee.html',
  './admin.html',
  './contacts.html',
  './locator.html',
  './css/styles.css',
  './js/employeeParser.js',
  './js/storage.js',
  './js/auth.js',
  './js/employee.js',
  './js/collection.js',
  './js/analytical.js',
  './js/hourly.js',
  './js/portfolio.js',
  './js/disbursement.js',
  './js/contacts.js',
  './js/admin.js',
  './js/parser.js',
  './js/renderer.js',
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Playfair+Display:wght@700&display=swap'
];

// Install — best-effort cache of core assets (failures don't block install).
self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return Promise.all(ASSETS.map(function (u) {
        return cache.add(u).catch(function () { /* swallow individual failures */ });
      }));
    }).then(function () {
      return self.skipWaiting();
    })
  );
});

// Activate — delete old caches.
self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    }).then(function () {
      return self.clients.claim();
    })
  );
});

// Fetch — only intercept navigations (so JS/CSS go straight through the browser's
// native HTTP cache and we can never break script loading with a failed respondWith).
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;

  var url = req.url;
  if (url.indexOf('/api/') !== -1) return;

  // Only take control of navigations. Everything else falls through to the browser.
  if (req.mode !== 'navigate') return;

  e.respondWith(
    fetch(req).then(function (res) {
      var clone = res.clone();
      caches.open(CACHE_NAME).then(function (cache) { cache.put(req, clone); });
      return res;
    }).catch(function () {
      return caches.match(req).then(function (m) {
        if (m) return m;
        for (var i = 0; i < NAV_FALLBACKS.length; i++) {
          var hit = caches.match(NAV_FALLBACKS[i]);
          if (hit) return hit;
        }
        return new Response('<h1>Offline</h1>', { headers: { 'Content-Type': 'text/html' } });
      });
    })
  );
});
