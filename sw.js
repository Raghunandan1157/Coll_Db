var CACHE_NAME = 'coll-report-v8';
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
  './js/admin.js',
  './js/parser.js',
  './js/renderer.js',
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Playfair+Display:wght@700&display=swap'
];

// Install — cache all core assets
self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(ASSETS);
    }).then(function () {
      return self.skipWaiting();
    })
  );
});

// Activate — delete old caches
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

// Fetch — network first for API calls, cache first for assets
self.addEventListener('fetch', function (e) {
  var url = e.request.url;

  // Never cache GitHub API calls or data fetches
  if (url.includes('api.github.com') || url.includes('raw.githubusercontent.com')) {
    return;
  }

  // For everything else: try network, fall back to cache
  e.respondWith(
    fetch(e.request).then(function (res) {
      // Clone response and update cache
      var clone = res.clone();
      caches.open(CACHE_NAME).then(function (cache) {
        cache.put(e.request, clone);
      });
      return res;
    }).catch(function () {
      return caches.match(e.request);
    })
  );
});
