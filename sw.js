/* CD Collection: service worker (offline cache + update detection).
   tools/deploy.sh stamps BUILD on every release; a changed BUILD is what tells devices an update exists.
   Written in ES2017 with no optional chaining so it runs on iOS 12 Safari. */
var BUILD = 4;
var CACHE = 'cd-collection-b' + BUILD;
var BASE = '/cd-collection/';
var FILES = ['', 'index.html', 'collection.js', 'apple-touch-icon.png', 'manifest.json'];

self.addEventListener('install', function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) {
    return Promise.all(FILES.map(function (f) {
      // cache-bust so the precache never stores a stale HTTP-cached copy
      return fetch(BASE + f + '?b=' + BUILD, { cache: 'no-store' }).then(function (res) {
        if (res.ok) return cache.put(BASE + f, res);
      }).catch(function () {});
    }));
  }));
});

self.addEventListener('activate', function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) {
      return k.indexOf('cd-collection-') === 0 && k !== CACHE;
    }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  var own = url.origin === self.location.origin && url.pathname.indexOf(BASE) === 0;
  var fonts = url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com';
  if (!own && !fonts) return;
  if (own && url.pathname.slice(-6) === '/sw.js') return;
  event.respondWith(caches.open(CACHE).then(function (cache) {
    var key = own ? new Request(url.origin + url.pathname) : req;
    return cache.match(key).then(function (hit) {
      if (own) {
        // app files: serve the cached release; only a new BUILD replaces them
        if (hit) return hit;
        return fetch(req).then(function (res) { if (res.ok) cache.put(key, res.clone()); return res; });
      }
      // fonts: cache-first, fill on first use
      return hit || fetch(req).then(function (res) { cache.put(req, res.clone()); return res; });
    });
  }));
});

self.addEventListener('message', function (event) {
  var d = event.data || {};
  if (d.type === 'SKIP_WAITING') self.skipWaiting();
  if (d.type === 'GET_BUILD' && event.ports && event.ports[0]) event.ports[0].postMessage({ build: BUILD });
});
