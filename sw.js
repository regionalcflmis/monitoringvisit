/**
 * SW.JS — service worker
 * ------------------------------------------------------------------
 * Two jobs:
 *  1. Makes the site installable (PWABuilder/Chrome require a service
 *     worker before they'll let you build/install an app).
 *  2. Caches the app SHELL (html/css/js/icons) so the form opens instantly
 *     and still loads on a weak connection.
 *
 * ⚠️ It deliberately does NOT cache API calls to Apps Script. Submitting a
 * report needs a live connection (photos upload to Drive, PDF/Excel get
 * generated, email goes out). Offline form-filling with queued submission
 * is a much bigger feature — not included here.
 */

var CACHE = 'cfl-monitoring-v1';
var SHELL = [
  './',
  './index.html',
  './dashboard.html',
  './home.html',
  './admin.html',
  './mysubmissions.html',
  './api.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE).then(function (cache) {
      // addAll fails the whole install if any single file 404s, so add
      // them individually and tolerate misses.
      return Promise.all(SHELL.map(function (url) {
        return cache.add(url).catch(function () {
          console.warn('SW: could not cache', url);
        });
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;

  // Never intercept Apps Script API traffic — it must always hit network.
  if (req.method !== 'GET' || req.url.indexOf('script.google.com') > -1) {
    return;
  }

  // Network-first for the shell, so officers always get the latest form
  // when online, but still get a cached copy when the network is poor.
  event.respondWith(
    fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); });
      return res;
    }).catch(function () {
      return caches.match(req).then(function (hit) {
        return hit || caches.match('./index.html');
      });
    })
  );
});
