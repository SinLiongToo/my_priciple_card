/* 工作的管見 service worker — makes the card app installable and usable offline.

   Strategy: NETWORK-FIRST, falling back to the cache only when the network
   fails. Online users always get the latest index.html (no "stuck on an old
   version" problem); offline users get the last copy they saw. The app shell
   is pre-cached on install so it opens offline even on the first offline visit.

   Bump CACHE when the precache list changes; old caches are removed on activate. */
const CACHE = 'my-two-cents-v1';
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      // Cache what we can; one failed file shouldn't block installation.
      .then(cache => Promise.all(SHELL.map(u => cache.add(u).catch(() => {}))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (!/^https?:$/.test(url.protocol)) return;

  event.respondWith(
    fetch(req)
      .then(res => {
        // Cache successful same-origin responses and CDN assets (e.g. Google Fonts; opaque ok).
        if (res && (res.ok || res.type === 'opaque')) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() =>
        caches.match(req, { ignoreSearch: true }).then(hit => {
          if (hit) return hit;
          if (req.mode === 'navigate') return caches.match('./index.html');
          return Response.error();
        })
      )
  );
});
