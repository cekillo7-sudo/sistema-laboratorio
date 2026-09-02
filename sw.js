self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open('lab-store').then((cache) => {
      return cache.addAll([
        '/',
        '/recepcion',
        '/generador'
      ]);
    })
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});