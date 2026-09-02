const CACHE_NAME = 'lab-gold-v2';
const ASSETS_TO_CACHE = [
  '/',
  '/recepcion',
  '/generador',
  '/manifest.json'
];

// Instalación del Service Worker y almacenamiento en caché inicial
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
  self.skipWaiting();
});

// Activación y limpieza de cachés antiguas
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Interceptación de peticiones (Estrategia mixta: Red primero, fallback a caché)
self.addEventListener('fetch', (event) => {
  // Solo manejamos peticiones GET para evitar interferir con los POST de guardar datos
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Si la red responde correctamente, clonamos y guardamos en caché dinámicamente
        if (response && response.status === 200) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }
        return response;
      })
      .catch(() => {
        // Si no hay red, buscamos en el caché
        return caches.match(event.request);
      })
  );
});