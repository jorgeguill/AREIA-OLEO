// Service Worker do módulo Painel do Forno.
// Caminhos RELATIVOS ao escopo onde o SW for registrado — assim o módulo
// funciona tanto em "/" (standalone) quanto sob uma subpasta do SIGA
// (ex.: "/ferramentas/painel-forno/").
const CACHE_NAME = 'painel-forno-v2';
const ASSETS = ['./', './index.html', './manifest.json', './painel-forno.js', './painel-forno.css'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS).catch(() => { /* offline ok */ });
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) => {
      return Promise.all(names.map((name) => {
        if (name !== CACHE_NAME) return caches.delete(name);
      }));
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      return fetch(event.request)
        .then((response) => {
          if (!response || response.status !== 200) return response;
          const toCache = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, toCache);
          });
          return response;
        })
        .catch(() => cached || new Response('Offline', { status: 503 }));
    })
  );
});
