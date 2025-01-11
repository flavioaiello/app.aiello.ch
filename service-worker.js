// Importiere Workbox von der CDN
importScripts('https://storage.googleapis.com/workbox-cdn/releases/6.5.4/workbox-sw.js');

if (workbox) {
  console.log('Workbox erfolgreich geladen');

  // Konfiguriere Workbox
  workbox.setConfig({
    debug: false
  });

  // Vorab gecachte URLs
  const precacheResources = [
    '/',
    '/index.html',
    '/assets/js/app.js',
    '/manifest.json',
    '/service-worker.js',
    '/assets/icons/192x192.png',
    '/assets/icons/512x512.png'
  ];

  // Pre-Caching der Ressourcen
  workbox.precaching.precacheAndRoute(precacheResources.map(url => ({ url, revision: null })));

  workbox.routing.registerRoute(
    ({ url }) => url.origin === 'https://esm.sh',
    new workbox.strategies.NetworkFirst({
      cacheName: 'esm-sh-cache',
    })
  );

  // Caching-Strategie für Bootstrap und externe CSS/JS
  workbox.routing.registerRoute(
    ({url}) => url.origin === 'https://cdn.jsdelivr.net' ||
               url.origin === 'https://unpkg.com',
    new workbox.strategies.StaleWhileRevalidate({
      cacheName: 'external-resources',
      plugins: [
        new workbox.expiration.ExpirationPlugin({
          maxEntries: 50,
          maxAgeSeconds: 30 * 24 * 60 * 60, // 30 Tage
        }),
      ],
    })
  );

  // Caching-Strategie für lokale CSS/JS
  workbox.routing.registerRoute(
    ({request}) => request.destination === 'style' ||
                   request.destination === 'script',
    new workbox.strategies.StaleWhileRevalidate({
      cacheName: 'static-resources',
      plugins: [
        new workbox.expiration.ExpirationPlugin({
          maxEntries: 50,
          maxAgeSeconds: 30 * 24 * 60 * 60, // 30 Tage
        }),
      ],
    })
  );

  // Caching-Strategie für HTML-Dokumente
  workbox.routing.registerRoute(
    ({request}) => request.destination === 'document',
    new workbox.strategies.NetworkFirst({
      cacheName: 'pages',
      plugins: [
        new workbox.expiration.ExpirationPlugin({
          maxEntries: 10,
          maxAgeSeconds: 7 * 24 * 60 * 60, // 7 Tage
        }),
      ],
    })
  );

  // Caching-Strategie für Bilder (Icons)
  workbox.routing.registerRoute(
    ({request}) => request.destination === 'image',
    new workbox.strategies.CacheFirst({
      cacheName: 'images',
      plugins: [
        new workbox.expiration.ExpirationPlugin({
          maxEntries: 50,
          maxAgeSeconds: 30 * 24 * 60 * 60, // 30 Tage
        }),
      ],
    })
  );

  // Fallback für Offline-Nutzung (optional)
  workbox.routing.setDefaultHandler(
    new workbox.strategies.NetworkFirst({
      cacheName: 'fallback',
    })
  );

} else {
  console.log('Workbox konnte nicht geladen werden');
}
