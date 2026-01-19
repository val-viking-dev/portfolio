/*
 * ========================================
 * GARDEN-HELPER v1.9.2 - SERVICE WORKER
 * ========================================
 * 
 * Service Worker pour le mode hors ligne (PWA)
 * Gère le cache des ressources et les stratégies de réseau
 * 
 * @author    Valentin
 * @role      Développeur / Concepteur d'Application
 * @date      Janvier 2026
 * @version   1.9.2
 * ========================================
 */

// Nom du cache (incrémenter à chaque version)
const CACHE_NAME = 'garden-helper-v1.9.2';
const urlsToCache = [
  '/',
  '/index.html',
  '/app.js',
  '/vegetables-data.js',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png'
];

// Installation du Service Worker
self.addEventListener('install', event => {
  console.log('[SW] Installation...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[SW] Mise en cache des fichiers');
        return cache.addAll(urlsToCache);
      })
      .then(() => self.skipWaiting())
  );
});

// Activation du Service Worker
self.addEventListener('activate', event => {
  console.log('[SW] Activation...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Suppression ancien cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Interception des requêtes (stratégie Cache First)
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Retourne depuis le cache si disponible
        if (response) {
          return response;
        }
        
        // Sinon, récupère depuis le réseau
        return fetch(event.request).then(response => {
          // Vérifie si la réponse est valide
          if (!response || response.status !== 200 || response.type !== 'basic') {
            return response;
          }
          
          // Clone la réponse
          const responseToCache = response.clone();
          
          // Ajoute au cache
          caches.open(CACHE_NAME)
            .then(cache => {
              cache.put(event.request, responseToCache);
            });
          
          return response;
        });
      })
      .catch(() => {
        // Page hors-ligne de secours
        return caches.match('/index.html');
      })
  );
});
