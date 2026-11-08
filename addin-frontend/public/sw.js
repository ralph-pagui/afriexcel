// Service Worker ultra-résilient pour Assistant IA Excel
// Permet une ouverture instantanée (< 50ms) même avec une connexion 2G/3G lente ou intermittente
const CACHE_NAME = 'afriexcel-shell-v3';

const ESSENTIAL_ASSETS = [
  '/',
  '/index.html',
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return Promise.allSettled(
        ESSENTIAL_ASSETS.map((url) =>
          fetch(url, { mode: url.includes('appsforoffice') ? 'cors' : 'same-origin' })
            .then((res) => {
              if (res.ok) return cache.put(url, res);
            })
            .catch((err) => console.log('[SW] Pré-cache ignoré pour', url, err))
        )
      );
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((k) => {
          if (k !== CACHE_NAME) {
            console.log('[SW] Nettoyage ancien cache:', k);
            return caches.delete(k);
          }
        })
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Uniquement les requêtes GET
  if (req.method !== 'GET') return;

  // Ignorer les appels API Gemini externes (ils ont leur propre système de retry)
  if (url.hostname.includes('googleapis.com') || url.hostname.includes('generativelanguage')) {
    return;
  }

  // 1. Assets avec hash de build (/assets/*) -> CACHE-FIRST (chargement instantané 0ms)
  if (url.pathname.includes('/assets/')) {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req)
          .then((networkRes) => {
            if (networkRes.ok) {
              const clone = networkRes.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
            }
            return networkRes;
          })
          .catch(() => cached || new Response('', { status: 408 }));
      })
    );
    return;
  }

  // 2. Navigation et index.html -> STRATÉGIE NETWORK-FIRST AVEC TIMEOUT RAPIDE (1200ms)
  // Si le réseau est lent ou défaillant, sert immédiatement le cache local
  // Cela empêche Excel d'afficher l'erreur jaune "ADD-IN ERROR"
  event.respondWith(
    new Promise((resolve) => {
      let resolved = false;

      // Timeout rapide : si le réseau ne répond pas en 1.2s, on bascule sur le cache
      const timeoutId = setTimeout(() => {
        if (!resolved) {
          caches.match(req).then((cached) => {
            if (cached && !resolved) {
              resolved = true;
              console.log('[SW] Réseau lent (>1.2s), chargement instantané depuis le cache local :', req.url);
              resolve(cached);
            }
          });
        }
      }, 1200);

      fetch(req)
        .then((networkRes) => {
          clearTimeout(timeoutId);
          if (networkRes && networkRes.ok) {
            const clone = networkRes.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          }
          if (!resolved) {
            resolved = true;
            resolve(networkRes);
          }
        })
        .catch(async () => {
          clearTimeout(timeoutId);
          if (!resolved) {
            // En cas d'erreur de connexion ou hors-ligne, chercher dans le cache
            const cached = await caches.match(req);
            if (cached) {
              resolved = true;
              console.log('[SW] Réseau hors-ligne/erreur, secours depuis le cache local :', req.url);
              resolve(cached);
              return;
            }
            if (req.mode === 'navigate') {
              const fallback = await caches.match('/index.html');
              if (fallback) {
                resolved = true;
                resolve(fallback);
                return;
              }
            }
            resolved = true;
            resolve(new Response('Hors ligne', { status: 503, statusText: 'Offline' }));
          }
        });
    })
  );
});
