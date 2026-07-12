// IMPORTANT : incrémenter ce nom à CHAQUE nouveau déploiement pour forcer les
// appareils des utilisateurs à récupérer la nouvelle version (sinon l'ancienne
// version reste affichée à cause du cache du navigateur).
const CACHE_NAME = 'noteexpress-cache-v2'

const CORE_ASSETS = ['/', '/manifest.webmanifest', '/icon-192.png', '/icon-512.png']

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS)).catch(() => {})
  )
  self.skipWaiting()
})

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  )
  self.clients.claim()
})

// Stratégie network-first : toujours essayer le réseau en premier pour avoir
// la dernière version ; ne retomber sur le cache qu'en cas d'échec (hors ligne).
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const clone = response.clone()
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone)).catch(() => {})
        return response
      })
      .catch(() => caches.match(event.request))
  )
})
