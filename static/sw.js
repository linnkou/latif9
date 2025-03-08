const CACHE_NAME = 'grades-app-v1';
const urlsToCache = [
  '/',
  '/static/manifest.json',
  '/static/icons/icon-72x72.png',
  '/static/icons/icon-96x96.png',
  '/static/icons/icon-128x128.png',
  '/static/icons/icon-144x144.png',
  '/static/icons/icon-152x152.png',
  '/static/icons/icon-192x192.png',
  '/static/icons/icon-384x384.png',
  '/static/icons/icon-512x512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', event => {
  // Only handle GET requests
  if (event.request.method !== 'GET') return;
  
  // Don't cache or handle API requests
  if (event.request.url.includes('/api/')) {
    return;
  }
  
  // Fix for malformed URLs
  if (event.request.url.includes('fetchRequest)')) {
    return;
  }
  
  // Handle the fetch event
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        if (response) {
          return response;
        }
        
        // Clone the request because it can only be used once
        const fetchRequest = event.request.clone();
        
        return fetch(fetchRequest)
          .then(response => {
            // Don't cache if not a valid response
            if (!response || response.status !== 200 || response.type !== 'basic') {
              return response;
            }
            
            // Clone the response because it can only be used once
            const responseToCache = response.clone();
            
            caches.open(CACHE_NAME)
              .then(cache => {
                cache.put(event.request, responseToCache);
              });
            
            return response;
          })
          .catch(() => {
            // Return cached index page for HTML requests when offline
            if (event.request.headers.get('accept').includes('text/html')) {
              return caches.match('/');
            }
            // Return a simple offline message for other requests
            return new Response('Offline', {
              status: 200,
              headers: new Headers({
                'Content-Type': 'text/plain'
              })
            });
          });
      })
  );
});