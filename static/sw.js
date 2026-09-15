const CACHE_NAME = 'onetrack-v1';
const STATIC_ASSETS = [
  '/',
  '/login',
  '/leaderboard',
  '/daily',
  '/static/manifest.json',
  '/static/images/icon-192x192.png',
  '/static/images/icon-512x512.png',
  '/static/images/shell-logo.png',
  '/static/images/favicon.ico'
];

const OFFLINE_PAGE = `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>OneTrack - Offline</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#0f172a;color:#e2e8f0;font-family:Inter,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center;padding:20px}
.box{max-width:400px}
.icon{font-size:64px;margin-bottom:16px}
h1{font-size:24px;font-weight:700;margin-bottom:8px}
p{color:#94a3b8;margin-bottom:24px}
.btn{display:inline-block;padding:12px 24px;background:#f59e0b;color:#0f172a;font-weight:600;border-radius:12px;text-decoration:none;font-size:16px}
</style>
</head>
<body>
<div class="box">
<div class="icon">&#128244;</div>
<h1>You're Offline</h1>
<p>No internet connection. Cached pages are still available.</p>
<a href="/" class="btn">Try Again</a>
</div>
</body>
</html>`;

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(STATIC_ASSETS).catch(() => {});
    }).then(() => {
      return caches.open(CACHE_NAME).then(cache => {
        return cache.put('/offline', new Response(OFFLINE_PAGE, {
          headers: { 'Content-Type': 'text/html' }
        }));
      });
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Skip non-GET and API/auth requests
  if (e.request.method !== 'GET') return;
  if (url.pathname.startsWith('/api/')) return;
  if (url.pathname.startsWith('/form') && e.request.method === 'GET') {
    // Network-first for form page
    e.respondWith(
      fetch(e.request).then(resp => {
        const clone = resp.clone();
        caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        return resp;
      }).catch(() => caches.match(e.request).then(r => r || caches.match('/offline')))
    );
    return;
  }

  // Cache-first for static assets, network-first for pages
  if (url.pathname.startsWith('/static/')) {
    e.respondWith(
      caches.match(e.request).then(r => r || fetch(e.request).then(resp => {
        const clone = resp.clone();
        caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        return resp;
      }))
    );
  } else {
    e.respondWith(
      fetch(e.request).then(resp => {
        const clone = resp.clone();
        caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        return resp;
      }).catch(() => caches.match(e.request).then(r => r || caches.match('/offline')))
    );
  }
});

// Background Sync: queue failed POSTs and retry when online
self.addEventListener('sync', e => {
  if (e.tag === 'sync-entries') {
    e.waitUntil(syncPendingEntries());
  }
});

// Also retry on online event (for browsers without sync support)
self.addEventListener('message', e => {
  if (e.data === 'retry-sync') {
    syncPendingEntries();
  }
});

async function syncPendingEntries() {
  const db = await openDB();
  const tx = db.transaction('pending-sync', 'readwrite');
  const store = tx.objectStore('pending-sync');
  const all = await getAllFromStore(store);

  for (const entry of all) {
    try {
      const resp = await fetch('/api/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ entries: [entry.data] })
      });
      if (resp.ok) {
        store.delete(entry.id);
      }
    } catch (err) {
      // Will retry next sync event
    }
  }
}

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('onetrack-sync', 1);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('pending-sync')) {
        db.createObjectStore('pending-sync', { keyPath: 'id', autoIncrement: true });
      }
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}

function getAllFromStore(store) {
  return new Promise((resolve, reject) => {
    const req = store.getAll();
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}
