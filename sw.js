// Boutique 1981 — Scanner POS service worker
// ══════════════════════════════════════════════════════════════════════════
// Item 7 (wishlist): makes the app installable and usable with ZERO
// connectivity, ahead of outage season.
//
// Scope, deliberately narrow: this only caches the APP SHELL — the HTML
// page itself, the manifest, the icons, and the fixed CDN libraries the
// page loads. It NEVER intercepts a request to an Apps Script URL (product
// sync, customer sync, scan/sale pushes) — those already have their own
// offline-tolerant retry queue inside Boutique_Scanner.html
// (_scanPushQueue / flushScanQueue), which is a better fit than a service
// worker cache: it queues actual writes and retries them, rather than
// serving back stale data. Caching them here would risk masking a real
// sync failure or serving yesterday's product list as if it were current.
//
// Bump CACHE_NAME whenever the shell's own file list changes, so old
// clients pick up the new set on next activate instead of being stuck.
const CACHE_NAME = 'boutique1981-shell-v1';
const SHELL_URLS = [
  './Boutique_Scanner.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
  'https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js',
  'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
  'https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(SHELL_URLS))
      .catch(() => { /* a CDN hiccup during install shouldn't block installability */ })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

function isShellUrl(url) {
  return SHELL_URLS.some((u) => url === u || url.endsWith(u.replace('./', '/')));
}

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Never touch writes (sale/scan pushes, price edits, etc.) — let them
  // fail straight through to the app's own queue when offline, exactly as
  // they did before this service worker existed.
  if (req.method !== 'GET') return;
  if (!isShellUrl(req.url)) return; // Apps Script reads, anything dynamic: pass through untouched

  // Network-first, cache fallback: keeps "always freshest shell when
  // online" (the page's own Cache-Control meta tags already ask for
  // that), while still rendering with zero connectivity.
  event.respondWith(
    fetch(req)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(req))
  );
});
