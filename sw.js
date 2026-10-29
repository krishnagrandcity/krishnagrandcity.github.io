/* KGC app service worker
 * - The page itself (index.html) is fetched from the network first, so new releases
 *   show up as soon as you are online; the last copy is kept for offline opening.
 * - Fonts and the chart / Excel / PDF libraries are cached so the app opens offline.
 * - Society data (the backend API on Render) is never cached or intercepted, so
 *   members, payments and logins always come live from the server.
 */
const SHELL_CACHE = 'kgc-shell-v1';
const CDN_CACHE   = 'kgc-cdn-v1';
const SHELL_FILES = ['./', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'favicon-32.png'];
const CDN_HOSTS   = ['cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then(c => c.addAll(SHELL_FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== SHELL_CACHE && k !== CDN_CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 1) The page: network first, fall back to the last saved copy when offline
  if (req.mode === 'navigate' && url.origin === self.location.origin) {
    event.respondWith(
      fetch(req)
        .then(res => {
          if (res.ok) { const copy = res.clone(); caches.open(SHELL_CACHE).then(c => c.put('./', copy)); }
          return res;
        })
        .catch(() => caches.match('./').then(r => r || offlinePage()))
    );
    return;
  }

  // 2) Our own static files (icons, manifest): cached copy first
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then(hit => hit || fetch(req).then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(SHELL_CACHE).then(c => c.put(req, copy)); }
        return res;
      }))
    );
    return;
  }

  // 3) Fonts and libraries from public CDNs: use the saved copy, refresh it in the background
  if (CDN_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.open(CDN_CACHE).then(cache =>
        cache.match(req).then(hit => {
          const fresh = fetch(req).then(res => { if (res.ok || res.type === 'opaque') cache.put(req, res.clone()); return res; })
                                  .catch(() => hit);
          return hit || fresh;
        })
      )
    );
  }
  // 4) Anything else (the backend API, etc.) is left alone and always goes to the network
});

function offlinePage() {
  return new Response(
    '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>KGC — offline</title><body style="font-family:system-ui,sans-serif;background:#0f1f3d;color:#fff;display:flex;' +
    'align-items:center;justify-content:center;min-height:100vh;margin:0;text-align:center;padding:24px">' +
    '<div><h1 style="font-size:22px">You are offline</h1><p style="opacity:.8">Connect to the internet and open the app again.</p></div>',
    { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
  );
}
