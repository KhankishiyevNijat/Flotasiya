// Offline support: the app shell is cached on install; pages are fetched from the
// network first so updates arrive, with the cached copy used when offline.
const CACHE = "flotasiya-v5";
const SHELL = ["./", "index.html", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "jspdf.umd.min.js", "jspdf.plugin.autotable.min.js", "report-font.ttf", "report-font-bold.ttf"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const sameOrigin = new URL(req.url).origin === self.location.origin;
  const save = res => {
    if (res && (res.ok || res.type === "opaque")) {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
    }
    return res;
  };
  if (sameOrigin) {
    // Network first, cache as fallback.
    e.respondWith(
      fetch(req).then(save).catch(() =>
        caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("index.html"))
      )
    );
  } else {
    // Fonts and other outside files: cache first.
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(save)));
  }
});
