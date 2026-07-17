// 発注管理 Service Worker — ネットワーク優先＋オフライン時キャッシュ
const CACHE = "hatchu-v1";
const SHELL = ["./", "./index.html", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png", "./manifest.webmanifest"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  // GET かつ同一オリジンのみ扱う（GoogleシートAPI等の外部通信はそのまま素通し）
  if (e.request.method !== "GET" || url.origin !== self.location.origin) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(e.request).then((m) => m || caches.match("./index.html")))
  );
});
