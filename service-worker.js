/* Mohanpur Construction service worker — installable PWA + offline shell.
   Bump CACHE_VERSION whenever you change index.html or assets. */
const CACHE_VERSION = "mohanpur-v2";

const PRECACHE = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./icons/home-icon.png",
  "./photos/Surjendu.jpg",
  "./photos/Madhumita.jpg",
  "./photos/Subhendu.jpg",
  "./photos/Mamata.jpg"
];

self.addEventListener("install", (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_VERSION).then((c) =>
      Promise.all(PRECACHE.map((url) => c.add(url).catch(() => null)))
    )
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;

  const url = req.url;
  if (url.includes("script.google.com") || url.includes("script.googleusercontent.com")) {
    return;
  }

  e.respondWith(
    fetch(req)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE_VERSION).then((c) => c.put(req, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(req).then((hit) => hit || caches.match("./index.html")))
  );
});
