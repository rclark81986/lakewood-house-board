// Minimal service worker: network-first for pages, never touches Google Apps Script.
const CACHE = "house-board-v1";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // script.google.com etc. pass straight through
  if (req.mode !== "navigate") return;
  e.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put("./", copy)).catch(() => {});
      return res;
    }).catch(() => caches.match("./"))
  );
});
