/* OKV SACCO Subscribe — Service Worker
   Caches the installed-app shell so it still OPENS offline (login screen
   resumes a cached session; dashboard reads/writes local IndexedDB and
   syncs to the server whenever a connection is available).

   IMPORTANT: config.js and the marketing pages (index/demo/pricing/signup)
   are served network-first — they hold pricing/contact info and copy that
   can change, so a visitor should never get stuck on a stale cached copy.
   Bump CACHE_NAME any time you want to force everyone's cache to refresh. */

const CACHE_NAME = "okv-subscribe-shell-v9";
const APP_SHELL = [
  "./login.html", "./install.html", "./dashboard.html", "./super-admin.html", "./reset-password.html",
  "./db.js", "./manifest.json", "./icons/icon-192.png", "./icons/icon-512.png"
];
const NETWORK_FIRST = [
  "config.js", "index.html", "demo.html", "pricing.html", "signup.html", "/",
  "privacy-policy.html", "terms-of-service.html", "refund-policy.html"
];

self.addEventListener("install", (event)=>{
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache)=> cache.addAll(APP_SHELL)).then(()=> self.skipWaiting())
  );
});
self.addEventListener("activate", (event)=>{
  event.waitUntil(
    caches.keys().then((keys)=> Promise.all(keys.filter(k=> k!==CACHE_NAME).map(k=> caches.delete(k))))
      .then(()=> self.clients.claim())
  );
});
self.addEventListener("fetch", (event)=>{
  if(event.request.method !== "GET") return;
  const url = event.request.url;
  if(url.includes("script.google.com")) return; // never cache API calls — always go live

  const isNetworkFirst = NETWORK_FIRST.some(p=> url.endsWith(p));
  if(isNetworkFirst){
    event.respondWith(
      fetch(event.request).then((response)=>{
        caches.open(CACHE_NAME).then((cache)=> cache.put(event.request, response.clone()));
        return response;
      }).catch(()=> caches.match(event.request))
    );
    return;
  }

  // App shell: stale-while-revalidate — instant + works offline, refreshes in background.
  event.respondWith(
    caches.open(CACHE_NAME).then(async (cache)=>{
      const cached = await cache.match(event.request);
      const networkFetch = fetch(event.request).then((response)=>{
        if(response && (response.ok || response.type === "opaque")) cache.put(event.request, response.clone());
        return response;
      }).catch(()=> cached);
      return cached || networkFetch;
    })
  );
});
