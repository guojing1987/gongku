/* Service Worker - 设备资料库 */
const CACHE = 'gongku-v3';
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return; // 跨域(API/cloud)不拦截
  // 每次导航优先走网络，保证拿最新版本；失败退回缓存
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(res => {
      const c = res.clone();
      if (res && res.ok) { caches.open(CACHE).then(ca => ca.put(e.request, c)); }
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match('../index.html'))));
    return;
  }
  // 静态资源：stale-while-revalidate
  e.respondWith(caches.match(e.request).then(cached => {
    const fp = fetch(e.request).then(res => {
      if (res && res.ok) { const c = res.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); }
      return res;
    }).catch(() => cached);
    return cached || fp;
  }));
});
