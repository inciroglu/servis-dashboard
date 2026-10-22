// Performans Yönetim Platformu · uygulama olarak yüklenebilmesi için servis çalışanı.
// Önbellek tutmaz: her istek doğrudan ağdan gelir, böylece veriler her zaman günceldir.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== self.location.origin) return;
  e.respondWith(fetch(e.request));
});
