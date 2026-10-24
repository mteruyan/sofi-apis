// Service worker: hace la app instalable y recibe avisos push. No cachea nada (datos de salud y ubicación).
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});

// Aviso push: nunca incluye datos del paciente (solo tipo de servicio, distrito y hora)
self.addEventListener('push', (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) {}
  e.waitUntil(self.registration.showNotification(d.titulo || 'Sanitas MAD', {
    body: d.cuerpo || '',
    tag: d.tag || 'mad',
    renotify: true,
    requireInteraction: true,
    vibrate: [300, 150, 300, 150, 300],
    icon: '/static/pwa/proveedor-192.png',
    badge: '/static/pwa/proveedor-192.png',
    data: { url: d.url || '/proveedor' },
  }));
});

// Al tocar el aviso: abre (o trae al frente) el panel
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || '/proveedor';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((lista) => {
    for (const c of lista) {
      if ('focus' in c) return c.focus();
    }
    return self.clients.openWindow(url);
  }));
});
