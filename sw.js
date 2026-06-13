/* Service Worker — מערכת הזימונים של אביאל
   מטרה: טעינה מהירה והתקנה למסך הבית. לא מאחסן קריאות API (תמיד נתונים טריים). */
const CACHE = 'aviel-booking-v1';
const SHELL = [
  './',
  './index.html',
  './icon-192.png',
  './icon-512.png',
  './manifest.json'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  // רק GET ורק אותו origin (האפליקציה עצמה) — קריאות ל-Apps Script עוברות ישר לרשת
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  // ניווט: network-first עם נפילה לקאש (כדי לקבל גרסה עדכנית כשיש רשת)
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then(r => {
        const copy = r.clone();
        caches.open(CACHE).then(c => c.put('./index.html', copy));
        return r;
      }).catch(() => caches.match('./index.html'))
    );
    return;
  }
  // נכסים סטטיים: cache-first
  e.respondWith(caches.match(req).then(c => c || fetch(req)));
});
