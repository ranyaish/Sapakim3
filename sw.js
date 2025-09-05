// sw.js — SW מינימלי
self.addEventListener('install', (e) => {
  // skip waiting כדי שה-SW יופעל מיידית אחרי רענון
  self.skipWaiting();
});
self.addEventListener('activate', (e) => {
  // משתלט על כל הטאבים בדומיין הזה
  e.waitUntil(self.clients.claim());
});
// לא חוסמים רשת, פשוט pass-through
self.addEventListener('fetch', () => {});
