/* ============================================================
   Generated-images gallery — persistent via IndexedDB.
   Full-res 1920x1080 images are too big for localStorage, so the
   library of generated/edited images lives in IndexedDB (large quota,
   survives reopening the add-in). Each item: { id, dataUrl, prompt,
   model, ts }. All methods are async and never throw (degrade to a
   no-op / empty list if IndexedDB is unavailable).
   ============================================================ */
(function (global) {
  'use strict';
  const DB = 'doodleImages', STORE = 'gallery', VERSION = 1;

  function openDb() {
    return new Promise((resolve, reject) => {
      let req;
      try { req = indexedDB.open(DB, VERSION); } catch (e) { return reject(e); }
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function list() {
    try {
      const db = await openDb();
      return await new Promise((resolve) => {
        const r = db.transaction(STORE, 'readonly').objectStore(STORE).getAll();
        r.onsuccess = () => resolve((r.result || []).sort((a, b) => (b.ts || 0) - (a.ts || 0)));
        r.onerror = () => resolve([]);
      });
    } catch (_) { return []; }
  }
  async function add(item) {
    try {
      const db = await openDb();
      return await new Promise((resolve) => {
        const t = db.transaction(STORE, 'readwrite');
        t.objectStore(STORE).put(item);
        t.oncomplete = () => resolve(true);
        t.onerror = () => resolve(false);
      });
    } catch (_) { return false; }
  }
  async function remove(id) {
    try {
      const db = await openDb();
      return await new Promise((resolve) => {
        const t = db.transaction(STORE, 'readwrite');
        t.objectStore(STORE).delete(id);
        t.oncomplete = () => resolve(true);
        t.onerror = () => resolve(false);
      });
    } catch (_) { return false; }
  }

  global.DoodleGallery = { list, add, remove };
})(window);
