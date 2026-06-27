/* ============================================================
   Image reference library (style references) — persistent.
   localStorage['doodle.imgRefs'] = [{ id, name, thumb(dataURL) }].
   Uploads are downscaled so they fit the storage quota — a style
   reference doesn't need full resolution.
   (Generated images live in the session gallery, in memory — see
   image-gen.js — because full-res 1920x1080 would blow the quota.)
   ============================================================ */
(function (global) {
  'use strict';
  const KEY = 'doodle.imgRefs';
  const MAX_PX = 1280;   // bigger: a reference can also be the image-to-edit

  function list() {
    try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch (_) { return []; }
  }
  function persist(items) {
    try { localStorage.setItem(KEY, JSON.stringify(items)); return true; }
    catch (e) { console.warn('imgRefs cheio?', e); return false; }
  }

  // Downscale a data URL to <= MAX_PX on the long edge, as JPEG.
  function downscale(dataUrl) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, MAX_PX / Math.max(img.naturalWidth, img.naturalHeight));
        const w = Math.max(1, Math.round(img.naturalWidth * scale));
        const h = Math.max(1, Math.round(img.naturalHeight * scale));
        const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
        cv.getContext('2d').drawImage(img, 0, 0, w, h);
        try { resolve(cv.toDataURL('image/jpeg', 0.85)); } catch (_) { resolve(dataUrl); }
      };
      img.onerror = () => resolve(dataUrl);
      img.src = dataUrl;
    });
  }

  async function add(dataUrl, name) {
    const thumb = await downscale(dataUrl);
    const items = list();
    const id = 'r' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
    items.unshift({ id, name: name || 'Referência', thumb });
    persist(items);
    return id;
  }
  function remove(id) { return persist(list().filter((x) => x.id !== id)); }
  function get(id) { return list().find((x) => x.id === id) || null; }

  global.DoodleImgRefs = { list, add, remove, get };
})(window);
