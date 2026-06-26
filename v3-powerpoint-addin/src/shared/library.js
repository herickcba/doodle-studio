/* ============================================================
   Strokes library — persistent, global across decks.
   Stored in localStorage['doodle.library'] as an array of
   { id, name, kind:'png'|'gif', thumb(small PNG dataURL), payload,
     gif?:{duration,loop,fps,holdMs,easing} }.
   GIFs store only the strokes + animation options (regenerated on
   insert) — never the heavy GIF bytes. Shared by the task pane
   (save + browse) and the big-canvas dialog (save).
   ============================================================ */
(function (global) {
  'use strict';
  const KEY = 'doodle.library';

  function list() {
    try { return JSON.parse(localStorage.getItem(KEY) || '[]'); }
    catch (_) { return []; }
  }
  function persist(items) {
    try { localStorage.setItem(KEY, JSON.stringify(items)); return true; }
    catch (e) { console.warn('Biblioteca cheia ou indisponível:', e); return false; }
  }
  // extra = optional { kind:'gif', gif:{...} } merged into the item.
  function save(name, payload, thumb, extra) {
    const items = list();
    const id = 'd' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
    const item = Object.assign({ id, name: name || 'Doodle', kind: 'png', thumb: thumb || '', payload }, extra || {});
    items.unshift(item);
    return persist(items) ? id : null;
  }
  function remove(id) { return persist(list().filter((x) => x.id !== id)); }
  function get(id) { return list().find((x) => x.id === id) || null; }
  function clearAll() { return persist([]); }

  global.DoodleLibrary = { list, save, remove, get, clearAll };
})(window);
