/* ============================================================
   Strokes library — persistent, global across decks.
   Stored in localStorage['doodle.library'] as an array of
   { id, name, thumb(small PNG dataURL), payload }. Shared by the
   task pane (save + browse) and the big-canvas dialog (save).
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
  function save(name, payload, thumb) {
    const items = list();
    const id = 'd' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
    items.unshift({ id, name: name || 'Doodle', thumb: thumb || '', payload });
    return persist(items) ? id : null;
  }
  function remove(id) { return persist(list().filter((x) => x.id !== id)); }
  function get(id) { return list().find((x) => x.id === id) || null; }
  function clearAll() { return persist([]); }

  global.DoodleLibrary = { list, save, remove, get, clearAll };
})(window);
