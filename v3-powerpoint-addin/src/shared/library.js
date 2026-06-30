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
  const CAP = 40;            // máximo de itens guardados (evita estourar a quota)

  function list() {
    try { return JSON.parse(localStorage.getItem(KEY) || '[]'); }
    catch (_) { return []; }
  }
  // Grava limitando a CAP itens. Se a quota estourar, descarta os mais antigos
  // e tenta de novo — assim salvar nunca falha em silêncio por biblioteca cheia.
  function persist(items) {
    let arr = items.slice(0, CAP);
    for (let tries = 0; tries < 8; tries++) {
      try { localStorage.setItem(KEY, JSON.stringify(arr)); return true; }
      catch (e) {
        if (arr.length <= 1) { console.warn('Biblioteca: quota estourada:', e); return false; }
        arr = arr.slice(0, Math.max(1, Math.floor(arr.length * 0.7)));   // poda os mais antigos
      }
    }
    return false;
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

  // Save the 3 versions of a drawing in one go: a looping GIF, a non-looping
  // GIF, and the static PNG of the final state. GIFs store only strokes +
  // animation options (regenerated on insert) — no heavy bytes.
  function saveGifSet(payload, thumb, gifOpts) {
    const o = gifOpts || {};
    const a = save('GIF', payload, thumb, { kind: 'gif', gif: Object.assign({}, o, { loop: true }) });
    const b = save('GIF', payload, thumb, { kind: 'gif', gif: Object.assign({}, o, { loop: false }) });
    const c = save('Desenho', payload, thumb, { kind: 'png' });
    return !!(a && b && c);   // true só se as 3 versões salvaram
  }

  global.DoodleLibrary = { list, save, saveGifSet, remove, get, clearAll };
})(window);
