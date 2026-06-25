/* ============================================================
   Shared UI controls — palette, color, sliders, fault buttons.
   Used by both the task pane and the big-canvas dialog (same IDs).
   Exposes window.DoodleUI.
   ============================================================ */
(function (global) {
  'use strict';

  const PALETTE = ['#FD5E6D', '#436AE1', '#EEECE6', '#FFFFFF', '#000000'];
  const $ = (id) => document.getElementById(id);

  function buildPalette() {
    const wrap = $('palette');
    if (!wrap) return;
    PALETTE.forEach((hex) => {
      const sw = document.createElement('div');
      sw.className = 'swatch';
      sw.style.background = hex;
      sw.dataset.hex = hex;
      sw.title = hex;
      sw.addEventListener('click', () => selectColor(hex));
      wrap.appendChild(sw);
    });
  }

  function selectColor(hex) {
    hex = (hex || '').toUpperCase();
    if (!/^#[0-9A-F]{6}$/.test(hex)) return;
    Doodle.state.color = hex;
    if ($('hexInput')) $('hexInput').value = hex;
    if ($('pickBtn')) $('pickBtn').style.background = hex;
    if ($('colorPicker')) $('colorPicker').value = hex;
    document.querySelectorAll('.swatch').forEach((s) =>
      s.classList.toggle('active', s.dataset.hex.toUpperCase() === hex));
  }

  function bindSlider(id, valId, fmt, apply, rerender) {
    const el = $(id), label = $(valId);
    if (!el) return;
    const update = () => {
      const v = +el.value;
      if (label) label.textContent = fmt(v);
      apply(v);
      if (rerender) Doodle.render();
    };
    el.addEventListener('input', update);
    update();
  }

  function bindFaults() {
    const grid = $('faultGrid');
    if (!grid) return;
    grid.addEventListener('click', (e) => {
      const b = e.target.closest('.fault-btn');
      if (!b) return;
      const f = b.dataset.fault;
      const ft = Doodle.state.faultTypes;
      if (f === 'none') { ft.length = 0; }
      else {
        const i = ft.indexOf(f);
        if (i >= 0) ft.splice(i, 1); else ft.push(f);
      }
      [...grid.children].forEach((btn) => {
        const bf = btn.dataset.fault;
        const on = bf === 'none' ? ft.length === 0 : ft.includes(bf);
        btn.classList.toggle('active', on);
      });
      Doodle.render();
    });
  }

  /* Wire all standard controls and attach the engine to a canvas.
     opts: { canvasId, onChange } */
  function setup(opts) {
    buildPalette();
    selectColor(PALETTE[0]);
    Doodle.attach($(opts.canvasId), { onChange: opts.onChange });

    if ($('hexInput')) $('hexInput').addEventListener('change', (e) => selectColor(e.target.value.trim()));
    if ($('colorPicker')) $('colorPicker').addEventListener('input', (e) => selectColor(e.target.value));

    // Brush size is fixed at 8px (no control); opacity and ink faults removed.
    Doodle.state.size = 8;
    Doodle.state.opacity = 1;
    Doodle.state.faultTypes = [];

    bindSlider('mouseSmooth', 'mouseSmoothVal', (v) => v + '%', (v) => Doodle.state.mouseSmoothing = v, true);
    bindSlider('vectorSmooth', 'vectorSmoothVal', (v) => v + '%', (v) => Doodle.state.vectorSmoothing = v, true);

    // Paste (Cmd/Ctrl+V) an image to use as the slide backdrop reference.
    window.addEventListener('paste', async (e) => {
      const items = e.clipboardData && e.clipboardData.items;
      if (!items) return;
      for (const it of items) {
        if (it.type && it.type.indexOf('image') === 0) {
          e.preventDefault();
          setBackdrop(await blobToDataURL(it.getAsFile()));
          break;
        }
      }
    });
  }

  /* Backdrop helpers — stored as a data URL so it survives crossing into
     the big-canvas dialog window (a blob: URL would not). */
  function blobToDataURL(blob) {
    return new Promise((res) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(blob); });
  }
  function setBackdrop(dataUrl) {
    window.DoodleBackdrop = dataUrl;
    const bg = document.getElementById('drawBg');
    if (bg) { bg.style.backgroundImage = 'url("' + dataUrl + '")'; bg.style.backgroundSize = 'contain'; }
  }
  async function loadBackdropFromClipboard() {
    try {
      if (!(navigator.clipboard && navigator.clipboard.read)) return { ok: false, reason: 'unsupported' };
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const t = item.types.find((x) => x.indexOf('image/') === 0);
        if (t) { setBackdrop(await blobToDataURL(await item.getType(t))); return { ok: true }; }
      }
      return { ok: false, reason: 'no-image' };
    } catch (e) { return { ok: false, reason: 'denied', error: String(e) }; }
  }

  global.DoodleUI = { setup, selectColor, PALETTE, loadBackdropFromClipboard, setBackdrop };
})(window);
