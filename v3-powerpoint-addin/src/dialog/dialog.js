/* ============================================================
   Big-canvas dialog controller.
   Reads the slide backdrop from localStorage (same origin as the
   task pane), lets the user draw large, then hands the drawing back
   to the task pane (which inserts it). Works standalone too.
   ============================================================ */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);

  let zoom = 1;
  function setZoom(z) {
    zoom = Math.max(0.25, Math.min(4, z));
    const w = document.querySelector('.big-wrap');
    if (w) w.style.transform = 'scale(' + zoom + ')';
    const lbl = $('zoomLabel'); if (lbl) lbl.textContent = Math.round(zoom * 100) + '%';
  }

  // Reset a save button's label back to its original text (so a new drawing
  // doesn't keep showing a stale "✓ salvo").
  function resetLabel(id) { const b = $(id); if (b && b.dataset.label) b.textContent = b.dataset.label; }

  function updateButtons() {
    const empty = Doodle.isEmpty();
    $('insertBtn').disabled = empty;
    $('gifBtn').disabled = empty;
    $('saveGifBtn').disabled = empty;
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
    $('saveAllBtn').disabled = empty;
    $('saveSelBtn').disabled = !Doodle.getSelectedStroke();
    ['saveGifBtn', 'saveAllBtn', 'saveSelBtn'].forEach(resetLabel);
  }

  function saveToLibrary(strokes, label) {
    if (!strokes || !strokes.length) return;
    const thumb = Doodle.thumbnailOf(strokes);
    DoodleLibrary.save(label, Doodle.payloadOf(strokes), thumb);
  }

  // Save the 3 versions (loop GIF + no-loop GIF + static PNG) into the library.
  function saveGifToLibrary() {
    const strokes = Doodle.state.strokes;
    if (!strokes.length) return;
    DoodleLibrary.saveGifSet(Doodle.payload(), Doodle.thumbnailOf(strokes), gifOpts());
  }

  function loadBackdrop() {
    let img = '';
    try { img = localStorage.getItem('doodle.slideImage') || ''; } catch (_) {}
    const bg = $('drawBg');
    if (img) {
      bg.style.backgroundImage = `url("${img}")`;
      bg.style.backgroundSize = 'contain';
    }
    // opacity slider drives the backdrop dimming
    const op = $('bgOpacity');
    const apply = () => { bg.style.opacity = (+op.value) / 100; };
    op.addEventListener('input', apply);
    apply();
  }

  // Current animation settings (no loop control — see saveGifSet for the 3 versions).
  function gifOpts() {
    return {
      duration: (+$('gifDur').value) / 10,
      fps: +$('gifFps').value,
      holdMs: +$('gifHold').value,
      easing: $('gifEasing').value,
    };
  }

  function finish(kind, gif) {
    let p = null;
    if (kind === 'inserted') {
      p = Doodle.payload();
      if (gif) {
        const o = gifOpts();
        p.asGif = true; p.gifLoop = false;   // inserted GIFs go in without loop
        p.gifDuration = o.duration; p.gifEasing = o.easing; p.gifFps = o.fps; p.gifHold = o.holdMs;
      }
    }
    if (window.opener) {            // standalone window.open() (modo navegador)
      try { localStorage.setItem('doodle.result', JSON.stringify(p)); } catch (_) {}
      try { window.close(); } catch (_) {}
      return;
    }
    // Office Dialog: entrega o payload DENTRO da mensagem (sem localStorage),
    // eliminando a falha silenciosa quando a gravação no localStorage falhava.
    try {
      Office.context.ui.messageParent(JSON.stringify({ kind: kind, payload: p }));
    } catch (e) {
      // payload grande demais p/ a mensagem: cai no localStorage + sinal curto
      try { localStorage.setItem('doodle.result', JSON.stringify(p)); } catch (_) {}
      try { Office.context.ui.messageParent(kind); }
      catch (_) { try { window.close(); } catch (_) {} }
    }
  }

  function boot() {
    loadBackdrop();
    DoodleUI.setup({ canvasId: 'drawCanvas', onChange: updateButtons });
    $('undoBtn').addEventListener('click', () => Doodle.undo());
    $('redoBtn').addEventListener('click', () => Doodle.redo());
    $('clearBtn').addEventListener('click', () => Doodle.clear());
    $('insertBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) finish('inserted', false); });
    // Insert as GIF (no loop) AND stash the 3 versions in the library.
    $('gifBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) { saveGifToLibrary(); finish('inserted', true); } });
    $('cancelBtn').addEventListener('click', () => finish('cancel'));
    ['saveGifBtn', 'saveAllBtn', 'saveSelBtn'].forEach((id) => { const b = $(id); if (b) b.dataset.label = b.textContent; });
    const gd = $('gifDur'), gdv = $('gifDurVal');
    if (gd) gd.addEventListener('input', () => { gdv.textContent = ((+gd.value) / 10).toFixed(1) + 's'; });
    const gh = $('gifHold'), ghv = $('gifHoldVal');
    if (gh) gh.addEventListener('input', () => { ghv.textContent = gh.value + 'ms'; });
    $('zoomInBtn').addEventListener('click', () => setZoom(zoom * 1.2));
    $('zoomOutBtn').addEventListener('click', () => setZoom(zoom / 1.2));
    $('zoomFitBtn').addEventListener('click', () => setZoom(1));
    $('saveAllBtn').addEventListener('click', () => { saveToLibrary(Doodle.state.strokes, 'Desenho'); $('saveAllBtn').textContent = '✓ Salvo na biblioteca'; });
    $('saveSelBtn').addEventListener('click', () => { const s = Doodle.getSelectedStroke(); if (s) { saveToLibrary([s], 'Traço'); $('saveSelBtn').textContent = '✓ Traço salvo'; } });
    $('saveGifBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) { saveGifToLibrary(); $('saveGifBtn').textContent = '✓ 3 versões salvas'; } });
    $('bgBtn').addEventListener('click', async () => {
      const res = await DoodleUI.loadBackdropFromClipboard();
      $('bgBtn').textContent = res.ok ? '✓ Fundo do slide' : '📋 Copie o slide e cole (Cmd+V)';
    });
    updateButtons();
  }

  // UI loads immediately; office.js is deferred and only needed when the user
  // clicks Inserir (messageParent), by which point it has loaded.
  boot();
})();
