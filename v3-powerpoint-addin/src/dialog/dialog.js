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

  function updateButtons() {
    const empty = Doodle.isEmpty();
    $('insertBtn').disabled = empty;
    $('gifBtn').disabled = empty;
    $('saveGifBtn').disabled = empty;
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
    $('saveAllBtn').disabled = empty;
    $('saveSelBtn').disabled = !Doodle.getSelectedStroke();
  }

  function saveToLibrary(strokes, label) {
    if (!strokes || !strokes.length) return;
    const thumb = Doodle.thumbnailOf(strokes);
    DoodleLibrary.save(label, Doodle.payloadOf(strokes), thumb);
  }

  // Save the current drawing as a GIF item (strokes + animation opts) AND a
  // static PNG of the final state, both into the shared library.
  function saveGifToLibrary() {
    const strokes = Doodle.state.strokes;
    if (!strokes.length) return;
    const thumb = Doodle.thumbnailOf(strokes);            // final-state preview
    const payload = Doodle.payload();
    const gif = {
      duration: (+$('gifDur').value) / 10,
      loop: $('gifLoop').checked,
      fps: +$('gifFps').value,
      holdMs: +$('gifHold').value,
      easing: $('gifEasing').value,
    };
    DoodleLibrary.save('GIF', payload, thumb, { kind: 'gif', gif });
    DoodleLibrary.save('Desenho', payload, thumb, { kind: 'png' });   // static final state
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

  function finish(kind, gif) {
    if (kind === 'inserted') {
      const p = Doodle.payload();
      if (gif) {
        p.asGif = true; p.gifLoop = $('gifLoop').checked;
        p.gifDuration = (+$('gifDur').value) / 10;
        p.gifEasing = $('gifEasing').value; p.gifFps = +$('gifFps').value; p.gifHold = +$('gifHold').value;
      }
      try { localStorage.setItem('doodle.result', JSON.stringify(p)); } catch (_) {}
    }
    if (window.opener) {            // standalone window.open()
      try { window.close(); } catch (_) {}
      return;
    }
    try { Office.context.ui.messageParent(kind); }
    catch (e) { try { window.close(); } catch (_) {} }
  }

  function boot() {
    loadBackdrop();
    DoodleUI.setup({ canvasId: 'drawCanvas', onChange: updateButtons });
    $('undoBtn').addEventListener('click', () => Doodle.undo());
    $('redoBtn').addEventListener('click', () => Doodle.redo());
    $('clearBtn').addEventListener('click', () => Doodle.clear());
    $('insertBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) finish('inserted', false); });
    $('gifBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) finish('inserted', true); });
    $('cancelBtn').addEventListener('click', () => finish('cancel'));
    const gd = $('gifDur'), gdv = $('gifDurVal');
    if (gd) gd.addEventListener('input', () => { gdv.textContent = ((+gd.value) / 10).toFixed(1) + 's'; });
    const gh = $('gifHold'), ghv = $('gifHoldVal');
    if (gh) gh.addEventListener('input', () => { ghv.textContent = gh.value + 'ms'; });
    $('zoomInBtn').addEventListener('click', () => setZoom(zoom * 1.2));
    $('zoomOutBtn').addEventListener('click', () => setZoom(zoom / 1.2));
    $('zoomFitBtn').addEventListener('click', () => setZoom(1));
    $('saveAllBtn').addEventListener('click', () => { saveToLibrary(Doodle.state.strokes, 'Desenho'); $('saveAllBtn').textContent = '✓ Salvo na biblioteca'; });
    $('saveSelBtn').addEventListener('click', () => { const s = Doodle.getSelectedStroke(); if (s) { saveToLibrary([s], 'Traço'); $('saveSelBtn').textContent = '✓ Traço salvo'; } });
    $('saveGifBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) { saveGifToLibrary(); $('saveGifBtn').textContent = '✓ GIF + PNG salvos'; } });
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
