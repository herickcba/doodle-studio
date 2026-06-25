/* ============================================================
   Big-canvas dialog controller.
   Reads the slide backdrop from localStorage (same origin as the
   task pane), lets the user draw large, then hands the drawing back
   to the task pane (which inserts it). Works standalone too.
   ============================================================ */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);

  function updateButtons() {
    $('insertBtn').disabled = Doodle.isEmpty();
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
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

  function finish(kind) {
    if (kind === 'inserted') {
      try { localStorage.setItem('doodle.result', JSON.stringify(Doodle.payload())); } catch (_) {}
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
    $('insertBtn').addEventListener('click', () => { if (!Doodle.isEmpty()) finish('inserted'); });
    $('cancelBtn').addEventListener('click', () => finish('cancel'));
    $('bgBtn').addEventListener('click', async () => {
      const res = await DoodleUI.loadBackdropFromClipboard();
      $('bgBtn').textContent = res.ok ? '✓ Fundo do slide' : '📋 Copie o slide e cole (Cmd+V)';
    });
    updateButtons();
  }

  if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(() => boot());
  else boot();
})();
