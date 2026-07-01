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
    if (sending) return;               // não reabilita botões no meio de um envio
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

  /* --- Envio pro PowerPoint ---------------------------------------
     office.js carrega deferred; se o usuário desenhar rápido e clicar
     Inserir antes de ele carregar, Office ainda não existe. Em vez de
     falhar em silêncio (o bug do "às vezes não insere"), esperamos a
     API ficar pronta (poll até 15s) com feedback visível, e um clique
     só dispara um envio (dedup). ------------------------------------ */
  let sending = false;

  function officeReady() {
    return typeof Office !== 'undefined' && Office.context && Office.context.ui
      && typeof Office.context.ui.messageParent === 'function';
  }

  function showStatus(msg, isError) {
    const el = $('dlgStatus');
    if (!el) return;
    el.textContent = msg || '';
    el.className = 'dlg-status' + (isError ? ' err' : '');
  }

  function setBusy(busy) {
    sending = busy;
    $('cancelBtn').disabled = busy;
    if (busy) {
      $('insertBtn').disabled = true;
      $('gifBtn').disabled = true;
    } else {
      updateButtons();
    }
  }

  function finish(kind, gif) {
    if (sending) return;                 // dedup: um clique por vez
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
    setBusy(true);
    if (kind === 'inserted') showStatus('Inserindo…');
    sendToHost(kind, p, 0);
  }

  function sendToHost(kind, p, waited) {
    if (!officeReady()) {
      if (waited >= 15000) {
        // resgate: guarda o desenho e orienta — nunca falha mudo
        try { localStorage.setItem('doodle.result', JSON.stringify(p)); } catch (_) {}
        showStatus('Sem conexão com o PowerPoint — feche esta janela e tente de novo.', true);
        setBusy(false);
        return;
      }
      if (waited >= 600) showStatus('Conectando ao PowerPoint…');
      setTimeout(() => sendToHost(kind, p, waited + 150), 150);
      return;
    }
    // Entrega o payload DENTRO da mensagem (sem localStorage); o host fecha a janela.
    try {
      Office.context.ui.messageParent(JSON.stringify({ kind: kind, payload: p }));
    } catch (e) {
      // payload grande demais p/ a mensagem: cai no localStorage + sinal curto
      let saved = false;
      try { localStorage.setItem('doodle.result', JSON.stringify(p)); saved = true; } catch (_) {}
      if (!saved && kind === 'inserted') {
        showStatus('Desenho grande demais para transferir — salve na biblioteca e insira pelo painel.', true);
        setBusy(false);
        return;
      }
      try { Office.context.ui.messageParent(kind); }
      catch (_) { showStatus('Falha ao comunicar com o PowerPoint — tente de novo.', true); setBusy(false); }
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
