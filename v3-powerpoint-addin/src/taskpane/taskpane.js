/* ============================================================
   Task pane controller — orchestrates Office, hosts the quick-draw
   canvas, and launches the big-canvas dialog. Works standalone too.
   ============================================================ */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const insertBtn = $('insertBtn');
  let currentSlideImg = '';

  function setStatus(msg, kind) {
    $('status').textContent = msg;
    $('status').className = 'status' + (kind ? ' ' + kind : '');
  }

  function updateButtons() {
    insertBtn.disabled = Doodle.isEmpty();
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
  }

  function boot() {
    DoodleUI.setup({ canvasId: 'drawCanvas', onChange: updateButtons });
    $('undoBtn').addEventListener('click', () => Doodle.undo());
    $('redoBtn').addEventListener('click', () => Doodle.redo());
    $('clearBtn').addEventListener('click', () => Doodle.clear());
    $('bgBtn').addEventListener('click', loadBackground);
    insertBtn.addEventListener('click', insertOwnDrawing);
    $('bigBtn').addEventListener('click', openBigCanvas);
    updateButtons();
  }

  // Try, in order: the (preview) slide-image API → the clipboard (slide
  // copied by the user) → guidance to copy + paste.
  async function loadBackground() {
    if (OfficeBridge.inPowerPoint()) {
      const url = await OfficeBridge.getActiveSlideImage();
      if (url) { DoodleUI.setBackdrop(url); currentSlideImg = url; setStatus('Slide carregado ✓', 'ok'); return; }
    }
    const res = await DoodleUI.loadBackdropFromClipboard();
    if (res.ok) { setStatus('Fundo carregado da área de transferência ✓', 'ok'); return; }
    setStatus('Copie o slide (clique nele → Cmd+C) e clique aqui de novo — ou cole com Cmd+V.', 'warn');
  }

  async function insertPNG(png, successMsg) {
    if (!png) { setStatus('Nada desenhado ainda.', 'warn'); return; }
    try {
      const res = await OfficeBridge.insertDoodle(png);
      setStatus(res.mode === 'inserted' ? successMsg
        : 'Modo navegador: PNG baixado (no PowerPoint, vai pro slide).', 'ok');
    } catch (e) {
      console.error(e);
      setStatus('Erro ao inserir: ' + (e && e.message ? e.message : e), 'warn');
    }
  }

  function insertOwnDrawing() {
    insertPNG(Doodle.exportTransparentPNG(), 'Inserido no slide ✓');
  }

  async function openBigCanvas() {
    // Pass whatever backdrop we already have (API or pasted) into the dialog.
    const backdrop = window.DoodleBackdrop || currentSlideImg || '';
    setStatus('Tela grande aberta — desenhe e clique Inserir lá.', 'ok');
    OfficeBridge.openDrawDialog(backdrop, (payload) => {
      if (!payload || !payload.strokes || !payload.strokes.length) {
        setStatus('Tela grande fechada sem desenho.', 'warn');
        return;
      }
      const png = Doodle.renderExternalPNG(payload.strokes, payload.config);
      insertPNG(png, 'Inserido no slide (tela grande) ✓');
    });
  }

  boot();

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => {
      if (info.host === Office.HostType.PowerPoint) {
        $('hostLabel').textContent = 'PowerPoint';
        setStatus('Copie o slide (Cmd+C) e clique "Fundo do slide" para vê-lo atrás.', '');
      } else {
        $('hostLabel').textContent = 'Modo navegador';
      }
    });
  } else {
    $('hostLabel').textContent = 'Modo navegador';
  }
})();
