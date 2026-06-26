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
    const empty = Doodle.isEmpty();
    insertBtn.disabled = empty;
    $('gifBtn').disabled = empty;
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
    // Warm up slide-size detection (~8MB, background, once) when drawing starts,
    // so the first insert is already exact — never on open.
    if (!empty && OfficeBridge.inPowerPoint()) OfficeBridge.detectSlideSize();
  }

  function boot() {
    DoodleUI.setup({ canvasId: 'drawCanvas', onChange: updateButtons });
    $('undoBtn').addEventListener('click', () => Doodle.undo());
    $('redoBtn').addEventListener('click', () => Doodle.redo());
    $('clearBtn').addEventListener('click', () => Doodle.clear());
    $('bgBtn').addEventListener('click', loadBackground);
    insertBtn.addEventListener('click', insertOwnDrawing);
    $('gifBtn').addEventListener('click', insertGif);
    $('bigBtn').addEventListener('click', openBigCanvas);
    const gd = $('gifDur'), gdv = $('gifDurVal');
    gd.addEventListener('input', () => { gdv.textContent = ((+gd.value) / 10).toFixed(1) + 's'; });
    updateButtons();
  }

  async function insertGif() {
    if (Doodle.isEmpty()) { setStatus('Nada desenhado ainda.', 'warn'); return; }
    const duration = (+$('gifDur').value) / 10;
    $('gifBtn').disabled = true; insertBtn.disabled = true;
    setStatus('Gerando GIF…');
    try {
      const gif = await Doodle.makeAnimatedGif({ duration: duration, fps: 12, holdMs: 600 });
      if (!gif) { setStatus('Nada para animar.', 'warn'); return; }
      setStatus('Inserindo GIF…');
      const res = await OfficeBridge.insertDoodle(gif);
      setStatus(res.mode === 'inserted' ? 'GIF animado inserido ✓'
        : 'Modo navegador: GIF baixado (no PowerPoint, vai pro slide).', 'ok');
    } catch (e) {
      console.error(e);
      setStatus('Erro ao gerar/inserir o GIF: ' + (e && e.message ? e.message : e), 'warn');
    } finally {
      const empty = Doodle.isEmpty();
      $('gifBtn').disabled = empty; insertBtn.disabled = empty;
    }
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

  async function insertPNGs(pngs, successMsg) {
    if (!pngs || !pngs.length) { setStatus('Nada desenhado ainda.', 'warn'); return; }
    insertBtn.disabled = true;
    setStatus(pngs.length > 1 ? `Inserindo ${pngs.length} traços…` : 'Inserindo…');
    try {
      let mode = 'inserted';
      for (const png of pngs) {
        mode = (await OfficeBridge.insertDoodle(png)).mode;
        if (mode !== 'inserted') break;
      }
      setStatus(mode === 'inserted' ? successMsg
        : 'Modo navegador: PNG baixado (no PowerPoint, vai pro slide).', 'ok');
    } catch (e) {
      console.error(e);
      setStatus('Erro ao inserir: ' + (e && e.message ? e.message : e), 'warn');
    } finally {
      insertBtn.disabled = Doodle.isEmpty();
    }
  }

  function insertOwnDrawing() {
    insertPNGs(Doodle.exportPNGs(Doodle.state.insertSeparate), 'Inserido no slide ✓');
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
      const pngs = Doodle.renderExternalPNGs(payload.strokes, payload.config, payload.insertSeparate);
      insertPNGs(pngs, 'Inserido no slide (tela grande) ✓');
    });
  }

  boot();

  // office.js is loaded deferred (after this script), so poll until it's ready.
  // The UI and drawing already work; this only enables the Office features.
  function wireOffice() {
    if (typeof Office === 'undefined' || !Office.onReady) return false;
    Office.onReady((info) => {
      const inPP = info && info.host === Office.HostType.PowerPoint;
      $('hostLabel').textContent = inPP ? 'PowerPoint' : 'Modo navegador';
      if (inPP) setStatus('Copie o slide (Cmd+C) e clique "Fundo do slide" para vê-lo atrás.', '');
    });
    return true;
  }
  if (!wireOffice()) {
    const t = setInterval(() => { if (wireOffice()) clearInterval(t); }, 60);
    setTimeout(() => clearInterval(t), 10000);
  }
})();
