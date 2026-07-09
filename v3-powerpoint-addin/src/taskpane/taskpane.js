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
    $('undoBtn').disabled = Doodle.state.strokes.length === 0;
    $('redoBtn').disabled = Doodle.state.redo.length === 0;
    $('saveAllBtn').disabled = empty;
    $('saveSelBtn').disabled = !Doodle.getSelectedStroke();
    $('saveGifBtn').disabled = empty;
    const gb = $('saveGifBtn'); if (gb.dataset.label) gb.textContent = gb.dataset.label;  // reset stale "✓" on new drawing
    // Warm up slide-size detection (~8MB, background, once) when drawing starts,
    // so the first insert is already exact — never on open.
    if (!empty && OfficeBridge.inPowerPoint()) OfficeBridge.detectSlideSize();
  }

  /* ---------------- Library (save + browse) ---------------- */
  function renderLibrary() {
    const grid = $('libGrid');
    const items = DoodleLibrary.list();
    grid.innerHTML = '';
    for (const it of items) {
      const isGif = it.kind === 'gif';
      const cell = document.createElement('div');
      cell.className = 'lib-item';
      const img = document.createElement('img');
      img.src = it.thumb; img.alt = it.name || 'Doodle';
      cell.appendChild(img);
      if (isGif) {
        const looping = !(it.gif && it.gif.loop === false);
        const badge = document.createElement('span');
        badge.className = 'lib-badge';
        badge.textContent = looping ? 'GIF ↻' : 'GIF 1×';
        badge.title = looping ? 'GIF em loop' : 'GIF sem loop';
        cell.appendChild(badge);
      }
      const del = document.createElement('button');
      del.className = 'lib-del'; del.textContent = '×'; del.title = 'Excluir';
      del.addEventListener('click', (e) => {
        e.stopPropagation();
        DoodleLibrary.remove(it.id);
        renderLibrary();
      });
      cell.appendChild(del);
      cell.title = isGif ? 'Inserir GIF no slide no tamanho salvo' : 'Inserir no slide no tamanho salvo';
      cell.addEventListener('click', () => isGif ? insertGifFromLibrary(it) : insertFromLibrary(it));
      grid.appendChild(cell);
    }
  }

  // Insert a saved item straight into the slide, at the resolution it was
  // saved (full 1920×1080 frame coords → scaled to the slide by insertDoodle).
  async function insertFromLibrary(it) {
    const strokes = it && it.payload && it.payload.strokes;
    if (!strokes || !strokes.length) { setStatus('Item da biblioteca vazio.', 'warn'); return; }
    const pngs = Doodle.renderExternalPNGs(strokes, it.payload.config, false);
    await insertPNGs(pngs, 'Inserido no slide da biblioteca ✓');
  }

  // Regenerate the saved GIF (strokes + stored animation opts) and insert it.
  async function insertGifFromLibrary(it) {
    const strokes = it && it.payload && it.payload.strokes;
    if (!strokes || !strokes.length) { setStatus('Item da biblioteca vazio.', 'warn'); return; }
    setStatus('Gerando GIF…');
    const o = it.gif || {};
    const separate = !!(it.payload && it.payload.insertSeparate);
    const gifs = await Doodle.renderExternalGifs(strokes, it.payload.config, separate, {
      duration: o.duration || 2.5, loop: o.loop !== false, fps: o.fps || 12,
      holdMs: o.holdMs != null ? o.holdMs : 600, easing: o.easing || 'linear',
    });
    await insertPNGs(gifs, gifs.length > 1 ? `${gifs.length} GIFs inseridos da biblioteca ✓` : 'GIF inserido da biblioteca ✓');
  }

  function saveToLibrary(strokes, label) {
    if (!strokes || !strokes.length) { setStatus('Nada para salvar.', 'warn'); return; }
    const thumb = Doodle.thumbnailOf(strokes);
    const id = DoodleLibrary.save(label, Doodle.payloadOf(strokes), thumb);
    if (!id) { setStatus('Não foi possível salvar (armazenamento cheio).', 'warn'); return; }
    renderLibrary();
    if (DoodleLibrary.wasPruned()) setStatus('Salvo ✓ — biblioteca cheia: itens antigos foram removidos.', 'warn');
    else setStatus('Salvo na biblioteca ✓', 'ok');
  }

  // Save the 3 versions (loop GIF + no-loop GIF + static PNG). The sidebar has
  // no animation controls, so it uses sensible defaults (tune in the big canvas).
  const DEFAULT_GIF = { duration: 2.5, fps: 12, holdMs: 600, easing: 'linear' };
  function saveGifToLibrary() {
    const strokes = Doodle.state.strokes;
    if (!strokes.length) { setStatus('Nada para salvar.', 'warn'); return; }
    const ok = DoodleLibrary.saveGifSet(Doodle.payload(), Doodle.thumbnailOf(strokes), DEFAULT_GIF);
    renderLibrary();
    if (ok) {
      $('saveGifBtn').textContent = '✓ 3 versões salvas';
      if (DoodleLibrary.wasPruned()) setStatus('GIF salvo ✓ — biblioteca cheia: itens antigos foram removidos.', 'warn');
      else setStatus('GIF salvo na biblioteca (loop, sem loop e PNG) ✓', 'ok');
    } else {
      setStatus('Biblioteca cheia — apague alguns itens e tente de novo.', 'warn');
    }
  }

  function toggleLibrary() {
    const sec = $('libSec');
    const open = sec.classList.toggle('collapsed') === false;
    $('libHead').setAttribute('aria-expanded', String(open));
  }

  function boot() {
    DoodleUI.setup({ canvasId: 'drawCanvas', onChange: updateButtons });
    $('undoBtn').addEventListener('click', () => Doodle.undo());
    $('redoBtn').addEventListener('click', () => Doodle.redo());
    $('clearBtn').addEventListener('click', () => Doodle.clear());
    $('bgBtn').addEventListener('click', loadBackground);
    insertBtn.addEventListener('click', insertOwnDrawing);
    $('bigBtn').addEventListener('click', openBigCanvas);
    $('libHead').addEventListener('click', toggleLibrary);
    $('saveAllBtn').addEventListener('click', () => saveToLibrary(Doodle.state.strokes, 'Desenho'));
    $('saveSelBtn').addEventListener('click', () => {
      const sel = Doodle.getSelectedStroke();
      if (sel) saveToLibrary([sel], 'Traço');
    });
    $('saveGifBtn').dataset.label = $('saveGifBtn').textContent;
    $('saveGifBtn').addEventListener('click', saveGifToLibrary);
    renderLibrary();
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

  async function insertPNGs(pngs, successMsg) {
    if (!pngs || !pngs.length) { setStatus('Nada desenhado ainda.', 'warn'); return; }
    insertBtn.disabled = true;
    setStatus('Inserindo…');
    let i = 0;
    try {
      let mode = 'inserted';
      for (const png of pngs) {
        i++;
        if (pngs.length > 1) setStatus(`Inserindo ${i}/${pngs.length}…`);
        mode = (await OfficeBridge.insertDoodle(png)).mode;
        if (mode !== 'inserted') break;
      }
      setStatus(mode === 'inserted' ? successMsg
        : 'Modo navegador: PNG baixado (no PowerPoint, vai pro slide).', 'ok');
    } catch (e) {
      console.error(e);
      const where = pngs.length > 1 ? ` (item ${i}/${pngs.length})` : '';
      setStatus('Erro ao inserir' + where + ': ' + (e && e.message ? e.message : e), 'warn');
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
    setStatus('Abrindo tela grande…');
    OfficeBridge.openDrawDialog(backdrop, async (payload) => {
      renderLibrary();   // the dialog may have saved into the shared library
      if (!payload || !payload.strokes || !payload.strokes.length) {
        setStatus('Tela grande fechada sem inserir.', '');
        return;
      }
      if (payload.asGif) {
        setStatus('Gerando GIF…');
        const gifs = await Doodle.renderExternalGifs(payload.strokes, payload.config, payload.insertSeparate,
          { duration: payload.gifDuration || 2.5, loop: payload.gifLoop !== false,
            fps: payload.gifFps || 12, holdMs: payload.gifHold != null ? payload.gifHold : 600,
            easing: payload.gifEasing || 'linear' });
        await insertPNGs(gifs, gifs.length > 1 ? `${gifs.length} GIFs inseridos (tela grande) ✓` : 'GIF inserido (tela grande) ✓');
      } else {
        const pngs = Doodle.renderExternalPNGs(payload.strokes, payload.config, payload.insertSeparate);
        await insertPNGs(pngs, 'Inserido no slide (tela grande) ✓');
      }
    }, {
      onOpen: () => setStatus('Tela grande aberta — desenhe e clique Inserir lá.', 'ok'),
      onError: (msg) => setStatus(msg, 'warn'),
    });
  }

  boot();

  // ---- Versão + aviso de update ----------------------------------------
  // O painel sempre roda a versão publicada (web); quem fica pra trás é a
  // FAIXA (.ppam instalado). O version.json diz a última versão lançada —
  // se este painel for mais antigo (webview cacheado), avisa pra recarregar.
  const CBA_VERSION = '1.5.0';
  (function versionLine() {
    const el = $('verLine');
    if (!el) return;
    el.textContent = 'CBA Studio v' + CBA_VERSION;
    fetch('https://doodle-studio-sigma.vercel.app/download/version.json', { cache: 'no-store' })
      .then((r) => (r.ok ? r.json() : null))
      .then((v) => {
        if (!v || !v.version || v.version === CBA_VERSION) return;
        el.classList.add('warn');
        el.innerHTML = 'CBA Studio v' + CBA_VERSION + ' · <b>v' + String(v.version).replace(/[^0-9.]/g, '') +
          ' disponível</b> — rode o instalador do site de novo.';
      })
      .catch(() => { /* offline/CORS: fica só a versão local */ });
  })();

  // office.js is loaded deferred, so it executes after this script.
  // The UI and drawing already work; this only enables the Office features.
  function wireOffice() {
    if (typeof Office === 'undefined' || !Office.onReady) return false;
    Office.onReady((info) => {
      const inPP = info && info.host === Office.HostType.PowerPoint;
      const hl = $('hostLabel'); if (hl) hl.textContent = inPP ? 'PowerPoint' : 'Modo navegador';
      if (inPP) setStatus('Copie o slide (Cmd+C) e clique "Fundo do slide" para vê-lo atrás.', '');
    });
    return true;
  }
  if (!wireOffice()) {
    const tag = document.getElementById('officeJs');
    if (tag) tag.addEventListener('load', wireOffice, { once: true });
    // rede/CDN pode falhar o load event silenciosamente — uma reconferência única
    setTimeout(() => { wireOffice(); }, 10000);
  }
})();
