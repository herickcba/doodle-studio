/* ============================================================
   Image module — generate / edit images with Nano Banana (Gemini),
   via our same-origin serverless functions (the API key stays on the
   server). Generated images live in an in-session gallery (in memory);
   style references persist (downscaled) via DoodleImgRefs.
   ============================================================ */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);

  const CONS = [
    { k: 'imagem', label: 'Imagem' }, { k: 'estilo', label: 'Estilo' },
    { k: 'universo', label: 'Universo' }, { k: 'textura', label: 'Textura' },
    { k: 'personagem', label: 'Personagem' }, { k: 'composicao', label: 'Composição' },
  ];

  let activeModel = 'nano2';
  let activeRef = null;          // ref id or null
  const chips = new Set();       // active consistency keys
  let gallery = [];              // [{ id, dataUrl, prompt, model }]
  let busy = false;

  function setStatus(msg, kind) {
    const el = $('imgStatus'); if (!el) return;
    el.textContent = msg; el.className = 'status' + (kind ? ' ' + kind : '');
  }
  function setBusy(b) {
    busy = b;
    $('imgGenBtn').disabled = b;
    $('imgGenBtn').textContent = b ? 'Gerando…' : '✦ Melhorar + Gerar';
  }

  async function api(path, body) {
    const resp = await fetch(path, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || ('Erro ' + resp.status));
    return data;
  }

  // Dispatch: if a local key is set (test mode) call Gemini directly from the
  // browser; otherwise call our serverless functions (key on the server).
  function localKey() { return (window.DoodleGemini && DoodleGemini.getKey()) || ''; }
  function doImprove(p) { const k = localKey(); return k ? DoodleGemini.improvePrompt(k, p) : api('/api/improve-prompt', p); }
  function doGenerate(p) { const k = localKey(); return k ? DoodleGemini.generateImage(k, p) : api('/api/generate-image', p); }
  function doEdit(p) { const k = localKey(); return k ? DoodleGemini.editImage(k, p) : api('/api/edit-image', p); }

  // A light (JPEG, <=1280px) copy of an image — used only for DISPLAY inside
  // the edit dialog, so we never push a multi-MB image through localStorage
  // (which would blow the quota and open the editor blank). The actual edit
  // uses the full-res base; the overlay is scaled back up to it.
  function downscaleForDisplay(dataUrl, maxPx) {
    maxPx = maxPx || 1280;
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const s = Math.min(1, maxPx / Math.max(img.naturalWidth, img.naturalHeight));
        const w = Math.max(1, Math.round(img.naturalWidth * s)), h = Math.max(1, Math.round(img.naturalHeight * s));
        const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
        cv.getContext('2d').drawImage(img, 0, 0, w, h);
        try { resolve(cv.toDataURL('image/jpeg', 0.85)); } catch (_) { resolve(dataUrl); }
      };
      img.onerror = () => resolve(dataUrl);
      img.src = dataUrl;
    });
  }

  // Draw `overlayUrl` (a transparent scribble layer) on top of `baseUrl`.
  function compositeOver(baseUrl, overlayUrl) {
    return new Promise((resolve) => {
      const b = new Image();
      b.onload = () => {
        const cv = document.createElement('canvas'); cv.width = b.naturalWidth; cv.height = b.naturalHeight;
        const cx = cv.getContext('2d'); cx.drawImage(b, 0, 0);
        const o = new Image();
        o.onload = () => { cx.drawImage(o, 0, 0, cv.width, cv.height); resolve(cv.toDataURL('image/png')); };
        o.onerror = () => resolve(baseUrl);
        o.src = overlayUrl;
      };
      b.onerror = () => resolve(baseUrl);
      b.src = baseUrl;
    });
  }

  // Cover-fit any generated image onto an exact 1920x1080 canvas.
  function rescale1080(dataUrl) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const cv = document.createElement('canvas'); cv.width = 1920; cv.height = 1080;
        const cx = cv.getContext('2d');
        const s = Math.max(1920 / img.naturalWidth, 1080 / img.naturalHeight);
        const w = img.naturalWidth * s, h = img.naturalHeight * s;
        cx.drawImage(img, (1920 - w) / 2, (1080 - h) / 2, w, h);
        resolve(cv.toDataURL('image/png'));
      };
      img.onerror = () => resolve(dataUrl);
      img.src = dataUrl;
    });
  }

  function renderChips() {
    const wrap = $('imgChips'); wrap.innerHTML = '';
    for (const c of CONS) {
      const b = document.createElement('button');
      b.className = 'chip' + (chips.has(c.k) ? ' on' : '');
      b.textContent = c.label;
      b.addEventListener('click', () => { chips.has(c.k) ? chips.delete(c.k) : chips.add(c.k); renderChips(); });
      wrap.appendChild(b);
    }
  }

  function renderRefs() {
    const wrap = $('imgRefs'); wrap.innerHTML = '';
    const items = DoodleImgRefs.list();
    if (activeRef && !items.find((x) => x.id === activeRef)) activeRef = null;
    for (const it of items) {
      const cell = document.createElement('div');
      cell.className = 'ref-item' + (activeRef === it.id ? ' on' : '');
      cell.title = 'Usar como referência de estilo';
      const img = document.createElement('img'); img.src = it.thumb; cell.appendChild(img);
      const del = document.createElement('button');
      del.className = 'lib-del'; del.textContent = '×'; del.title = 'Excluir';
      del.addEventListener('click', (e) => { e.stopPropagation(); DoodleImgRefs.remove(it.id); renderRefs(); });
      cell.appendChild(del);
      cell.addEventListener('click', () => { activeRef = (activeRef === it.id) ? null : it.id; renderRefs(); });
      wrap.appendChild(cell);
    }
  }

  function galleryCard(item) {
    const card = document.createElement('div');
    card.className = 'gen-card';
    const img = document.createElement('img'); img.src = item.dataUrl; card.appendChild(img);
    const del = document.createElement('button');
    del.className = 'lib-del'; del.textContent = '×'; del.title = 'Excluir da biblioteca';
    del.addEventListener('click', (e) => { e.stopPropagation(); removeItem(item); });
    card.appendChild(del);
    const row = document.createElement('div'); row.className = 'gen-actions';
    const mk = (label, title, fn) => { const b = document.createElement('button'); b.className = 'btn'; b.textContent = label; b.title = title; b.addEventListener('click', fn); return b; };
    row.appendChild(mk('Inserir', 'Inserir no slide (preenche)', () => insert(item)));
    row.appendChild(mk('Editar', 'Editar na tela grande', () => edit(item)));
    row.appendChild(mk('★ Ref', 'Salvar como referência de estilo', () => saveRef(item)));
    card.appendChild(row);
    return card;
  }
  function renderGallery() {
    const wrap = $('imgGallery'); wrap.innerHTML = '';
    for (const it of gallery) wrap.appendChild(galleryCard(it));
  }

  // Add a generated/edited image to the persistent library + the live list.
  async function addToGallery(item) {
    gallery.unshift(item);
    renderGallery();
    try { if (window.DoodleGallery) await DoodleGallery.add(item); } catch (_) {}
  }
  async function removeItem(item) {
    gallery = gallery.filter((g) => g.id !== item.id);
    renderGallery();
    try { if (window.DoodleGallery) await DoodleGallery.remove(item.id); } catch (_) {}
    setStatus('Imagem removida da biblioteca.', '');
  }

  async function generate() {
    if (busy) return;
    const prompt = $('imgPrompt').value.trim();
    if (!prompt) { setStatus('Escreva um prompt.', 'warn'); return; }
    const ref = activeRef ? DoodleImgRefs.get(activeRef) : null;
    const refB64 = ref ? ref.thumb.split(',')[1] : null;
    const refMime = ref ? ref.thumb.substring(5, ref.thumb.indexOf(';')) : null;
    const n = Math.max(1, Math.min(4, +$('imgBatch').value || 1));
    setBusy(true);
    try {
      setStatus('Melhorando o prompt…');
      const improved = await doImprove({ prompt, consistency: [...chips], hasRef: !!ref });
      const fp = improved.finalPrompt || prompt;
      for (let i = 1; i <= n; i++) {
        setStatus('Gerando ' + i + '/' + n + '…');
        const p = n > 1 ? fp + ' — variação ' + i + ', enquadramento ligeiramente diferente' : fp;
        const r = await doGenerate({ prompt: p, model: activeModel, refImageBase64: refB64, refMimeType: refMime });
        const full = await rescale1080('data:' + (r.mimeType || 'image/png') + ';base64,' + r.imageBase64);
        await addToGallery({ id: 'g' + Date.now() + '_' + i, dataUrl: full, prompt: prompt, model: activeModel, ts: Date.now() });
      }
      setStatus(n > 1 ? n + ' imagens geradas ✓' : 'Imagem gerada ✓', 'ok');
    } catch (e) {
      setStatus(e && e.message ? e.message : 'Erro ao gerar.', 'warn');
    } finally { setBusy(false); }
  }

  async function insert(item) {
    if (!window.OfficeBridge || !OfficeBridge.inPowerPoint()) { setStatus('Inserir funciona dentro do PowerPoint.', 'warn'); return; }
    setStatus('Inserindo no slide…');
    try { await OfficeBridge.insertImage(item.dataUrl); setStatus('Inserida no slide ✓', 'ok'); }
    catch (e) { setStatus(e && e.message ? e.message : 'Erro ao inserir.', 'warn'); }
  }

  async function saveRef(item) {
    await DoodleImgRefs.add(item.dataUrl, 'Gerada');
    renderRefs();
    setStatus('Salva como referência de estilo ✓', 'ok');
  }

  async function edit(item) {
    if (!window.OfficeBridge || !OfficeBridge.openImageEditDialog) { setStatus('Edição indisponível aqui.', 'warn'); return; }
    const display = await downscaleForDisplay(item.dataUrl);   // light copy for the editor canvas
    OfficeBridge.openImageEditDialog(display, async (result) => {
      if (!result || !result.prompt) { setStatus('Edição cancelada.', 'warn'); return; }
      setBusy(true);
      setStatus('Aplicando edição…');
      try {
        const base = item.dataUrl;
        // composite the (small) scribble overlay over the full-res base here,
        // so the model sees the marks in context without blowing localStorage.
        const markup = result.markupDataUrl ? await compositeOver(base, result.markupDataUrl) : null;
        const r = await doEdit({
          prompt: result.prompt, model: activeModel,
          baseImageBase64: base.split(',')[1], baseMimeType: base.substring(5, base.indexOf(';')),
          markupImageBase64: markup ? markup.split(',')[1] : null,
          markupMimeType: 'image/png',
        });
        const full = await rescale1080('data:' + (r.mimeType || 'image/png') + ';base64,' + r.imageBase64);
        await addToGallery({ id: 'e' + Date.now(), dataUrl: full, prompt: result.prompt, model: activeModel, ts: Date.now() });
        setStatus('Edição aplicada ✓', 'ok');
      } catch (e) { setStatus(e && e.message ? e.message : 'Erro ao editar.', 'warn'); }
      finally { setBusy(false); }
    });
  }

  function refreshKeyUI() {
    const has = !!localKey();
    const dot = $('keyDot'), mode = $('keyMode');
    if (dot) dot.classList.toggle('on', has);
    if (mode) mode.textContent = has ? 'teste (chave local)' : 'serviço';
  }

  function boot() {
    if (!$('imgGenBtn')) return;
    renderChips(); renderRefs(); renderGallery();
    // load the persistent library (IndexedDB) — survives reopening the add-in
    if (window.DoodleGallery) DoodleGallery.list().then((items) => { gallery = items || []; renderGallery(); }).catch(() => {});
    // key (test mode) wiring
    if ($('geminiKey')) $('geminiKey').value = localKey();
    if ($('saveKeyBtn')) $('saveKeyBtn').addEventListener('click', () => {
      DoodleGemini.setKey($('geminiKey').value); refreshKeyUI();
      setStatus(localKey() ? 'Chave salva — modo teste ativo ✓' : 'Chave vazia — usando o serviço.', 'ok');
    });
    if ($('clearKeyBtn')) $('clearKeyBtn').addEventListener('click', () => {
      DoodleGemini.clearKey(); $('geminiKey').value = ''; refreshKeyUI(); setStatus('Chave removida — usando o serviço.', '');
    });
    refreshKeyUI();
    document.querySelectorAll('#imgModel .seg-btn').forEach((b) => {
      b.addEventListener('click', () => {
        activeModel = b.dataset.model;
        document.querySelectorAll('#imgModel .seg-btn').forEach((x) => x.classList.toggle('active', x === b));
      });
    });
    $('imgGenBtn').addEventListener('click', generate);
    $('imgRefBtn').addEventListener('click', () => $('imgRefFile').click());
    $('imgRefFile').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      const rd = new FileReader();
      rd.onload = async () => { await DoodleImgRefs.add(rd.result, f.name); renderRefs(); setStatus('Referência adicionada ✓', 'ok'); };
      rd.readAsDataURL(f);
      e.target.value = '';
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
