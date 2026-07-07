/* ============================================================
   Office bridge — talks to PowerPoint via Office.js.
   Degrades gracefully when not running inside PowerPoint
   (standalone browser): getActiveSlideImage -> null,
   insertDoodle -> triggers a PNG download instead.
   Exposes window.OfficeBridge.
   ============================================================ */
(function (global) {
  'use strict';

  // Standard 16:9 slide size in points (13.333in x 7.5in).
  const SLIDE_W_PT = 960, SLIDE_H_PT = 540;

  function inPowerPoint() {
    return typeof Office !== 'undefined'
      && Office.context
      && Office.context.host === Office.HostType.PowerPoint
      && typeof PowerPoint !== 'undefined';
  }

  /* Render the currently active slide to a PNG data URL.
     Uses preview APIs (getActiveSlide / getImageAsBase64); returns
     null if unavailable so the caller can fall back to a blank frame. */
  async function getActiveSlideImage() {
    if (!inPowerPoint()) return null;
    try {
      return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getActiveSlide();
        const img = slide.getImageAsBase64({});   // preview API
        await context.sync();
        return 'data:image/png;base64,' + img.value;
      });
    } catch (e) {
      console.warn('[OfficeBridge] getActiveSlideImage unavailable:', e && e.message);
      return null;
    }
  }

  /* ---- Slide size (points) ----
     The exact size lives in <p:sldSz> of ppt/presentation.xml, but the only
     way to read it on older builds is getFileAsync, which downloads the WHOLE
     presentation (can be hundreds of MB). So we NEVER block on it:
       - inserts use a synchronous best-known size (memory → localStorage cache
         → 16:9 default) so they're instant;
       - the exact size is detected ONCE in the background and cached in
         localStorage (per document), so reopening never downloads again. */
  let _slideSize = null;     // {w,h} once known (exact)
  let _detecting = null;     // in-flight detection promise (dedup)
  let _detectTried = false;  // stop retrying the heavy download after one attempt

  function _sizeKey() {
    let url = '';
    try { url = (Office.context.document && Office.context.document.url) || ''; } catch (_) {}
    return 'doodle.slideSize::' + url;
  }
  function _loadCachedSize() {
    try {
      const v = localStorage.getItem(_sizeKey());
      if (v) { const o = JSON.parse(v); if (o && o.w && o.h) return o; }
    } catch (_) {}
    return null;
  }

  const SLICE = 4194304; // 4 MB

  function _getFile() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: SLICE }, (res) => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error);
      });
    });
  }
  function _getSlice(file, i) {
    return new Promise((resolve, reject) => {
      file.getSliceAsync(i, (s) => {
        s.status === Office.AsyncResultStatus.Succeeded ? resolve(s.value.data) : reject(s.error);
      });
    });
  }
  function _toBytes(d) {
    if (typeof d === 'string') { const b = atob(d), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
    return d instanceof Uint8Array ? d : Uint8Array.from(d);
  }
  /* Read an absolute byte range [start, start+len) by fetching only the slices
     that overlap it — not the whole file. `cache` (Map opcional) evita rebaixar
     a mesma fatia em leituras repetidas (auditoria lê ~90 rels vizinhos). */
  async function _readRange(file, start, len, cache) {
    const end = start + len, first = Math.floor(start / SLICE), last = Math.floor((end - 1) / SLICE);
    const out = new Uint8Array(len); let w = 0;
    for (let i = first; i <= last; i++) {
      let bytes = cache && cache.get(i);
      if (!bytes) {
        bytes = _toBytes(await _getSlice(file, i));
        if (cache) { if (cache.size >= 24) cache.clear(); cache.set(i, bytes); }
      }
      const sStart = i * SLICE;
      const from = Math.max(start, sStart) - sStart;
      const to = Math.min(end, sStart + bytes.length) - sStart;
      const chunk = bytes.subarray(from, to);
      out.set(chunk, w); w += chunk.length;
    }
    return out.subarray(0, w);
  }

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Response(new Blob([bytes]).stream().pipeThrough(ds));
    return new Uint8Array(await stream.arrayBuffer());
  }

  /* Read one ZIP entry by fetching only: the tail (End-Of-Central-Directory +
     central directory) and the entry's local header + data. ~2-3 slices total. */
  async function _readEntryViaSlices(file, name) {
    const u32 = (b, o) => b[o] | (b[o+1]<<8) | (b[o+2]<<16) | (b[o+3]*0x1000000);
    const u16 = (b, o) => b[o] | (b[o+1]<<8);
    const size = file.size, nSlices = file.sliceCount;
    // 1) tail slice -> EOCD -> central directory location
    const lastStart = (nSlices - 1) * SLICE;
    const tail = await _readRange(file, lastStart, size - lastStart);
    let e = -1;
    for (let i = tail.length - 22; i >= 0; i--) { if (u32(tail, i) === 0x06054b50) { e = i; break; } }
    if (e < 0) throw new Error('no EOCD');
    const cdOff = u32(tail, e + 16), cdSize = u32(tail, e + 12);
    // 2) central directory -> find the entry
    const cd = await _readRange(file, cdOff, cdSize);
    let off = 0, hit = null;
    while (off + 46 <= cd.length) {
      if (u32(cd, off) !== 0x02014b50) break;
      const method = u16(cd, off + 10), csize = u32(cd, off + 20);
      const nlen = u16(cd, off + 28), elen = u16(cd, off + 30), clen = u16(cd, off + 32), lho = u32(cd, off + 42);
      if (new TextDecoder().decode(cd.subarray(off + 46, off + 46 + nlen)) === name) { hit = { method, csize, lho }; break; }
      off += 46 + nlen + elen + clen;
    }
    if (!hit) throw new Error('entry not found: ' + name);
    // 3) local header + compressed data
    const head = await _readRange(file, hit.lho, 30);
    const dataStart = hit.lho + 30 + u16(head, 26) + u16(head, 28);
    const cdata = await _readRange(file, dataStart, hit.csize);
    const out = hit.method === 0 ? cdata : await inflateRaw(cdata);
    return new TextDecoder().decode(out);
  }

  /* Instant: best-known size, never downloads. exact=false means it's the
     16:9 fallback (insert will be off-scale only on non-16:9 / custom decks). */
  function getSlideSizeSync() {
    if (_slideSize) return { w: _slideSize.w, h: _slideSize.h, exact: true };
    const c = _loadCachedSize();
    if (c) { _slideSize = c; return { w: c.w, h: c.h, exact: true }; }
    return { w: SLIDE_W_PT, h: SLIDE_H_PT, exact: false };
  }

  /* Background, one-time, deduped: read the real size from the .pptx and cache
     it. Heavy on big decks, but never blocks an insert or the open. */
  function detectSlideSize() {
    if (_slideSize) return Promise.resolve(_slideSize);
    if (_detecting) return _detecting;
    if (_detectTried || !inPowerPoint()) return Promise.resolve(null);
    _detecting = (async () => {
      let file = null;
      try {
        file = await _getFile();                                   // does not download slices yet
        const xml = await _readEntryViaSlices(file, 'ppt/presentation.xml'); // ~2-3 slices only
        const m = xml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        if (m) {
          _slideSize = { w: +m[1] / 12700, h: +m[2] / 12700 };
          try { localStorage.setItem(_sizeKey(), JSON.stringify(_slideSize)); } catch (_) {}
        }
      } catch (e) {
        console.warn('[OfficeBridge] slide size detect failed:', e && e.message);
      } finally {
        if (file) { try { file.closeAsync(() => {}); } catch (_) {} }
        _detectTried = true; _detecting = null;
      }
      return _slideSize;
    })();
    return _detecting;
  }

  /* ---------- Auditoria de imagens (leitura do zip por slices) ---------- */

  const _zu32 = (b, o) => b[o] | (b[o+1]<<8) | (b[o+2]<<16) | (b[o+3]*0x1000000);
  const _zu16 = (b, o) => b[o] | (b[o+1]<<8);

  /* Read the whole central directory once -> list of entries. */
  async function _readCentralDir(file, cache) {
    const size = file.size, lastStart = (file.sliceCount - 1) * SLICE;
    const tail = await _readRange(file, lastStart, size - lastStart, cache);
    let e = -1;
    for (let i = tail.length - 22; i >= 0; i--) { if (_zu32(tail, i) === 0x06054b50) { e = i; break; } }
    if (e < 0) throw new Error('no EOCD');
    const cdOff = _zu32(tail, e + 16), cdSize = _zu32(tail, e + 12);
    const cd = await _readRange(file, cdOff, cdSize, cache);
    const entries = []; let off = 0;
    while (off + 46 <= cd.length) {
      if (_zu32(cd, off) !== 0x02014b50) break;
      const method = _zu16(cd, off + 10), csize = _zu32(cd, off + 20), usize = _zu32(cd, off + 24);
      const nlen = _zu16(cd, off + 28), elen = _zu16(cd, off + 30), clen = _zu16(cd, off + 32), lho = _zu32(cd, off + 42);
      entries.push({ name: new TextDecoder().decode(cd.subarray(off + 46, off + 46 + nlen)), method, csize, usize, lho });
      off += 46 + nlen + elen + clen;
    }
    return entries;
  }

  async function _entryBytes(file, ent, cache) {
    const head = await _readRange(file, ent.lho, 30, cache);
    const dataStart = ent.lho + 30 + _zu16(head, 26) + _zu16(head, 28);
    const cdata = await _readRange(file, dataStart, ent.csize, cache);
    return ent.method === 0 ? cdata : await inflateRaw(cdata);
  }
  const _entryText = async (file, ent, cache) => new TextDecoder().decode(await _entryBytes(file, ent, cache));

  /* List pictures in the deck: name, extension, compressed bytes and the
     slides that use each one (position + sldId for navigation).
     progress(msg) is optional. Heavy on big decks — only on demand. */
  async function getImageAudit(progress) {
    if (!inPowerPoint()) throw new Error('Disponível apenas dentro do PowerPoint.');
    const say = progress || function () {};
    let file = null;
    try {
      say('Lendo o arquivo…');
      file = await _getFile();
      const cache = new Map();
      const entries = await _readCentralDir(file, cache);
      const byName = {}; entries.forEach((en) => { byName[en.name] = en; });

      const media = entries.filter((en) => /^ppt\/media\//.test(en.name) &&
        /\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg|webp)$/i.test(en.name));
      if (!media.length) return { total: 0, items: [] };

      // ordem real dos slides: presentation.xml (sldIdLst) + seus rels
      say('Mapeando slides…');
      let fileToPos = {}; // 'slide3.xml' -> {pos, sldId}
      try {
        const pXml = await _entryText(file, byName['ppt/presentation.xml'], cache);
        const pRels = await _entryText(file, byName['ppt/_rels/presentation.xml.rels'], cache);
        const ridToFile = {};
        pRels.replace(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="slides\/([^"]+)"[^>]*>/g,
          (_, rid, f) => { ridToFile[rid] = f; return _; });
        let pos = 0;
        pXml.replace(/<p:sldId [^>]*id="(\d+)"[^>]*r:id="([^"]+)"/g, (_, sldId, rid) => {
          pos += 1;
          if (ridToFile[rid]) fileToPos[ridToFile[rid]] = { pos, sldId: +sldId };
          return _;
        });
      } catch (_) { fileToPos = {}; }

      // rels de cada slide: media -> slides que usam
      const usedBy = {}; // 'image3.png' -> [{pos, sldId, slideFile}]
      const relEntries = entries.filter((en) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(en.name));
      for (let i = 0; i < relEntries.length; i++) {
        say('Cruzando imagens ' + (i + 1) + '/' + relEntries.length + '…');
        const slideFile = relEntries[i].name.replace(/^ppt\/slides\/_rels\//, '').replace(/\.rels$/, '');
        let xml = '';
        try { xml = await _entryText(file, relEntries[i], cache); } catch (_) { continue; }
        xml.replace(/Target="\.\.\/media\/([^"]+)"/g, (_, m) => {
          (usedBy[m] = usedBy[m] || []).push(Object.assign({ slideFile }, fileToPos[slideFile] || {}));
          return _;
        });
      }

      const items = media.map((en) => {
        const base = en.name.replace(/^ppt\/media\//, '');
        return {
          name: en.name, base,
          ext: (base.match(/\.([a-z0-9]+)$/i) || [,''])[1].toLowerCase(),
          bytes: en.csize, usize: en.usize,
          slides: (usedBy[base] || []).filter((s, i2, a) => a.findIndex((x) => x.slideFile === s.slideFile) === i2),
        };
      }).sort((a, b) => b.bytes - a.bytes);
      return { total: items.reduce((s, it) => s + it.bytes, 0), items };
    } finally {
      if (file) { try { file.closeAsync(() => {}); } catch (_) {} }
    }
  }

  /* Navigate to a slide by its OOXML sldId (goToByIdAsync). Resolves false
     when the API is unavailable (list keeps working without navigation). */
  function goToSlide(sldId) {
    return new Promise((resolve) => {
      try {
        Office.context.document.goToByIdAsync(sldId, Office.GoToType.Slide,
          (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded));
      } catch (_) { resolve(false); }
    });
  }

  /* Re-encode one audited image at its DISPLAYED size and insert it on the
     same slide, same position (the original must be deleted by hand — the
     API can't replace it). Returns {smaller:false, bytes} when re-encoding
     wouldn't save space. */
  async function optimizeImage(item, progress) {
    if (!inPowerPoint()) throw new Error('Disponível apenas dentro do PowerPoint.');
    const say = progress || function () {};
    let file = null;
    let src, place = null;
    try {
      say('Lendo a imagem…');
      file = await _getFile();
      const cache = new Map();
      const entries = await _readCentralDir(file, cache);
      const byName = {}; entries.forEach((en) => { byName[en.name] = en; });
      const ent = byName[item.name];
      if (!ent) throw new Error('Imagem não encontrada (o arquivo mudou? Analise de novo).');
      src = await _entryBytes(file, ent, cache);

      // tamanho/posição exibidos: acha o <p:pic> que usa esse media no 1º slide
      const use = item.slides && item.slides[0];
      if (use && byName['ppt/slides/' + use.slideFile]) {
        try {
          const rels = await _entryText(file, byName['ppt/slides/_rels/' + use.slideFile + '.rels'], cache);
          let rid = null;
          rels.replace(new RegExp('<Relationship [^>]*Id="([^"]+)"[^>]*Target="\\.\\./media/' +
            item.base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"', 'g'), (_, r) => { rid = r; return _; });
          if (rid) {
            const sx = await _entryText(file, byName['ppt/slides/' + use.slideFile], cache);
            const pic = sx.split('<p:pic>').find((seg) => seg.indexOf('r:embed="' + rid + '"') >= 0);
            if (pic) {
              const mOff = pic.match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
              const mExt = pic.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
              if (mExt) place = {
                left: mOff ? +mOff[1] / 12700 : 0, top: mOff ? +mOff[2] / 12700 : 0,
                w: +mExt[1] / 12700, h: +mExt[2] / 12700,
              };
            }
          }
        } catch (_) {}
      }
    } finally {
      if (file) { try { file.closeAsync(() => {}); } catch (_) {} }
    }

    say('Reprocessando…');
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[item.ext];
    if (!mime) throw new Error('Formato .' + item.ext + ' não dá para converter no navegador.');
    const bmp = await createImageBitmap(new Blob([src], { type: mime }));

    // alvo = tamanho exibido em px @2x (nitidez em retina), cap 1920, nunca upscale
    let tw = bmp.width, th = bmp.height;
    if (place && place.w > 0) {
      tw = Math.round(place.w / 72 * 96 * 2); th = Math.round(tw * bmp.height / bmp.width);
    }
    const cap = 1920 / Math.max(tw, th);
    if (cap < 1) { tw = Math.round(tw * cap); th = Math.round(th * cap); }
    if (tw >= bmp.width) { tw = bmp.width; th = bmp.height; }

    const cv = document.createElement('canvas'); cv.width = tw; cv.height = th;
    cv.getContext('2d').drawImage(bmp, 0, 0, tw, th);

    // PNG só se tiver alfa de verdade; senão JPG 0.85
    let hasAlpha = false;
    if (item.ext === 'png' || item.ext === 'webp' || item.ext === 'gif') {
      const d = cv.getContext('2d').getImageData(0, 0, tw, th).data;
      for (let i = 3; i < d.length; i += 64) { if (d[i] < 250) { hasAlpha = true; break; } }
    }
    const outUrl = cv.toDataURL(hasAlpha ? 'image/png' : 'image/jpeg', 0.85);
    const outBytes = Math.floor((outUrl.length - outUrl.indexOf(',') - 1) * 3 / 4);
    if (outBytes >= item.bytes) return { smaller: false, bytes: outBytes };

    say('Inserindo a versão otimizada…');
    const use = item.slides && item.slides[0];
    if (use && use.sldId) { await goToSlide(use.sldId); await new Promise((r) => setTimeout(r, 350)); }
    const base64 = outUrl.split(',')[1];
    const pos = place || { left: 40, top: 40, w: 300, h: 300 * th / tw };
    const ok = await new Promise((resolve) => {
      try {
        Office.context.document.setSelectedDataAsync(base64, {
          coercionType: Office.CoercionType.Image,
          imageLeft: pos.left, imageTop: pos.top, imageWidth: pos.w, imageHeight: pos.h,
        }, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? true : res));
      } catch (e) { resolve(e); }
    });
    if (ok !== true) throw new Error((ok && ok.error && ok.error.message) || 'Falha ao inserir.');
    return { smaller: true, bytes: outBytes, saved: item.bytes - outBytes, format: hasAlpha ? 'PNG' : 'JPG' };
  }

  /* Insert the doodle PNG onto the active slide, positioned so it lands
     where it was drawn over the slide backdrop.
     png = { dataUrl, bbox:{x,y,w,h}, frame:{w,h} } from Doodle.exportTransparentPNG(). */
  async function insertDoodle(png) {
    if (!png) throw new Error('Nada para inserir.');

    if (!inPowerPoint()) {
      // Standalone fallback: download the PNG.
      const a = document.createElement('a');
      a.href = png.dataUrl;
      a.download = 'doodle.png';
      a.click();
      return { mode: 'download' };
    }

    const base64 = png.dataUrl.split(',')[1];
    const slide = getSlideSizeSync();              // instant, never downloads
    if (!slide.exact) detectSlideSize();           // refine in background for next inserts
    const left = png.bbox.x / png.frame.w * slide.w;
    const top = png.bbox.y / png.frame.h * slide.h;
    const width = png.bbox.w / png.frame.w * slide.w;
    const height = png.bbox.h / png.frame.h * slide.h;

    // Primary: classic Common API image insertion — broadly supported on
    // PowerPoint Windows/Mac/web, including older builds. Positions in points.
    const inserted = await new Promise((resolve) => {
      try {
        Office.context.document.setSelectedDataAsync(base64, {
          coercionType: Office.CoercionType.Image,
          imageLeft: left, imageTop: top, imageWidth: width, imageHeight: height,
        }, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? true : res));
      } catch (e) { resolve(e); }
    });
    if (inserted === true) return { mode: 'inserted' };

    // Fallback: application-specific API (newer builds only).
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getActiveSlide();
        const shape = slide.shapes.addImage(base64);
        try { shape.left = left; shape.top = top; shape.width = width; shape.height = height; } catch (_) {}
        await context.sync();
      });
      return { mode: 'inserted' };
    } catch (e2) {
      const msg = (inserted && inserted.error && inserted.error.message)
        || (e2 && e2.message) || 'Falha ao inserir a imagem.';
      throw new Error(msg);
    }
  }

  /* Insert a full image (e.g. a generated 1920x1080) onto the active slide,
     sized to fill the slide (16:9 image on a 16:9 slide). Accepts a data URL
     or raw base64. Falls back to a download outside PowerPoint. */
  async function insertImage(dataUrl) {
    if (!inPowerPoint()) {
      const a = document.createElement('a');
      a.href = dataUrl.indexOf(',') >= 0 ? dataUrl : 'data:image/png;base64,' + dataUrl;
      a.download = 'imagem.png'; a.click();
      return { mode: 'download' };
    }
    const base64 = dataUrl.indexOf(',') >= 0 ? dataUrl.split(',')[1] : dataUrl;
    const slide = getSlideSizeSync();
    if (!slide.exact) detectSlideSize();
    const o = { coercionType: Office.CoercionType.Image, imageLeft: 0, imageTop: 0, imageWidth: slide.w, imageHeight: slide.h };
    const inserted = await new Promise((resolve) => {
      try {
        Office.context.document.setSelectedDataAsync(base64, o,
          (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? true : res));
      } catch (e) { resolve(e); }
    });
    if (inserted === true) return { mode: 'inserted' };
    try {
      await PowerPoint.run(async (context) => {
        const sl = context.presentation.getActiveSlide();
        const shape = sl.shapes.addImage(base64);
        try { shape.left = 0; shape.top = 0; shape.width = slide.w; shape.height = slide.h; } catch (_) {}
        await context.sync();
      });
      return { mode: 'inserted' };
    } catch (e2) {
      const msg = (inserted && inserted.error && inserted.error.message) || (e2 && e2.message) || 'Falha ao inserir a imagem.';
      throw new Error(msg);
    }
  }

  /* Open the big-canvas drawing dialog.
     The slide backdrop and the returned drawing are passed through
     localStorage (same origin) to avoid Office message size limits.
     onResult(payload) — payload { strokes, config }, ou null se o usuário
     cancelou/fechou sem inserir (o painel usa isso pra resetar o status).
     hooks (opcional): { onOpen(), onError(msg) } — feedback visível de
     abertura/falha (nunca só console.warn). */
  function openDrawDialog(slideImageUrl, onResult, hooks) {
    const onOpen = (hooks && hooks.onOpen) || function () {};
    const onError = (hooks && hooks.onError) || function (m) { console.warn('[OfficeBridge]', m); };
    try {
      localStorage.setItem('doodle.slideImage', slideImageUrl || '');
      localStorage.removeItem('doodle.result');
    } catch (_) {}
    const url = new URL('../dialog/dialog.html', location.href).href;

    const canOfficeDialog = inPowerPoint()
      && Office.context.ui && typeof Office.context.ui.displayDialogAsync === 'function';

    if (!canOfficeDialog) {
      // Standalone (browser): open a normal window, poll localStorage.
      const w = window.open(url, 'doodleDialog', 'width=1280,height=820');
      if (!w) { onError('O navegador bloqueou a janela — permita pop-ups para este site.'); return; }
      onOpen();
      const timer = setInterval(() => {
        let res = null;
        try { res = localStorage.getItem('doodle.result'); } catch (_) {}
        if (res) {
          clearInterval(timer);
          try { localStorage.removeItem('doodle.result'); } catch (_) {}
          try { w.close(); } catch (_) {}
          let parsed = null;
          try { parsed = JSON.parse(res); } catch (_) {}
          onResult(parsed);
        } else if (w.closed) {
          clearInterval(timer);
          onResult(null);          // fechou sem inserir → painel reseta o status
        }
      }, 400);
      return;
    }

    Office.context.ui.displayDialogAsync(url, { height: 82, width: 82, displayInIframe: false }, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        const err = asyncResult.error || {};
        console.warn('[OfficeBridge] dialog failed:', err);
        onError(err.code === 12007
          ? 'Já existe uma janela do CBA Studio aberta — feche-a e tente de novo.'
          : 'Não consegui abrir a tela grande (' + (err.message || err.code || 'erro') + ').');
        return;
      }
      onOpen();
      const dialog = asyncResult.value;
      let done = false;                          // só o 1º resultado conta (dedup)
      const settle = (payload) => {
        if (done) return;
        done = true;
        try { dialog.close(); } catch (_) {}
        onResult(payload);
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (done) return;
        // Novo formato: a mensagem É o resultado { kind, payload } — sem localStorage.
        let parsed = null;
        try { parsed = JSON.parse(arg.message); } catch (_) {}
        if (parsed && parsed.kind) {
          settle(parsed.kind === 'inserted' ? (parsed.payload || null) : null);
          return;
        }
        // Fallback (compat / payload grande): string 'inserted'/'cancel' + localStorage.
        const msg = arg.message;
        if (msg === 'inserted') {
          let res = null;
          try { res = localStorage.getItem('doodle.result'); localStorage.removeItem('doodle.result'); } catch (_) {}
          let parsedRes = null;
          try { parsedRes = res ? JSON.parse(res) : null; } catch (_) {}
          settle(parsedRes);
        } else if (msg === 'cancel') {
          settle(null);
        }
      });
      // Usuário fechou no X (12006) ou o dialog caiu: avisa o painel.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        if (done) return;
        done = true;
        onResult(null);
      });
    });
  }

  /* Open the image-edit dialog: shows `imageDataUrl`, lets the user mark it up
     with a solid brush and type an edit prompt. onResult receives
     { prompt, markupDataUrl }. Image + result pass through localStorage.
     hooks (opcional): { onError(msg) }. */
  function openImageEditDialog(imageDataUrl, onResult, initialPrompt, hooks) {
    try {
      localStorage.setItem('doodle.editImage', imageDataUrl || '');
      localStorage.setItem('doodle.editPrompt', initialPrompt || '');
      localStorage.removeItem('doodle.editResult');
    } catch (_) {}
    const url = new URL('../dialog/imgedit.html', location.href).href;
    const canOfficeDialog = inPowerPoint()
      && Office.context.ui && typeof Office.context.ui.displayDialogAsync === 'function';

    const onError = (hooks && hooks.onError) || function (m) { console.warn('[OfficeBridge]', m); };

    if (!canOfficeDialog) {
      const w = window.open(url, 'imgEditDialog', 'width=1280,height=820');
      if (!w) { onError('O navegador bloqueou a janela — permita pop-ups para este site.'); return; }
      const timer = setInterval(() => {
        let res = null;
        try { res = localStorage.getItem('doodle.editResult'); } catch (_) {}
        if (res) {
          clearInterval(timer);
          try { localStorage.removeItem('doodle.editResult'); } catch (_) {}
          try { w.close(); } catch (_) {}
          let parsed = null;
          try { parsed = JSON.parse(res); } catch (_) {}
          if (parsed) onResult(parsed);
        } else if (w.closed) { clearInterval(timer); }
      }, 400);
      return;
    }

    Office.context.ui.displayDialogAsync(url, { height: 82, width: 82, displayInIframe: false }, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        const err = asyncResult.error || {};
        console.warn('[OfficeBridge] imgedit dialog failed:', err);
        onError(err.code === 12007
          ? 'Já existe uma janela do CBA Studio aberta — feche-a e tente de novo.'
          : 'Não consegui abrir o editor (' + (err.message || err.code || 'erro') + ').');
        return;
      }
      const dialog = asyncResult.value;
      let done = false;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (done) return;
        if (arg.message === 'edit') {
          done = true;
          let res = null;
          try { res = localStorage.getItem('doodle.editResult'); localStorage.removeItem('doodle.editResult'); } catch (_) {}
          try { dialog.close(); } catch (_) {}
          let parsed = null;
          try { parsed = res ? JSON.parse(res) : null; } catch (_) {}
          if (parsed) onResult(parsed);
        } else if (arg.message === 'cancel') {
          done = true;
          try { dialog.close(); } catch (_) {}
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { done = true; });
    });
  }

  /* ============================================================
     TYPOGRAPHY module — apply text styles + insert the style anchor.
     ============================================================ */

  function apiSupported(version) {
    try {
      return !!(Office.context && Office.context.requirements
        && Office.context.requirements.isSetSupported('PowerPointApi', version));
    } catch (_) { return false; }
  }

  function _abToBase64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(bin);
  }

  /* Best-effort: is the (single) selected shape's fill the brand blue?
     Fully isolated — never throws, returns false on any uncertainty so it
     can't break the apply that follows. */
  async function _selectionBgIsBlue() {
    try {
      return await PowerPoint.run(async (context) => {
        const sel = context.presentation.getSelectedTextRangeOrNullObject();
        const shapes = context.presentation.getSelectedShapes();
        sel.load('isNullObject');
        shapes.load('items');
        await context.sync();
        let shape = null;
        if (!sel.isNullObject) { try { shape = sel.getParentTextFrame().getParentShape(); } catch (_) {} }
        else if (shapes.items && shapes.items.length === 1) shape = shapes.items[0];
        if (!shape) return false;
        shape.fill.load('type');
        await context.sync();
        if (String(shape.fill.type).toLowerCase().indexOf('solid') < 0) return false;
        shape.fill.load('foregroundColor');
        await context.sync();
        const fg = (shape.fill.foregroundColor || '').toString().toUpperCase().replace('#', '');
        return fg.indexOf('436AE1') >= 0;
      });
    } catch (_) { return false; }
  }

  /* Apply a typographic style to the current selection.
     style = { font, size, color, periodColor, bgAware, sample }.
     Targets, in order: a selected text sub-range -> selected shape(s) ->
     a brand-new text box on the active slide. Sets font/bold/size/color and
     colors the final period; line spacing is NOT settable here (use the anchor). */
  async function applyTextStyle(style) {
    if (!inPowerPoint()) throw new Error('Disponível apenas dentro do PowerPoint.');
    if (!apiSupported('1.6')) {
      throw new Error('Seu PowerPoint não suporta a API 1.6. Atualize, ou use a referência (Pincel de Formatação).');
    }

    // background-aware override (Hero on a blue object -> white text)
    let textColor = style.color, periodColor = style.periodColor;
    if (style.bgAware && await _selectionBgIsBlue()) { textColor = '#FFFFFF'; periodColor = '#FC5E6D'; }

    return await PowerPoint.run(async (context) => {
      const sel = context.presentation.getSelectedTextRangeOrNullObject();
      sel.load('text');
      const shapes = context.presentation.getSelectedShapes();
      shapes.load('items');
      await context.sync();

      let ranges = [], mode = '';
      if (!sel.isNullObject && sel.text && sel.text.trim().length) {
        ranges = [sel]; mode = 'text';
      } else if (shapes.items && shapes.items.length) {
        ranges = shapes.items.map((sh) => sh.textFrame.textRange);
        mode = 'shape';
      } else {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const box = slide.shapes.addTextBox((style.sample || 'Texto') + '.',
          { left: 48, top: 72, width: 864, height: 220 });
        ranges = [box.textFrame.textRange]; mode = 'new';
      }

      ranges.forEach((r) => r.load('text'));
      await context.sync();

      ranges.forEach((r) => {
        r.font.name = style.font;
        r.font.bold = !!style.bold;
        r.font.size = style.size;
        r.font.color = textColor;
        if (periodColor) {
          const t = r.text || '';
          if (t.length && t.charAt(t.length - 1) === '.') {
            try { r.getSubstring(t.length - 1, 1).font.color = periodColor; } catch (_) {}
          }
        }
      });
      await context.sync();
      return { applied: ranges.length, mode };
    });
  }

  /* Insert the "style anchor" reference slide (exact styles incl. line
     spacing, baked in OOXML) so the user can copy a box or Format-Painter it. */
  async function insertStylesReference() {
    if (!inPowerPoint()) throw new Error('Disponível apenas dentro do PowerPoint.');
    const url = new URL('../../assets/estilos-bg.pptx', location.href).href;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error('Não foi possível carregar a referência (' + resp.status + ').');
    const b64 = _abToBase64(await resp.arrayBuffer());
    await PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(b64, { formatting: 'KeepSourceFormatting' });
      await context.sync();
    });
  }

  global.OfficeBridge = {
    inPowerPoint, getActiveSlideImage, insertDoodle, insertImage, openDrawDialog,
    openImageEditDialog, getSlideSizeSync, detectSlideSize, SLIDE_W_PT, SLIDE_H_PT,
    apiSupported, applyTextStyle, insertStylesReference,
    getImageAudit, goToSlide, optimizeImage,
  };
})(window);
