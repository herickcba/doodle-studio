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
     that overlap it — not the whole file. */
  async function _readRange(file, start, len) {
    const end = start + len, first = Math.floor(start / SLICE), last = Math.floor((end - 1) / SLICE);
    const out = new Uint8Array(len); let w = 0;
    for (let i = first; i <= last; i++) {
      const bytes = _toBytes(await _getSlice(file, i));
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

  /* Open the big-canvas drawing dialog.
     The slide backdrop and the returned drawing are passed through
     localStorage (same origin) to avoid Office message size limits.
     onResult receives the parsed payload { strokes, config }. */
  function openDrawDialog(slideImageUrl, onResult) {
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
      const timer = setInterval(() => {
        let res = null;
        try { res = localStorage.getItem('doodle.result'); } catch (_) {}
        if (res) {
          clearInterval(timer);
          try { localStorage.removeItem('doodle.result'); } catch (_) {}
          try { w && w.close(); } catch (_) {}
          onResult(JSON.parse(res));
        } else if (w && w.closed) {
          clearInterval(timer);
        }
      }, 400);
      return;
    }

    Office.context.ui.displayDialogAsync(url, { height: 82, width: 82, displayInIframe: false }, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('[OfficeBridge] dialog failed:', asyncResult.error);
        return;
      }
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const msg = arg.message;
        if (msg === 'inserted') {
          let res = null;
          try { res = localStorage.getItem('doodle.result'); localStorage.removeItem('doodle.result'); } catch (_) {}
          dialog.close();
          if (res) onResult(JSON.parse(res));
        } else if (msg === 'cancel') {
          dialog.close();
        }
      });
    });
  }

  global.OfficeBridge = {
    inPowerPoint, getActiveSlideImage, insertDoodle, openDrawDialog,
    getSlideSizeSync, detectSlideSize, SLIDE_W_PT, SLIDE_H_PT,
  };
})(window);
