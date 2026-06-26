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

  function getCompressedBytes() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
        const file = res.value, n = file.sliceCount, parts = new Array(n);
        let got = 0, failed = false;
        for (let i = 0; i < n; i++) {
          file.getSliceAsync(i, (s) => {
            if (failed) return;
            if (s.status !== Office.AsyncResultStatus.Succeeded) { failed = true; file.closeAsync(() => {}); return reject(s.error); }
            parts[s.value.index] = s.value.data;
            if (++got === n) {
              file.closeAsync(() => {});
              let len = 0; for (const p of parts) len += p.length;
              const out = new Uint8Array(len); let o = 0;
              for (const p of parts) { out.set(p, o); o += p.length; }
              resolve(out);
            }
          });
        }
      });
    });
  }

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Response(new Blob([bytes]).stream().pipeThrough(ds));
    return new Uint8Array(await stream.arrayBuffer());
  }

  async function unzipEntry(b, name) {
    const u32 = (o) => b[o] | (b[o+1]<<8) | (b[o+2]<<16) | (b[o+3]*0x1000000);
    const u16 = (o) => b[o] | (b[o+1]<<8);
    let eocd = -1;
    for (let i = b.length - 22; i >= 0; i--) { if (u32(i) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error('zip: no EOCD');
    let off = u32(eocd + 16); const count = u16(eocd + 10);
    for (let k = 0; k < count; k++) {
      if (u32(off) !== 0x02014b50) throw new Error('zip: bad CDH');
      const method = u16(off + 10), csize = u32(off + 20);
      const nlen = u16(off + 28), elen = u16(off + 30), clen = u16(off + 32), lho = u32(off + 42);
      const fname = new TextDecoder().decode(b.subarray(off + 46, off + 46 + nlen));
      if (fname === name) {
        const lnlen = u16(lho + 26), lelen = u16(lho + 28);
        const ds = lho + 30 + lnlen + lelen;
        const data = b.subarray(ds, ds + csize);
        const out = method === 0 ? data : await inflateRaw(data);
        return new TextDecoder().decode(out);
      }
      off += 46 + nlen + elen + clen;
    }
    throw new Error('zip: entry not found');
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
      try {
        const bytes = await getCompressedBytes();
        const xml = await unzipEntry(bytes, 'ppt/presentation.xml');
        const m = xml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        if (m) {
          _slideSize = { w: +m[1] / 12700, h: +m[2] / 12700 };
          try { localStorage.setItem(_sizeKey(), JSON.stringify(_slideSize)); } catch (_) {}
        }
      } catch (e) {
        console.warn('[OfficeBridge] slide size detect failed:', e && e.message);
      } finally {
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
