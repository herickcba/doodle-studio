/* ============================================================
   Doodle Engine — extracted from CBA Studio v2.html
   Pure vector drawing core (no animation, no timeline).
   Frame coordinate space is fixed at 1920x1080 (16:9).
   Exposes window.Doodle.
   ============================================================ */
(function (global) {
  'use strict';

  const FRAME_W = 1920, FRAME_H = 1080;

  const state = {
    color: '#FD5E6D',
    opacity: 1,            // 0..1
    size: 8,               // 1..60
    textures: ['giz'],     // MVP: giz only (solido kept as fallback)
    mouseSmoothing: 50,    // 0..100 → Gaussian window
    vectorSmoothing: 0,    // 0..100 → RDP + Chaikin
    faultTypes: [],        // subset of ['quebra','falha','respingo']
    faultLevel: 16,        // 0..100
    gapSize: 54,           // 0..100
    strokes: [],
    redo: [],
    current: null,
    isDrawing: false,
    pointerStart: null,
    insertSeparate: false,   // insert each stroke as its own image vs one merged image
    selectedStrokeIdx: -1,   // index of the click-selected stroke (-1 = none)
  };

  let canvas = null, ctx = null, onChange = null;

  /* ---------------- Vectorization (pure) ---------------- */
  function bboxDiag(pts) {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const p of pts) {
      if (p.x < minX) minX = p.x; if (p.x > maxX) maxX = p.x;
      if (p.y < minY) minY = p.y; if (p.y > maxY) maxY = p.y;
    }
    return Math.hypot(maxX - minX, maxY - minY);
  }

  function smoothingParams(mouse, vector, raw) {
    const diag = bboxDiag(raw);
    const n = raw.length;
    const gwinTarget = Math.round((mouse / 100) * 34);
    const gwinCap = Math.max(0, Math.floor(n * 0.12));
    const gwin = Math.min(gwinTarget, gwinCap);
    const epsTarget = Math.pow(vector / 100, 1.35) * 70;
    const epsCap = Math.max(0, diag * 0.10);
    const eps = Math.min(epsTarget, epsCap);
    let iters;
    if (vector <= 5) iters = 0;
    else if (vector <= 30) iters = 1;
    else if (vector <= 60) iters = 2;
    else if (vector <= 85) iters = 3;
    else iters = 4;
    return { gwin, eps, iters };
  }

  function perpDist(p, a, b) {
    const dx = b.x - a.x, dy = b.y - a.y;
    if (dx === 0 && dy === 0) return Math.hypot(p.x - a.x, p.y - a.y);
    const t = ((p.x - a.x) * dx + (p.y - a.y) * dy) / (dx * dx + dy * dy);
    const cx = a.x + t * dx, cy = a.y + t * dy;
    return Math.hypot(p.x - cx, p.y - cy);
  }

  function rdp(points, epsilon) {
    if (points.length < 3 || epsilon <= 0) return points.slice();
    const end = points.length - 1;
    let idx = 0, dmax = 0;
    for (let i = 1; i < end; i++) {
      const d = perpDist(points[i], points[0], points[end]);
      if (d > dmax) { dmax = d; idx = i; }
    }
    if (dmax > epsilon) {
      const left = rdp(points.slice(0, idx + 1), epsilon);
      const right = rdp(points.slice(idx), epsilon);
      return left.slice(0, -1).concat(right);
    }
    return [points[0], points[end]];
  }

  function gaussianSmooth(pts, windowSize) {
    if (pts.length < 3 || windowSize < 1) return pts.slice();
    const half = Math.floor(windowSize / 2);
    const sigma = Math.max(1, windowSize / 3);
    const w = new Array(half * 2 + 1);
    for (let i = -half; i <= half; i++) w[i + half] = Math.exp(-(i * i) / (2 * sigma * sigma));
    const out = new Array(pts.length);
    for (let i = 0; i < pts.length; i++) {
      let sx = 0, sy = 0, sw = 0;
      for (let k = -half; k <= half; k++) {
        const j = i + k;
        if (j < 0 || j >= pts.length) continue;
        const ww = w[k + half];
        sx += pts[j].x * ww; sy += pts[j].y * ww; sw += ww;
      }
      out[i] = { x: sx / sw, y: sy / sw };
    }
    out[0] = pts[0];
    out[out.length - 1] = pts[pts.length - 1];
    return out;
  }

  function chaikin(pts, iters) {
    if (iters <= 0 || pts.length < 3) return pts;
    let out = pts;
    for (let it = 0; it < iters; it++) {
      const next = [out[0]];
      for (let i = 0; i < out.length - 1; i++) {
        const p = out[i], q = out[i + 1];
        next.push({ x: 0.75 * p.x + 0.25 * q.x, y: 0.75 * p.y + 0.25 * q.y });
        next.push({ x: 0.25 * p.x + 0.75 * q.x, y: 0.25 * p.y + 0.75 * q.y });
      }
      next.push(out[out.length - 1]);
      out = next;
    }
    return out;
  }

  function vectorize(stroke) {
    const key = state.mouseSmoothing * 1000 + state.vectorSmoothing;
    if (stroke._cache && stroke._cache.k === key) return stroke._cache.pts;
    const raw = stroke.raw;
    if (raw.length < 3) {
      const pts = raw.slice();
      stroke._cache = { k: key, pts };
      return pts;
    }
    const { gwin, eps, iters } = smoothingParams(state.mouseSmoothing, state.vectorSmoothing, raw);
    const filtered = gwin > 0 ? gaussianSmooth(raw, gwin) : raw;
    const simplified = eps > 0 ? rdp(filtered, eps) : filtered;
    const pts = chaikin(simplified, iters);
    stroke._cache = { k: key, pts };
    return pts;
  }

  function anchorCount(stroke) {
    const raw = stroke.raw;
    if (raw.length < 3) return raw.length;
    const { gwin, eps } = smoothingParams(state.mouseSmoothing, state.vectorSmoothing, raw);
    const filtered = gwin > 0 ? gaussianSmooth(raw, gwin) : raw;
    const simplified = eps > 0 ? rdp(filtered, eps) : filtered;
    return simplified.length;
  }

  function densify(pts, stepsPerSeg = 20) {
    if (pts.length < 2) return pts.slice();
    const out = [pts[0]];
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[Math.max(0, i - 1)];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[Math.min(pts.length - 1, i + 2)];
      const c1x = p1.x + (p2.x - p0.x) / 6;
      const c1y = p1.y + (p2.y - p0.y) / 6;
      const c2x = p2.x - (p3.x - p1.x) / 6;
      const c2y = p2.y - (p3.y - p1.y) / 6;
      for (let s = 1; s <= stepsPerSeg; s++) {
        const t = s / stepsPerSeg, u = 1 - t;
        out.push({
          x: u * u * u * p1.x + 3 * u * u * t * c1x + 3 * u * t * t * c2x + t * t * t * p2.x,
          y: u * u * u * p1.y + 3 * u * u * t * c1y + 3 * u * t * t * c2y + t * t * t * p2.y,
        });
      }
    }
    return out;
  }

  function walkPath(dense, step, fn) {
    let carry = 0;
    for (let i = 0; i < dense.length - 1; i++) {
      const a = dense[i], b = dense[i + 1];
      const dx = b.x - a.x, dy = b.y - a.y;
      const d = Math.hypot(dx, dy);
      if (d < 1e-6) continue;
      const ux = dx / d, uy = dy / d;
      let t = step - carry;
      while (t <= d) { fn(a.x + ux * t, a.y + uy * t, ux, uy); t += step; }
      carry = d - (t - step);
    }
  }

  function rand(x, y) {
    const v = Math.sin(x * 127.1 + y * 311.7) * 43758.5453;
    return v - Math.floor(v);
  }

  /* ---------------- Stroke rendering ---------------- */
  function strokeBezier(c, pts) {
    if (pts.length === 0) return;
    if (pts.length === 1) {
      c.beginPath();
      c.arc(pts[0].x, pts[0].y, Math.max(0.5, c.lineWidth / 2), 0, Math.PI * 2);
      c.fill();
      return;
    }
    c.beginPath();
    c.moveTo(pts[0].x, pts[0].y);
    if (pts.length === 2) { c.lineTo(pts[1].x, pts[1].y); c.stroke(); return; }
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[Math.max(0, i - 1)];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[Math.min(pts.length - 1, i + 2)];
      const c1x = p1.x + (p2.x - p0.x) / 6;
      const c1y = p1.y + (p2.y - p0.y) / 6;
      const c2x = p2.x - (p3.x - p1.x) / 6;
      const c2y = p2.y - (p3.y - p1.y) / 6;
      c.bezierCurveTo(c1x, c1y, c2x, c2y, p2.x, p2.y);
    }
    c.stroke();
  }

  const TEXTURES = {
    solido(c, pts, s) {
      c.lineCap = 'round'; c.lineJoin = 'round';
      c.strokeStyle = s.color; c.globalAlpha = s.opacity; c.lineWidth = s.size;
      strokeBezier(c, pts);
    },
    giz(c, pts, s) {
      c.lineCap = 'round'; c.lineJoin = 'round';
      c.strokeStyle = s.color;
      c.globalAlpha = s.opacity * 0.35;
      c.lineWidth = s.size * 0.85;
      strokeBezier(c, pts);

      const dense = densify(pts, 16);
      c.fillStyle = s.color;
      c.globalAlpha = s.opacity * 0.7;
      walkPath(dense, Math.max(1, s.size * 0.25), (x, y, ux, uy) => {
        const nx = -uy, ny = ux;
        for (let k = 0; k < 4; k++) {
          const off = (rand(x + k * 3.1, y + k * 7.7) - 0.5) * s.size * 1.4;
          const jx = (rand(x * 2 + k, y) - 0.5) * 0.6;
          const jy = (rand(x, y * 2 + k) - 0.5) * 0.6;
          const px = x + nx * off + jx;
          const py = y + ny * off + jy;
          const r = 0.6 + rand(x + k, y - k) * Math.max(0.8, s.size * 0.12);
          c.beginPath();
          c.arc(px, py, r, 0, Math.PI * 2);
          c.fill();
        }
      });
    },
  };

  /* ---------------- Ink faults ---------------- */
  function buildFaultPlan(pts, types, level, gapSize, seed) {
    if (!types || !types.length || level <= 0 || pts.length < 3) {
      return { segments: [{ pts, alpha: 1 }], splatter: 0 };
    }
    const segments = [];
    const N = pts.length;
    const rate = Math.pow(level / 100, 1.25) * 0.55;
    const g = Math.max(0, Math.min(100, gapSize || 0)) / 100;
    const chunkMax = 2 + Math.round(Math.pow(g, 1.2) * 16);
    let i = 0;
    while (i < N - 1) {
      const sx = pts[i].x + seed * 0.13;
      const sy = pts[i].y + seed * 0.19;
      const chunk = 2 + Math.floor(rand(sx * 0.37, sy * 0.71 + i) * (chunkMax - 1));
      const end = Math.min(i + chunk, N - 1);
      const seg = pts.slice(i, end + 1);
      const r = rand(sx + 11, sy * 2.3 + i * 3.1);
      if (r < rate) {
        const pickIdx = Math.floor(rand(sx + 17 + seed, sy + 29) * types.length);
        const picked = types[pickIdx];
        if (picked === 'quebra') {
          // drop — clean gap
        } else if (picked === 'falha') {
          const a = 0.08 + rand(sx, sy + i) * 0.28;
          segments.push({ pts: seg, alpha: a });
        } else if (picked === 'respingo') {
          segments.push({ pts: seg, alpha: 1 });
        }
      } else {
        segments.push({ pts: seg, alpha: 1 });
      }
      i = end;
    }
    return { segments, splatter: types.includes('respingo') ? level : 0 };
  }

  function drawSplatter(c, pts, stroke, level) {
    const dense = densify(pts, 12);
    const density = Math.pow(level / 100, 1.1) * 0.35;
    const reach = stroke.size * 3.5;
    c.fillStyle = stroke.color;
    walkPath(dense, 3, (x, y, ux, uy) => {
      for (let k = 0; k < 3; k++) {
        const rr = rand(x + k * 1.9 + stroke.seed, y + k * 2.7);
        if (rr >= density) continue;
        const ang = rand(x + k * 3.1, y - k * 1.7 + stroke.seed) * Math.PI * 2;
        const dist = 0.2 + rand(x * 2, y * 2 + k) * reach;
        const sz = 0.5 + rand(x - k, y + k + stroke.seed) * Math.max(1, stroke.size * 0.35);
        c.globalAlpha = stroke.opacity * (0.25 + rand(x * 3 + stroke.seed, y) * 0.55);
        c.beginPath();
        c.arc(x + Math.cos(ang) * dist, y + Math.sin(ang) * dist, sz, 0, Math.PI * 2);
        c.fill();
      }
    });
  }

  function drawStroke(c, stroke, ptsOverride) {
    const pts = ptsOverride || vectorize(stroke);
    const textures = (stroke.textures && stroke.textures.length)
      ? stroke.textures : [stroke.texture || 'giz'];
    const seed = stroke.seed || 0;
    let faultTypes = state.faultTypes;
    let faultLevel = state.faultLevel;
    if ((!faultTypes || !faultTypes.length) && stroke.faultTypes && stroke.faultTypes.length) {
      faultTypes = stroke.faultTypes;
      faultLevel = stroke.faultLevel != null ? stroke.faultLevel : faultLevel;
    }
    const plan = buildFaultPlan(pts, faultTypes, faultLevel, state.gapSize, seed);
    c.save();
    for (const seg of plan.segments) {
      if (seg.pts.length < 2) continue;
      const sub = Object.assign({}, stroke, { opacity: stroke.opacity * seg.alpha });
      for (const tex of textures) {
        c.save();
        (TEXTURES[tex] || TEXTURES.solido)(c, seg.pts, sub);
        c.restore();
      }
    }
    if (plan.splatter > 0) drawSplatter(c, pts, stroke, plan.splatter);
    c.restore();
  }

  /* ---------------- Canvas render + input ---------------- */
  function render() {
    if (!ctx) return;
    ctx.clearRect(0, 0, FRAME_W, FRAME_H);
    for (const s of state.strokes) drawStroke(ctx, s);
    if (state.current) drawStroke(ctx, state.current);
    highlightSelected();
    if (onChange) onChange();
  }

  // Soft blue halo under/over the click-selected stroke.
  function highlightSelected() {
    const i = state.selectedStrokeIdx;
    if (i < 0 || i >= state.strokes.length) return;
    const s = state.strokes[i];
    const pts = vectorize(s);
    if (!pts || pts.length < 2) return;
    ctx.save();
    ctx.globalAlpha = 0.28;
    ctx.strokeStyle = '#436AE1';
    ctx.lineWidth = Math.max(6, s.size + 10);
    ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (let k = 1; k < pts.length; k++) ctx.lineTo(pts[k].x, pts[k].y);
    ctx.stroke();
    ctx.restore();
  }

  /* ---------------- Click-to-select a stroke ---------------- */
  function distPointToSeg(p, a, b) {
    const dx = b.x - a.x, dy = b.y - a.y;
    const len2 = dx * dx + dy * dy;
    let t = len2 ? ((p.x - a.x) * dx + (p.y - a.y) * dy) / len2 : 0;
    t = t < 0 ? 0 : t > 1 ? 1 : t;
    return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy));
  }
  function distToStroke(p, stroke) {
    const pts = vectorize(stroke);
    if (!pts || !pts.length) return Infinity;
    if (pts.length === 1) return Math.hypot(p.x - pts[0].x, p.y - pts[0].y);
    let min = Infinity;
    for (let i = 1; i < pts.length; i++) {
      const d = distPointToSeg(p, pts[i - 1], pts[i]);
      if (d < min) min = d;
    }
    return min;
  }
  // Pick the nearest stroke under p (frame coords); -1 if none within reach.
  function handleStrokeClick(p) {
    let best = -1, bestD = Infinity;
    for (let i = 0; i < state.strokes.length; i++) {
      const d = distToStroke(p, state.strokes[i]);
      const reach = Math.max(22, state.strokes[i].size * 2.5);
      if (d <= reach && d < bestD) { bestD = d; best = i; }
    }
    state.selectedStrokeIdx = best;
    return best;
  }
  function getSelectedStroke() {
    const i = state.selectedStrokeIdx;
    return (i >= 0 && i < state.strokes.length) ? state.strokes[i] : null;
  }
  function clearSelection() { state.selectedStrokeIdx = -1; render(); }

  function getPos(e) {
    const rect = canvas.getBoundingClientRect();
    const sx = FRAME_W / rect.width;
    const sy = FRAME_H / rect.height;
    return { x: (e.clientX - rect.left) * sx, y: (e.clientY - rect.top) * sy };
  }

  function onDown(e) {
    e.preventDefault();
    try { canvas.setPointerCapture(e.pointerId); } catch (_) {}
    state.isDrawing = true;
    const p = getPos(e);
    state.pointerStart = { x: e.clientX, y: e.clientY };
    state.current = {
      raw: [p], color: state.color, opacity: state.opacity,
      size: state.size, textures: state.textures.slice(),
      seed: Math.floor(Math.random() * 1e6), _cache: null,
    };
  }

  function onMove(e) {
    if (!state.isDrawing) return;
    e.preventDefault();
    const p = getPos(e);
    if (e.shiftKey) {                           // straight line
      state.current.raw = [state.current.raw[0], p];
      state.current._cache = null;
      render();
      return;
    }
    const last = state.current.raw[state.current.raw.length - 1];
    const dx = p.x - last.x, dy = p.y - last.y;
    if (dx * dx + dy * dy < 1.5) return;
    state.current.raw.push(p);
    state.current._cache = null;
    render();
  }

  function onUp(e) {
    if (!state.isDrawing) return;
    state.isDrawing = false;
    const cur = state.current;
    // A tap that barely moved is a selection click, not a stroke.
    let moved = 999;
    if (cur && state.pointerStart && e && e.clientX != null) {
      moved = Math.hypot(e.clientX - state.pointerStart.x, e.clientY - state.pointerStart.y);
    }
    const wasClick = cur && cur.raw.length <= 2 && moved < 6;
    if (wasClick) {
      state.current = null;
      handleStrokeClick(cur.raw[0]);
    } else if (cur && cur.raw.length > 0) {
      state.strokes.push(cur);
      state.current = null;
      state.redo = [];
      state.selectedStrokeIdx = -1;   // a fresh stroke clears any selection
    }
    state.pointerStart = null;
    render();
  }

  function attach(canvasEl, opts) {
    canvas = canvasEl;
    canvas.width = FRAME_W; canvas.height = FRAME_H;
    ctx = canvas.getContext('2d');
    onChange = (opts && opts.onChange) || null;
    canvas.addEventListener('pointerdown', onDown);
    canvas.addEventListener('pointermove', onMove);
    canvas.addEventListener('pointerup', onUp);
    canvas.addEventListener('pointercancel', onUp);
    window.addEventListener('pointerup', onUp);
    render();
  }

  /* ---------------- History ---------------- */
  function undo() {
    if (!state.strokes.length) return;
    state.redo.push(state.strokes.pop());
    state.selectedStrokeIdx = -1;
    render();
  }
  function redo() {
    if (!state.redo.length) return;
    state.strokes.push(state.redo.pop());
    render();
  }
  function clear() {
    if (!state.strokes.length) return;
    state.redo = state.strokes.splice(0).concat(state.redo).slice(0, 200);
    state.selectedStrokeIdx = -1;
    render();
  }

  /* ---------------- Export: transparent PNG cropped to content ----------------
     Returns { dataUrl, bbox } where bbox = { x, y, w, h } in FRAME (1920x1080)
     coordinates, or null if there is nothing drawn. */
  function bboxOfStrokes(strokes) {
    if (!strokes || !strokes.length) return null;
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    let maxSize = 1;
    for (const s of strokes) {
      if (s.size > maxSize) maxSize = s.size;
      const pts = vectorize(s);
      for (const p of pts) {
        if (p.x < minX) minX = p.x; if (p.x > maxX) maxX = p.x;
        if (p.y < minY) minY = p.y; if (p.y > maxY) maxY = p.y;
      }
    }
    // pad for stroke width + texture grain / splatter reach
    const pad = maxSize * 4 + 12;
    minX = Math.max(0, minX - pad);
    minY = Math.max(0, minY - pad);
    maxX = Math.min(FRAME_W, maxX + pad);
    maxY = Math.min(FRAME_H, maxY + pad);
    return { x: minX, y: minY, w: Math.max(1, maxX - minX), h: Math.max(1, maxY - minY) };
  }

  function pngOfStrokes(strokes) {
    const bbox = bboxOfStrokes(strokes);
    if (!bbox) return null;
    const off = document.createElement('canvas');
    off.width = Math.round(bbox.w);
    off.height = Math.round(bbox.h);
    const oc = off.getContext('2d');
    oc.translate(-bbox.x, -bbox.y);
    for (const s of strokes) drawStroke(oc, s);
    return { dataUrl: off.toDataURL('image/png'), bbox: bbox, frame: { w: FRAME_W, h: FRAME_H } };
  }

  function contentBBox() { return bboxOfStrokes(state.strokes); }
  function exportTransparentPNG() { return pngOfStrokes(state.strokes); }

  /* Render an arbitrary set of strokes (e.g. coming from the big-canvas
     dialog) using the given config, WITHOUT disturbing live state. */
  function renderExternalPNG(strokes, config) {
    const saved = {
      ms: state.mouseSmoothing, vs: state.vectorSmoothing,
      ft: state.faultTypes, fl: state.faultLevel, gs: state.gapSize,
    };
    if (config) {
      state.mouseSmoothing = config.mouseSmoothing;
      state.vectorSmoothing = config.vectorSmoothing;
      state.faultTypes = config.faultTypes || [];
      state.faultLevel = config.faultLevel;
      state.gapSize = config.gapSize;
    }
    const png = pngOfStrokes(strokes);
    state.mouseSmoothing = saved.ms; state.vectorSmoothing = saved.vs;
    state.faultTypes = saved.ft; state.faultLevel = saved.fl; state.gapSize = saved.gs;
    return png;
  }

  /* Export one PNG per stroke (separate) or one merged PNG (together).
     Returns an array of { dataUrl, bbox, frame }. */
  function exportPNGs(separate) {
    if (separate) return state.strokes.map((s) => pngOfStrokes([s])).filter(Boolean);
    const p = pngOfStrokes(state.strokes);
    return p ? [p] : [];
  }

  /* Same, for an arbitrary stroke set (e.g. from the big-canvas dialog),
     using the given config, without disturbing live state. */
  function renderExternalPNGs(strokes, config, separate) {
    const saved = {
      ms: state.mouseSmoothing, vs: state.vectorSmoothing,
      ft: state.faultTypes, fl: state.faultLevel, gs: state.gapSize,
    };
    if (config) {
      state.mouseSmoothing = config.mouseSmoothing;
      state.vectorSmoothing = config.vectorSmoothing;
      state.faultTypes = config.faultTypes || [];
      state.faultLevel = config.faultLevel;
      state.gapSize = config.gapSize;
    }
    let out;
    if (separate) out = strokes.map((s) => pngOfStrokes([s])).filter(Boolean);
    else { const p = pngOfStrokes(strokes); out = p ? [p] : []; }
    state.mouseSmoothing = saved.ms; state.vectorSmoothing = saved.vs;
    state.faultTypes = saved.ft; state.faultLevel = saved.fl; state.gapSize = saved.gs;
    return out;
  }

  /* Compact, serializable snapshot (strokes + config) for passing
     between windows. Drops the vectorization cache. */
  function payloadOf(strokes) {
    return {
      insertSeparate: state.insertSeparate,
      strokes: strokes.map((s) => ({
        raw: s.raw, color: s.color, opacity: s.opacity,
        size: s.size, textures: s.textures, seed: s.seed,
      })),
      config: {
        mouseSmoothing: state.mouseSmoothing, vectorSmoothing: state.vectorSmoothing,
        faultTypes: state.faultTypes.slice(), faultLevel: state.faultLevel, gapSize: state.gapSize,
      },
    };
  }
  function payload() { return payloadOf(state.strokes); }

  /* Append saved strokes (from the library) onto the live canvas. */
  function loadStrokes(arr) {
    if (!arr || !arr.length) return;
    for (const s of arr) {
      state.strokes.push({
        raw: s.raw, color: s.color, opacity: s.opacity,
        size: s.size, textures: (s.textures && s.textures.slice()) || ['giz'],
        seed: s.seed != null ? s.seed : Math.floor(Math.random() * 1e6),
        _cache: null,
      });
    }
    state.redo = [];
    state.selectedStrokeIdx = -1;
    render();
  }

  /* Small PNG preview of a stroke set, cropped to its bbox. */
  function thumbnailOf(strokes, maxPx) {
    maxPx = maxPx || 132;
    const bbox = bboxOfStrokes(strokes);
    if (!bbox) return '';
    const scale = Math.min(1, maxPx / Math.max(bbox.w, bbox.h));
    const W = Math.max(1, Math.round(bbox.w * scale));
    const H = Math.max(1, Math.round(bbox.h * scale));
    const off = document.createElement('canvas');
    off.width = W; off.height = H;
    const oc = off.getContext('2d');
    oc.setTransform(scale, 0, 0, scale, -bbox.x * scale, -bbox.y * scale);
    for (const s of strokes) drawStroke(oc, s);
    return off.toDataURL('image/png');
  }

  /* Thumbnail de strokes externos (presets/biblioteca) usando a config DELES,
     sem mexer no estado vivo — mesmo padrão do renderExternalPNG. */
  function thumbnailExternal(strokes, config, maxPx) {
    const saved = {
      ms: state.mouseSmoothing, vs: state.vectorSmoothing,
      ft: state.faultTypes, fl: state.faultLevel, gs: state.gapSize,
    };
    if (config) {
      state.mouseSmoothing = config.mouseSmoothing;
      state.vectorSmoothing = config.vectorSmoothing;
      state.faultTypes = config.faultTypes || [];
      state.faultLevel = config.faultLevel;
      state.gapSize = config.gapSize;
    }
    const t = thumbnailOf(strokes.map((s) => Object.assign({}, s, { _cache: null })), maxPx);
    state.mouseSmoothing = saved.ms; state.vectorSmoothing = saved.vs;
    state.faultTypes = saved.ft; state.faultLevel = saved.fl; state.gapSize = saved.gs;
    return t;
  }

  function isEmpty() { return state.strokes.length === 0; }

  /* ============================================================
     ANIMATION (progressive reveal by arc length) + GIF89a encoder.
     Ported from the web version (v2.html), no dependencies.
     ============================================================ */
  const EASING = {
    linear: (t) => t,
    easeIn: (t) => t * t,
    easeOut: (t) => 1 - (1 - t) * (1 - t),
    easeInOut: (t) => t * t * (3 - 2 * t),
  };
  function strokeArcLen(stroke) {
    const key = state.mouseSmoothing * 1000 + state.vectorSmoothing;
    if (stroke._arc && stroke._arc.k === key) return stroke._arc.len;
    const dense = densify(vectorize(stroke), 10);
    let len = 0;
    for (let i = 0; i < dense.length - 1; i++) len += Math.hypot(dense[i+1].x - dense[i].x, dense[i+1].y - dense[i].y);
    stroke._arc = { k: key, len };
    return len;
  }
  function strokeDur(s) { return Math.max(strokeArcLen(s), 4); }
  function truncatePts(pts, targetLen) {
    if (pts.length < 2 || targetLen <= 0) return pts.length ? [pts[0]] : [];
    const out = [pts[0]]; let acc = 0;
    for (let i = 0; i < pts.length - 1; i++) {
      const a = pts[i], b = pts[i + 1], seg = Math.hypot(b.x - a.x, b.y - a.y);
      if (acc + seg < targetLen) { out.push(b); acc += seg; }
      else { const t = seg < 1e-6 ? 0 : (targetLen - acc) / seg; out.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t }); return out; }
    }
    return out;
  }
  /* Reveal lengths for a group of strokes — all strokes animate TOGETHER on the
     same 0..1 timeline (each progresses independently from 0 to its full length).
     The big-canvas GIF has no fine per-stroke timing (that's the v2 web app), so
     drawing every stroke at once is the default. Pure — no mutation. */
  function strokeRevealsFor(strokes, progress, easing) {
    const ease = EASING[easing] || EASING.linear;
    const durs = strokes.map(strokeDur);
    const t = ease(Math.max(0, Math.min(1, progress)));
    return strokes.map((s, i) => t * durs[i]);
  }

  function gifBuildPalette(rgba, maxColors) {
    let samples = [];
    for (let i = 0; i < rgba.length; i += 4) if (rgba[i+3] >= 128) samples.push([rgba[i], rgba[i+1], rgba[i+2]]);
    if (samples.length === 0) return [[0, 0, 0]];
    if (samples.length > 20000) { const s = [], stp = samples.length / 20000; for (let k = 0; k < 20000; k++) s.push(samples[Math.floor(k * stp)]); samples = s; }
    let boxes = [samples];
    while (boxes.length < maxColors) {
      let bi = -1, brange = -1, bchan = 0;
      for (let i = 0; i < boxes.length; i++) {
        const b = boxes[i]; if (b.length < 2) continue;
        const mn = [255,255,255], mx = [0,0,0];
        for (const p of b) for (let c = 0; c < 3; c++) { if (p[c] < mn[c]) mn[c] = p[c]; if (p[c] > mx[c]) mx[c] = p[c]; }
        for (let c = 0; c < 3; c++) { const r = mx[c] - mn[c]; if (r > brange) { brange = r; bi = i; bchan = c; } }
      }
      if (bi < 0) break;
      const box = boxes[bi]; box.sort((a, b) => a[bchan] - b[bchan]); const mid = box.length >> 1;
      boxes.splice(bi, 1, box.slice(0, mid), box.slice(mid));
    }
    return boxes.map((b) => { let r = 0, g = 0, bl = 0; for (const p of b) { r += p[0]; g += p[1]; bl += p[2]; } const n = Math.max(1, b.length); return [Math.round(r/n), Math.round(g/n), Math.round(bl/n)]; });
  }
  function gifNearest(pal, r, g, b, cache) {
    const qr = r & 0xF0, qg = g & 0xF0, qb = b & 0xF0, key = (qr << 16) | (qg << 8) | qb;
    const hit = cache.get(key); if (hit !== undefined) return hit;
    let best = 0, bd = Infinity;
    for (let i = 0; i < pal.length; i++) { const dr = pal[i][0]-qr, dg = pal[i][1]-qg, db = pal[i][2]-qb, d = dr*dr+dg*dg+db*db; if (d < bd) { bd = d; best = i; } }
    cache.set(key, best); return best;
  }
  const LZW_DICT = new Int32Array(1 << 20);
  function lzwEncode(minCodeSize, indices) {
    const initBits = minCodeSize + 1, clearCode = 1 << minCodeSize, eofCode = clearCode + 1;
    const maxbits = 12, maxmaxcode = 1 << maxbits, dict = LZW_DICT; dict.fill(-1);
    const out = []; let nBits = initBits, maxcode = (1 << nBits) - 1, freeEnt = clearCode + 2, clearFlg = false, cur = 0, curBits = 0;
    const MAXCODE = (n) => (1 << n) - 1;
    const output = (code) => {
      cur |= code << curBits; curBits += nBits;
      while (curBits >= 8) { out.push(cur & 0xff); cur >>= 8; curBits -= 8; }
      if (freeEnt > maxcode || clearFlg) { if (clearFlg) { nBits = initBits; maxcode = MAXCODE(nBits); clearFlg = false; } else { nBits++; maxcode = (nBits === maxbits) ? maxmaxcode : MAXCODE(nBits); } }
    };
    const clBlock = () => { dict.fill(-1); freeEnt = clearCode + 2; clearFlg = true; output(clearCode); };
    output(clearCode); let ent = indices[0];
    for (let i = 1; i < indices.length; i++) { const c = indices[i], fcode = (ent << 8) | c, got = dict[fcode]; if (got !== -1) { ent = got; continue; } output(ent); ent = c; if (freeEnt < maxmaxcode) dict[fcode] = freeEnt++; else clBlock(); }
    output(ent); output(eofCode); if (curBits > 0) out.push(cur & 0xff); return out;
  }

  /* Build an animated, transparent GIF of the progressive reveal, cropped to the
     content bbox. Returns { dataUrl, bbox, frame } like the PNG exporters. */
  // Peso do traço SÓ no GIF (escolhido em comparação visual): a transparência
  // binária do GIF deixa cada grão 100% opaco e o traço lê mais denso que o
  // pincel; com peso 70%, menos grãos sobrevivem ao corte e o GIF recupera o
  // arejado do PNG. Não afeta o desenho ao vivo, o PNG nem a biblioteca.
  const GIF_STROKE_WEIGHT = 0.7;
  async function makeAnimatedGif(strokes, opts) {
    opts = opts || {};
    if (!strokes || !strokes.length) return null;
    strokes = strokes.map((s) => Object.assign({}, s, {
      opacity: (s.opacity != null ? s.opacity : 1) * GIF_STROKE_WEIGHT,
      _cache: null, _arc: null,
    }));
    const fps = opts.fps || 12, duration = opts.duration || 2.5;
    const holdMs = opts.holdMs != null ? opts.holdMs : 600, easing = opts.easing || 'linear';
    const loop = opts.loop !== false;   // default: loop forever
    const bbox = bboxOfStrokes(strokes); if (!bbox) return null;
    const CAP = 720, scale = Math.min(1, CAP / Math.max(bbox.w, bbox.h));
    const W = Math.max(1, Math.round(bbox.w * scale)), H = Math.max(1, Math.round(bbox.h * scale));
    const base = document.createElement('canvas'); base.width = W; base.height = H; const bctx = base.getContext('2d');
    const gif = document.createElement('canvas'); gif.width = W; gif.height = H; const gctx = gif.getContext('2d');
    const durs = strokes.map(strokeDur);
    const drawScaled = (c, stroke, pts) => { c.setTransform(scale, 0, 0, scale, -bbox.x * scale, -bbox.y * scale); drawStroke(c, stroke, pts); c.setTransform(1, 0, 0, 1, 0, 0); };
    let bakedUpTo = 0;
    const frameData = (p) => {
      const reveals = strokeRevealsFor(strokes, p, easing);
      while (bakedUpTo < strokes.length && reveals[bakedUpTo] >= durs[bakedUpTo]) { drawScaled(bctx, strokes[bakedUpTo]); bakedUpTo++; }
      gctx.setTransform(1,0,0,1,0,0); gctx.clearRect(0,0,W,H); gctx.drawImage(base,0,0);
      for (let i = bakedUpTo; i < strokes.length; i++) { const r = reveals[i]; if (r <= 0) continue; const s = strokes[i]; const pts = r >= durs[i] ? vectorize(s) : truncatePts(vectorize(s), r); if (pts.length >= 1) drawScaled(gctx, s, pts); }
      return gctx.getImageData(0,0,W,H).data;
    };
    const N = Math.max(2, Math.round(duration * fps)), baseDelay = Math.max(2, Math.round(100 / fps)), holdCenti = Math.round(holdMs / 10);
    for (const s of strokes) drawScaled(bctx, s);
    const palette = gifBuildPalette(bctx.getImageData(0,0,W,H).data, 255);
    bctx.clearRect(0,0,W,H);
    const realCount = palette.length, transIndex = realCount;
    let tableSize = 2; while (tableSize < realCount + 1) tableSize <<= 1;
    const minCodeSize = Math.max(2, Math.round(Math.log2(tableSize))), gctSizeField = Math.round(Math.log2(tableSize)) - 1;
    const bytes = [];
    const write16 = (v) => { bytes.push(v & 0xff, (v >> 8) & 0xff); };
    const writeStr = (s) => { for (let i = 0; i < s.length; i++) bytes.push(s.charCodeAt(i)); };
    writeStr('GIF89a'); write16(W); write16(H);
    bytes.push(0x80 | ((minCodeSize - 1) << 4) | gctSizeField, transIndex, 0);
    for (let i = 0; i < tableSize; i++) { if (i < realCount) bytes.push(palette[i][0], palette[i][1], palette[i][2]); else bytes.push(0, 0, 0); }
    if (loop) { bytes.push(0x21, 0xFF, 0x0B); writeStr('NETSCAPE2.0'); bytes.push(0x03, 0x01, 0x00, 0x00, 0x00); }
    const idx = new Uint8Array(W * H), cache = new Map();
    // Emit one frame. disposal: 1 = leave in place (reveal is additive),
    // 2 = restore to background (clears before the next frame).
    const emitFrame = (data, delay, disposal) => {
      for (let i = 0, pix = 0; i < data.length; i += 4, pix++) idx[pix] = data[i+3] < 128 ? transIndex : gifNearest(palette, data[i], data[i+1], data[i+2], cache);
      bytes.push(0x21, 0xF9, 0x04, ((disposal || 1) << 2) | 0x01, delay & 0xff, (delay >> 8) & 0xff, transIndex, 0x00);
      bytes.push(0x2C); write16(0); write16(0); write16(W); write16(H); bytes.push(0x00); bytes.push(minCodeSize);
      const lzw = lzwEncode(minCodeSize, idx);
      for (let o = 0; o < lzw.length; o += 255) { const len = Math.min(255, lzw.length - o); bytes.push(len); for (let j = 0; j < len; j++) bytes.push(lzw[o + j]); }
      bytes.push(0x00);
    };
    // Poster: make the FIRST frame the final (completed) state, so static
    // renderers (PDF export, thumbnails) show the finished drawing. We reuse
    // frameData(1) so the poster pixels are identical to the reveal's last
    // frame, then reset the bake state so the reveal still starts from empty.
    // disposal=2 clears the poster before the reveal. Seamless in a loop
    // (merges with the end hold); a brief flash on a non-looping GIF.
    if (opts.finalFirst !== false) {
      emitFrame(frameData(1), baseDelay, 2);
      bctx.clearRect(0, 0, W, H); bakedUpTo = 0;
    }
    for (let f = 0; f < N; f++) {
      const p = N === 1 ? 1 : f / (N - 1), data = frameData(p);
      const delay = f === N - 1 ? baseDelay + holdCenti : baseDelay;
      emitFrame(data, delay, 1);
      if (f % 4 === 0) await new Promise((r) => setTimeout(r)); // keep the webview responsive
    }
    bytes.push(0x3B);
    const blob = new Blob([new Uint8Array(bytes)], { type: 'image/gif' });
    const dataUrl = await new Promise((res) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(blob); });
    return { dataUrl, bbox, frame: { w: FRAME_W, h: FRAME_H } };
  }

  /* One GIF per stroke (separate) or one GIF for the whole drawing (together). */
  async function makeGifGroups(strokes, separate, opts) {
    const groups = separate ? strokes.map((s) => [s]) : [strokes];
    const out = [];
    for (const g of groups) { const gif = await makeAnimatedGif(g, opts); if (gif) out.push(gif); }
    return out;
  }
  function exportGifs(separate, opts) { return makeGifGroups(state.strokes, separate, opts); }
  async function renderExternalGifs(strokes, config, separate, opts) {
    const saved = { ms: state.mouseSmoothing, vs: state.vectorSmoothing, ft: state.faultTypes, fl: state.faultLevel, gs: state.gapSize };
    if (config) {
      state.mouseSmoothing = config.mouseSmoothing; state.vectorSmoothing = config.vectorSmoothing;
      state.faultTypes = config.faultTypes || []; state.faultLevel = config.faultLevel; state.gapSize = config.gapSize;
    }
    const out = await makeGifGroups(strokes, separate, opts);
    state.mouseSmoothing = saved.ms; state.vectorSmoothing = saved.vs;
    state.faultTypes = saved.ft; state.faultLevel = saved.fl; state.gapSize = saved.gs;
    return out;
  }

  global.Doodle = {
    state, FRAME_W, FRAME_H,
    attach, render, undo, redo, clear,
    exportTransparentPNG, contentBBox,
    renderExternalPNG, renderExternalPNGs, exportPNGs, payload, payloadOf, isEmpty,
    exportGifs, renderExternalGifs,
    vectorize, anchorCount,
    handleStrokeClick, getSelectedStroke, clearSelection, loadStrokes, thumbnailOf,
    thumbnailExternal,
  };
})(window);
