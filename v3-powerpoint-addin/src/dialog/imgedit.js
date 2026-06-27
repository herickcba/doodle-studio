/* ============================================================
   Image-edit dialog — show an image, mark it up with a SOLID brush
   (no chalk texture), type an edit prompt, and hand { prompt,
   markupDataUrl } back to the task pane (which calls /api/edit-image).
   ============================================================ */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);

  const img = $('editImg');
  const canvas = $('markCanvas');
  const ctx = canvas.getContext('2d');
  let strokes = [];          // [{ color, size, pts:[{x,y}] }]
  let cur = null, drawing = false;

  function loadImage() {
    let src = '';
    try { src = localStorage.getItem('doodle.editImage') || ''; } catch (_) {}
    img.onload = () => {
      canvas.width = img.naturalWidth || 1920;
      canvas.height = img.naturalHeight || 1080;
      redraw();
    };
    img.src = src;
  }

  function redraw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    for (const s of strokes) drawStroke(s);
    if (cur) drawStroke(cur);
  }
  function drawStroke(s) {
    if (s.pts.length < 1) return;
    ctx.save();
    ctx.strokeStyle = s.color; ctx.fillStyle = s.color;
    ctx.lineWidth = s.size; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    if (s.pts.length === 1) {
      ctx.beginPath(); ctx.arc(s.pts[0].x, s.pts[0].y, s.size / 2, 0, Math.PI * 2); ctx.fill();
    } else {
      ctx.beginPath(); ctx.moveTo(s.pts[0].x, s.pts[0].y);
      for (let i = 1; i < s.pts.length; i++) ctx.lineTo(s.pts[i].x, s.pts[i].y);
      ctx.stroke();
    }
    ctx.restore();
  }

  function pos(e) {
    const r = canvas.getBoundingClientRect();
    return { x: (e.clientX - r.left) * (canvas.width / r.width), y: (e.clientY - r.top) * (canvas.height / r.height) };
  }
  function onDown(e) {
    e.preventDefault();
    try { canvas.setPointerCapture(e.pointerId); } catch (_) {}
    drawing = true;
    cur = { color: $('brushColor').value, size: +$('brushSize').value, pts: [pos(e)] };
    redraw();
  }
  function onMove(e) {
    if (!drawing) return;
    cur.pts.push(pos(e)); redraw();
  }
  function onUp() {
    if (!drawing) return;
    drawing = false;
    if (cur && cur.pts.length) strokes.push(cur);
    cur = null; redraw();
  }

  function finish(kind) {
    if (kind === 'edit') {
      const prompt = $('editPrompt').value.trim();
      // Return ONLY the transparent scribble overlay (small — mostly transparent,
      // compresses tiny). The task pane composites it over the full-res base, so
      // we never push a multi-MB image through localStorage (quota would blow).
      let markupDataUrl = null;
      if (strokes.length) { try { markupDataUrl = canvas.toDataURL('image/png'); } catch (_) {} }
      try { localStorage.setItem('doodle.editResult', JSON.stringify({ prompt, markupDataUrl })); }
      catch (e) {
        // last resort: drop the overlay so at least the prompt-only edit fires
        try { localStorage.setItem('doodle.editResult', JSON.stringify({ prompt, markupDataUrl: null })); } catch (_) {}
      }
    }
    if (window.opener) { try { window.close(); } catch (_) {} return; }
    try { Office.context.ui.messageParent(kind); } catch (e) { try { window.close(); } catch (_) {} }
  }

  function syncApply() {
    $('applyBtn').disabled = !$('editPrompt').value.trim();
  }

  function boot() {
    loadImage();
    canvas.addEventListener('pointerdown', onDown);
    canvas.addEventListener('pointermove', onMove);
    canvas.addEventListener('pointerup', onUp);
    canvas.addEventListener('pointercancel', onUp);
    window.addEventListener('pointerup', onUp);
    $('undoBtn').addEventListener('click', () => { strokes.pop(); redraw(); });
    $('clearBtn').addEventListener('click', () => { strokes = []; redraw(); });
    $('editPrompt').addEventListener('input', syncApply);
    $('applyBtn').addEventListener('click', () => { if ($('editPrompt').value.trim()) finish('edit'); });
    $('cancelBtn').addEventListener('click', () => finish('cancel'));
    syncApply();
  }
  boot();
})();
