/* ============================================================
   Typography module — B+G typographic styles.
   Clickable preview cards (no dropdown). Clicking applies
   font/weight/size/color + colored final period to the selection
   via Office.js; exact line spacing comes from the "style anchor"
   reference slide (insertStylesReference). Works standalone too
   (previews render; apply needs PowerPoint).
   ============================================================ */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const FONT_STACK = "'Avenir Next','Avenir',-apple-system,'Segoe UI',sans-serif";

  // Each style mirrors the spec; previewPx scales the on-card sample.
  const STYLES = [
    { key: 'hero', name: 'Hero Statement', spec: 'Avenir Next Bold · 120pt · 0.8x',
      sample: 'Ideia grande', font: 'Avenir Next', size: 120, lh: 0.8,
      color: '#FC5E6D', periodColor: '#436AE1', bgAware: true, previewPx: 30 },
    { key: 'mega', name: 'Mega Statement', spec: 'Avenir Next Bold · 80pt · 0.9x',
      sample: 'Mega statement', font: 'Avenir Next', size: 80, lh: 0.9,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 23 },
    { key: 'h1', name: 'H1 Título de Página', spec: 'Avenir Next Bold · 60pt · 0.9x',
      sample: 'Título de página', font: 'Avenir Next', size: 60, lh: 0.9,
      color: '#FC5E6D', periodColor: null, bgAware: false, previewPx: 19 },
  ];

  function setStatus(msg, kind) {
    const el = $('typeStatus');
    if (!el) return;
    el.textContent = msg;
    el.className = 'status' + (kind ? ' ' + kind : '');
  }

  function renderCards() {
    const wrap = $('typeCards');
    if (!wrap) return;
    wrap.innerHTML = '';
    for (const s of STYLES) {
      const card = document.createElement('div');
      card.className = 'type-card';

      const prev = document.createElement('div');
      prev.className = 'type-preview';
      prev.style.cssText = 'color:' + s.color + ';font-size:' + s.previewPx + 'px;line-height:'
        + s.lh + ';font-family:' + FONT_STACK + ';font-weight:700;';
      prev.textContent = s.sample;
      const dot = document.createElement('span');
      dot.textContent = '.';
      dot.style.color = s.periodColor || s.color;
      prev.appendChild(dot);
      card.appendChild(prev);

      const row = document.createElement('div');
      row.className = 'type-row';
      const meta = document.createElement('div');
      meta.innerHTML = '<div class="type-name"></div><div class="type-spec"></div>';
      meta.querySelector('.type-name').textContent = s.name;
      meta.querySelector('.type-spec').textContent = s.spec;
      row.appendChild(meta);

      const btn = document.createElement('button');
      btn.className = 'btn type-apply';
      btn.textContent = 'Aplicar';
      btn.addEventListener('click', () => applyStyle(s, btn));
      row.appendChild(btn);

      card.appendChild(row);
      wrap.appendChild(card);
    }
  }

  async function applyStyle(style, btn) {
    if (!window.OfficeBridge || !OfficeBridge.inPowerPoint()) {
      setStatus('Aplicar funciona dentro do PowerPoint. (Esta é a prévia do estilo.)', 'warn');
      return;
    }
    btn.disabled = true;
    setStatus('Aplicando ' + style.name + '…');
    try {
      const r = await OfficeBridge.applyTextStyle(style);
      const where = r.mode === 'text' ? 'ao texto selecionado'
        : r.mode === 'shape' ? (r.applied > 1 ? 'aos objetos selecionados' : 'ao objeto selecionado')
        : 'numa caixa de texto nova';
      setStatus(style.name + ' aplicado ' + where + ' ✓ — para a entrelinha exata, use a referência.', 'ok');
    } catch (e) {
      setStatus(e && e.message ? e.message : 'Erro ao aplicar.', 'warn');
    } finally {
      btn.disabled = false;
    }
  }

  async function insertReference() {
    const btn = $('insertRefBtn');
    if (!window.OfficeBridge || !OfficeBridge.inPowerPoint()) {
      setStatus('A referência é inserida no PowerPoint. (Modo navegador: indisponível.)', 'warn');
      return;
    }
    btn.disabled = true;
    setStatus('Inserindo slide de referência…');
    try {
      await OfficeBridge.insertStylesReference();
      setStatus('Referência inserida ✓ — copie uma caixa para seu slide ou use o Pincel de Formatação. Apague o slide depois.', 'ok');
    } catch (e) {
      setStatus(e && e.message ? e.message : 'Erro ao inserir a referência.', 'warn');
    } finally {
      btn.disabled = false;
    }
  }

  function boot() {
    renderCards();
    const ref = $('insertRefBtn');
    if (ref) ref.addEventListener('click', insertReference);
    // When the module is shown, hint if "apply to selection" isn't available.
    if (window.DoodleModules) {
      DoodleModules.onShow.type = () => {
        if (window.OfficeBridge && OfficeBridge.inPowerPoint() && !OfficeBridge.apiSupported('1.6')) {
          setStatus('Seu PowerPoint não tem a API 1.6: use a referência (Pincel de Formatação) para aplicar os estilos.', 'warn');
        }
      };
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
