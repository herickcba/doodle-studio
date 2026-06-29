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
      sample: 'Ideia grande', font: 'Avenir Next', size: 120, lh: 0.8, bold: true,
      color: '#FC5E6D', periodColor: '#436AE1', bgAware: true, previewPx: 30 },
    { key: 'mega', name: 'Mega Statement', spec: 'Avenir Next Bold · 80pt · 0.9x',
      sample: 'Mega statement', font: 'Avenir Next', size: 80, lh: 0.9, bold: true,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 24 },
    { key: 'h1', name: 'H1 Título de Página', spec: 'Avenir Next Bold · 60pt · 0.9x',
      sample: 'Título de página', font: 'Avenir Next', size: 60, lh: 0.9, bold: true,
      color: '#FC5E6D', periodColor: null, bgAware: false, previewPx: 19 },
    { key: 'label_sec', name: 'Label de Seção', spec: 'Avenir Next Bold · 60pt · 0.9x',
      sample: 'Label de seção', font: 'Avenir Next', size: 60, lh: 0.9, bold: true,
      color: '#FC5E6D', periodColor: null, bgAware: false, previewPx: 19 },
    { key: 'corpo', name: 'Corpo de Texto', spec: 'Avenir Next Regular · 44pt · 1.15x',
      sample: 'Corpo de texto', font: 'Avenir Next', size: 44, lh: 1.15, bold: false,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 16 },
    { key: 'h3', name: 'H3 · Corpo Descritivo', spec: 'Avenir Next Bold · 34pt · 0.95x',
      sample: 'Corpo descritivo', font: 'Avenir Next', size: 34, lh: 0.95, bold: true,
      color: '#FC5E6D', periodColor: null, bgAware: false, previewPx: 15 },
    { key: 'h4', name: 'H4 · Subtítulo de Pilar', spec: 'Avenir Next Bold · 28pt · 1.0x',
      sample: 'Subtítulo de pilar', font: 'Avenir Next', size: 28, lh: 1.0, bold: true,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 14 },
    { key: 'h5', name: 'H5 · Texto Descritivo', spec: 'Avenir Next Regular · 24pt · 1.0x',
      sample: 'Texto descritivo', font: 'Avenir Next', size: 24, lh: 1.0, bold: false,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 14 },
    { key: 'corpo_pil', name: 'Corpo de Pilar', spec: 'Avenir Next Regular · 20pt · 1.3x',
      sample: 'Corpo de pilar', font: 'Avenir Next', size: 20, lh: 1.3, bold: false,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 13 },
    { key: 'eyebrow', name: 'Label / Eyebrow', spec: 'Avenir Next Bold · 18pt · 1.0x',
      sample: 'Label / eyebrow', font: 'Avenir Next', size: 18, lh: 1.0, bold: true,
      color: '#FC5E6D', periodColor: null, bgAware: false, previewPx: 13 },
    { key: 'caption', name: 'Legenda / Caption', spec: 'Avenir Next Regular · 16pt · 1.3x',
      sample: 'Legenda / caption', font: 'Avenir Next', size: 16, lh: 1.3, bold: false,
      color: '#436AE1', periodColor: null, bgAware: false, previewPx: 13 },
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
        + s.lh + ';font-family:' + FONT_STACK + ';font-weight:' + (s.bold ? 700 : 400) + ';';
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

  // Copy the macro code to the clipboard, stripping the `Attribute VB_Name`
  // line (valid in a .bas import, but errors when pasted into the VBE).
  async function copyMacro() {
    const btn = $('copyMacroBtn');
    try {
      const resp = await fetch('../../assets/BG-Tipografia.bas');
      const text = (await resp.text()).split('\n').filter((l) => l.indexOf('Attribute VB_Name') !== 0).join('\n').replace(/^\n+/, '');
      await navigator.clipboard.writeText(text);
      if (btn) { btn.textContent = '✓ Copiado'; setTimeout(() => { btn.textContent = 'Copiar código'; }, 2000); }
    } catch (e) {
      setStatus('Não foi possível copiar — use "Baixar .bas".', 'warn');
    }
  }

  function boot() {
    renderCards();
    const ref = $('insertRefBtn');
    if (ref) ref.addEventListener('click', insertReference);
    const copyBtn = $('copyMacroBtn');
    if (copyBtn) copyBtn.addEventListener('click', copyMacro);
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
