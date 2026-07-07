/* CBA Studio — Otimização (auditoria de mídia do .pptx)
   Tab própria: lista as imagens do deck por peso; clicar numa imagem a
   SELECIONA (navega até o slide + mostra a pré-visualização no card), pra
   saber exatamente qual é. "Otimizar" reencoda no tamanho exibido e insere
   por cima da original (a API não substitui — a original é apagada à mão).
   Análise só sob demanda (a leitura baixa o arquivo por fatias). */
(function () {
  'use strict';

  const view = document.getElementById('view-optimize');
  if (!view) return;

  const runBtn = document.getElementById('imgaRun');
  const summary = document.getElementById('imgaSummary');
  const list = document.getElementById('imgaList');
  const detail = document.getElementById('imgaDetail');
  const preview = document.getElementById('imgaPreview');
  const meta = document.getElementById('imgaMeta');
  const levelSeg = document.getElementById('imgaLevel');
  const optBtn = document.getElementById('imgaOpt');
  const optStatus = document.getElementById('imgaOptStatus');

  const fmt = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB');
  const inPP = () => !!(window.OfficeBridge && window.OfficeBridge.inPowerPoint());
  const isVector = (it) => it.ext === 'svg' || it.ext === 'emf' || it.ext === 'wmf';

  let items = [];
  let selected = null;   // item atualmente selecionado
  let level = 'suave';

  // segmented control de nível
  levelSeg.querySelectorAll('.seg-btn').forEach((b) => {
    b.addEventListener('click', () => {
      levelSeg.querySelectorAll('.seg-btn').forEach((x) => x.classList.remove('active'));
      b.classList.add('active');
      level = b.dataset.level;
    });
  });

  function rowLabel(it) {
    if (!it.slides.length) return 'sem uso direto';
    return 'slide ' + it.slides.map((s) => s.pos || '?').join(', ');
  }

  function renderList() {
    list.innerHTML = '';
    items.forEach((it) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'imga-item' + (selected === it ? ' on' : '');
      row.innerHTML =
        '<span class="imga-badge">' + (it.ext || '?').toUpperCase() + '</span>' +
        '<span class="imga-info"><b>' + fmt(it.bytes) + '</b> · ' + rowLabel(it) + '</span>';
      row.addEventListener('click', () => select(it));
      list.appendChild(row);
    });
  }

  async function select(it) {
    selected = it;
    renderList();
    detail.classList.remove('hidden');
    optStatus.textContent = '';
    optBtn.disabled = isVector(it);
    meta.innerHTML =
      '<b>' + fmt(it.bytes) + '</b> · ' + (it.ext || '?').toUpperCase() +
      ' · ' + rowLabel(it);
    preview.innerHTML = '<span class="imga-preview-ph">Carregando…</span>';

    // navega até o slide (quando a API suporta)
    const nav = it.slides.find((s) => s.sldId);
    if (nav) { try { await window.OfficeBridge.goToSlide(nav.sldId); } catch (_) {} }

    if (isVector(it)) {
      preview.innerHTML = '<span class="imga-preview-ph">Vetor (' + it.ext.toUpperCase() + ') — sem conversão</span>';
      return;
    }
    try {
      const thumb = await window.OfficeBridge.getImageThumbnail(it, (m) => {
        if (selected === it) preview.innerHTML = '<span class="imga-preview-ph">' + m + '</span>';
      });
      if (selected !== it) return; // trocou de imagem enquanto carregava
      if (thumb.url) {
        preview.innerHTML = '';
        const img = document.createElement('img');
        img.src = thumb.url; img.alt = it.base;
        preview.appendChild(img);
        if (thumb.w) meta.innerHTML += ' · ' + thumb.w + '×' + thumb.h + ' px';
      } else {
        preview.innerHTML = '<span class="imga-preview-ph">Sem pré-visualização</span>';
      }
    } catch (e) {
      if (selected === it) preview.innerHTML = '<span class="imga-preview-ph">Falha ao carregar</span>';
    }
  }

  optBtn.addEventListener('click', async () => {
    if (!selected) return;
    optBtn.disabled = true;
    try {
      const r = await window.OfficeBridge.optimizeImage(selected, level, (m) => { optStatus.textContent = m; });
      if (!r.smaller) {
        optStatus.textContent = 'Nesse nível a imagem não fica menor que ' + fmt(selected.bytes) + '. Tente um nível mais forte.';
      } else {
        optStatus.textContent = 'Versão ' + r.format + ' de ' + fmt(r.bytes) + ' inserida por cima (economia de ' +
          fmt(r.saved) + '). Apague a original: ela fica logo atrás da nova.';
      }
    } catch (e) {
      optStatus.textContent = 'Não deu: ' + (e && e.message ? e.message : e);
    } finally {
      optBtn.disabled = isVector(selected);
    }
  });

  runBtn.addEventListener('click', async () => {
    if (!inPP()) { summary.textContent = 'Disponível apenas dentro do PowerPoint.'; return; }
    runBtn.disabled = true;
    list.innerHTML = '';
    detail.classList.add('hidden');
    selected = null;
    summary.textContent = 'Analisando…';
    try {
      const audit = await window.OfficeBridge.getImageAudit((m) => { summary.textContent = m; });
      items = audit.items;
      if (!items.length) { summary.textContent = 'Nenhuma imagem no arquivo.'; return; }
      summary.textContent = items.length + ' imagem(ns) · total ' + fmt(audit.total) + '. Clique numa imagem para vê-la e otimizá-la.';
      renderList();
    } catch (e) {
      summary.textContent = 'Falha na análise: ' + (e && e.message ? e.message : e);
    } finally {
      runBtn.disabled = false;
    }
  });
})();
