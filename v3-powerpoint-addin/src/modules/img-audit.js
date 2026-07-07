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
  const selectedBtn = document.getElementById('imgaSelected');
  const bulk = document.getElementById('imgaBulk');
  const downloadAllBtn = document.getElementById('imgaDownloadAll');
  const optAllBtn = document.getElementById('imgaOptAll');
  const allStatus = document.getElementById('imgaAllStatus');
  const list = document.getElementById('imgaList');
  const detail = document.getElementById('imgaDetail');
  const preview = document.getElementById('imgaPreview');
  const meta = document.getElementById('imgaMeta');
  const levelSeg = document.getElementById('imgaLevel');
  const optBtn = document.getElementById('imgaOpt');
  const replaceBtn = document.getElementById('imgaReplace');
  const replaceFile = document.getElementById('imgaReplaceFile');
  const optStatus = document.getElementById('imgaOptStatus');

  const fmt = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB');
  const inPP = () => !!(window.OfficeBridge && window.OfficeBridge.inPowerPoint());
  const isVector = (it) => it.ext === 'svg' || it.ext === 'emf' || it.ext === 'wmf';

  let items = [];
  let selected = null;   // item atualmente selecionado
  let level = 'fiel';

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
        if (thumb.w) meta.innerHTML += ' · ' + thumb.w + '×' + thumb.h + ' px' + (thumb.cropped ? ' (recortada)' : '');
      } else {
        preview.innerHTML = '<span class="imga-preview-ph">Sem pré-visualização</span>';
      }
    } catch (e) {
      if (selected === it) preview.innerHTML = '<span class="imga-preview-ph">Falha ao carregar</span>';
    }
  }

  optBtn.addEventListener('click', async () => {
    if (!selected) return;
    optBtn.disabled = true; replaceBtn.disabled = true;
    try {
      const r = await window.OfficeBridge.optimizeImage(selected, level, (m) => { optStatus.textContent = m; });
      if (!r.smaller) {
        optStatus.textContent = 'Nesse nível a imagem não fica menor que ' + fmt(selected.bytes) + '. Tente um nível mais forte.';
      } else {
        const kept = r.cropped || r.rotated ? ' Recorte/proporção mantidos.' : '';
        const many = r.usages > 1 ? ' em ' + r.usages + ' slides' : '';
        optStatus.textContent = 'Colada por cima' + many + ' — versão ' + r.format + ' de ' + fmt(r.bytes) + '.' +
          kept + ' A original (' + fmt(selected.bytes) + ') fica atrás; apague-a para reduzir o arquivo.';
      }
    } catch (e) {
      optStatus.textContent = 'Não deu: ' + (e && e.message ? e.message : e);
    } finally {
      optBtn.disabled = isVector(selected); replaceBtn.disabled = false;
    }
  });

  // Substituir por um arquivo do PC (cola por cima, mesmo lugar/tamanho)
  replaceBtn.addEventListener('click', () => { if (selected) replaceFile.click(); });
  replaceFile.addEventListener('change', async () => {
    const f = replaceFile.files && replaceFile.files[0];
    replaceFile.value = ''; // permite reescolher o mesmo arquivo
    if (!f || !selected) return;
    const dataUrl = await new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => res(fr.result);
      fr.onerror = () => rej(fr.error);
      fr.readAsDataURL(f);
    }).catch(() => null);
    if (!dataUrl) { optStatus.textContent = 'Não consegui ler o arquivo.'; return; }
    optBtn.disabled = true; replaceBtn.disabled = true;
    optStatus.textContent = 'Substituindo…';
    try {
      const r = await window.OfficeBridge.replaceImage(selected, dataUrl, (m) => { optStatus.textContent = m; });
      const many = r.usages > 1 ? ' em ' + r.usages + ' slides' : '';
      optStatus.textContent = '"' + f.name + '" colada por cima' + many +
        '. A original fica atrás; apague-a se não precisar mais.';
    } catch (e) {
      optStatus.textContent = 'Não deu: ' + (e && e.message ? e.message : e);
    } finally {
      optBtn.disabled = isVector(selected); replaceBtn.disabled = false;
    }
  });

  // Inspecionar a imagem SELECIONADA no slide (resolve → abre o card de detalhe)
  selectedBtn.addEventListener('click', async () => {
    if (!inPP()) { summary.textContent = 'Disponível apenas dentro do PowerPoint.'; return; }
    selectedBtn.disabled = true;
    const prev = selectedBtn.textContent;
    selectedBtn.textContent = 'Procurando a seleção…';
    try {
      const res = await window.OfficeBridge.getSelectedImage(items);
      if (res.none) { summary.textContent = 'Selecione uma imagem no slide e clique de novo.'; return; }
      if (res.multi) { summary.textContent = 'Selecione só UMA imagem no slide.'; return; }
      if (res.unmatched) { summary.textContent = 'Não identifiquei a seleção como imagem do arquivo (se colou agora, salve e tente).'; return; }
      const existing = items.find((it) => it.base === res.item.base) || res.item;
      await select(existing);
    } catch (e) {
      summary.textContent = 'Não deu: ' + (e && e.message ? e.message : e);
    } finally {
      selectedBtn.disabled = false;
      selectedBtn.textContent = prev;
    }
  });

  runBtn.addEventListener('click', async () => {
    if (!inPP()) { summary.textContent = 'Disponível apenas dentro do PowerPoint.'; return; }
    runBtn.disabled = true;
    list.innerHTML = '';
    detail.classList.add('hidden');
    bulk.classList.add('hidden');
    allStatus.classList.add('hidden');
    selected = null;
    summary.textContent = 'Analisando…';
    try {
      const audit = await window.OfficeBridge.getImageAudit((m) => { summary.textContent = m; });
      items = audit.items;
      if (!items.length) { summary.textContent = 'Nenhuma imagem no arquivo.'; return; }
      summary.textContent = items.length + ' imagem(ns) · total ' + fmt(audit.total) + '. Clique numa imagem para vê-la, ou baixe/otimize todas.';
      renderList();
      const rasterN = items.filter((it) => !isVector(it)).length;
      allLabel = 'Otimizar todas' + (rasterN ? ' (' + rasterN + ')' : '');
      disarmAll();
      optAllBtn.disabled = rasterN === 0;
      bulk.classList.remove('hidden');
    } catch (e) {
      summary.textContent = 'Falha na análise: ' + (e && e.message ? e.message : e);
    } finally {
      runBtn.disabled = false;
    }
  });

  // confirmação em 2 cliques (window.confirm não roda no webview do add-in)
  let allArmed = false, allTimer = null, allLabel = 'Otimizar todas';
  function disarmAll() {
    allArmed = false;
    if (allTimer) { clearTimeout(allTimer); allTimer = null; }
    optAllBtn.classList.remove('primary');
    optAllBtn.textContent = allLabel;
  }
  optAllBtn.addEventListener('click', async () => {
    if (!items.length) return;
    if (!allArmed) {
      allArmed = true;
      optAllBtn.classList.add('primary');
      optAllBtn.textContent = 'Confirmar: substituir todas? (Cmd+Z desfaz)';
      allTimer = setTimeout(disarmAll, 4500);
      return;
    }
    disarmAll();
    optAllBtn.disabled = true; downloadAllBtn.disabled = true;
    runBtn.disabled = true;
    allStatus.classList.remove('hidden');
    allStatus.textContent = 'Otimizando…';
    try {
      const r = await window.OfficeBridge.optimizeAll(items, level, (m) => { allStatus.textContent = m; });
      const parts = [r.done + ' de ' + r.total + ' colada(s) por cima'];
      if (r.skipped) parts.push(r.skipped + ' já no menor tamanho');
      if (r.failed) parts.push(r.failed + ' com erro');
      let msg = parts.join(' · ') + '. As originais ficam atrás' +
        (r.saved ? ' (~' + fmt(r.saved) + ')' : '') + ' — apague-as para reduzir o arquivo.';
      if (r.failed && r.lastError) msg += ' (erro: ' + r.lastError + ')';
      allStatus.textContent = msg;
    } catch (e) {
      allStatus.textContent = 'Falha: ' + (e && e.message ? e.message : e);
    } finally {
      optAllBtn.disabled = false; downloadAllBtn.disabled = false;
      runBtn.disabled = false;
    }
  });

  // Baixar todas as imagens num ZIP
  downloadAllBtn.addEventListener('click', async () => {
    if (!items.length) return;
    downloadAllBtn.disabled = true; optAllBtn.disabled = true;
    allStatus.classList.remove('hidden');
    allStatus.textContent = 'Preparando o ZIP…';
    try {
      const r = await window.OfficeBridge.exportAllImages(items, (m) => { allStatus.textContent = m; });
      allStatus.textContent = 'Baixado: ' + r.count + ' imagem(ns) (' + fmt(r.bytes) +
        ') em imagens-do-deck.zip. Reduza fora e recoloque com "Substituir".';
    } catch (e) {
      allStatus.textContent = 'Falha ao baixar: ' + (e && e.message ? e.message : e);
    } finally {
      downloadAllBtn.disabled = false;
      optAllBtn.disabled = items.filter((it) => !isVector(it)).length === 0;
    }
  });
})();
