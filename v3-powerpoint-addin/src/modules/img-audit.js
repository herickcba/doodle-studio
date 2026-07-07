/* CBA Studio — Imagens do arquivo (auditoria de mídia do .pptx)
   Lista as imagens do deck com extensão + peso (ordenado por peso),
   clique navega até o slide, "Otimizar" insere uma versão reencodada
   no tamanho exibido (a original é apagada à mão — a API não substitui).
   Análise só sob demanda (botão) — a leitura baixa o arquivo por slices. */
(function () {
  'use strict';

  const sec = document.getElementById('imgaSec');
  if (!sec) return;
  const head = document.getElementById('imgaHead');
  const body = document.getElementById('imgaBody');
  const runBtn = document.getElementById('imgaRun');
  const list = document.getElementById('imgaList');
  const statusEl = document.getElementById('imgaStatus');

  head.addEventListener('click', () => {
    const collapsed = sec.classList.toggle('collapsed');
    head.setAttribute('aria-expanded', String(!collapsed));
  });

  const say = (m) => { statusEl.textContent = m || ''; };
  const fmt = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB');

  let items = [];

  function render() {
    list.innerHTML = '';
    items.forEach((it, idx) => {
      const row = document.createElement('div');
      row.className = 'imga-item';

      const badge = document.createElement('span');
      badge.className = 'imga-badge';
      badge.textContent = (it.ext || '?').toUpperCase();

      const info = document.createElement('button');
      info.className = 'imga-info';
      info.type = 'button';
      const slideLabel = it.slides.length
        ? 'slide ' + it.slides.map((s) => s.pos || '?').join(', ')
        : 'sem uso direto';
      info.innerHTML = '<b>' + fmt(it.bytes) + '</b> · ' + slideLabel;
      info.title = it.base + (it.slides.length ? ' — clique para ir ao slide' : ' (layout/master ou não usado)');
      const nav = it.slides.find((s) => s.sldId);
      if (nav) {
        info.addEventListener('click', async () => {
          const ok = await window.OfficeBridge.goToSlide(nav.sldId);
          if (!ok) say('Navegação indisponível nesta versão do PowerPoint.');
        });
      } else { info.disabled = true; }

      const opt = document.createElement('button');
      opt.className = 'btn imga-opt';
      opt.type = 'button';
      opt.textContent = 'Otimizar';
      opt.title = 'Insere uma versão menor (no tamanho exibido) sobre a original — depois apague a original.';
      opt.addEventListener('click', async () => {
        opt.disabled = true;
        try {
          const r = await window.OfficeBridge.optimizeImage(it, say);
          if (!r.smaller) {
            say('Essa imagem já está eficiente (' + fmt(it.bytes) + ') — reencodar não ajudaria.');
          } else {
            say('Versão ' + r.format + ' de ' + fmt(r.bytes) + ' inserida por cima (economia de ' +
              fmt(r.saved) + '). Apague a original: ela fica logo atrás da nova.');
          }
        } catch (e) {
          say('Não deu: ' + (e && e.message ? e.message : e));
        } finally { opt.disabled = false; }
      });

      row.appendChild(badge);
      row.appendChild(info);
      row.appendChild(opt);
      list.appendChild(row);
      void idx;
    });
  }

  runBtn.addEventListener('click', async () => {
    if (!(window.OfficeBridge && window.OfficeBridge.inPowerPoint())) {
      say('Disponível apenas dentro do PowerPoint.');
      return;
    }
    runBtn.disabled = true;
    list.innerHTML = '';
    try {
      const audit = await window.OfficeBridge.getImageAudit(say);
      items = audit.items;
      if (!items.length) { say('Nenhuma imagem no arquivo.'); return; }
      say(items.length + ' imagem(ns) · total ' + fmt(audit.total) + '. Clique para ir ao slide.');
      render();
    } catch (e) {
      say('Falha na análise: ' + (e && e.message ? e.message : e));
    } finally { runBtn.disabled = false; }
  });
})();
