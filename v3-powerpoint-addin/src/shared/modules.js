/* ============================================================
   Module navigation — top tabs switch the active view.
   Each module's UI lives in a <div class="view" id="view-KEY">;
   the tab's data-view picks which one is visible. Lightweight,
   no framework: just toggles a `.hidden` class + active tab.
   ============================================================ */
(function (global) {
  'use strict';

  function show(view) {
    document.querySelectorAll('.view').forEach((v) => {
      v.classList.toggle('hidden', v.id !== 'view-' + view);
    });
    document.querySelectorAll('.modtab').forEach((t) => {
      const on = t.dataset.view === view;
      t.classList.toggle('active', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    if (global.DoodleModules.onShow[view]) {
      try { global.DoodleModules.onShow[view](); } catch (_) {}
    }
  }

  function boot() {
    document.querySelectorAll('.modtab').forEach((t) => {
      t.addEventListener('click', () => show(t.dataset.view));
    });
  }

  // Modules can register a callback to run when their view is shown.
  global.DoodleModules = { show, onShow: {} };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})(window);
