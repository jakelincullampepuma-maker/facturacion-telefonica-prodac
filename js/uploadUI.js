/* Selección compartida: sólo metadatos. Ninguna lectura ni petición de archivos. */
(function () {
  'use strict';
  if (window.UploadUI) return;
  const bound = new WeakMap();
  const MAX_BYTES = 100 * 1024 * 1024;
  const stages = ['pdf_read_ms', 'excel_read_ms', 'inventory_match_ms', 'reconciliation_ms', 'workbook_ms'];
  const diagnostics = window.UploadDiagnostics = {
    selections: [], runs: [], listeners: {},
    snapshot() { return JSON.parse(JSON.stringify({selections: this.selections, runs: this.runs, listeners: this.listeners})); }
  };
  const keep = (list, item, limit) => { list.push(item); if (list.length > limit) list.shift(); };
  const yieldToBrowser = () => new Promise(resolve => {
    if (document.visibilityState === 'visible') requestAnimationFrame(() => setTimeout(resolve, 0));
    else setTimeout(resolve, 0);
  });
  function startRun(module) {
    const run = {module, started: performance.now(), status: 'processing', timings: Object.fromEntries(stages.map(key => [key, 0]))};
    keep(diagnostics.runs, run, 20);
    return {
      run,
      async measure(stage, task) {
        const start = performance.now();
        try { return await task(); }
        finally { run.timings[stage] = (run.timings[stage] || 0) + performance.now() - start; }
      },
      merge(timings) { for (const [key, value] of Object.entries(timings || {})) if (Number.isFinite(value)) run.timings[key] = value; },
      finish(status = 'success') { run.status = status; run.total_ms = performance.now() - run.started; }
    };
  }
  function create({module, cards, getState, required, button, onChange, onError}) {
    let busy = false;
    const ready = () => required.every(key => Boolean(getState()[key]));
    function updateButton() { button.disabled = busy || !ready(); }
    function render(card) {
      const file = getState()[card.stateKey];
      card.name.textContent = file ? file.name : (card.emptyLabel || 'Arrastra o haz clic aquí');
      card.name.style.color = file ? 'var(--text-main)' : 'var(--text-muted)';
      for (const el of [card.check, card.rm]) if (el) {
        el.hidden = !file;
        el.classList.toggle('d-none', !file);
      }
      if (card.check) card.check.style.display = file ? 'block' : 'none';
      card.input.closest('.upload-file-card')?.classList.toggle('empty-border', !file);
    }
    function select(card, file, action = 'change') {
      if (busy) return;
      const start = performance.now();
      let error = '';
      if (file) {
        const extensions = card.input.accept.split(',').map(s => s.trim().toLowerCase()).filter(s => s.startsWith('.'));
        if (extensions.length && !extensions.some(ext => file.name.toLowerCase().endsWith(ext))) error = 'Seleccione un archivo ' + extensions.join(' o ') + ' válido.';
        else if (!file.size) error = 'El archivo está vacío.';
        else if (file.size > MAX_BYTES) error = 'El archivo supera el límite de 100 MB.';
      }
      getState()[card.stateKey] = error ? null : file;
      if (!file || error) card.input.value = '';
      render(card);
      onChange?.(card, getState()[card.stateKey]);
      updateButton();
      if (error) onError?.(error);
      if (action === 'change' || action === 'drop') keep(diagnostics.selections, {
        module, field: card.stateKey, bytes: file?.size || 0,
        valid: !error, action, selection_ms: performance.now() - start
      }, 300);
    }
    for (const card of cards) {
      if (bound.has(card.input)) throw new Error('La tarjeta de carga ya tiene un controlador.');
      bound.set(card.input, true);
      diagnostics.listeners[module + ':' + card.stateKey] = 1;
      card.input.addEventListener('change', () => select(card, card.input.files[0] || null));
      card.drop.addEventListener('click', event => {
        if (!busy && event.target !== card.input && !event.target.closest('button')) card.input.click();
      });
      card.drop.addEventListener('keydown', event => {
        if (!busy && event.target === card.drop && ['Enter', ' '].includes(event.key)) { event.preventDefault(); card.input.click(); }
      });
      card.drop.addEventListener('dragover', event => { event.preventDefault(); if (!busy) card.drop.classList.add('drag-over'); });
      card.drop.addEventListener('dragleave', () => card.drop.classList.remove('drag-over'));
      card.drop.addEventListener('drop', event => {
        event.preventDefault(); card.drop.classList.remove('drag-over');
        if (event.dataTransfer.files[0]) select(card, event.dataTransfer.files[0], 'drop');
      });
      card.rm?.addEventListener('click', event => { event.stopPropagation(); select(card, null, 'remove'); });
    }
    const controls = {
      ready,
      refresh() { cards.forEach(render); updateButton(); },
      clear() { if (!busy) cards.forEach(card => select(card, null, 'clear')); },
      setBusy(value) {
        busy = value;
        for (const card of cards) {
          card.input.disabled = busy;
          if (card.rm) card.rm.disabled = busy;
          card.drop.setAttribute('aria-disabled', String(busy));
          card.drop.style.pointerEvents = busy ? 'none' : '';
        }
        button.setAttribute('aria-busy', String(busy));
        updateButton();
      }
    };
    controls.refresh();
    return controls;
  }
  window.UploadUI = {create, startRun, yieldToBrowser, MAX_BYTES};
})();
