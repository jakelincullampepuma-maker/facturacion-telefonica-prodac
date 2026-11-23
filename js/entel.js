/**
 * js/entel.js
 * Modulo Entel - Conciliacion Telefonica PRODAC / PRODIMIN
 * 3 archivos de entrada: Excel Entel, Inventario Celulares General, PDF Entel.
 * Procesamiento dinámico local con SheetJS, PDF.js y JSZip.
 */
document.addEventListener('DOMContentLoaded', () => {

  const state = {
    excelFile:     null,
    inventoryFile: null,
    pdfFile:       null,
    excelBlob: null,
    filename: null,
    lastResult:    null,
    filterStatus:  'all',
    searchQuery:   '',
    processing: false,
  };

  function q(id) { return document.getElementById(id); }

  // Elementos de tarjetas de carga
  const CARDS = [
    { stateKey: 'excelFile',     input: q('input-excel-file'),     drop: q('dropzone-excel'),     name: q('name-excel-file'),     check: q('check-icon-excel'),     rm: q('btn-remove-excel') },
    { stateKey: 'inventoryFile', input: q('input-inventory-file'), drop: q('dropzone-inventory'), name: q('name-inventory-file'), check: q('check-icon-inventory'), rm: q('btn-remove-inventory') },
    { stateKey: 'pdfFile',       input: q('input-pdf-file'),       drop: q('dropzone-pdf'),       name: q('name-pdf-file'),       check: q('check-icon-pdf'),       rm: q('btn-remove-pdf') },
  ];

  const btnProcess    = q('btn-process-files');
  const btnClear      = q('btn-clear-files');
  const btnExport     = q('btn-export-excel');
  const alertBox      = q('entel-alert-container');

  const secResults    = q('section-resultados-entel');

  const kpis = {
    lineas:        q('kpi-val-lineas'),
    subLineas:     q('kpi-sub-lineas'),
    conciliadas:   q('kpi-val-conciliadas'),
    subConcil:     q('kpi-sub-conciliadas'),
    observaciones: q('kpi-val-observaciones'),
    subObs:        q('kpi-sub-observaciones'),
    excelPlan:     q('kpi-val-excel-plan'),
    pdfPlan:       q('kpi-val-pdf-plan'),
    equipos:       q('kpi-val-equipos'),
    totalFactura:  q('kpi-val-total-factura'),
    diferencia:    q('kpi-val-diferencia'),
    subDif:        q('kpi-sub-diferencia'),
  };

  const tbodyLines       = q('tbody-lineas-entel');
  const btnDlObs         = q('btn-download-obs');
  const searchInput      = q('input-search-lines');
  const filterSelect     = q('select-filter-status');
  const badgeFilterCount = q('badge-filtered-count');

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
  }

  function invalidateResult() {
    state.excelBlob = null;
    state.filename = null;
    state.lastResult = null;
    btnExport.disabled = true;
    secResults.classList.add('d-none');
    ModuleUI.setPeriod();
  }

  function setBusy(busy) {
    state.processing = busy;
    btnClear.disabled = busy;
    uploadControls.setBusy(busy);
  }

  // Navegacion lateral
  AppSidebar.bind({results: 'section-resultados-entel', missing: 'section-sin-cc-entel', download: 'btn-export-excel'});

  // ------------------------------------------------------------------
  // Alertas
  // ------------------------------------------------------------------
  function showAlert(msg, type = 'danger') {
    if (!alertBox) return;
    alertBox.innerHTML = msg;
  }

  function clearAlert() {
    if (alertBox) alertBox.innerHTML = '';
  }

  // ------------------------------------------------------------------
  // Formateo numerico en Soles
  // ------------------------------------------------------------------
  function fmt(n) {
    if (n === undefined || n === null || isNaN(n)) return 'No detectado';
    return 'S/ ' + Number(n).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ------------------------------------------------------------------
  // Dropzones
  // ------------------------------------------------------------------
  const uploadControls = UploadUI.create({module: 'entel', cards: ModuleUI.files(CARDS), getState: () => state,
    required: ['excelFile', 'inventoryFile', 'pdfFile'], button: btnProcess,
    onChange: () => { invalidateResult(); clearAlert(); },
    onError: message => showAlert(escapeHtml(message))
  });

  // ------------------------------------------------------------------
  // Limpiar
  // ------------------------------------------------------------------
  btnClear?.addEventListener('click', () => {
    if (state.processing) return;
    uploadControls.clear();

    state.excelBlob = null;
    state.filename = null;
    state.lastResult    = null;
    if (btnExport) btnExport.disabled = true;
    if (secResults) secResults.classList.add('d-none');
    ModuleUI.setPeriod();
    clearAlert();
  });

  // ------------------------------------------------------------------
  // Procesar archivos
  // ------------------------------------------------------------------
  btnProcess?.addEventListener('click', async () => {
    if (state.processing) return;
    clearAlert();

    const missing = [
      [state.excelFile,     'Excel de facturación de Entel (.xlsx)'],
      [state.inventoryFile, 'Inventario Celulares General (.xlsx)'],
      [state.pdfFile,       'Recibo PDF mensual de Entel (.pdf)'],
    ].find(([f]) => !f);

    if (missing) {
      showAlert(`Debe cargar: <strong>${missing[1]}</strong>.`, 'danger');
      return;
    }

    invalidateResult();
    const measurement = UploadUI.startRun('entel');
    let outcome = 'error';
    setBusy(true);
    btnProcess.disabled = true;
    const oldBtnContent = btnProcess.innerHTML;
    btnProcess.innerHTML = `
      <svg class="spin-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10" stroke-opacity=".2"/>
        <path d="M12 2a10 10 0 0 1 10 10"/>
      </svg>
      <span>Procesando…</span>`;

    try {
      await UploadUI.yieldToBrowser();
      const data = await EntelProcessor.process(state.excelFile, state.inventoryFile, state.pdfFile, measurement);
      state.excelBlob = data.excel_blob;
      state.filename = data.filename;
      state.lastResult = data;
      renderResults(data);
      outcome = 'success';

    } catch (err) {
      showAlert('No se pudieron procesar los archivos: ' + escapeHtml(err.message || String(err)), 'danger');
    } finally {
      setBusy(false);
      measurement.finish(outcome);
      btnProcess.innerHTML = oldBtnContent;
    }
  });

  // ------------------------------------------------------------------
  // Renderizar resultados
  // ------------------------------------------------------------------
  function renderResults(data) {
    if (!secResults) return;
    secResults.classList.remove('d-none');

    ModuleUI.setPeriod(data.periodo);
    const isCuadra = !!data.cuadra;
    const missingCount = (data.lines || []).filter(line => !line.cc_dpto).length;
    ModuleUI.status({cuadra:isCuadra, missing:missingCount, observations:data.obs_lines_count || data.cc_discrepancies?.length || 0,
      difference:fmt(data.diferencia), detail:isCuadra ? 'Los importes de PLAN coinciden con el recibo oficial.' : ''});

    // KPIs
    if (kpis.lineas)      kpis.lineas.textContent = data.line_count || 0;
    if (kpis.subLineas)   kpis.subLineas.textContent = `En hoja ${data.sheet_used || 'proveedor'}`;

    if (kpis.conciliadas) kpis.conciliadas.textContent = data.conciliado_count || 0;
    if (kpis.subConcil)   kpis.subConcil.textContent = 'Con CC Dpto asignado';

    const obsCount = (data.obs_lines_count !== undefined) ? data.obs_lines_count : ((data.sin_cc_count || 0) + (data.no_inv_count || 0));
    if (kpis.observaciones) kpis.observaciones.textContent = obsCount;
    if (kpis.subObs)        kpis.subObs.textContent = (data.sin_cc_count > 0 ? `${data.sin_cc_count} sin CC` : 'Discrepancias o notas');

    if (kpis.excelPlan)    kpis.excelPlan.textContent = fmt(data.total_excel);
    if (kpis.pdfPlan)      kpis.pdfPlan.textContent = fmt(data.total_pdf_plan);
    if (kpis.equipos)      kpis.equipos.textContent = fmt(data.total_pdf_equipos);
    if (kpis.totalFactura) kpis.totalFactura.textContent = fmt(data.total_pdf_factura);

    if (kpis.diferencia) {
      kpis.diferencia.textContent = fmt(data.diferencia);
      ModuleUI.difference('kpi-val-diferencia', data.diferencia);
    }
    if (kpis.subDif) {
      kpis.subDif.textContent = isCuadra ? 'Tolerancia máx S/ 0.01' : 'Diferencia detectable';
    }

    // Renderizar tablas
    renderLinesTable();
    renderCCTable(data.cc_summary || []);
    renderObsTable(data.obs_lines || [], data.cc_discrepancies || []);

    // Habilitar boton de descarga
    if (btnExport) btnExport.disabled = false;

    // Scroll suave a resultados
    secResults.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ------------------------------------------------------------------
  // Tabla Detalle de Lineas con Búsqueda y Filtros
  // ------------------------------------------------------------------
  function renderLinesTable() {
    if (!tbodyLines || !state.lastResult) return;
    const lines = state.lastResult.lines || [];
    const qText = (state.searchQuery || '').trim().toLowerCase();
    const stFilter = state.filterStatus;

    const filtered = lines.filter(item => {
      // Filtro por estado
      if (stFilter === 'conciliado' && !item.estado.toLowerCase().startsWith('conciliado')) return false;
      if (stFilter === 'sin_cc' && item.estado !== 'Sin CC') return false;
      if (stFilter === 'observacion' && !item.observacion && item.estado === 'Conciliado') return false;

      // Busqueda por texto
      if (qText) {
        const fullStr = [
          item.phone,
          item.plan,
          item.usuario,
          item.area,
          item.departamento,
          item.cc_dpto,
          item.prs_code,
          item.observacion,
        ].join(' ').toLowerCase();
        if (!fullStr.includes(qText)) return false;
      }
      return true;
    });

    if (badgeFilterCount) {
      badgeFilterCount.textContent = `${filtered.length} de ${lines.length} líneas`;
    }

    if (filtered.length === 0) {
      tbodyLines.innerHTML = `<tr><td colspan="10" style="text-align:center;padding:1.5rem;color:var(--text-muted);">No se encontraron líneas con los criterios seleccionados.</td></tr>`;
      return;
    }

    ModuleUI.rows('tbody-lineas-entel', filtered.map(item => [
      item.phone, item.plan || '-', fmt(item.cargo_fijo), item.empresa || '-',
      item.usuario || 'Sin info', (item.area || '-') + ' / ' + (item.departamento || '-'),
      item.cc_dpto || 'Sin CC', item.prs_code || '-', item.estado, item.observacion || '-'
    ]));
  }

  // Busqueda y filtros
  searchInput?.addEventListener('input', e => {
    state.searchQuery = e.target.value;
    renderLinesTable();
  });

  filterSelect?.addEventListener('change', e => {
    state.filterStatus = e.target.value;
    renderLinesTable();
  });

  // ------------------------------------------------------------------
  // Tabla Resumen por Centro de Costo (CC Dpto)
  // ------------------------------------------------------------------
  function renderCCTable(ccSummary) {
    const values = ccSummary.map(item => [item.cc, item.count, fmt(item.monto), item.pct.toFixed(1) + '%']);
    if (values.length) values.push(['Total general',ccSummary.reduce((n,c)=>n+c.count,0),fmt(ccSummary.reduce((n,c)=>n+c.monto,0)),'100.0%']);
    ModuleUI.rows('tbody-resumen-cc', values);
  }

  function renderObsTable(obsLines, ccDiscrepancies) {
    const review = [...obsLines];
    for (const discrepancy of ccDiscrepancies) {
      if (review.some(line => line.phone === discrepancy.phone)) continue;
      const line = state.lastResult.lines.find(line => line.phone === discrepancy.phone);
      if (line) review.push({...line, observacion:`CC del inventario: ${discrepancy.inv_cc}; CC anterior: ${discrepancy.hoja1_cc}`});
    }
    ModuleUI.review({body:'tbody-observaciones-entel', values:review.map(item => [item.phone, item.usuario || 'Sin info', item.cc_dpto || 'Sin CC',fmt(item.cargo_fijo),item.observacion || item.estado]), empty:'empty-obs-msg', button:'btn-download-obs'});
  }

  // ------------------------------------------------------------------
  // Descarga de Excel conciliado
  // ------------------------------------------------------------------
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  btnExport?.addEventListener('click', () => {
    if (!state.excelBlob) {
      showAlert('Procese los archivos antes de descargar.', 'warning');
      return;
    }
    download(state.excelBlob, state.filename);
  });

  btnDlObs?.addEventListener('click', () => {
    if (!state.lastResult) return;
    const columns = ['phone', 'usuario', 'cc_dpto', 'cargo_fijo', 'estado', 'observacion'];
    const csv = value => {
      let text = String(value ?? '');
      if (/^[=+@-]/.test(text)) text = "'" + text;
      return '"' + text.replace(/"/g, '""') + '"';
    };
    const rows = [['Número', 'Usuario', 'Centro de costo', 'Cargo Fijo', 'Estado', 'Observación'],
      ...state.lastResult.obs_lines.map(l => columns.map(c => l[c]))];
    download(new Blob(['\ufeff' + rows.map(r => r.map(csv).join(';')).join('\r\n')], {type: 'text/csv;charset=utf-8'}), 'Observaciones Entel.csv');
  });

});
