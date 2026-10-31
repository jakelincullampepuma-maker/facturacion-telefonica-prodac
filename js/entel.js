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
  const periodText    = q('period-display-text');
  const bannerCard    = q('conciliation-banner');
  const bannerTitle   = q('banner-title-text');
  const bannerSub     = q('banner-sub-text');
  const bannerPill    = q('banner-pill-badge');
  const bannerCircle  = q('banner-status-circle');

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
  const tbodyCC          = q('tbody-resumen-cc');
  const tbodyObs         = q('tbody-observaciones-entel');
  const btnDlObs         = q('btn-download-obs');
  const emptyObsMsg      = q('empty-obs-msg');
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
  }

  function setBusy(busy) {
    btnClear.disabled = busy;
    CARDS.forEach(c => {
      c.input.disabled = busy;
      c.rm.disabled = busy;
      c.drop.style.pointerEvents = busy ? 'none' : '';
    });
  }

  // Navegacion lateral
  function setActiveNav(btn) {
    document.querySelectorAll('.sidebar-nav .nav-item').forEach(b => b.classList.remove('active'));
    if (btn) btn.classList.add('active');
  }

  q('nav-dashboard')?.addEventListener('click', () => { setActiveNav(q('nav-dashboard')); window.scrollTo({ top: 0, behavior: 'smooth' }); });
  q('nav-upload')?.addEventListener('click',    () => { setActiveNav(q('nav-upload'));    q('section-carga-archivos')?.scrollIntoView({ behavior: 'smooth' }); });
  q('nav-results')?.addEventListener('click',   () => {
    setActiveNav(q('nav-results'));
    if (secResults && !secResults.classList.contains('d-none')) secResults.scrollIntoView({ behavior: 'smooth' });
    else showAlert('Primero procese los archivos para visualizar los resultados.', 'warning');
  });
  q('nav-resumen')?.addEventListener('click', () => {
    setActiveNav(q('nav-resumen'));
    if (secResults && !secResults.classList.contains('d-none')) q('section-resumen-cc')?.scrollIntoView({ behavior: 'smooth' });
    else showAlert('Primero procese los archivos.', 'warning');
  });
  q('nav-obs')?.addEventListener('click', () => {
    setActiveNav(q('nav-obs'));
    if (secResults && !secResults.classList.contains('d-none')) q('section-observaciones-entel')?.scrollIntoView({ behavior: 'smooth' });
    else showAlert('Primero procese los archivos.', 'warning');
  });
  q('nav-export')?.addEventListener('click', () => {
    if (btnExport && !btnExport.disabled) btnExport.click();
    else showAlert('Primero procese los archivos para habilitar la descarga.', 'warning');
  });

  // ------------------------------------------------------------------
  // Alertas
  // ------------------------------------------------------------------
  function showAlert(msg, type = 'danger') {
    if (!alertBox) return;
    const colors = {
      danger:  '#FEF2F2;border-color:#FCA5A5;color:#991B1B;',
      warning: '#FFFBEB;border-color:#FCD34D;color:#92400E;',
      info:    '#EFF6FF;border-color:#93C5FD;color:#1E40AF;',
      success: '#ECFDF5;border-color:#6EE7B7;color:#065F46;',
    };
    alertBox.innerHTML = `
      <div style="padding:0.9rem 1.1rem;border-radius:0.6rem;border:1.5px solid;display:flex;align-items:flex-start;gap:0.75rem;font-size:0.875rem;background:${colors[type] || colors.danger}">
        <div style="flex:1;">${msg}</div>
        <button type="button" onclick="this.parentElement.remove()" style="background:none;border:none;cursor:pointer;color:inherit;font-weight:700;line-height:1;padding:0;font-size:1.1rem;">&times;</button>
      </div>`;
    alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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
  CARDS.forEach(card => {
    const { stateKey, input, drop, name, check, rm } = card;
    if (!drop || !input) return;

    drop.addEventListener('click', e => {
      if (!e.target.closest('.btn-remove-card-file')) input.click();
    });
    drop.addEventListener('dragover', e => {
      e.preventDefault();
      drop.classList.add('drag-over');
    });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
    drop.addEventListener('drop', e => {
      e.preventDefault();
      drop.classList.remove('drag-over');
      if (e.dataTransfer.files?.[0]) assignFile(e.dataTransfer.files[0], card);
    });
    input.addEventListener('change', e => {
      if (e.target.files?.[0]) assignFile(e.target.files[0], card);
    });
    rm?.addEventListener('click', e => {
      e.stopPropagation();
      invalidateResult();
      state[stateKey] = null;
      input.value = '';
      name.textContent = 'Arrastra o haz clic aquí';
      name.style.color = 'var(--text-muted)';
      check.classList.add('d-none');
      rm.classList.add('d-none');
    });
  });

  function assignFile(file, card) {
    const { stateKey, name, check, rm } = card;
    invalidateResult();
    state[stateKey] = file;
    name.textContent = file.name;
    name.style.color = 'var(--text-main)';
    check.classList.remove('d-none');
    rm.classList.remove('d-none');
    clearAlert();
  }

  // ------------------------------------------------------------------
  // Limpiar
  // ------------------------------------------------------------------
  btnClear?.addEventListener('click', () => {
    CARDS.forEach(card => {
      state[card.stateKey] = null;
      if (card.input) card.input.value = '';
      if (card.name) {
        card.name.textContent = 'Arrastra o haz clic aquí';
        card.name.style.color = 'var(--text-muted)';
      }
      card.check?.classList.add('d-none');
      card.rm?.classList.add('d-none');
    });

    state.excelBlob = null;
    state.filename = null;
    state.lastResult    = null;
    if (btnExport) btnExport.disabled = true;
    if (secResults) secResults.classList.add('d-none');
    if (periodText) periodText.textContent = 'Periodo: Pendiente de carga';
    clearAlert();
    showAlert('Se han limpiado todos los archivos cargados.', 'info');
  });

  // ------------------------------------------------------------------
  // Procesar archivos
  // ------------------------------------------------------------------
  btnProcess?.addEventListener('click', async () => {
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
    setBusy(true);
    btnProcess.disabled = true;
    const oldBtnContent = btnProcess.innerHTML;
    btnProcess.innerHTML = `
      <svg class="spin-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10" stroke-opacity=".2"/>
        <path d="M12 2a10 10 0 0 1 10 10"/>
      </svg>
      <span>Procesando archivos...</span>`;

    try {
      const data = await EntelProcessor.process(state.excelFile, state.inventoryFile, state.pdfFile);
      state.excelBlob = data.excel_blob;
      state.filename = data.filename;
      state.lastResult = data;
      renderResults(data);

      if (data.cuadra) {
        showAlert('Los importes de PLAN coinciden. ' + (data.obs_lines_count ? 'Hay ' + data.obs_lines_count + ' línea(s) con observaciones; revise el detalle.' : 'Sin observaciones.'), data.obs_lines_count ? 'warning' : 'success');
      } else {
        showAlert(`Archivos procesados. Se encontró una diferencia de S/ ${data.diferencia.toFixed(2)} entre el Excel y el PDF.`, 'warning');
      }

    } catch (err) {
      console.error(err);
      showAlert('No se pudieron procesar los archivos: ' + escapeHtml(err.message || String(err)), 'danger');
    } finally {
      setBusy(false);
      btnProcess.disabled = false;
      btnProcess.innerHTML = oldBtnContent;
    }
  });

  // ------------------------------------------------------------------
  // Renderizar resultados
  // ------------------------------------------------------------------
  function renderResults(data) {
    if (!secResults) return;
    secResults.classList.remove('d-none');

    // Periodo
    if (periodText && data.periodo) {
      periodText.textContent = `Periodo: ${data.periodo}`;
    }

    // Banner de conciliacion
    const isCuadra = !!data.cuadra;
    if (bannerCard) {
      bannerCard.style.borderColor = isCuadra ? '#10B981' : '#F59E0B';
      bannerCard.style.background  = isCuadra ? 'rgba(16, 185, 129, 0.05)' : 'rgba(245, 158, 11, 0.05)';
    }
    if (bannerCircle) {
      bannerCircle.style.background = isCuadra ? '#10B981' : '#F59E0B';
      bannerCircle.innerHTML = isCuadra
        ? `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#FFFFFF" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>`
        : `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#FFFFFF" stroke-width="3"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`;
    }
    if (bannerTitle) {
      bannerTitle.textContent = isCuadra ? 'La conciliación de planes cuadra' : 'Revisar conciliación con el PDF';
    }
    if (bannerSub) {
      bannerSub.textContent = data.mensaje_conciliacion || (isCuadra
        ? 'El total de cargos fijos de las líneas coincide con el concepto PLAN del recibo.'
        : `Diferencia de ${fmt(data.diferencia)} entre la suma del Excel y el recibo PDF.`);
    }
    if (bannerPill) {
      bannerPill.textContent = isCuadra ? 'CONCILIADO' : 'REVISAR';
      bannerPill.style.background = isCuadra ? '#10B981' : '#F59E0B';
      bannerPill.style.color = '#FFFFFF';
    }

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
      kpis.diferencia.className = `kpi-value-bold ${isCuadra ? 'text-green' : 'text-danger'}`;
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

    tbodyLines.innerHTML = filtered.map(item => {
      const hasObs = !!item.observacion;
      const isOk = item.estado === 'Conciliado';
      const statusBadge = isOk
        ? `<span style="background:rgba(16,185,129,0.12);color:#059669;padding:0.2rem 0.55rem;border-radius:12px;font-weight:600;font-size:0.75rem;">Conciliado</span>`
        : `<span style="background:rgba(239,68,68,0.12);color:#DC2626;padding:0.2rem 0.55rem;border-radius:12px;font-weight:600;font-size:0.75rem;">${escapeHtml(item.estado)}</span>`;

      return `
        <tr>
          <td style="font-family:monospace;font-weight:600;text-align:center;">${escapeHtml(item.phone)}</td>
          <td>${item.plan ? escapeHtml(item.plan) : '-'}</td>
          <td style="text-align:right;font-weight:600;">${fmt(item.cargo_fijo)}</td>
          <td style="text-align:center;">${escapeHtml(item.empresa || '-')}</td>
          <td>${item.usuario ? escapeHtml(item.usuario) : '<span style="color:#94A3B8;">Sin info</span>'}</td>
          <td style="font-size:0.8rem;">${item.area ? escapeHtml(item.area) : '-'} / ${item.departamento ? escapeHtml(item.departamento) : '-'}</td>
          <td style="text-align:center;font-weight:600;color:${item.cc_dpto ? 'inherit' : '#DC2626'};">${item.cc_dpto ? escapeHtml(item.cc_dpto) : 'Sin CC'}</td>
          <td style="text-align:center;color:#64748B;">${item.prs_code ? escapeHtml(item.prs_code) : '-'}</td>
          <td style="text-align:center;">${statusBadge}</td>
          <td style="font-size:0.8rem;color:${hasObs ? '#B45309' : '#64748B'};">${item.observacion ? escapeHtml(item.observacion) : '-'}</td>
        </tr>
      `;
    }).join('');
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
    if (!tbodyCC) return;
    if (!ccSummary || ccSummary.length === 0) {
      tbodyCC.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:1.5rem;color:var(--text-muted);">Sin datos de centro de costo.</td></tr>`;
      return;
    }

    const totalLines = ccSummary.reduce((acc, c) => acc + c.count, 0);
    const totalMonto = ccSummary.reduce((acc, c) => acc + c.monto, 0);

    const rowsHtml = ccSummary.map(item => {
      const isSinCC = item.cc === 'Sin Centro de Costo';
      return `
        <tr style="${isSinCC ? 'background:rgba(239,68,68,0.05);color:#DC2626;' : ''}">
          <td style="font-weight:600;text-align:center;">${escapeHtml(item.cc)}</td>
          <td style="text-align:right;">${item.count}</td>
          <td style="text-align:right;font-weight:600;">${fmt(item.monto)}</td>
          <td style="text-align:right;">${item.pct.toFixed(1)}%</td>
        </tr>
      `;
    }).join('');

    const totalRowHtml = `
      <tr style="font-weight:800;background:var(--bg-subtle, #F1F5F9);border-top:2px solid var(--border-color);">
        <td style="text-align:center;">Total general</td>
        <td style="text-align:right;">${totalLines}</td>
        <td style="text-align:right;">${fmt(totalMonto)}</td>
        <td style="text-align:right;">100.0%</td>
      </tr>
    `;

    tbodyCC.innerHTML = rowsHtml + totalRowHtml;
  }

  // ------------------------------------------------------------------
  // Tabla de Observaciones y Discrepancias
  // ------------------------------------------------------------------
  function renderObsTable(obsLines, ccDiscrepancies) {
    if (!tbodyObs) return;

    if (!obsLines || obsLines.length === 0) {
      tbodyObs.innerHTML = '';
      emptyObsMsg?.classList.remove('d-none');
      btnDlObs?.classList.add('d-none');
      return;
    }

    emptyObsMsg?.classList.add('d-none');
    btnDlObs?.classList.remove('d-none');

    tbodyObs.innerHTML = obsLines.map(item => `
      <tr>
        <td style="font-family:monospace;font-weight:600;text-align:center;">${escapeHtml(item.phone)}</td>
        <td>${item.usuario ? escapeHtml(item.usuario) : '<span style="color:#94A3B8;">Sin info</span>'}</td>
        <td style="text-align:center;font-weight:600;">${item.cc_dpto ? escapeHtml(item.cc_dpto) : '<span style="color:#DC2626;">Sin CC</span>'}</td>
        <td style="text-align:right;">${fmt(item.cargo_fijo)}</td>
        <td style="color:#92400E;font-size:0.83rem;">${escapeHtml(item.observacion || item.estado)}</td>
      </tr>
    `).join('');
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
