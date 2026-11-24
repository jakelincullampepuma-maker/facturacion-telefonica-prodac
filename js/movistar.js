/**
 * js/movistar.js
 * Módulo Movistar - PRODAC
 * Conciliación mensual dinámica basada en Recibo PDF e Inventario Celulares General.
 * Soporta validación opcional con Fact, Traf y Cargos.
 */
document.addEventListener('DOMContentLoaded', () => {

  // Determinar la URL del servidor backend
  let API_BASE = '';
  if (window.location.protocol === 'file:') {
    API_BASE = 'http://localhost:8080';
  } else if (window.location.hostname.endsWith('github.io')) {
    API_BASE = localStorage.getItem('PRODAC_API_URL') || 'https://facturacion-telefonica-prodac.onrender.com';
  }

  const state = {
    pdfFile:       null,
    inventoryFile: null,
    factFile:      null,
    trafFile:      null,
    cargosFile:    null,
    downloadToken: null,
    processing: false,
  };

  function q(id) { return document.getElementById(id); }
  function escapeHTML(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }
  function invalidateResults() {
    state.downloadToken = null;
    if (btnExport) btnExport.disabled = true;
    secResults?.classList.add('d-none');
    ModuleUI.setPeriod();
  }

  // Configuración de tarjetas de carga
  const CARDS = [
    { stateKey: 'pdfFile',       input: q('input-pdf-file'),       drop: q('dropzone-pdf'),       name: q('name-pdf-file'),       check: q('check-icon-pdf'),       rm: q('btn-remove-pdf')       },
    { stateKey: 'inventoryFile', input: q('input-inventory-file'), drop: q('dropzone-inventory'), name: q('name-inventory-file'), check: q('check-icon-inventory'), rm: q('btn-remove-inventory') },
    { stateKey: 'factFile',      input: q('input-fact-file'),      drop: q('dropzone-fact'),      name: q('name-fact-file'),      check: q('check-icon-fact'),      rm: q('btn-remove-fact')      },
    { stateKey: 'trafFile',      input: q('input-traf-file'),      drop: q('dropzone-traf'),      name: q('name-traf-file'),      check: q('check-icon-traf'),      rm: q('btn-remove-traf')      },
    { stateKey: 'cargosFile',    input: q('input-cargos-file'),    drop: q('dropzone-cargos'),    name: q('name-cargos-file'),    check: q('check-icon-cargos'),    rm: q('btn-remove-cargos')    },
  ];

  const inputPwd     = q('input-movistar-password');
  const btnTogglePwd = q('btn-toggle-password');
  const iconEyeOpen  = q('icon-eye-open');
  const iconEyeClose = q('icon-eye-closed');
  const btnProcess   = q('btn-process-files');
  const btnClear     = q('btn-clear-files');
  const btnExport    = q('btn-export-excel');
  const alertBox     = q('movistar-alert-container');

  const secResults   = q('section-resultados-movistar');

  // KPIs
  const kpis = {
    lineas:      q('kpi-val-lineas'),
    subLineas:   q('kpi-sub-lineas'),
    monto:       q('kpi-val-monto'),
    subMonto:    q('kpi-sub-monto'),
    desc:        q('kpi-val-desc'),
    subDesc:     q('kpi-sub-desc'),
    montoF:      q('kpi-val-monto-f'),
    igv:         q('kpi-val-igv'),
    totalAfecto: q('kpi-val-total-afecto'),
    inafectos:   q('kpi-val-inafectos'),
    subInaf:     q('kpi-sub-inafectos'),
    redondeo:    q('kpi-val-redondeo'),
    totalPdf:    q('kpi-val-total-pdf'),
    subCalc:     q('kpi-sub-calc'),
    dif:         q('kpi-val-diferencia'),
    subDif:      q('kpi-sub-diferencia'),
  };

  // Badges desglose
  const badges = {
    d12m:        q('badge-desc-12m'),
    d3m:         q('badge-desc-3m'),
    dperm:       q('badge-desc-perm'),
    d18m:        q('badge-desc-18m'),
    inafEquipos: q('badge-inaf-equipos'),
    inafVentas:  q('badge-inaf-ventas'),
  };

  const optWarnContainer = q('movistar-optional-warnings-container');
  const btnDlSinCC = q('btn-download-sin-cc');

  // Navegación lateral
  AppSidebar.bind({results: 'section-resultados-movistar', missing: 'section-sin-cc-movistar', download: 'btn-export-excel'});

  // ------------------------------------------------------------------
  // Alertas
  // ------------------------------------------------------------------
  function showAlert(msg, type='danger') {
    if (!alertBox) return;
    alertBox.innerHTML = msg;
  }
  function clearAlert() { if (alertBox) alertBox.innerHTML = ''; }

  // ------------------------------------------------------------------
  // Formateo monetario
  // ------------------------------------------------------------------
  function fmt(n) {
    if (n === undefined || n === null || isNaN(n)) return 'S/ 0.00';
    return 'S/ ' + Number(n).toLocaleString('es-PE', {minimumFractionDigits:2, maximumFractionDigits:2});
  }

  // ------------------------------------------------------------------
  // Manejo de Dropzones
  // ------------------------------------------------------------------
  const uploadControls = UploadUI.create({module: 'movistar', cards: ModuleUI.files(CARDS), getState: () => state,
    required: ['pdfFile', 'inventoryFile'], button: btnProcess,
    onChange: () => { invalidateResults(); clearAlert(); },
    onError: message => showAlert(escapeHTML(message))
  });

  // ------------------------------------------------------------------
  // Toggle password
  // ------------------------------------------------------------------
  btnTogglePwd?.addEventListener('click', () => {
    const isPwd = inputPwd.getAttribute('type') === 'password';
    inputPwd.setAttribute('type', isPwd ? 'text' : 'password');
    iconEyeOpen.classList.toggle('d-none', !isPwd);
    iconEyeClose.classList.toggle('d-none', isPwd);
  });

  // ------------------------------------------------------------------
  // Procesar archivos
  // ------------------------------------------------------------------
  btnProcess?.addEventListener('click', async () => {
    if (state.processing) return;
    clearAlert();

    // Validar archivos obligatorios
    if (!state.pdfFile) {
      showAlert('Debe cargar el <strong>Recibo PDF Movistar</strong>.');
      return;
    }
    if (!state.inventoryFile) {
      showAlert('Debe cargar el archivo <strong>Inventario Celulares General</strong>.');
      return;
    }

    const password = inputPwd?.value?.trim() || '';
    if ((state.factFile || state.trafFile || state.cargosFile) && !password) {
      q('details-optional-val').open = true;
      showAlert('Ingrese la contraseña mensual para validar los Excel opcionales, o retírelos para procesar sólo PDF e inventario.', 'warning');
      inputPwd?.focus();
      return;
    }

    const formData = new FormData();
    formData.append('pdf_file',       state.pdfFile);
    formData.append('inventory_file', state.inventoryFile);

    // Archivos opcionales
    if (state.factFile)   formData.append('fact_file',   state.factFile);
    if (state.trafFile)   formData.append('traf_file',   state.trafFile);
    if (state.cargosFile) formData.append('cargos_file', state.cargosFile);
    if (password)         formData.append('password',    password);

    const measurement = UploadUI.startRun('movistar');
    let outcome = 'error';
    state.processing = true;
    uploadControls.setBusy(true);
    btnClear.disabled = true;
    if (inputPwd) inputPwd.disabled = true;
    invalidateResults();
    if (inputPwd) { inputPwd.value = ''; inputPwd.type = 'password'; }
    btnProcess.innerHTML = `<svg class="spin-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10" stroke-opacity=".2"/><path d="M12 2a10 10 0 0 1 10 10"/></svg><span>Procesando…</span>`;

    try {
      await UploadUI.yieldToBrowser();
      const resp = await measurement.measure('request_ms', () => fetch(`${API_BASE}/api/movistar/process`, { method: 'POST', body: formData }));
      if (inputPwd) inputPwd.value = '';  // limpiar contraseña inmediatamente por seguridad

      const data = await resp.json();
      if (!resp.ok || !data.success) {
        showAlert(escapeHTML(data.error || 'Error al procesar los archivos de Movistar.'));
        return;
      }

      state.downloadToken = data.download_token;
      measurement.merge(data.timings);
      renderResults(data);
      outcome = 'success';
      secResults.scrollIntoView({behavior:'smooth', block:'start'});
    } catch(err) {
      if (inputPwd) inputPwd.value = '';
      if (window.location.hostname.endsWith('github.io')) {
        showAlert(`No se pudo conectar con el servidor de procesamiento. Verifique la conexión o inicie el servidor local en <a href="http://localhost:8080/movistar" style="color:inherit;font-weight:700;text-decoration:underline;">http://localhost:8080/movistar</a>.`);
      } else {
        showAlert(`No se pudo procesar la solicitud. Verifique que el servidor esté activo y vuelva a intentarlo. Si abrió el HTML directamente, utilice <a href="http://localhost:8080/movistar" style="color:inherit;font-weight:700;text-decoration:underline;">http://localhost:8080/movistar</a>.`);
      }
    } finally {
      state.processing = false;
      uploadControls.setBusy(false);
      btnClear.disabled = false;
      if (inputPwd) inputPwd.disabled = false;
      measurement.finish(outcome);
      btnProcess.innerHTML = `<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"/></svg><span>Procesar archivos</span>`;
    }
  });

  // ------------------------------------------------------------------
  // Renderizar resultados
  // ------------------------------------------------------------------
  function renderResults(data) {
    secResults.classList.remove('d-none');
    if (btnExport) btnExport.disabled = false;
    ModuleUI.setPeriod(data.periodo);

    const cuadra = data.cuadra === true && data.estado === 'CUADRA';
    ModuleUI.status({cuadra, missing:data.sin_cc_count || 0,
      observations:(data.warnings?.length || 0) + (data.optional_warnings?.length || 0),
      difference:fmt(data.diferencia), detail:cuadra ? 'Los importes calculados coinciden con el recibo oficial.' : ''});

    // KPIs
    if (kpis.lineas)      kpis.lineas.textContent      = data.line_count || 0;
    if (kpis.subLineas)   kpis.subLineas.textContent   = `${data.line_count || 0} anexos en el PDF`;
    if (kpis.monto)       kpis.monto.textContent       = fmt(data.cargo_mensual_sin_igv);
    if (kpis.subMonto)    kpis.subMonto.textContent    = 'Incluye descuentos incorporados';
    q('subtotal-comparison').textContent = `PDF: ${fmt(data.subtotal_pdf)} · Diferencia: ${fmt(data.diferencia_subtotal)}`;
    q('kpi-val-total-calculado').textContent = fmt(data.total_calculado);
    if (kpis.desc)        kpis.desc.textContent        = fmt(data.descuento_total);
    if (kpis.subDesc)     kpis.subDesc.textContent     = `${data.discount_breakdown?.total_desc_count || 0} descuentos en DESC`;
    if (kpis.montoF)      kpis.montoF.textContent      = fmt(data.cargo_final_sin_igv);
    if (kpis.igv)         kpis.igv.textContent         = fmt(data.igv);
    if (kpis.totalAfecto) kpis.totalAfecto.textContent = fmt(data.total_afecto);
    if (kpis.inafectos)   kpis.inafectos.textContent   = fmt(data.cargos_inafectos);
    if (kpis.subInaf)     kpis.subInaf.textContent     = `${data.inafectos_breakdown?.count_equipos || 0} equipos, ${data.inafectos_breakdown?.count_ventas || 0} ventas`;
    if (kpis.redondeo)    kpis.redondeo.textContent    = fmt(data.redondeo);
    if (kpis.totalPdf)    kpis.totalPdf.textContent    = fmt(data.pdf_total);
    if (kpis.subCalc)     kpis.subCalc.textContent     = `Recibo ${data.recibo_numero || ''} · A pagar: ${fmt(data.total_pagar ?? data.pdf_total)}`;
    if (kpis.dif) {
      kpis.dif.textContent = fmt(data.diferencia);
      ModuleUI.difference('kpi-val-diferencia', data.diferencia);
    }
    if (kpis.subDif)      kpis.subDif.textContent      = cuadra ? 'Estado: CUADRA' : 'Revisar diferencias';

    // Desglose de badges
    const db = data.discount_breakdown || {};
    if (badges.d12m)   badges.d12m.textContent   = `12 meses: ${db.count_12m || 0} (${fmt(db.monto_12m)})`;
    if (badges.d3m)    badges.d3m.textContent    = `3 meses: ${db.count_3m || 0} (${fmt(db.monto_3m)})`;
    if (badges.dperm)  badges.dperm.textContent  = `Permanentes: ${db.count_perm || 0} (${fmt(db.monto_perm)})`;
    if (badges.d18m)   badges.d18m.textContent   = `18m incorporados en MONTO: ${db.count_18m || 0}`;

    const ib = data.inafectos_breakdown || {};
    const warningsBox = q('movistar-warnings-container');
    warningsBox.classList.toggle('d-none', !data.warnings?.length);
    warningsBox.innerHTML = (data.warnings || []).map(w => `<div class="warning-item-box">${escapeHTML(w)}</div>`).join('');
    if (badges.inafEquipos) badges.inafEquipos.textContent = `Equipos inafectos: ${ib.count_equipos || 0}`;
    if (badges.inafVentas)  badges.inafVentas.textContent  = `Ventas inafectas: ${ib.count_ventas || 0}`;

    if (optWarnContainer) {
      optWarnContainer.classList.toggle('d-none', !data.optional_warnings?.length);
      optWarnContainer.innerHTML = (data.optional_warnings || []).map(w => `<div class="warning-item-box">${escapeHTML(w)}</div>`).join('');
    }
    renderCC(data.resumen_cc || []);
    renderSinCC(data.sin_cc_lines || [], [...(data.warnings || []), ...(data.optional_warnings || [])]);
  }

  function renderCC(list) {
    ModuleUI.rows('tbody-resumen-cc', list.map(item => [item.cc, item.lineas, fmt(item.monto), item.pct + '%']));
  }

  function renderSinCC(lines, notes) {
    ModuleUI.review({body:'tbody-sin-cc-movistar', values:lines.map(item => [item.phone, item.plan || 'Sin plan', fmt(item.monto_f), 'Sin CC']), notes, button:'btn-download-sin-cc'});
  }

  // ------------------------------------------------------------------
  // Descargas
  // ------------------------------------------------------------------
  btnExport?.addEventListener('click', () => {
    if (!state.downloadToken) { showAlert('Primero procese los archivos.'); return; }
    window.location.href = `${API_BASE}/api/movistar/download-excel?token=${state.downloadToken}`;
  });

  btnDlSinCC?.addEventListener('click', () => {
    if (state.downloadToken) window.location.href = `${API_BASE}/api/movistar/download-observations?token=${state.downloadToken}`;
  });

  // ------------------------------------------------------------------
  // Limpiar
  // ------------------------------------------------------------------
  btnClear?.addEventListener('click', () => {
    if (state.processing) return;
    uploadControls.clear();
    if (inputPwd) inputPwd.value = '';
    state.downloadToken = null;
    if (btnExport) btnExport.disabled = true;
    secResults?.classList.add('d-none');
    ModuleUI.setPeriod();
    if (optWarnContainer) {
      optWarnContainer.classList.add('d-none');
      optWarnContainer.innerHTML = '';
    }
    clearAlert();
  });
});
