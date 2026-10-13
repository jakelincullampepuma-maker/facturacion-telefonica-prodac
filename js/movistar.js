/**
 * js/movistar.js  v2.0
 * Modulo Movistar - PRODAC
 * 5 archivos de entrada: Fact, Traf, Cargos, Inventario, PDF
 * Descuentos extraidos automaticamente del PDF.
 */
document.addEventListener('DOMContentLoaded', () => {

  // Determinar la URL del servidor backend
  let API_BASE = '';
  if (window.location.protocol === 'file:') {
    API_BASE = 'http://localhost:8080';
  } else if (window.location.hostname.endsWith('github.io')) {
    // Si corre en GitHub Pages, conecta al backend desplegado en la nube
    API_BASE = localStorage.getItem('PRODAC_API_URL') || 'https://facturacion-telefonica-prodac.onrender.com';
  }

  const state = {
    factFile:      null,
    trafFile:      null,
    cargosFile:    null,
    inventoryFile: null,
    pdfFile:       null,
    downloadToken: null,
  };

  function q(id) { return document.getElementById(id); }

  // Elementos de archivos
  const CARDS = [
    { stateKey:'factFile',      input:q('input-fact-file'),      drop:q('dropzone-fact'),      name:q('name-fact-file'),      check:q('check-icon-fact'),      rm:q('btn-remove-fact')      },
    { stateKey:'trafFile',      input:q('input-traf-file'),      drop:q('dropzone-traf'),      name:q('name-traf-file'),      check:q('check-icon-traf'),      rm:q('btn-remove-traf')      },
    { stateKey:'cargosFile',    input:q('input-cargos-file'),    drop:q('dropzone-cargos'),    name:q('name-cargos-file'),    check:q('check-icon-cargos'),    rm:q('btn-remove-cargos')    },
    { stateKey:'inventoryFile', input:q('input-inventory-file'), drop:q('dropzone-inventory'), name:q('name-inventory-file'), check:q('check-icon-inventory'), rm:q('btn-remove-inventory') },
    { stateKey:'pdfFile',       input:q('input-pdf-file'),       drop:q('dropzone-pdf'),       name:q('name-pdf-file'),       check:q('check-icon-pdf'),       rm:q('btn-remove-pdf')       },
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
  const periodText   = q('period-display-text');
  const bannerCard   = q('conciliation-banner');
  const bannerTitle  = q('banner-title-text');
  const bannerSub    = q('banner-sub-text');
  const bannerPill   = q('banner-pill-badge');
  const bannerCircle = q('banner-status-circle');
  const kpis = {
    lineas:   q('kpi-val-lineas'),
    subLineas:q('kpi-sub-lineas'),
    desc:     q('kpi-val-desc-total'),
    subDesc:  q('kpi-sub-desc-count'),
    monto:    q('kpi-val-cargo-sin-igv'),
    montoF:   q('kpi-val-cargo-final-sin-igv'),
    igv:      q('kpi-val-igv'),
    equipos:  q('kpi-val-equipos'),
    totalPdf: q('kpi-val-total-pdf'),
    subCalc:  q('kpi-sub-calc'),
    dif:      q('kpi-val-diferencia'),
    subDif:   q('kpi-sub-diferencia'),
  };
  const tbodyCC    = q('tbody-resumen-cc');
  const tbodySinCC = q('tbody-sin-cc-movistar');
  const btnDlSinCC = q('btn-download-sin-cc');
  const emptyMsg   = q('empty-sin-cc-msg');

  // Nav
  function setActiveNav(btn) {
    document.querySelectorAll('.sidebar-nav .nav-item').forEach(b => b.classList.remove('active'));
    if (btn) btn.classList.add('active');
  }
  q('nav-dashboard')?.addEventListener('click', () => { setActiveNav(q('nav-dashboard')); window.scrollTo({top:0,behavior:'smooth'}); });
  q('nav-upload')?.addEventListener('click',    () => { setActiveNav(q('nav-upload'));    q('section-carga-archivos')?.scrollIntoView({behavior:'smooth'}); });
  q('nav-results')?.addEventListener('click',   () => {
    setActiveNav(q('nav-results'));
    if (secResults && !secResults.classList.contains('d-none')) secResults.scrollIntoView({behavior:'smooth'});
    else showAlert('Primero procese los archivos para ver resultados.','warning');
  });
  q('nav-sin-cc')?.addEventListener('click', () => {
    setActiveNav(q('nav-sin-cc'));
    if (secResults && !secResults.classList.contains('d-none')) q('section-sin-cc-movistar')?.scrollIntoView({behavior:'smooth'});
    else showAlert('Primero procese los archivos.','warning');
  });
  q('nav-export')?.addEventListener('click', () => {
    if (btnExport && !btnExport.disabled) btnExport.click();
    else showAlert('Primero procese los archivos para habilitar la descarga.','warning');
  });

  // ------------------------------------------------------------------
  // Alertas
  // ------------------------------------------------------------------
  function showAlert(msg, type='danger') {
    if (!alertBox) return;
    const colors = { danger:'#FEF2F2;border-color:#FCA5A5;color:#991B1B', warning:'#FFFBEB;border-color:#FCD34D;color:#92400E', info:'#EFF6FF;border-color:#93C5FD;color:#1E40AF' };
    alertBox.innerHTML = `<div style="padding:0.85rem 1rem;border-radius:0.6rem;border:1.5px solid;display:flex;align-items:flex-start;gap:0.75rem;font-size:0.875rem;background:${colors[type]||colors.danger}">
      <div style="flex:1;">${msg}</div>
      <button type="button" onclick="this.parentElement.remove()" style="background:none;border:none;cursor:pointer;color:inherit;font-weight:700;line-height:1;padding:0;">&times;</button>
    </div>`;
    alertBox.scrollIntoView({behavior:'smooth',block:'nearest'});
  }
  function clearAlert() { if (alertBox) alertBox.innerHTML = ''; }

  // ------------------------------------------------------------------
  // Formateo
  // ------------------------------------------------------------------
  function fmt(n) {
    if (n===undefined||n===null||isNaN(n)) return 'S/ 0.00';
    return 'S/ '+Number(n).toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2});
  }

  // ------------------------------------------------------------------
  // Dropzones
  // ------------------------------------------------------------------
  CARDS.forEach(card => {
    const {stateKey, input, drop, name, check, rm} = card;
    if (!drop || !input) return;

    drop.addEventListener('click', e => { if (!e.target.closest('.btn-remove-card-file')) input.click(); });
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
    drop.addEventListener('drop', e => {
      e.preventDefault(); drop.classList.remove('drag-over');
      if (e.dataTransfer.files?.[0]) assignFile(e.dataTransfer.files[0], card);
    });
    input.addEventListener('change', e => { if (e.target.files?.[0]) assignFile(e.target.files[0], card); });
    rm?.addEventListener('click', e => {
      e.stopPropagation();
      state[stateKey] = null;
      input.value = '';
      name.textContent = 'Arrastra o haz clic aqui';
      name.style.color = 'var(--text-muted)';
      check.classList.add('d-none');
      rm.classList.add('d-none');
    });
  });

  function assignFile(file, card) {
    const {stateKey, name, check, rm} = card;
    state[stateKey] = file;
    name.textContent = file.name;
    name.style.color = 'var(--text-main)';
    check.classList.remove('d-none');
    rm.classList.remove('d-none');
    clearAlert();
  }

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
  // Procesar
  // ------------------------------------------------------------------
  btnProcess?.addEventListener('click', async () => {
    clearAlert();
    const missing = [
      [state.factFile,      'Archivo Fact (Fact_*.xlsx)'],
      [state.trafFile,      'Archivo Traf (Traf_*.xlsx)'],
      [state.cargosFile,    'Archivo Cargos (Cargos_*.xlsx)'],
      [state.inventoryFile, 'Inventario Celulares General'],
      [state.pdfFile,       'Recibo PDF mensual de Movistar'],
    ].find(([f]) => !f);
    if (missing) { showAlert(`Debe cargar: <strong>${missing[1]}</strong>.`); return; }

    const password = inputPwd?.value?.trim() || '';
    if (!password) { showAlert('Debe ingresar la <strong>contrasena</strong> de los archivos del mes.'); inputPwd?.focus(); return; }

    const formData = new FormData();
    formData.append('fact_file',      state.factFile);
    formData.append('traf_file',      state.trafFile);
    formData.append('cargos_file',    state.cargosFile);
    formData.append('inventory_file', state.inventoryFile);
    formData.append('pdf_file',       state.pdfFile);
    formData.append('password',       password);

    btnProcess.disabled = true;
    btnProcess.innerHTML = `<svg class="spin-icon" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10" stroke-opacity=".2"/><path d="M12 2a10 10 0 0 1 10 10"/></svg><span>Procesando...</span>`;

    try {
      const resp = await fetch(`${API_BASE}/api/movistar/process`, { method:'POST', body:formData });
      if (inputPwd) inputPwd.value = '';  // limpiar password inmediatamente

      const data = await resp.json();
      if (!resp.ok || !data.success) { showAlert(data.error || 'Error al procesar los archivos.'); return; }

      state.downloadToken = data.download_token;
      renderResults(data);
      showAlert('Archivos procesados correctamente.','info');
      secResults.scrollIntoView({behavior:'smooth',block:'start'});
    } catch(err) {
      if (inputPwd) inputPwd.value = '';
      if (window.location.hostname.endsWith('github.io')) {
        showAlert(`No se pudo conectar con el servidor de procesamiento. Para usar la aplicación en línea, despliegue el backend en Render o abra la versión local con <strong>iniciar_servidor.bat</strong> en <a href="http://localhost:8080/movistar.html" style="color:inherit;font-weight:700;text-decoration:underline;">http://localhost:8080/movistar.html</a>.`);
      } else {
        showAlert(`Error de conexión con el servidor. Verifique que el servidor esté iniciado ejecutando <strong>iniciar_servidor.bat</strong> o abra <a href="http://localhost:8080/movistar.html" style="color:inherit;font-weight:700;text-decoration:underline;">http://localhost:8080/movistar.html</a>.`);
      }
    } finally {
      btnProcess.disabled = false;
      btnProcess.innerHTML = `<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"/></svg><span>Procesar archivos</span>`;
    }
  });

  // ------------------------------------------------------------------
  // Render
  // ------------------------------------------------------------------
  function renderResults(data) {
    secResults.classList.remove('d-none');
    if (btnExport) btnExport.disabled = false;
    if (periodText) periodText.textContent = `Periodo: ${data.periodo || 'Detectado'}`;

    const cuadra = data.cuadra;
    if (cuadra) {
      bannerCard.style.borderLeftColor = '#10B981';
      bannerCircle.style.background    = '#10B981';
      bannerCircle.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
      bannerTitle.textContent = 'La conciliacion cuadra';
      bannerSub.textContent   = `Todos los totales coinciden con el recibo PDF (${data.periodo}).`;
      bannerPill.textContent  = 'CUADRA';
      bannerPill.style.cssText= 'background:rgba(16,185,129,0.15);color:#065F46;padding:0.3rem 0.85rem;border-radius:999px;font-weight:700;font-size:0.78rem;';
    } else {
      bannerCard.style.borderLeftColor = '#EF4444';
      bannerCircle.style.background    = '#EF4444';
      bannerCircle.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
      bannerTitle.textContent = 'La conciliacion no cuadra';
      const causas = data.causas_no_cuadra?.join('; ') || '';
      bannerSub.textContent   = `Diferencia de ${fmt(data.diferencia)}. ${causas}`;
      bannerPill.textContent  = 'NO CUADRA';
      bannerPill.style.cssText= 'background:rgba(239,68,68,0.15);color:#991B1B;padding:0.3rem 0.85rem;border-radius:999px;font-weight:700;font-size:0.78rem;';
    }

    // KPIs
    if (kpis.lineas)   kpis.lineas.textContent   = data.line_count || 0;
    if (kpis.subLineas) kpis.subLineas.textContent = `${data.anexos_count||data.line_count||0} anexos en el PDF`;
    if (kpis.desc)     kpis.desc.textContent      = fmt(data.descuento_total);
    if (kpis.subDesc)  kpis.subDesc.textContent   = `${data.disc_count_12m||0} descuentos de 12 m en PDF`;
    if (kpis.monto)    kpis.monto.textContent     = fmt(data.cargo_mensual_sin_igv);
    if (kpis.montoF)   kpis.montoF.textContent    = fmt(data.cargo_mensual_final_sin_igv);
    if (kpis.igv)      kpis.igv.textContent       = fmt(data.igv);
    if (kpis.equipos)  kpis.equipos.textContent   = fmt(data.equipos);
    if (kpis.totalPdf) kpis.totalPdf.textContent  = fmt(data.total_pdf);
    if (kpis.subCalc)  kpis.subCalc.textContent   = `Calculado: ${fmt(data.total_calculado)}`;
    if (kpis.dif) {
      kpis.dif.textContent  = fmt(data.diferencia);
      kpis.dif.className    = `kpi-value-bold ${Math.abs(data.diferencia)<=0.01?'text-green':'text-red'}`;
    }
    if (kpis.subDif) kpis.subDif.textContent = cuadra ? 'La conciliacion cuadra' : 'Revisar diferencias';

    // Warnings CC
    if (data.cc_warnings?.length) showAlert('Advertencia: '+data.cc_warnings.join(' | '),'warning');

    renderCC(data.cc_summary||[]);
    renderSinCC(data.sin_cc_lines||[]);
  }

  function renderCC(list) {
    if (!tbodyCC) return;
    tbodyCC.innerHTML = '';
    if (!list.length) {
      tbodyCC.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--text-muted);padding:1.5rem;">Sin datos</td></tr>`;
      return;
    }
    list.forEach(item => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><strong>${item.cc}</strong></td><td>${item.count}</td><td>${fmt(item.monto)}</td><td>${item.pct}%</td>`;
      tbodyCC.appendChild(tr);
    });
  }

  function renderSinCC(lines) {
    if (!tbodySinCC) return;
    tbodySinCC.innerHTML = '';
    if (!lines.length) {
      emptyMsg?.classList.remove('d-none');
      btnDlSinCC?.classList.add('d-none');
    } else {
      emptyMsg?.classList.add('d-none');
      btnDlSinCC?.classList.remove('d-none');
      lines.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td style="font-family:monospace;font-weight:600">${item.phone}</td><td>${item.plan||'Sin plan'}</td><td>${fmt(item.monto_f)}</td><td><span style="background:rgba(239,68,68,0.12);color:#991B1B;padding:0.15rem 0.6rem;border-radius:999px;font-size:0.75rem;font-weight:600;">Sin CC</span></td>`;
        tbodySinCC.appendChild(tr);
      });
    }
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
    CARDS.forEach(c => c.rm?.click());
    if (inputPwd) inputPwd.value = '';
    state.downloadToken = null;
    if (btnExport) btnExport.disabled = true;
    secResults?.classList.add('d-none');
    if (periodText) periodText.textContent = 'Periodo: Pendiente de carga';
    clearAlert();
  });
});
