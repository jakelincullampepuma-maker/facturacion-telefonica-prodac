/** Claro: diseño original, tres entradas obligatorias y resultado sobre el Excel cargado. */
const empty = () => ({ files: { claro: null, inv: null, pdf: null }, report: null, version: 0 });
const state = { company: 'PRODAC', PRODAC: empty(), PRODIMIN: empty(), busy: false };
const byId = id => document.getElementById(id);
const put = (id, value) => { const el = byId(id); if (el) el.textContent = value ?? '—'; };
const money = value => new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' }).format(value ?? 0);
const current = () => state[state.company];
let uploadControls;

document.addEventListener('DOMContentLoaded', () => {
  for (const company of ['PRODAC', 'PRODIMIN']) byId(`tab-${company.toLowerCase()}`).addEventListener('click', () => switchCompany(company));
  uploadControls = UploadUI.create({module: 'claro', getState: () => current().files,
    required: ['claro', 'inv', 'pdf'], button: byId('btn-process-files'),
    cards: ['claro', 'inv', 'pdf'].map(key => ({stateKey: key, input: byId(`input-${key}-file`),
      drop: byId(`dropzone-${key}`), name: byId(`name-${key}-file`), check: byId(`check-icon-${key}`),
      rm: byId(`btn-remove-${key}`), emptyLabel: 'Seleccionar archivo'})),
    onChange: () => {
      current().version++; current().report = null;
      byId('process-message').hidden = true;
      resetDashboard(false); updateFiles();
    },
    onError: message => { put('process-message', message); byId('process-message').hidden = false; }
  });
  byId('btn-process-files').addEventListener('click', runConciliation);
  byId('btn-clear-files').addEventListener('click', () => {
    uploadControls.clear();
  });
  byId('btn-export-excel').addEventListener('click', downloadExcelReport);
  AppSidebar.bind({results: 'claroResultsSection', missing: 'claroMissingCcSection', download: 'btn-export-excel'});
  byId('btn-ver-todos-detalle').addEventListener('click', () => showDetails(''));
  byId('btn-close-full-view').addEventListener('click', () => byId('full-detail-section').classList.add('d-none'));
  byId('filter-search-input').addEventListener('input', filterDetails);
  switchCompany('PRODAC');
});
function switchCompany(company) {
  // Invalidar también las respuestas pendientes de la empresa que se abandona.
  current().version++;
  current().report = null;
  state.company = company;
  current().version++;
  current().report = null;
  for (const key of Object.keys(current().files)) byId(`input-${key}-file`).value = '';
  for (const value of ['PRODAC', 'PRODIMIN']) byId(`tab-${value.toLowerCase()}`).classList.toggle('active', value === company);
  put('label-active-company', company);
  put('label-full-company', company);
  byId('process-message').hidden = true;
  updateFiles();
  resetDashboard(false);
}
function updateFiles() {
  uploadControls.refresh();
  byId('btn-export-excel').disabled = !current().report;
}
function resetDashboard(clearTables = true) {
  byId('claroResultsSection').hidden = true;
  byId('claroMissingCcSection').hidden = true;
  ModuleUI.setPeriod();
  byId('btn-export-excel').disabled = true;
  if (!clearTables) return;
  put('banner-title-text', '');
  put('banner-sub-text', '');
  put('banner-pill-badge', '');
  for (const id of ['lineas', 'total-lineas', 'ajustes', 'total', 'cargos-mes', 'diferencia']) put(`kpi-val-${id}`, '');
  put('kpi-sub-lineas', 'Números únicos y registros de cargos');
  put('kpi-sub-diferencia', 'Pendiente de procesar');
  put('warnings', '');
  byId('warnings').hidden = true;
  for (const id of ['tbody-resumen-cc', 'tbody-full-detail-list', 'tbody-sin-cc']) fillTable(id, []);
  byId('filter-search-input').value = '';
  byId('conciliation-banner').classList.remove('banner-error', 'banner-warning');
  byId('full-detail-section').classList.add('d-none');
  byId('btn-export-excel').disabled = true;
}
async function runConciliation() {
  if (state.busy || !Object.values(current().files).every(Boolean)) return;
  const company = state.company, slot = current(), version = slot.version;
  const measurement = UploadUI.startRun('claro');
  let outcome = 'error';
  state.busy = true; slot.report = null; resetDashboard(false);
  uploadControls.setBusy(true); updateFiles();
  for (const name of ['prodac', 'prodimin']) byId('tab-' + name).disabled = true;
  byId('btn-clear-files').disabled = true;
  byId('process-message').hidden = true;
  put('process-button-label', 'Procesando…');
  byId('btn-process-files').setAttribute('aria-busy', 'true');
  const form = new FormData();
  form.append('company', company); form.append('fact_file', slot.files.claro);
  form.append('inventory_file', slot.files.inv); form.append('pdf_file', slot.files.pdf);
  try {
    await UploadUI.yieldToBrowser();
    resetDashboard();
    const response = await measurement.measure('request_ms', () => fetch('/api/claro/process', { method: 'POST', body: form }));
    if (!(response.headers.get('content-type') || '').includes('application/json')) throw new Error('Abra la aplicación con iniciar_servidor.bat.');
    const report = await response.json();
    if (!response.ok || !report.success) throw new Error(report.error || 'No se pudieron procesar los archivos.');
    if (slot.version !== version) return;
    measurement.merge(report.timings);
    slot.report = report;
    if (state.company === company) displayResults(report);
    outcome = 'success';
  } catch (error) {
    if (state.company === company && slot.version === version) {
      slot.report = null;
      resetDashboard();
      put('process-message', error.message);
      byId('process-message').hidden = false;
    }
  } finally {
    state.busy = false;
    uploadControls.setBusy(false);
    for (const name of ['prodac', 'prodimin']) byId('tab-' + name).disabled = false;
    byId('btn-clear-files').disabled = false;
    measurement.finish(outcome);
    put('process-button-label', 'Procesar archivos');
    byId('btn-process-files').removeAttribute('aria-busy');
    updateFiles();
  }
}
function fillTable(id, rows) {
  ModuleUI.rows(id, rows);
}
function displayResults(r) {
  ModuleUI.setPeriod(r.periodo);
  ModuleUI.status({cuadra:r.cuadra, missing:r.lineas_sin_cc, observations:r.warnings.length,
    difference:money(r.diferencia), detail:r.cuadra ? 'El total final coincide con Cargos del Mes del PDF.' : ''});
  put('kpi-val-lineas', r.numeros_unicos);
  put('kpi-sub-lineas', `${r.registros_cargos} registros · ${r.lineas_con_cc} con CC · ${r.lineas_sin_cc} sin CC`);
  for (const [id, key] of [['total-lineas','suma_lineas'],['ajustes','ajustes_abonos'],['total','total_calculado'],['cargos-mes','cargos_del_mes'],['diferencia','diferencia']]) put(`kpi-val-${id}`, money(r[key]));
  put('kpi-sub-diferencia', r.cuadra ? 'La conciliación cuadra' : 'Existe diferencia');
  ModuleUI.difference('kpi-val-diferencia', r.diferencia);
  put('warnings', r.warnings.join(' · '));
  byId('warnings').hidden = !r.warnings.length;
  fillTable('tbody-resumen-cc', r.summary.map(x => [x.cc, x.lineas, money(x.total), r.total_calculado ? `${(x.total/r.total_calculado*100).toFixed(1)}%` : '0%']));
  fillTable('tbody-full-detail-list', r.details.map((x,i) => [i+1,x.numero,x.usuario,x.cc,x.plan,money(x.costo),x.observacion || 'OK']));
  ModuleUI.review({body:'tbody-sin-cc', values:r.details.filter(x => x.cc === 'SIN CC - REVISAR' || x.observacion).map(x => [x.numero, x.usuario, x.cc, money(x.costo), x.observacion || 'SIN CC - REVISAR']), notes:r.warnings});
  byId('claroMissingCcSection').hidden = false;
  byId('filter-search-input').value = ''; filterDetails();
  updateFiles();
  byId('claroResultsSection').hidden = false;
}
function showDetails(term) { if (!current().report) return; byId('full-detail-section').classList.remove('d-none'); byId('filter-search-input').value = term; filterDetails(); byId('full-detail-section').scrollIntoView({ behavior:'smooth' }); }
function filterDetails() { const term = byId('filter-search-input').value.toLowerCase(); for (const row of byId('tbody-full-detail-list').rows) row.hidden = !row.textContent.toLowerCase().includes(term); }
function downloadExcelReport() {
  const report = current().report; if (!report) return;
  const a = document.createElement('a'); a.href = `/api/claro/download-excel?token=${encodeURIComponent(report.download_token)}`;
  a.download = report.download_filename; document.body.append(a); a.click(); a.remove();
}
