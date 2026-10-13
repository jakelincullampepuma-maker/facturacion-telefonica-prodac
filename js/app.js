/**
 * app.js - App Celulares PRODAC
 * Implementación exacta del algoritmo de conciliador-claro-mensual.prodac-5408.chatgpt.site
 * con soporte para pestañas independientes PRODAC y PRODIMIN.
 */

// Helpers exactos del conciliador
const normalize = v => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const phone = v => String(v ?? '').replace(/\D/g, '').slice(-9);
const money = v => new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' }).format(Number(v) || 0);

// Estado por empresa
const AppState = {
  activeCompany: 'PRODAC',
  PRODAC: {
    inputs: { claro: null, inventario: null, pdf: null },
    report: null
  },
  PRODIMIN: {
    inputs: { claro: null, inventario: null, pdf: null },
    report: null
  }
};

document.addEventListener('DOMContentLoaded', () => {
  initCompanyTabs();
  initSidebar();
  initDropzones();
  initButtons();
});

function initCompanyTabs() {
  const tabProdac = document.getElementById('tab-prodac');
  const tabProdimin = document.getElementById('tab-prodimin');

  tabProdac?.addEventListener('click', () => switchCompany('PRODAC'));
  tabProdimin?.addEventListener('click', () => switchCompany('PRODIMIN'));
}

function switchCompany(company) {
  AppState.activeCompany = company;

  document.getElementById('tab-prodac')?.classList.toggle('active', company === 'PRODAC');
  document.getElementById('tab-prodimin')?.classList.toggle('active', company === 'PRODIMIN');
  document.getElementById('label-active-company').textContent = company;

  updateFilesVisualState();

  const cur = AppState[company];
  if (cur.report) {
    displayResults(cur.report);
  } else {
    hideResults();
  }
}

function initSidebar() {
  document.getElementById('nav-dashboard')?.addEventListener('click', () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  document.getElementById('nav-upload')?.addEventListener('click', () => {
    document.getElementById('section-carga-archivos')?.scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('nav-results')?.addEventListener('click', () => {
    document.getElementById('results-section-container')?.scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('nav-sin-cc')?.addEventListener('click', () => {
    const fullSec = document.getElementById('full-detail-section');
    if (fullSec) {
      fullSec.classList.remove('d-none');
      fullSec.scrollIntoView({ behavior: 'smooth' });
    }
  });

  document.getElementById('nav-export')?.addEventListener('click', () => {
    downloadExcelReport();
  });
}

/**
 * Dropzones y carga de los 3 archivos
 */
function initDropzones() {
  setupFileBox('claro', 'dropzone-claro', 'input-claro-file', 'name-claro-file', 'check-icon-claro', 'card-wrap-claro', 'btn-remove-claro');
  setupFileBox('inventario', 'dropzone-inv', 'input-inv-file', 'name-inv-file', 'check-icon-inv', 'card-wrap-inv', 'btn-remove-inv');
  setupFileBox('pdf', 'dropzone-pdf', 'input-pdf-file', 'name-pdf-file', 'check-icon-pdf', 'card-wrap-pdf', 'btn-remove-pdf');
}

function setupFileBox(key, dropId, inputId, nameId, checkId, cardId, removeId) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const removeBtn = document.getElementById(removeId);

  drop?.addEventListener('click', (e) => {
    if (e.target.closest('.btn-remove-card-file')) return;
    input.click();
  });

  input?.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      setCompanyFile(key, e.target.files[0]);
    }
  });

  drop?.addEventListener('dragover', (e) => {
    e.preventDefault();
    drop.classList.add('drag-over');
  });

  drop?.addEventListener('dragleave', () => {
    drop.classList.remove('drag-over');
  });

  drop?.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setCompanyFile(key, e.dataTransfer.files[0]);
    }
  });

  removeBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    input.value = '';
    removeCompanyFile(key);
  });
}

function setCompanyFile(key, file) {
  const active = AppState.activeCompany;
  AppState[active].inputs[key] = file;
  updateFilesVisualState();
}

function removeCompanyFile(key) {
  const active = AppState.activeCompany;
  AppState[active].inputs[key] = null;
  updateFilesVisualState();
}

function updateFilesVisualState() {
  const active = AppState.activeCompany;
  const currentInputs = AppState[active].inputs;

  // 01 Claro
  const fClaro = currentInputs.claro;
  document.getElementById('name-claro-file').textContent = fClaro ? fClaro.name : 'Selecciona el archivo con “Factura Detallada”';
  document.getElementById('check-icon-claro').style.display = fClaro ? 'block' : 'none';
  document.getElementById('card-wrap-claro').classList.toggle('empty-border', !fClaro);

  // 02 Inventario
  const fInv = currentInputs.inventario;
  document.getElementById('name-inv-file').textContent = fInv ? fInv.name : 'Selecciona el archivo con Número y CC Dpto.';
  document.getElementById('check-icon-inv').style.display = fInv ? 'block' : 'none';
  document.getElementById('card-wrap-inv').classList.toggle('empty-border', !fInv);

  // 03 PDF
  const fPdf = currentInputs.pdf;
  document.getElementById('name-pdf-file').textContent = fPdf ? fPdf.name : 'Selecciona el recibo mensual de Claro';
  document.getElementById('check-icon-pdf').style.display = fPdf ? 'block' : 'none';
  document.getElementById('card-wrap-pdf').classList.toggle('empty-border', !fPdf);

  const canRun = !!(fClaro && fInv && fPdf);
  document.getElementById('btn-process-files').disabled = !canRun;
}

/**
 * Acciones: Procesar, Limpiar, Descargar
 */
function initButtons() {
  document.getElementById('btn-process-files')?.addEventListener('click', runConciliation);

  document.getElementById('btn-clear-files')?.addEventListener('click', () => {
    const active = AppState.activeCompany;
    AppState[active].inputs = { claro: null, inventario: null, pdf: null };
    AppState[active].report = null;
    document.getElementById('input-claro-file').value = '';
    document.getElementById('input-inv-file').value = '';
    document.getElementById('input-pdf-file').value = '';
    updateFilesVisualState();
    hideResults();
  });

  document.getElementById('btn-export-excel')?.addEventListener('click', downloadExcelReport);

  // Ver todos
  document.getElementById('btn-ver-todos-detalle')?.addEventListener('click', () => {
    const fullSec = document.getElementById('full-detail-section');
    fullSec?.classList.remove('d-none');
    fullSec?.scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('btn-close-full-view')?.addEventListener('click', () => {
    document.getElementById('full-detail-section')?.classList.add('d-none');
  });

  document.getElementById('filter-search-input')?.addEventListener('input', (e) => {
    const term = e.target.value.toLowerCase().trim();
    const rows = document.querySelectorAll('#tbody-full-detail-list tr');
    rows.forEach(r => {
      r.style.display = r.textContent.toLowerCase().includes(term) ? '' : 'none';
    });
  });
}

// ==========================================
// ALGORITMO EXACTO DE CONCILIACIÓN
// ==========================================
async function readWorkbook(file) {
  return XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
}

function rowsOf(wb, sheet) {
  return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null, raw: true });
}

function findHeader(rows, required) {
  return rows.findIndex(r => required.every(term => r.some(c => normalize(c).includes(term))));
}

function parseClaro(wb) {
  const sheet = wb.SheetNames.find(n => normalize(n).includes('factura detallada'));
  if (!sheet) throw new Error('No encontré la hoja “Factura Detallada” en el Excel de Claro.');
  const rows = rowsOf(wb, sheet);
  const hi = findHeader(rows, ['linea', 'sin igv']);
  if (hi < 0) throw new Error('No encontré las columnas Línea y Total Cargos por Línea (Sin IGV).');

  const h = rows[hi].map(normalize);
  const ni = h.findIndex(x => x === 'linea' || x.includes('linea'));
  const ci = h.findIndex(x => x.includes('total cargos por linea') && x.includes('sin igv'));

  const data = [];
  for (const r of rows.slice(hi + 1)) {
    const n = phone(r[ni]);
    const c = Number(r[ci]);
    if (n.length === 9 && Number.isFinite(c)) {
      data.push({ numero: n, costo: c });
    }
  }

  if (!data.length) throw new Error('El Excel de Claro no contiene líneas válidas para procesar.');
  return data;
}

function parseInventory(wb, company) {
  // Busca hoja "prodac" o "prodimin", o usa la primera hoja
  const target = normalize(company);
  const sheet = wb.SheetNames.find(n => normalize(n) === target) || wb.SheetNames[0];
  const rows = rowsOf(wb, sheet);
  const hi = findHeader(rows, ['numero', 'cc dpto']);
  if (hi < 0) throw new Error(`No encontré las columnas Número y CC Dpto. en la hoja "${sheet}" del inventario.`);

  const h = rows[hi].map(normalize);
  const ni = h.findIndex(x => x === 'numero');
  const ci = h.findIndex(x => x.includes('cc dpto'));
  const ui = h.findIndex(x => x.includes('usuario') || x.includes('responsable') || x.includes('nombre'));

  const map = new Map();
  const userMap = new Map();
  const duplicates = new Set();

  for (const r of rows.slice(hi + 1)) {
    const n = phone(r[ni]);
    if (n.length !== 9) continue;
    if (map.has(n)) {
      duplicates.add(n);
    } else {
      map.set(n, String(r[ci] ?? '').trim());
      if (ui >= 0) userMap.set(n, String(r[ui] ?? '').trim());
    }
  }

  return { map, userMap, duplicates, sheetUsed: sheet };
}

async function parsePdf(file) {
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('PDF.js no está disponible.');
  }

  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  let text = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const content = await (await pdf.getPage(p)).getTextContent();
    text += ' ' + content.items.map(i => i.str).join(' ');
    if (/Cargos\s+del\s+Mes/i.test(text)) break;
  }

  const m = text.match(/Cargos\s+del\s+Mes\s*(?:S\/?\s*)?([\d,.]+)/i);
  if (!m) throw new Error('No pude identificar “Cargos del Mes” en el PDF.');
  let s = m[1].replace(/,/g, '');
  const value = Number(s);
  if (!Number.isFinite(value)) throw new Error('El monto “Cargos del Mes” del PDF no es válido.');
  return value;
}

async function runConciliation() {
  const active = AppState.activeCompany;
  const currentInputs = AppState[active].inputs;

  if (!currentInputs.claro || !currentInputs.inventario || !currentInputs.pdf) {
    alert('Por favor cargue los 3 archivos obligatorios (Excel Claro, Inventario General y Recibo PDF).');
    return;
  }

  const runBtn = document.getElementById('btn-process-files');
  runBtn.disabled = true;
  runBtn.innerHTML = `<span>Procesando archivos...</span>`;

  try {
    const [claroWb, invWb, pdfTotal] = await Promise.all([
      readWorkbook(currentInputs.claro),
      readWorkbook(currentInputs.inventario),
      parsePdf(currentInputs.pdf)
    ]);

    const data = parseClaro(claroWb);
    const inv = parseInventory(invWb, active);

    const summary = new Map();
    let missing = 0;

    const details = data.map(x => {
      const cc = inv.map.get(x.numero) || 'NO ENCONTRADO';
      const usuario = inv.userMap.get(x.numero) || 'No asignado';
      if (cc === 'NO ENCONTRADO' || !cc) missing++;
      summary.set(cc, (summary.get(cc) || 0) + x.costo);
      return {
        ...x,
        cc,
        usuario,
        duplicado: inv.duplicates.has(x.numero),
        observacion: cc === 'NO ENCONTRADO' || !cc ? 'Sin CC' : inv.duplicates.has(x.numero) ? 'Duplicado' : 'OK'
      };
    });

    const total = data.reduce((a, b) => a + b.costo, 0);
    const diff = Number((total - pdfTotal).toFixed(2));
    const ok = Math.abs(diff) < 0.01 && missing === 0;

    const report = {
      company: active,
      details,
      summary,
      total,
      pdfTotal,
      diff,
      missing,
      ok,
      inv
    };

    AppState[active].report = report;
    displayResults(report);

    alert(`¡Conciliación de ${active} procesada con éxito! Estado: ${ok ? 'CUADRA' : 'REVISAR'}`);
  } catch (err) {
    console.error(err);
    alert('Error al procesar: ' + (err.message || err));
  } finally {
    runBtn.disabled = false;
    runBtn.innerHTML = `<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"/></svg><span>Procesar archivos</span>`;
  }
}

function displayResults(report) {
  const { details, summary, total, pdfTotal, diff, missing, ok, company } = report;

  // Banner
  const banner = document.getElementById('conciliation-banner');
  const bannerCircle = document.getElementById('banner-status-circle');
  const bannerTitle = document.getElementById('banner-title-text');
  const bannerSub = document.getElementById('banner-sub-text');
  const bannerPill = document.getElementById('banner-pill-badge');

  bannerTitle.textContent = ok ? 'La conciliación cuadra' : 'Hay observaciones por revisar';
  bannerSub.textContent = ok 
    ? 'Los totales del archivo de Claro coinciden con el inventario.'
    : `${missing} línea(s) sin CC o una diferencia de ${money(diff)} pendiente.`;
  bannerPill.textContent = ok ? 'CUADRA' : 'REVISAR';

  if (ok) {
    banner.classList.remove('banner-error');
    bannerCircle.classList.remove('circle-alert');
    bannerPill.className = 'banner-pill-cuadra';
  } else {
    banner.classList.add('banner-error');
    bannerCircle.classList.add('circle-alert');
    bannerPill.className = 'banner-pill-cuadra pill-revisar';
  }

  // KPIs
  document.getElementById('kpi-val-lineas').textContent = details.length;
  document.getElementById('kpi-sub-lineas').textContent = `${details.length - missing} con CC · ${missing} por revisar`;

  document.getElementById('kpi-val-total-lineas').textContent = money(total);
  document.getElementById('kpi-val-cargos-mes').textContent = money(pdfTotal);

  const diffEl = document.getElementById('kpi-val-diferencia');
  diffEl.textContent = money(diff);
  diffEl.className = ok ? 'kpi-value-bold text-green' : 'kpi-value-bold text-red';
  document.getElementById('kpi-sub-diferencia').textContent = ok ? 'La conciliación cuadra' : 'Existe diferencia';

  // Resumen por CC Dpto
  const tbodyRes = document.getElementById('tbody-resumen-cc');
  if (tbodyRes) {
    tbodyRes.innerHTML = [...summary]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([cc, v]) => {
        const pct = total > 0 ? ((v / total) * 100).toFixed(1) : 0;
        const isBad = cc === 'NO ENCONTRADO' || !cc;
        return `
          <tr>
            <td><strong class="${isBad ? 'badge-status-sin' : ''}">${cc || 'SIN CC'}</strong></td>
            <td>${isBad ? 'Línea no inventariada' : 'Área / Centro de Costo'}</td>
            <td>-</td>
            <td style="font-family: ui-monospace, monospace; font-weight:600;">${money(v)}</td>
            <td>${pct}%</td>
          </tr>
        `;
      }).join('');
  }

  // Detalle y observaciones (Primeras 8 filas)
  const tbodyDet = document.getElementById('tbody-detalle-preview');
  if (tbodyDet) {
    tbodyDet.innerHTML = details.slice(0, 8).map(x => {
      const isOk = x.observacion === 'OK';
      return `
        <tr>
          <td style="font-family: ui-monospace, monospace; font-weight:600;">${x.numero}${x.duplicado ? ' (dup)' : ''}</td>
          <td>${x.usuario}</td>
          <td class="${isOk ? '' : 'badge-status-sin'}">${x.cc || 'SIN CC'}</td>
          <td style="font-family: ui-monospace, monospace;">${money(x.costo)}</td>
          <td><span class="${isOk ? 'badge-status-ok' : 'badge-status-sin'}">${x.observacion}</span></td>
        </tr>
      `;
    }).join('');
  }

  // Detalle completo
  const tbodyFull = document.getElementById('tbody-full-detail-list');
  if (tbodyFull) {
    tbodyFull.innerHTML = details.map((x, idx) => {
      const isOk = x.observacion === 'OK';
      return `
        <tr>
          <td>${idx + 1}</td>
          <td style="font-family: ui-monospace, monospace; font-weight:600;">${x.numero}</td>
          <td>${x.usuario}</td>
          <td class="${isOk ? '' : 'badge-status-sin'}"><strong>${x.cc || 'SIN CC'}</strong></td>
          <td>${isOk ? 'Asignado' : 'No inventariado'}</td>
          <td style="font-family: ui-monospace, monospace; font-weight:700;">${money(x.costo)}</td>
          <td><span class="${isOk ? 'badge-status-ok' : 'badge-status-sin'}">${x.observacion}</span></td>
        </tr>
      `;
    }).join('');
  }

  // Habilitar botón de descarga
  document.getElementById('btn-export-excel').disabled = false;
  document.getElementById('label-full-company').textContent = company;
}

function hideResults() {
  document.getElementById('btn-export-excel').disabled = true;
}

/**
 * Exportar Excel oficial según el conciliador
 */
function downloadExcelReport() {
  const active = AppState.activeCompany;
  const report = AppState[active].report;

  if (!report) {
    alert('Procese los archivos antes de descargar el Excel.');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Hoja 1: Detalle
  const det = [
    ['NUMERO', 'COSTO', 'CC', 'OBSERVACIÓN'],
    ...report.details.map(x => [
      x.numero,
      x.costo,
      x.cc,
      x.cc === 'NO ENCONTRADO' 
        ? 'Número no ubicado en inventario' 
        : x.duplicado 
          ? 'Número duplicado en inventario' 
          : 'OK'
    ])
  ];

  // Hoja 2: Resumen CC
  const res = [
    ['CC DPTO', 'TOTAL'],
    ...[...report.summary].sort((a, b) => a[0].localeCompare(b[0]))
  ];

  // Hoja 3: Conciliación
  const con = [
    ['CONTROL', 'VALOR'],
    ['Empresa', report.company],
    ['Líneas procesadas', report.details.length],
    ['Líneas sin CC', report.missing],
    ['Total por líneas sin IGV', report.total],
    ['Cargos del Mes PDF', report.pdfTotal],
    ['Diferencia', report.diff],
    ['Estado', report.ok ? 'CUADRA' : 'REVISAR']
  ];

  const s1 = XLSX.utils.aoa_to_sheet(det);
  const s2 = XLSX.utils.aoa_to_sheet(res);
  const s3 = XLSX.utils.aoa_to_sheet(con);

  s1['!cols'] = [{ wch: 14 }, { wch: 13 }, { wch: 22 }, { wch: 34 }];
  s2['!cols'] = [{ wch: 24 }, { wch: 16 }];
  s3['!cols'] = [{ wch: 28 }, { wch: 20 }];

  XLSX.utils.book_append_sheet(wb, s1, 'Detalle');
  XLSX.utils.book_append_sheet(wb, s2, 'Resumen CC');
  XLSX.utils.book_append_sheet(wb, s3, 'Conciliación');

  const fileName = `Conciliacion_Claro_${report.company}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
