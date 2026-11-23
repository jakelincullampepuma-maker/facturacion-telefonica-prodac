/* Procesamiento local Entel. Los documentos permanecen en el navegador. */
(function () {
  'use strict';
  const workerUrl = typeof document !== 'undefined' ? new URL('entelExcelWorker.js', document.currentScript.src).href : '';
  const libraries = {
    XLSX: 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js',
    pdfjsLib: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    JSZip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
  };
  const pendingLibraries = new Map();
  async function ensureLibraries() {
    await Promise.all(Object.entries(libraries).map(([name, src]) => {
      if (globalThis[name]) return;
      if (!pendingLibraries.has(name)) pendingLibraries.set(name, new Promise((resolve, reject) => {
        const script = document.createElement('script'); script.src = src; script.async = true;
        script.onload = resolve;
        script.onerror = () => { pendingLibraries.delete(name); script.remove(); reject(new Error('No se cargó ' + name + '. Revise su conexión e intente procesar de nuevo.')); };
        document.head.appendChild(script);
      }));
      return pendingLibraries.get(name);
    }));
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
  async function readWorkbooks(excelFile, inventoryFile, measurement) {
    const [original, inventoryBytes] = await measurement.measure('excel_read_ms', () => Promise.all([excelFile.arrayBuffer(), inventoryFile.arrayBuffer()]));
    let worker;
    // file:// no permite Workers de archivos locales en Chrome; el resto usa Worker.
    if (typeof Worker !== 'undefined' && location.protocol !== 'file:') {
      try {
        worker = new Worker(workerUrl);
        await measurement.measure('worker_init_ms', () => new Promise((resolve, reject) => {
          const timeout = setTimeout(() => { worker.terminate(); reject(new Error('Worker no disponible.')); }, 12000);
          worker.onmessage = event => { if (event.data.ready) { clearTimeout(timeout); resolve(); } };
          worker.onerror = event => { event.preventDefault(); clearTimeout(timeout); reject(new Error('Worker no disponible.')); };
        }));
      } catch { worker?.terminate(); worker = null; }
    }
    if (worker) {
      try {
        const parsed = await new Promise((resolve, reject) => {
          worker.onmessage = event => event.data.error ? reject(new Error(event.data.error)) : resolve(event.data);
          worker.onerror = event => reject(new Error(event.message || 'No se pudo leer el Excel en segundo plano.'));
          worker.postMessage({original, inventoryBytes}, [original, inventoryBytes]);
        });
        for (const [key, value] of Object.entries(parsed.timings)) measurement.run.timings[key] += value;
        parsed.worker = true;
        return parsed;
      } finally { worker.terminate(); }
    }
    await UploadUI.yieldToBrowser();
    let wb, invWb;
    wb = await measurement.measure('excel_read_ms', async () => {
      try { return XLSX.read(original, {type: 'array'}); } catch (e) { throw new Error('No se pudo leer el Excel de Entel: ' + e.message); }
    });
    await UploadUI.yieldToBrowser();
    invWb = await measurement.measure('excel_read_ms', async () => {
      try { return XLSX.read(inventoryBytes, {type: 'array'}); } catch (e) { throw new Error('No se pudo leer el inventario: ' + e.message); }
    });
    await UploadUI.yieldToBrowser();
    const detail = billing(wb), prev = previous(wb);
    const inv = await measurement.measure('inventory_match_ms', async () => inventory(invWb));
    return {original, detail, inv, prev, worker: false};
  }
  const norm = v => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[_\s]+/g, ' ').trim();
  const str = v => String(v ?? '').trim();
  function phone(v) {
    const s = str(v).replace(/\.0+$/, '').replace(/\D/g, '');
    return /^(?:51)?9\d{8}$/.test(s) ? s.slice(-9) : '';
  }
  // Sumas y comparaciones en céntimos, con redondeo comercial por línea.
  function cents(v, context = 'importe') {
    if (v === null || v === undefined || str(v) === '') throw new Error('Falta ' + context + '.');
    let s = str(v).replace(/S\/?\.?|\$|\s/gi, '');
    if (s.includes(',') && s.includes('.')) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    else if (s.includes(',')) s = s.replace(',', '.');
    if (!/^[+-]?\d+(?:\.\d+)?$/.test(s) || !Number.isFinite(Number(s))) throw new Error(context + ' no válido: ' + str(v));
    return Math.sign(Number(s)) * Math.round((Math.abs(Number(s)) + Number.EPSILON) * 100);
  }
  const rowsOf = (wb, name) => XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1, defval: null, raw: true, blankrows: true});
  const isPhoneHeader = h => /^(numero(?: entel| telefonico| celular)?|linea|telefono|celular|movil|n[°ºo.]? entel)$/.test(h);
  function findHeader(rows, kind) {
    // Los encabezados del proveedor pueden ocupar varias filas consecutivas.
    for (let r = 0; r < rows.length; r++) {
      const h = rows[r].map(norm);
      const m = {phone: h.findIndex(isPhoneHeader)};
      if (m.phone < 0) continue;
      if (kind === 'billing') {
        m.plan = h.findIndex(x => /plan|tarifa/.test(x));
        m.amount = -1;
        let end = r;
        for (let j = r; j < Math.min(r + 6, rows.length); j++) {
          if (j > r && rows[j].some(v => phone(v))) break;
          const idx = rows[j].map(norm).findIndex(x => /cargos? fijo|renta mensual|cargo_fijo|monto f/.test(x));
          if (idx >= 0) { m.amount = idx; end = j; break; }
        }
        if (m.amount >= 0) return {...m, row: end};
      } else {
        m.cc = h.findIndex(x => /cc.*(?:dpto|depto|departamento)|centro.*costo/.test(x) || (kind === 'previous' && x === 'cc'));
        m.operator = h.findIndex(x => /operador/.test(x));
        m.user = h.findIndex(x => /usuario|responsable|nombre/.test(x));
        m.area = h.findIndex(x => x === 'area');
        m.department = h.findIndex(x => /^(departamento|dpto|depto)$/.test(x));
        m.ccArea = h.findIndex(x => /^cc.*area/.test(x));
        if (kind === 'previous' ? m.cc >= 0 : m.operator >= 0 && m.cc >= 0) return {...m, row: r};
      }
    }
    return null;
  }
  function billing(wb) {
    const candidates = [];
    for (const name of wb.SheetNames) {
      if (/^(resultado entel(?: anterior.*)?|hoja1|resumen|ceco)$/.test(norm(name))) continue;
      const rows = rowsOf(wb, name), h = findHeader(rows, 'billing');
      if (!h) continue;
      const lines = rows.slice(h.row + 1).map((r, i) => ({r, row: h.row + i + 2})).filter(x => phone(x.r[h.phone]));
      if (lines.length) candidates.push({name, h, lines});
    }
    candidates.sort((a, b) => b.lines.length - a.lines.length);
    if (!candidates.length) throw new Error('No se encontró detalle de Entel con columnas Número y Cargo Fijo. Revise los encabezados del Excel.');
    return candidates[0];
  }
  function inventory(wb) {
    const entel = new Map(), others = new Map(), duplicates = new Set();
    for (const name of wb.SheetNames) {
      const rows = rowsOf(wb, name), h = findHeader(rows, 'inventory');
      if (!h) continue;
      for (const r of rows.slice(h.row + 1)) {
        const n = phone(r[h.phone]);
        if (!n) continue;
        const op = str(r[h.operator]);
        if (!norm(op).includes('entel')) { if (op) others.set(n, op); continue; }
        if (entel.has(n)) { duplicates.add(n); continue; }
        entel.set(n, {empresa: name, usuario: str(r[h.user]), area: str(r[h.area]), departamento: str(r[h.department]), cc_dpto: str(r[h.cc]), cc_area: str(r[h.ccArea])});
      }
    }
    if (!entel.size) throw new Error('No se encontraron líneas Entel en el inventario con encabezados Número, Operador y CC Dpto.');
    return {entel, others, duplicates};
  }
  function previous(wb) {
    const map = new Map(), name = wb.SheetNames.find(s => norm(s) === 'hoja1');
    if (!name) return map;
    const rows = rowsOf(wb, name), h = findHeader(rows, 'previous');
    if (!h) throw new Error('Hoja1 existe pero no se reconocen sus columnas Número y CC para validar discrepancias.');
    for (const r of rows.slice(h.row + 1)) if (phone(r[h.phone])) map.set(phone(r[h.phone]), str(r[h.cc]));
    return map;
  }
  async function pdfInfo(file) {
    const task = pdfjsLib.getDocument({data: await file.arrayBuffer()});
    let doc;
    try {
      doc = await task.promise;
      const texts = [], pageRows = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const content = await page.getTextContent();
        texts.push(content.items.map(i => i.str).join('\n'));
        const lines = [];
        for (const i of content.items.filter(i => i.str.trim())) {
          const y = i.transform[5];
          let line = lines.find(l => Math.abs(l.y - y) < 3);
          if (!line) { line = {y, items: []}; lines.push(line); }
          line.items.push(i);
        }
        pageRows.push(lines.sort((a, b) => b.y - a.y).map(l => l.items.sort((a, b) => a.transform[4] - b.transform[4]).map(i => i.str).join(' ')));
        page.cleanup();
        await UploadUI.yieldToBrowser();
      }
      const first = texts[0], all = texts.join('\n');
      let plan = null, equipment = null;
      // Tabla oficial: Plan, Prorrateo, Cargos, Descuentos, Equipos, Saldo.
      const six = first.match(/S\/\s*([\d,]+\.\d{2})\s+S\/\s*([\d,]+\.\d{2})\s+S\/\s*([\d,]+\.\d{2})\s+S\/\s*([\d,]+\.\d{2})\s+S\/\s*([\d,]+\.\d{2})\s+S\/\s*([\d,]+\.\d{2})/);
      if (six && /\bplan\b/i.test(first)) { plan = cents(six[1]); equipment = cents(six[5]); }
      if (plan === null) {
        for (const rows of pageRows) {
          const hi = rows.findIndex(r => /\bplan\b/i.test(r) && /equipos/i.test(r) && /prorrateo/i.test(r));
          if (hi < 0) continue;
          for (const row of rows.slice(hi + 1, hi + 8)) {
            const amounts = [...row.matchAll(/(?:S\/\s*)?(-?[\d,]+\.\d{2})/g)];
            if (amounts.length >= 6) { plan = cents(amounts[0][1]); equipment = cents(amounts[4][1]); break; }
          }
          if (plan !== null) break;
        }
      }
      if (plan === null) {
        const total = all.match(/Total\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})/);
        if (total && /\bplan\b/i.test(all)) { plan = cents(total[1]); equipment = cents(total[5]); }
      }
      // Mantener el formato tributario admitido por el motor Python.
      if (plan === null) {
        const tax = first.match(/Total\s*\(Incl\.\s*IGV\)[\s\S]*?Montos ya gravados[\s\S]*?(\d[\d,]*\.\d{2})[\s\S]*?(\d[\d,]*\.\d{2})/);
        if (tax) { plan = cents(tax[1]); equipment = cents(tax[2]); }
      }
      if (plan === null) throw new Error('No se pudo identificar el concepto PLAN del PDF de Entel. Use el recibo original con texto seleccionable.');
      const payable = first.match(/Total\s*a\s*pagar[\s\S]*?S\/\s*([\d,]+\.\d{2})/i);
      const dates = first.match(/\b\d{1,2}\/[A-Za-z]{3}\/\d{4}\b/g) || [];
      return {plan, equipos: equipment ?? 0, total: payable ? cents(payable[1]) : null,
        periodo: dates.length >= 3 ? dates[1] + ' - ' + dates[2] : dates.length ? 'Emisión: ' + dates[0] : 'Periodo no especificado',
        recibo_nro: first.match(/\b[A-Z0-9]{3,5}-\d{6,}\b/)?.[0] || 'No detectado'};
    } finally { if (doc) await doc.destroy(); else await task.destroy(); }
  }
  async function process(excelFile, inventoryFile, pdfFile, measurement = UploadUI.startRun('entel')) {
    await measurement.measure('libraries_ms', ensureLibraries);
    const {original, detail, inv, prev, worker} = await readWorkbooks(excelFile, inventoryFile, measurement);
    const pdf = await measurement.measure('pdf_read_ms', () => pdfInfo(pdfFile));
    const matchStart = performance.now();
    const counts = new Map();
    for (const {r} of detail.lines) { const n = phone(r[detail.h.phone]); counts.set(n, (counts.get(n) || 0) + 1); }
    const discrepancies = [];
    let sum = 0, sinCC = 0, missing = 0, matched = 0;
    const lines = [];
    for (const {r, row} of detail.lines) {
      const n = phone(r[detail.h.phone]), item = inv.entel.get(n), notes = [];
      const amount = cents(r[detail.h.amount], 'Cargo Fijo de ' + n + ', fila ' + row);
      sum += amount;
      let estado = 'Conciliado';
      if (!item) { missing++; estado = 'No encontrado'; notes.push(inv.others.has(n) ? 'Registrado en inventario bajo otro operador (' + inv.others.get(n) + ')' : 'No figura en el Inventario de Celulares'); }
      else if (!item.cc_dpto) { sinCC++; estado = 'Sin CC'; notes.push('Sin Centro de Costo en inventario'); }
      else matched++;
      if (item?.cc_dpto && prev.get(n) && item.cc_dpto !== prev.get(n)) {
        notes.push('CC difiere de Hoja1 previa (Inv: ' + item.cc_dpto + ' vs Hoja1: ' + prev.get(n) + ')');
        discrepancies.push({phone: n, inv_cc: item.cc_dpto, hoja1_cc: prev.get(n), usuario: item.usuario});
        if (estado === 'Conciliado') estado = 'Conciliado (CC cambiado)';
      }
      if (inv.duplicates.has(n)) { notes.push('Número duplicado en inventario: se utilizó la primera coincidencia; revisar asignación'); estado = 'Duplicado en inventario'; }
      if (counts.get(n) > 1) { notes.push('Número duplicado en la factura'); estado = 'Duplicado'; }
      lines.push({phone: n, plan: str(r[detail.h.plan]), cargo_fijo: amount / 100, empresa: item?.empresa || '', usuario: item?.usuario || '', area: item?.area || '', departamento: item?.departamento || '', cc_dpto: item?.cc_dpto || '', prs_code: item?.cc_area && item.cc_area !== item.cc_dpto ? item.cc_area : '', estado, observacion: notes.join('; ')});
      if (lines.length % 200 === 0) await UploadUI.yieldToBrowser();
    }
    measurement.run.timings.inventory_match_ms += performance.now() - matchStart;
    const reconcileStart = performance.now();
    const groups = new Map();
    for (const l of lines) {
      const cc = l.cc_dpto || 'Sin Centro de Costo';
      if (!groups.has(cc)) groups.set(cc, {cc, count: 0, cents: 0});
      const g = groups.get(cc); g.count++; g.cents += cents(l.cargo_fijo);
    }
    const summary = [...groups.values()].map(g => ({cc: g.cc, count: g.count, monto: g.cents / 100, pct: sum ? Math.round(g.cents / sum * 1000) / 10 : 0})).sort((a, b) => (a.cc === 'Sin Centro de Costo') - (b.cc === 'Sin Centro de Costo') || b.monto - a.monto);
    const diff = sum - pdf.plan, ok = Math.abs(diff) <= 1, obs = lines.filter(l => l.observacion || l.estado !== 'Conciliado');
    const dupCount = [...counts.values()].reduce((a, n) => a + Math.max(n - 1, 0), 0);
    const result = {success: true, periodo: pdf.periodo, recibo_nro: pdf.recibo_nro, sheet_used: detail.name, line_count: lines.length, conciliado_count: matched, sin_cc_count: sinCC, no_inv_count: missing, dup_count: dupCount, total_excel: sum / 100, total_pdf_plan: pdf.plan / 100, total_pdf_equipos: pdf.equipos / 100, total_pdf_factura: pdf.total === null ? null : pdf.total / 100, diferencia: diff / 100, cuadra: ok, es_valido_todo: ok && !sinCC && !missing && !dupCount && !lines.some(l => inv.duplicates.has(l.phone)), estado_conciliacion: ok ? 'CONCILIADO' : 'REVISAR', mensaje_conciliacion: ok ? 'CONCILIADO — Los importes de PLAN coinciden.' : 'REVISAR — El total del Excel difiere del PLAN del PDF.', cc_discrepancies: discrepancies, lines, cc_summary: summary, obs_lines: obs, obs_lines_count: obs.length};
    measurement.run.timings.reconciliation_ms += performance.now() - reconcileStart;
    await UploadUI.yieldToBrowser();
    result.excel_blob = await measurement.measure('workbook_ms', () => EntelWorkbook.create(original, result));
    result.timings = {...measurement.run.timings};
    result.excel_worker = worker;
    result.filename = excelFile.name.replace(/\.xlsx$/i, '') + ' - Conciliado.xlsx';
    return result;
  }
  globalThis.EntelProcessor = {process, cents, phone, billing, inventory, previous, pdfInfo, ensureLibraries};
})();
