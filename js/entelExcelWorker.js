/* SheetJS y la indexación de hojas se ejecutan fuera del hilo de la interfaz. */
importScripts('https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js', 'entelProcessor.js');
self.onmessage = event => {
  const {original, inventoryBytes} = event.data;
  try {
    let start = performance.now();
    let wb, invWb;
    try { wb = XLSX.read(original, {type: 'array'}); }
    catch (error) { throw new Error('No se pudo leer el Excel de Entel: ' + error.message); }
    try { invWb = XLSX.read(inventoryBytes, {type: 'array'}); }
    catch (error) { throw new Error('No se pudo leer el inventario: ' + error.message); }
    const detail = EntelProcessor.billing(wb), prev = EntelProcessor.previous(wb);
    const excelMs = performance.now() - start;
    start = performance.now();
    const inv = EntelProcessor.inventory(invWb);
    self.postMessage({original, detail, inv, prev, timings: {excel_read_ms: excelMs,
      inventory_match_ms: performance.now() - start}}, [original]);
  } catch (error) { self.postMessage({error: error.message}); }
};
self.postMessage({ready: true});
