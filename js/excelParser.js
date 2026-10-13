/**
 * excelParser.js - App Celulares PRODAC
 * Lectura y análisis de:
 * 1. Excel de Facturación Detallada de Claro (proporciona Número y Costo sin IGV)
 * 2. Excel del Inventario General (proporciona CC Dpto, Descripción y Usuario por hoja PRODAC / PRODIMIN)
 */

export const ExcelParser = {
  /**
   * Lee el Excel de facturación detallada de Claro.
   * Extrae cada línea telefónica y su costo sin IGV.
   * @param {File} file 
   * @returns {Promise<{success: boolean, lines?: Array, totalCost?: number, fileName?: string, error?: string}>}
   */
  async parseClaroBillingExcel(file) {
    return new Promise((resolve) => {
      try {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
              return resolve({ success: false, error: 'El Excel de Claro no contiene hojas válidas.' });
            }

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

            if (rawRows.length === 0) {
              return resolve({ success: false, error: 'La hoja de facturación de Claro está vacía.' });
            }

            const headers = Object.keys(rawRows[0]);
            const mapping = this.detectClaroBillingHeaders(headers);

            const lines = [];
            let totalCost = 0;

            rawRows.forEach((row, idx) => {
              const rawPhone = mapping.phone ? String(row[mapping.phone] || '').trim() : '';
              const phone = this.normalizePhone(rawPhone);

              let rawAmount = mapping.amount ? String(row[mapping.amount] || '').trim() : '0';
              // Limpiar símbolos de moneda y comas
              rawAmount = rawAmount.replace(/[S/$,\s]/g, '');
              let amount = parseFloat(rawAmount) || 0;

              if (phone) {
                totalCost += amount;
                lines.push({
                  idx: idx + 1,
                  phone,
                  rawPhone,
                  amount,
                  rawRow: row
                });
              }
            });

            if (lines.length === 0) {
              return resolve({ success: false, error: 'No se encontraron números telefónicos válidos en el archivo de Claro.' });
            }

            resolve({
              success: true,
              lines,
              totalCost: Number(totalCost.toFixed(2)),
              totalLines: lines.length,
              fileName: file.name
            });
          } catch (err) {
            resolve({ success: false, error: 'Error procesando Excel de Claro: ' + err.message });
          }
        };
        reader.onerror = () => resolve({ success: false, error: 'Error al leer el archivo Excel de Claro.' });
        reader.readAsArrayBuffer(file);
      } catch (err) {
        resolve({ success: false, error: 'Excepción al leer archivo: ' + err.message });
      }
    });
  },

  /**
   * Lee el Excel del Inventario General.
   * Soporta pestañas independientes (PRODAC / PRODIMIN) o filtrado por columna.
   * @param {File} file 
   * @param {'PRODAC' | 'PRODIMIN'} company 
   * @returns {Promise<{success: boolean, inventory?: Map, list?: Array, fileName?: string, error?: string}>}
   */
  async parseInventoryFile(file, company = 'PRODAC') {
    return new Promise((resolve) => {
      try {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
              return resolve({ success: false, error: 'El archivo de inventario no contiene hojas.' });
            }

            // Buscar hoja específica de la empresa (ej: "PRODAC" o "PRODIMIN")
            let selectedSheet = workbook.SheetNames[0];
            const targetClean = company.toUpperCase();

            for (const name of workbook.SheetNames) {
              if (name.toUpperCase().includes(targetClean)) {
                selectedSheet = name;
                break;
              }
            }

            const worksheet = workbook.Sheets[selectedSheet];
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

            if (rawRows.length === 0) {
              return resolve({ success: false, error: `La hoja "${selectedSheet}" del inventario está vacía.` });
            }

            const headers = Object.keys(rawRows[0]);
            const mapping = this.detectInventoryHeaders(headers);

            const inventoryMap = new Map(); // phone -> item
            const inventoryList = [];

            rawRows.forEach((row, idx) => {
              const rawPhone = mapping.phone ? String(row[mapping.phone] || '').trim() : '';
              const phone = this.normalizePhone(rawPhone);

              const ccDpto = mapping.ccDpto ? String(row[mapping.ccDpto] || '').trim() : 'CC-001';
              const description = mapping.description ? String(row[mapping.description] || '').trim() : 'General';
              const user = mapping.user ? String(row[mapping.user] || '').trim() : 'No asignado';

              if (phone) {
                const item = {
                  id: idx + 1,
                  phone,
                  rawPhone,
                  ccDpto: ccDpto || 'CC-SIN-ASIGNAR',
                  description: description || 'Área General',
                  user: user || 'No asignado'
                };
                inventoryMap.set(phone, item);
                inventoryList.push(item);
              }
            });

            const uniqueCCs = [...new Set(inventoryList.map(i => i.ccDpto))];

            resolve({
              success: true,
              inventoryMap,
              inventoryList,
              totalLines: inventoryList.length,
              uniqueCostCenters: uniqueCCs,
              sheetUsed: selectedSheet,
              fileName: file.name
            });
          } catch (err) {
            resolve({ success: false, error: 'Error procesando Inventario: ' + err.message });
          }
        };
        reader.onerror = () => resolve({ success: false, error: 'Error al leer el archivo de Inventario.' });
        reader.readAsArrayBuffer(file);
      } catch (err) {
        resolve({ success: false, error: 'Excepción al leer inventario: ' + err.message });
      }
    });
  },

  detectClaroBillingHeaders(headers) {
    const mapping = { phone: null, amount: null };
    const clean = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    headers.forEach(h => {
      const c = clean(h);
      if (!mapping.phone && /linea|telefono|celular|numero|nro|movil|servicio/.test(c)) {
        mapping.phone = h;
      }
      if (!mapping.amount && /sin.*igv|costo|cargo.*mes|cargo.*fijo|total.*sin|monto|importe|subtotal|cargos/.test(c)) {
        mapping.amount = h;
      }
    });

    if (!mapping.phone && headers.length > 0) mapping.phone = headers[0];
    if (!mapping.amount && headers.length > 1) mapping.amount = headers[headers.length - 1];
    return mapping;
  },

  detectInventoryHeaders(headers) {
    const mapping = { phone: null, ccDpto: null, description: null, user: null };
    const clean = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    headers.forEach(h => {
      const c = clean(h);
      if (!mapping.phone && /linea|telefono|celular|numero|nro|movil/.test(c)) {
        mapping.phone = h;
      }
      if (!mapping.ccDpto && /cc.*dpto|cc|centro.*costo|cod.*cc|codigo.*cc/.test(c)) {
        mapping.ccDpto = h;
      }
      if (!mapping.description && /descripcion|area|departamento|dpto|unidad|gerencia|nombre.*cc/.test(c)) {
        mapping.description = h;
      }
      if (!mapping.user && /usuario|responsable|asignado|colaborador|personal|nombre|titular/.test(c)) {
        mapping.user = h;
      }
    });

    if (!mapping.phone && headers.length > 0) mapping.phone = headers[0];
    if (!mapping.ccDpto && headers.length > 1) mapping.ccDpto = headers[1];
    if (!mapping.description && headers.length > 2) mapping.description = headers[2];
    if (!mapping.user && headers.length > 3) mapping.user = headers[3];

    return mapping;
  },

  normalizePhone(phoneStr) {
    if (!phoneStr) return '';
    let digits = String(phoneStr).replace(/\D/g, '');
    if (digits.length === 11 && digits.startsWith('519')) {
      digits = digits.substring(2);
    }
    if (digits.length === 9 && digits.startsWith('9')) {
      return digits;
    }
    if (digits.length >= 6) {
      return digits;
    }
    return '';
  }
};
