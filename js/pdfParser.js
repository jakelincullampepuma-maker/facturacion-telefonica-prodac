/**
 * pdfParser.js - App Celulares PRODAC
 * Extracción de "Cargos del Mes", N° de Recibo y Período desde el PDF de Claro
 */

export const PdfParser = {
  /**
   * Lee el recibo mensual de Claro en PDF y extrae "Cargos del Mes" para la conciliación.
   * @param {File} file 
   * @returns {Promise<{success: boolean, cargosDelMes?: number, invoiceNumber?: string, period?: string, fileName?: string, error?: string}>}
   */
  async parseClaroInvoicePdf(file) {
    try {
      if (typeof pdfjsLib === 'undefined') {
        return { success: false, error: 'Librería PDF.js no inicializada.' };
      }

      const arrayBuffer = await file.arrayBuffer();
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      const pdf = await loadingTask.promise;
      const totalPages = pdf.numPages;

      let fullRawText = '';

      for (let pageNum = 1; pageNum <= Math.min(totalPages, 5); pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ');
        fullRawText += '\n' + pageText;
      }

      const meta = this.extractCargosAndMeta(fullRawText, file.name);

      return {
        success: true,
        cargosDelMes: meta.cargosDelMes,
        invoiceNumber: meta.invoiceNumber,
        period: meta.period,
        totalPages,
        fileName: file.name
      };
    } catch (err) {
      console.error('Error parseando PDF de Claro:', err);
      return { success: false, error: 'Error al procesar el PDF de Claro: ' + err.message };
    }
  },

  extractCargosAndMeta(text, fileName) {
    let cargosDelMes = 0;
    let invoiceNumber = 'Detectado';
    let period = 'Agosto 2026';

    // 1. Buscar "Cargos del Mes" o "Cargos del Periodo" o "Total a Pagar"
    const regexCargos = /(?:cargos?\s*del\s*(?:mes|per[ií]odo)|total\s*(?:cargos|a\s*pagar|facturado)|importe\s*total)[\s:S/$.]*([\d,]+\.\d{2})/i;
    const matchCargos = text.match(regexCargos);

    if (matchCargos) {
      cargosDelMes = parseFloat(matchCargos[1].replace(/,/g, ''));
    } else {
      // Fallback: buscar montos con decimales cerca de palabras clave
      const matchAnyAmount = text.match(/(?:S\/\.?\s*)([\d,]+\.\d{2})/i);
      if (matchAnyAmount) {
        cargosDelMes = parseFloat(matchAnyAmount[1].replace(/,/g, ''));
      }
    }

    // 2. Nro de Recibo / Factura
    const invMatch = text.match(/(?:recibo|factura|n[uú]mero)[:\s#]*([A-Z0-9]{3,4}[-\s]\d{5,10})/i);
    if (invMatch) invoiceNumber = invMatch[1].replace(/\s+/g, '-');

    // 3. Período
    const perMatch = text.match(/(?:per[ií]odo|mes\s*de\s*facturaci[oó]n)[:\s]*([0-9a-zA-Z\s/\-al]+?)(?=\n|$)/i);
    if (perMatch) period = perMatch[1].trim();

    return {
      cargosDelMes: Number(cargosDelMes.toFixed(2)),
      invoiceNumber,
      period
    };
  }
};
