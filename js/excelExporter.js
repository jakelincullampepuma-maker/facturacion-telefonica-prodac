/**
 * excelExporter.js - App Celulares PRODAC
 * Generación del Excel Conciliado Oficial (Detalle + Resumen por CC Dpto en la misma hoja y Observaciones)
 */

export const ExcelExporter = {
  /**
   * Genera y descarga el archivo Excel conciliado
   * @param {Object} conciliationResult - Objeto devuelto por BillingMatcher
   */
  exportConsolidatedWorkbook(conciliationResult) {
    if (!conciliationResult || !conciliationResult.detailedList) {
      alert('No hay datos procesados para exportar.');
      return;
    }

    const wb = XLSX.utils.book_new();
    const { company, lineasProcesadas, totalPorLineas, cargosDelMes, diferencia, estadoCuadra, resumenPorCC, detailedList, unmatchedList } = conciliationResult;

    // ==========================================
    // HOJA 1: CONCILIADO Y RESUMEN
    // ==========================================
    const sheet1Data = [];

    sheet1Data.push([`CONCILIACIÓN MENSUAL DE FACTURACIÓN CLARO - ${company}`]);
    sheet1Data.push([`Fecha de Generación: ${new Date().toLocaleString('es-PE')}`]);
    sheet1Data.push([`Estado: ${estadoCuadra} | Total Líneas: ${lineasProcesadas} | Total sin IGV: S/ ${totalPorLineas.toFixed(2)} | Cargos del Mes: S/ ${cargosDelMes.toFixed(2)} | Diferencia: S/ ${diferencia.toFixed(2)}`]);
    sheet1Data.push([]);

    // SECCIÓN A: RESUMEN POR CENTRO DE COSTO
    sheet1Data.push(['--- 1. RESUMEN POR CENTRO DE COSTO (CC DPTO) ---']);
    sheet1Data.push(['Centro de Costo', 'Descripción', 'Líneas', 'Costo sin IGV (S/)', '% del Total']);

    resumenPorCC.forEach(cc => {
      sheet1Data.push([
        cc.ccDpto,
        cc.description,
        cc.lineas,
        cc.costoSinIgv,
        `${cc.porcentaje}%`
      ]);
    });

    sheet1Data.push(['TOTAL GENERAL', 'Consolidado', lineasProcesadas, totalPorLineas, '100.0%']);
    sheet1Data.push([]);
    sheet1Data.push([]);

    // SECCIÓN B: DETALLE DE LÍNEAS
    sheet1Data.push(['--- 2. DETALLE DE NÚMEROS, COSTOS Y CENTROS DE COSTO ---']);
    sheet1Data.push(['N°', 'Línea Telefónica', 'Usuario', 'Centro de Costo', 'Descripción Área', 'Cargo del Mes (S/)', 'Observación']);

    detailedList.forEach(item => {
      sheet1Data.push([
        item.n,
        item.linea,
        item.usuario,
        item.ccDpto,
        item.descripcion,
        item.cargoDelMes,
        item.observacion
      ]);
    });

    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    ws1['!cols'] = [
      { wch: 18 },
      { wch: 26 },
      { wch: 22 },
      { wch: 20 },
      { wch: 20 },
      { wch: 18 },
      { wch: 18 }
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'Conciliado y Resumen');

    // ==========================================
    // HOJA 2: LÍNEAS SIN CC Y OBSERVACIONES
    // ==========================================
    const sheet2Data = [];
    sheet2Data.push([`LÍNEAS SIN CC Y OBSERVACIONES - CLARO ${company}`]);
    sheet2Data.push([`Total líneas por revisar: ${unmatchedList.length}`]);
    sheet2Data.push([]);
    sheet2Data.push(['N°', 'Línea Telefónica', 'Empresa', 'Cargo del Mes (S/)', 'Estado']);

    if (unmatchedList.length === 0) {
      sheet2Data.push([1, 'Todas las líneas cuentan con Centro de Costo asignado', company, 0, 'CUADRA OK']);
    } else {
      unmatchedList.forEach((item, idx) => {
        sheet2Data.push([
          idx + 1,
          item.linea,
          company,
          item.cargoDelMes,
          'Sin CC - No encontrado en Inventario'
        ]);
      });
    }

    const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    ws2['!cols'] = [
      { wch: 8 },
      { wch: 20 },
      { wch: 16 },
      { wch: 20 },
      { wch: 40 }
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Sin CC y Observaciones');

    const fileName = `Conciliacion_Claro_${company}_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}.xlsx`;
    XLSX.writeFile(wb, fileName);
  }
};
