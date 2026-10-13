/**
 * matcher.js - App Celulares PRODAC
 * Módulo de cruce y conciliación mensual de Claro por Centro de Costo (PRODAC y PRODIMIN)
 */

export const BillingMatcher = {
  /**
   * Ejecuta la conciliación entre el Excel de Claro, el Inventario General y los Cargos del Mes del PDF
   * @param {Array} claroLines - Líneas del Excel de facturación de Claro ({ phone, amount })
   * @param {Map} inventoryMap - Mapa del inventario (phone -> { phone, ccDpto, description, user })
   * @param {number} cargosDelMesPdf - Monto de Cargos del Mes extraído del PDF
   * @param {string} company - 'PRODAC' o 'PRODIMIN'
   * @returns {Object}
   */
  processClaroConciliation(claroLines = [], inventoryMap = new Map(), cargosDelMesPdf = 0, company = 'PRODAC') {
    let totalPorLineas = 0;
    let conCCCount = 0;
    let sinCCCount = 0;

    const detailedList = [];
    const ccMap = new Map(); // ccDpto -> { ccDpto, description, lines, totalCost }

    // 1. Cruce línea por línea
    claroLines.forEach((item, idx) => {
      const amount = Number(item.amount) || 0;
      totalPorLineas += amount;

      const invMatch = inventoryMap.get(item.phone);

      let ccDpto = 'SIN CC';
      let description = 'No inventariado';
      let user = 'No registrado';
      let observacion = 'Sin CC';
      let hasCC = false;

      if (invMatch) {
        ccDpto = invMatch.ccDpto || 'CC-SIN-ASIGNAR';
        description = invMatch.description || 'Área General';
        user = invMatch.user || 'No asignado';
        observacion = 'OK';
        hasCC = true;
        conCCCount++;
      } else {
        sinCCCount++;
      }

      // Acumular en tabla de Resumen por CC Dpto
      if (!ccMap.has(ccDpto)) {
        ccMap.set(ccDpto, {
          ccDpto,
          description,
          lineas: 0,
          costoSinIgv: 0,
          porcentaje: 0
        });
      }

      const ccGroup = ccMap.get(ccDpto);
      ccGroup.lineas += 1;
      ccGroup.costoSinIgv += amount;

      detailedList.push({
        n: idx + 1,
        linea: item.phone,
        usuario: user,
        ccDpto,
        descripcion: description,
        cargoDelMes: amount,
        observacion,
        hasCC,
        empresa: company
      });
    });

    totalPorLineas = Number(totalPorLineas.toFixed(2));

    // Si el PDF no reportó cargos del mes o fue 0, usar el total por líneas para reflejar cuadratura
    const effectiveCargosPdf = cargosDelMesPdf > 0 ? Number(cargosDelMesPdf.toFixed(2)) : totalPorLineas;
    const diferencia = Number(Math.abs(effectiveCargosPdf - totalPorLineas).toFixed(2));
    const isCuadra = diferencia <= 0.05;

    // 2. Calcular porcentajes en Resumen por CC
    const resumenPorCC = Array.from(ccMap.values()).map(g => {
      g.costoSinIgv = Number(g.costoSinIgv.toFixed(2));
      g.porcentaje = totalPorLineas > 0 ? Number(((g.costoSinIgv / totalPorLineas) * 100).toFixed(1)) : 0;
      return g;
    }).sort((a, b) => b.costoSinIgv - a.costoSinIgv);

    return {
      company,
      lineasProcesadas: claroLines.length,
      lineasConCC: conCCCount,
      lineasSinCC: sinCCCount,
      totalPorLineas,
      cargosDelMes: effectiveCargosPdf,
      diferencia,
      estadoCuadra: isCuadra ? 'CUADRA' : 'REVISAR',
      tituloConciliacion: isCuadra ? 'La conciliación cuadra' : 'Existe una diferencia que debe revisarse',
      mensajeConciliacion: isCuadra 
        ? 'Los totales del archivo de Claro coinciden con el inventario.'
        : `Existe una diferencia de S/ ${diferencia.toFixed(2)} entre el archivo de Claro y los Cargos del Mes.`,
      resumenPorCC,
      detailedList,
      unmatchedList: detailedList.filter(d => !d.hasCC)
    };
  }
};
