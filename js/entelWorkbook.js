/* Añade una hoja OOXML al ZIP original sin reescribir las hojas del proveedor. */
(function () {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PACK = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const xml = s => String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const parse = text => {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('El XLSX contiene XML no válido.');
    return doc;
  };
  const serialize = doc => new XMLSerializer().serializeToString(doc);
  const elements = (doc, name) => [...doc.getElementsByTagNameNS('*', name)];
  function element(doc, ns, tag, attrs = {}) {
    const e = doc.createElementNS(ns, tag);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    return e;
  }
  async function create(original, result) {
    let zip;
    try { zip = await JSZip.loadAsync(original, {checkCRC32: true}); }
    catch (e) { throw new Error('El proveedor debe ser un XLSX original sin contraseña: ' + e.message); }
    const read = async name => {
      if (!zip.file(name)) throw new Error('Falta ' + name + ' en el XLSX.');
      return parse(await zip.file(name).async('string'));
    };
    const book = await read('xl/workbook.xml'), rels = await read('xl/_rels/workbook.xml.rels'), types = await read('[Content_Types].xml');
    if (book.documentElement.namespaceURI !== NS) throw new Error('El XLSX usa un formato OOXML no compatible. Guárdelo como Libro de Excel (.xlsx).');
    const sheets = elements(book, 'sheets')[0], sheetList = elements(book, 'sheet');
    // Conservar también una hoja Resultado Entel cargada previamente.
    const old = sheetList.find(s => s.getAttribute('name') === 'Resultado Entel');
    if (old) {
      let i = 1, name;
      do { name = 'Resultado Entel anterior ' + i++; } while (sheetList.some(s => s.getAttribute('name') === name));
      old.setAttribute('name', name);
    }
    const id = Math.max(0, ...sheetList.map(s => Number(s.getAttribute('sheetId')))) + 1;
    let rid = 'rIdEntel', suffix = 1;
    while (elements(rels, 'Relationship').some(r => r.getAttribute('Id') === rid)) rid = 'rIdEntel' + suffix++;
    let target = 'worksheets/resultadoEntel.xml';
    suffix = 1;
    while (zip.file('xl/' + target)) target = 'worksheets/resultadoEntel' + suffix++ + '.xml';
    const sheet = element(book, NS, 'sheet', {name: 'Resultado Entel', sheetId: id});
    sheet.setAttributeNS(REL, 'r:id', rid);
    sheets.insertBefore(sheet, sheets.firstChild);
    // Los nombres locales usan índices de hoja; compensar la nueva primera hoja.
    for (const n of elements(book, 'definedName')) if (n.hasAttribute('localSheetId')) n.setAttribute('localSheetId', Number(n.getAttribute('localSheetId')) + 1);
    // Mantener la pestaña activa original desplazada una posición. Marcar además
    // el resultado como seleccionado agruparía pestañas y Excel bloquearía Refresh.
    // Así no es necesario reescribir las vistas de las hojas originales.
    for (const v of elements(book, 'workbookView')) {
      v.setAttribute('activeTab', Number(v.getAttribute('activeTab') || 0) + 1);
      v.setAttribute('firstSheet', '0');
    }
    let calc = elements(book, 'calcPr')[0];
    if (!calc) { calc = element(book, NS, 'calcPr'); book.documentElement.appendChild(calc); }
    calc.setAttribute('fullCalcOnLoad', '1');
    calc.setAttribute('forceFullCalc', '1');
    rels.documentElement.appendChild(element(rels, PACK, 'Relationship', {Id: rid, Type: REL + '/worksheet', Target: target}));
    types.documentElement.appendChild(element(types, CT, 'Override', {PartName: '/xl/' + target, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'}));

    // Añadir estilos, manteniendo todos los índices originales.
    const styleRel = elements(rels, 'Relationship').find(r => r.getAttribute('Type') === REL + '/styles');
    let stylePath = styleRel ? styleRel.getAttribute('Target').replace(/^\//, '') : '';
    if (stylePath && !stylePath.startsWith('xl/')) stylePath = 'xl/' + stylePath;
    if (!stylePath) throw new Error('El XLSX no contiene una tabla de estilos compatible.');
    const styles = await read(stylePath);
    const xfs = elements(styles, 'cellXfs')[0], fonts = elements(styles, 'fonts')[0], fills = elements(styles, 'fills')[0];
    const append = (parent, tagText) => {
      const node = parse('<root xmlns="' + NS + '">' + tagText + '</root>').documentElement.firstChild;
      const index = parent.children.length;
      parent.appendChild(styles.importNode(node, true)); parent.setAttribute('count', parent.children.length);
      return index;
    };
    const normalFont = append(fonts, '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>');
    const boldFont = append(fonts, '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>');
    const fill = append(fills, '<fill><patternFill patternType="solid"><fgColor rgb="FFC0E6F5"/><bgColor indexed="64"/></patternFill></fill>');
    const style = (font, bg, format, align) => append(xfs, '<xf numFmtId="' + format + '" fontId="' + font + '" fillId="' + bg + '" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="' + align + '" vertical="center"/></xf>');
    const textStyle = style(normalFont, 0, 49, 'left');
    const phoneStyle = style(normalFont, 0, 49, 'right');
    const amountStyle = style(normalFont, 0, 4, 'right');
    const headerStyle = style(normalFont, 0, 0, 'left');
    const totalLabelStyle = style(boldFont, 0, 0, 'left');
    const totalAmountStyle = style(boldFont, 0, 4, 'right');
    const pivotHeaderStyle = style(boldFont, fill, 0, 'left');
    const pivotTotalStyle = style(boldFont, fill, 4, 'right');

    // Nombres, IDs y rutas únicos: el proveedor puede traer tablas y pivotes propios.
    const tableDocs = await Promise.all(Object.keys(zip.files).filter(n => /^xl\/tables\/[^/]+\.xml$/.test(n)).map(read));
    const usedNames = new Set([...tableDocs.flatMap(d => [d.documentElement.getAttribute('name'), d.documentElement.getAttribute('displayName')]), ...elements(book, 'definedName').map(n => n.getAttribute('name'))].map(n => (n || '').toLowerCase()));
    const tableId = Math.max(0, ...tableDocs.map(d => Number(d.documentElement.getAttribute('id')))) + 1;
    let tableName = 'EntelDatos' + tableId;
    while (usedNames.has(tableName.toLowerCase())) tableName += '_';
    const pivotDocs = await Promise.all(Object.keys(zip.files).filter(n => /^xl\/pivotTables\/[^/]+\.xml$/.test(n)).map(read));
    const cacheId = Math.max(0, ...elements(book, 'pivotCache').map(c => Number(c.getAttribute('cacheId')))) + 1;
    let pivotName = 'EntelResumen' + cacheId;
    while (pivotDocs.some(d => d.documentElement.getAttribute('name') === pivotName)) pivotName += '_';
    const part = prefix => {
      let i = 1;
      while (zip.file(prefix + i + '.xml')) i++;
      return prefix + i + '.xml';
    };
    const tablePath = part('xl/tables/entelDatos');
    const pivotPath = part('xl/pivotTables/entelResumen');
    const cachePath = part('xl/pivotCache/entelCache');
    const recordsPath = part('xl/pivotCache/entelRecords');
    const basename = path => path.slice(path.lastIndexOf('/') + 1);
    const relPath = path => path.slice(0, path.lastIndexOf('/') + 1) + '_rels/' + basename(path) + '.rels';
    const relXml = relationships => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + PACK + '">' + relationships.map(([id, type, path]) => '<Relationship Id="' + id + '" Type="' + REL + '/' + type + '" Target="' + xml(path) + '"/>').join('') + '</Relationships>';
    const addType = (path, type) => types.documentElement.appendChild(element(types, CT, 'Override', {PartName: '/' + path, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.' + type + '+xml'}));
    let cacheRid = 'rIdEntelCache', cacheSuffix = 1;
    while (elements(rels, 'Relationship').some(r => r.getAttribute('Id') === cacheRid)) cacheRid = 'rIdEntelCache' + cacheSuffix++;
    rels.documentElement.appendChild(element(rels, PACK, 'Relationship', {Id: cacheRid, Type: REL + '/pivotCacheDefinition', Target: cachePath.slice(3)}));
    let caches = elements(book, 'pivotCaches')[0];
    if (!caches) {
      caches = element(book, NS, 'pivotCaches');
      const later = ['smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
      const next = [...book.documentElement.children].find(e => later.includes(e.localName));
      book.documentElement.insertBefore(caches, next || null);
    }
    const cacheEntry = element(book, NS, 'pivotCache', {cacheId});
    cacheEntry.setAttributeNS(REL, 'r:id', cacheRid);
    caches.appendChild(cacheEntry);

    // La tabla dinámica nativa tiene su propia definición, caché y registros.
    // Las celdas G:H son su vista guardada; Excel las regenera al actualizar.
    const groups = [...new Set(result.lines.map(l => l.cc_dpto || 'SIN CC'))].sort((a, b) => a.localeCompare(b));
    const groupIndex = new Map(groups.map((cc, i) => [cc, i]));
    const amounts = groups.map(() => 0);
    for (const line of result.lines) amounts[groupIndex.get(line.cc_dpto || 'SIN CC')] += Math.round(line.cargo_fijo * 100);
    const count = result.lines.length, last = count + 1, totalRow = last + 1, pivotTotalRow = groups.length + 5;
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const sourceXml = declaration + '<table xmlns="' + NS + '" id="' + tableId + '" name="' + tableName + '" displayName="' + tableName + '" ref="A1:C' + totalRow + '" totalsRowCount="1"><autoFilter ref="A1:C' + last + '"/><tableColumns count="3"><tableColumn id="1" name="NUMERO" totalsRowLabel="Total general"/><tableColumn id="2" name="MONTO" totalsRowFunction="custom"><totalsRowFormula>SUM([MONTO])</totalsRowFormula></tableColumn><tableColumn id="3" name="CC"/></tableColumns><tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>';
    const cacheXml = declaration + '<pivotCacheDefinition xmlns="' + NS + '" xmlns:r="' + REL + '" r:id="rIdRecords" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="' + count + '" enableRefresh="1"><cacheSource type="worksheet"><worksheetSource name="' + tableName + '"/></cacheSource><cacheFields count="3"><cacheField name="NUMERO" numFmtId="49"><sharedItems/></cacheField><cacheField name="MONTO" numFmtId="4"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"/></cacheField><cacheField name="CC" numFmtId="49"><sharedItems count="' + groups.length + '">' + groups.map(cc => '<s v="' + xml(cc) + '"/>').join('') + '</sharedItems></cacheField></cacheFields></pivotCacheDefinition>';
    const recordsXml = declaration + '<pivotCacheRecords xmlns="' + NS + '" count="' + count + '">' + result.lines.map(l => '<r><s v="' + xml(l.phone) + '"/><n v="' + l.cargo_fijo + '"/><x v="' + groupIndex.get(l.cc_dpto || 'SIN CC') + '"/></r>').join('') + '</pivotCacheRecords>';
    const pivotXml = declaration + '<pivotTableDefinition xmlns="' + NS + '" name="' + pivotName + '" cacheId="' + cacheId + '" dataCaption="Valores" grandTotalCaption="Total general" rowHeaderCaption="Etiquetas de fila" updatedVersion="3" minRefreshableVersion="3" createdVersion="3" useAutoFormatting="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" itemPrintTitles="1" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="G4:H' + pivotTotalRow + '" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/><pivotField axis="axisRow" showAll="0" sortType="ascending"><items count="' + (groups.length + 1) + '">' + groups.map((_, i) => '<item x="' + i + '"/>').join('') + '<item t="default"/></items></pivotField></pivotFields><rowFields count="1"><field x="2"/></rowFields><rowItems count="' + (groups.length + 1) + '">' + groups.map((_, i) => '<i><x v="' + i + '"/></i>').join('') + '<i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Suma de MONTO" fld="1" subtotal="sum" baseField="0" baseItem="0" numFmtId="4"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
    zip.file(tablePath, sourceXml);
    zip.file(cachePath, cacheXml);
    zip.file(recordsPath, recordsXml);
    zip.file(pivotPath, pivotXml);
    zip.file(relPath('xl/' + target), relXml([['rIdTable', 'table', '../tables/' + basename(tablePath)], ['rIdPivot', 'pivotTable', '../pivotTables/' + basename(pivotPath)]]));
    zip.file(relPath(pivotPath), relXml([['rIdCache', 'pivotCacheDefinition', '../pivotCache/' + basename(cachePath)]]));
    zip.file(relPath(cachePath), relXml([['rIdRecords', 'pivotCacheRecords', basename(recordsPath)]]));
    addType(tablePath, 'table'); addType(pivotPath, 'pivotTable'); addType(cachePath, 'pivotCacheDefinition'); addType(recordsPath, 'pivotCacheRecords');

    const rows = new Map();
    function put(row, col, value, cellStyle, formula) {
      if (!rows.has(row)) rows.set(row, []);
      const ref = XLSX.utils.encode_col(col - 1) + row;
      const body = formula ? '<f>' + xml(formula) + '</f><v>' + value + '</v>' : typeof value === 'number' ? '<v>' + value + '</v>' : '<is><t xml:space="preserve">' + xml(value) + '</t></is>';
      rows.get(row).push('<c r="' + ref + '" s="' + cellStyle + '"' + (formula || typeof value === 'number' ? '' : ' t="inlineStr"') + '>' + body + '</c>');
    }
    ['NUMERO', 'MONTO', 'CC'].forEach((h, i) => put(1, i + 1, h, headerStyle));
    result.lines.forEach((l, i) => {
      put(i + 2, 1, l.phone, phoneStyle);
      put(i + 2, 2, l.cargo_fijo, amountStyle);
      put(i + 2, 3, l.cc_dpto || 'SIN CC', textStyle);
    });
    put(totalRow, 1, 'Total general', totalLabelStyle);
    put(totalRow, 2, result.total_excel, totalAmountStyle, 'SUM(' + tableName + '[MONTO])');
    put(4, 7, 'Etiquetas de fila', pivotHeaderStyle); put(4, 8, 'Suma de MONTO', pivotHeaderStyle);
    groups.forEach((cc, i) => { put(i + 5, 7, cc, textStyle); put(i + 5, 8, amounts[i] / 100, amountStyle); });
    put(pivotTotalRow, 7, 'Total general', pivotHeaderStyle);
    put(pivotTotalRow, 8, result.total_excel, pivotTotalStyle);
    const widths = [18, 16, 22, 5, 5, 5, 24, 22];
    const sheetXml = declaration + '<worksheet xmlns="' + NS + '" xmlns:r="' + REL + '"><dimension ref="A1:H' + Math.max(totalRow, pivotTotalRow) + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>' + widths.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols><sheetData>' + [...rows].sort((a, b) => a[0] - b[0]).map(([r, cells]) => '<row r="' + r + '">' + cells.join('') + '</row>').join('') + '</sheetData><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>';
    zip.file('xl/' + target, sheetXml);
    zip.file('xl/workbook.xml', serialize(book));
    zip.file('xl/_rels/workbook.xml.rels', serialize(rels));
    zip.file('[Content_Types].xml', serialize(types));
    zip.file(stylePath, serialize(styles));
    return zip.generateAsync({type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  window.EntelWorkbook = {create};
})();
