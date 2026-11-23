"""Excel Movistar: tres hojas y PivotTable nativa, sin documentos maestros.

Las clases OOXML de openpyxl crean definición, caché, registros y relaciones.
No se reutilizan cachés del mes anterior ni archivos corporativos como plantilla.
Referencia: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables
"""
import io
import zipfile
from decimal import Decimal
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import TableDefinition
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MONEY = '#,##0.00;[Red]-#,##0.00'
SOURCE_TABLE_NAME = 'MovistarDatos'


def node(tag, parent=None, **attrs):
    element = ET.Element(tag, {k: str(v) for k, v in attrs.items()})
    if parent is not None:
        parent.append(element)
    return element


def add_pivot(sheet, lines, groups):
    """Construye todas las partes de una PivotTable real con datos del mes actual."""
    definition = node('pivotTableDefinition', name='MovistarResumenCC', cacheId=1,
                      dataCaption='Valores', grandTotalCaption='Total general',
                      rowHeaderCaption='Etiquetas de fila', createdVersion=3,
                      updatedVersion=3, minRefreshableVersion=3)
    node('location', definition, ref=f'G4:H{len(groups) + 5}',
         firstHeaderRow=1, firstDataRow=1, firstDataCol=1)
    fields = node('pivotFields', definition, count=3)
    node('pivotField', fields, showAll=0)
    node('pivotField', fields, dataField=1, showAll=0)
    cc_field = node('pivotField', fields, axis='axisRow', showAll=0, sortType='ascending')
    items = node('items', cc_field, count=len(groups) + 1)
    for index in range(len(groups)):
        node('item', items, x=index)
    node('item', items, t='default')
    node('field', node('rowFields', definition, count=1), x=2)
    row_items = node('rowItems', definition, count=len(groups) + 1)
    for index in range(len(groups)):
        node('x', node('i', row_items), v=index)
    node('x', node('i', row_items, t='grand'))
    node('i', node('colItems', definition, count=1))
    node('dataField', node('dataFields', definition, count=1), name='Suma de MONTO',
         fld=1, subtotal='sum', baseField=0, baseItem=0, numFmtId=4)
    node('pivotTableStyleInfo', definition, name='PivotStyleLight16',
         showRowHeaders=1, showColHeaders=1, showRowStripes=0,
         showColStripes=0, showLastColumn=1)

    cache = node('pivotCacheDefinition', recordCount=len(lines), refreshOnLoad=1, missingItemsLimit=0,
                 enableRefresh=1, saveData=1, createdVersion=3,
                 refreshedVersion=3, minRefreshableVersion=3)
    node('worksheetSource', node('cacheSource', cache, type='worksheet'),
         name=SOURCE_TABLE_NAME)
    cache_fields = node('cacheFields', cache, count=3)
    node('sharedItems', node('cacheField', cache_fields, name='NUMERO', numFmtId=49))
    node('sharedItems', node('cacheField', cache_fields, name='MONTO', numFmtId=4),
         containsSemiMixedTypes=0, containsString=0, containsNumber=1)
    shared = node('sharedItems', node('cacheField', cache_fields, name='CC', numFmtId=49),
                  count=len(groups))
    indices = {cc: i for i, cc in enumerate(groups)}
    for cc in groups:
        node('s', shared, v=cc)
    records = node('pivotCacheRecords', count=len(lines))
    for line in lines:
        record = node('r', records)
        node('s', record, v=line['phone'])
        node('n', record, v=line['monto_f'])
        node('x', record, v=indices[line['cc']])
    pivot = TableDefinition.from_tree(definition)
    pivot.cache = CacheDefinition.from_tree(cache)
    pivot.cache.records = RecordList.from_tree(records)
    sheet.add_pivot(pivot)


def generate_movistar_excel(lines, inafectos, inv_map, pdf_summary):
    wb = openpyxl.Workbook()
    wb.active.title = 'CARGOS ADICIONALES'
    cm = wb.create_sheet('CARGO MENSUAL')
    h1 = wb.create_sheet('Hoja1')
    ca = wb.worksheets[0]
    cached = {s.title: {} for s in wb}

    def formula(sheet, cell, expression, value):
        sheet[cell] = expression
        cached[sheet.title][cell] = value

    def text(sheet, row, col, value):
        cell = sheet.cell(row, col, str(value))
        cell.data_type = 's'
        return cell

    for sheet, headers in [
        (ca, ['NUMERO', 'DESCRIPCION', 'MONTO']),
        (cm, ['N°', 'NUMERO', 'PLAN', 'MONTO', 'DESC', 'MONTO F'])
    ]:
        for col, header in enumerate(headers, 1):
            sheet.cell(1, col, header)
    for row, item in enumerate(inafectos, 2):
        text(ca, row, 1, item['phone']).number_format = '@'
        text(ca, row, 2, item['description']).alignment = Alignment(wrap_text=True)
        ca.cell(row, 3, float(item['monto']))
    end_ca = len(inafectos) + 1
    ca.cell(end_ca + 1, 2, 'TOTAL')
    formula(ca, f'C{end_ca + 1}', f'=SUM(C2:C{end_ca})' if inafectos else '=0',
            sum((i['monto'] for i in inafectos), Decimal('0')))
    ca.auto_filter.ref = f'A1:C{end_ca}'

    for row, line in enumerate(lines, 2):
        cm.cell(row, 1, row - 1)
        text(cm, row, 2, line['phone']).number_format = '@'
        text(cm, row, 3, line['plan']).alignment = Alignment(wrap_text=True)
        cm.cell(row, 4, float(line['monto']))
        cm.cell(row, 5, float(line['desc']))
        formula(cm, f'F{row}', f'=D{row}-E{row}', line['monto_f'])
    end_cm = len(lines) + 1
    cm.cell(end_cm + 1, 3, 'TOTAL')
    for col, key in [('D', 'monto'), ('E', 'desc'), ('F', 'monto_f')]:
        formula(cm, f'{col}{end_cm + 1}', f'=SUM({col}2:{col}{end_cm})',
                sum((l[key] for l in lines), Decimal('0')))
    cm.cell(end_cm + 2, 3, 'Subtotal PDF')
    cm.cell(end_cm + 2, 4, float(pdf_summary['subtotal']))
    cm.cell(end_cm + 3, 3, 'Diferencia')
    formula(cm, f'D{end_cm + 3}', f'=F{end_cm + 1}-D{end_cm + 2}',
            sum(l['monto_f'] for l in lines) - pdf_summary['subtotal'])
    cm.auto_filter.ref = f'A1:F{end_cm}'

    h1.append(['NUMERO', 'MONTO', 'CC'])
    groups = {}
    for row, line in enumerate(lines, 2):
        cc = inv_map.get(line['phone']) or 'SIN CC'
        line['cc'] = cc
        text(h1, row, 1, line['phone']).number_format = '@'
        h1.cell(row, 2, float(line['monto_f']))
        text(h1, row, 3, cc)
        groups[cc] = groups.get(cc, Decimal('0')) + line['monto_f']
    last = len(lines) + 1
    h1.cell(last + 1, 1, 'TOTAL')
    total_formula = f'SUM({SOURCE_TABLE_NAME}[MONTO])'
    formula(h1, f'B{last + 1}', '=' + total_formula, sum(groups.values()))
    source = Table(displayName=SOURCE_TABLE_NAME, ref=f'A1:C{last + 1}',
                   totalsRowCount=1, totalsRowShown=True,
                   autoFilter=AutoFilter(ref=f'A1:C{last}'),
                   tableColumns=[
                       TableColumn(id=1, name='NUMERO', totalsRowLabel='TOTAL'),
                       TableColumn(id=2, name='MONTO', totalsRowFunction='custom',
                                   totalsRowFormula=TableFormula(attr_text=total_formula)),
                       TableColumn(id=3, name='CC')],
                   tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=False,
                                                showColumnStripes=False,
                                                showFirstColumn=False, showLastColumn=False))
    h1.add_table(source)
    h1['G4'], h1['H4'] = 'Etiquetas de fila', 'Suma de MONTO'
    for row, cc in enumerate(sorted(groups), 5):
        text(h1, row, 7, cc)
        h1.cell(row, 8, float(groups[cc]))
    total_pivot = len(groups) + 5
    h1.cell(total_pivot, 7, 'Total general')
    h1.cell(total_pivot, 8, float(sum(groups.values())))
    add_pivot(h1, lines, sorted(groups))

    for sheet, header, widths, money_cols in [
        (ca, 1, {'A': 17, 'B': 85, 'C': 19}, [3]),
        (cm, 1, {'A': 7, 'B': 17, 'C': 65, 'D': 19, 'E': 19, 'F': 19}, [4, 5, 6]),
        (h1, 1, {'A': 17, 'B': 19, 'C': 24, 'D': 4, 'E': 4, 'F': 4, 'G': 25, 'H': 22}, [2, 8])
    ]:
        sheet.freeze_panes = f'A{header + 1}'
        sheet.sheet_view.showGridLines = False
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
        for row in sheet.iter_rows():
            for cell in row:
                cell.font = Font(name='Calibri', size=11)
                if cell.column in money_cols and cell.row > header:
                    cell.number_format = MONEY
        for cell in sheet[header]:
            if cell.value:
                cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='0369A1')
        sheet.row_dimensions[header].height = 24
    for sheet, row, first, end in [(ca, end_ca + 1, 1, 3), (cm, end_cm + 1, 1, 6),
                                   (h1, last + 1, 1, 3), (h1, 4, 7, 8),
                                   (h1, total_pivot, 7, 8)]:
        for col in range(first, end + 1):
            cell = sheet.cell(row, col)
            cell.font = Font(name='Calibri', bold=True, size=11)
            cell.fill = PatternFill('solid', fgColor='D9EAF7')
    wb.calculation = CalcProperties(calcMode='auto', fullCalcOnLoad=True)
    out = io.BytesIO()
    wb.save(out)

    # Conservar fórmulas y valores calculados para vistas previas antes de abrir Excel.
    result = io.BytesIO()
    with zipfile.ZipFile(out) as src, zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as dst:
        sheet_paths = {f'xl/worksheets/sheet{i}.xml': cached[s.title]
                       for i, s in enumerate(wb, 1)}
        for part in src.infolist():
            data = src.read(part.filename)
            if part.filename in sheet_paths:
                root = ET.fromstring(data)
                for cell in root.findall(f'.//{{{NS}}}c'):
                    value = sheet_paths[part.filename].get(cell.get('r'))
                    if value is not None:
                        old = cell.find(f'{{{NS}}}v')
                        if old is None:
                            old = ET.SubElement(cell, f'{{{NS}}}v')
                        old.text = str(value)
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            dst.writestr(part, data)
    return result.getvalue()
