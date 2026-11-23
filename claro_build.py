"""Hojas de trabajo y resultado con tablas estructuradas y pivotes nativos."""
import io
import zipfile
from collections import defaultdict
from decimal import Decimal
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Border, Font, PatternFill
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from claro_package import preserve_billing, S
from claro_workbook import add_pivot, text, MONEY, PIVOT_HEADER_ROW


def generate_claro_excel(report, original_bytes):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cached = {}
    headers = ['NUMERO', 'MONTO', 'CC']
    records = report['details']
    working = [{'numero': r['numero'], 'amount': r['amount'], 'cc': r['cc']} for r in records]
    if report['varios']:
        working.append({'numero': 'Varios', 'amount': report['ajustes_abonos'], 'cc': 'Varios'})
    final = [{'numero': r['numero'], 'amount': r['final_amount'], 'cc': r['cc']} for r in records]
    for title, source_name, items in [('Hoja2', 'ClaroTrabajo', working), ('Hoja1', 'ClaroFinal', final)]:
        sheet = wb.create_sheet(title)
        cached[title] = {}
        sheet.append(headers)
        groups = defaultdict(lambda: Decimal('0.00'))
        for row, item in enumerate(items, 2):
            text(sheet, row, 1, item['numero']).number_format = '@'
            sheet.cell(row, 2, float(item['amount'])).number_format = MONEY
            text(sheet, row, 3, item['cc'])
            groups[item['cc']] += item['amount']
        # Conservar el origen ampliable del pivote sin estilo ni filtros visibles.
        table = Table(displayName=source_name, ref=f'A1:C{len(items) + 1}',
                      tableColumns=[TableColumn(id=i, name=header) for i, header in enumerate(headers, 1)])
        table.tableStyleInfo = TableStyleInfo(name='', showFirstColumn=False, showLastColumn=False,
                                             showRowStripes=False, showColumnStripes=False)
        sheet.add_table(table)
        first = PIVOT_HEADER_ROW + 1
        sheet.cell(PIVOT_HEADER_ROW, 7, 'Etiquetas de fila')
        sheet.cell(PIVOT_HEADER_ROW, 8, 'Suma de MONTO')
        sorted_groups = sorted(groups)
        for row, cc in enumerate(sorted_groups, first):
            text(sheet, row, 7, cc)
            sheet.cell(row, 8, float(groups[cc])).number_format = MONEY
        last = len(groups) + first
        sheet.cell(last, 7, 'Total general')
        sheet.cell(last, 8, float(sum(groups.values()))).number_format = MONEY
        add_pivot(sheet, items, headers, sorted_groups, table_name=source_name, pivot_name='Resumen' + title)
        # openpyxl no asigna cacheId y crea el Target de records antes de
        # actualizar su _id: cada pivote necesita identificadores propios.
        sheet._pivots[0].cacheId = len(wb.worksheets)
        sheet._pivots[0].cache.records._id = len(wb.worksheets)
        if title == 'Hoja2' and report['varios']:
            sheet.cell(PIVOT_HEADER_ROW, 9, 'prorrateo')
            sheet.cell(PIVOT_HEADER_ROW, 10, 'total ajustado')
            by_cc = {r['cc']: r for r in report['summary']}
            eligible = [cc for cc in sorted(by_cc) if cc != 'SIN CC - REVISAR' and by_cc[cc]['antes'] != 0]
            general = 'SUMIF(ClaroTrabajo[CC],"Varios",ClaroTrabajo[MONTO])'
            denominator = 'SUMIFS(ClaroTrabajo[MONTO],ClaroTrabajo[CC],"<>Varios",ClaroTrabajo[CC],"<>SIN CC - REVISAR")'
            for row, cc in enumerate(sorted_groups, first):
                if cc == 'Varios':
                    continue
                if cc not in eligible:
                    expression = '=0'
                elif cc == eligible[-1]:
                    previous = f'SUM(I{first}:I{row - 1})' if row > first else '0'
                    expression = f'=ROUND({general}-{previous},2)'
                else:
                    expression = f'=IFERROR(ROUND(H{row}/{denominator}*{general},2),0)'
                sheet[f'I{row}'] = expression
                sheet[f'J{row}'] = f'=ROUND(H{row}+I{row},2)'
                cached[title][f'I{row}'] = by_cc[cc]['prorrateo']
                cached[title][f'J{row}'] = by_cc[cc]['total']
            sheet[f'I{last}'] = f'=SUM(I{first}:I{last - 1})'
            sheet[f'J{last}'] = f'=SUM(J{first}:J{last - 1})'
            cached[title][f'I{last}'] = report['ajustes_abonos']
            cached[title][f'J{last}'] = report['total_calculado']
            for row in range(first, last + 1):
                sheet.cell(row, 9).number_format = MONEY
                sheet.cell(row, 10).number_format = MONEY
        for col, width in {'A': 16, 'B': 16, 'C': 18, 'D': 14, 'E': 14, 'F': 14,
                           'G': 26, 'H': 23, 'I': 16, 'J': 18}.items():
            sheet.column_dimensions[col].width = width
        sheet.freeze_panes = 'A2'
        for row in sheet:
            for cell in row:
                cell.font = Font(name='Calibri', size=11, color='000000')
                cell.fill = PatternFill()
                cell.border = Border()
        for row in (PIVOT_HEADER_ROW, last):
            for col in (7, 8):
                cell = sheet.cell(row, col)
                cell.font = Font(name='Calibri', size=11, bold=True, color='000000')
                cell.fill = PatternFill('solid', fgColor='C6E7F4')
    wb.calculation = CalcProperties(calcMode='auto', fullCalcOnLoad=True)
    output = io.BytesIO()
    wb.save(output)
    cache_output = io.BytesIO()
    with zipfile.ZipFile(output) as src, zipfile.ZipFile(cache_output, 'w', zipfile.ZIP_DEFLATED) as dst:
        sheets = {f'xl/worksheets/sheet{i}.xml': cached[s.title] for i, s in enumerate(wb, 1)}
        for part in src.infolist():
            data = src.read(part.filename)
            if part.filename in sheets:
                root = ET.fromstring(data)
                for cell in root.findall(f'.//{{{S}}}c'):
                    if cell.get('r') in sheets[part.filename]:
                        value = cell.find(f'{{{S}}}v')
                        if value is None:
                            value = ET.SubElement(cell, f'{{{S}}}v')
                        value.text = str(sheets[part.filename][cell.get('r')])
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            dst.writestr(part, data)
    return preserve_billing(original_bytes, cache_output.getvalue())
