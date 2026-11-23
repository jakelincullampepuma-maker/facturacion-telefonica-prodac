"""Libro Claro con tabla estructurada y PivotTable OOXML nativa."""
import io
from collections import defaultdict
from decimal import Decimal
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import TableDefinition
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.properties import CalcProperties

SOURCE_TABLE = 'ClaroDatos'
MONEY = '#,##0.00;-#,##0.00'
PIVOT_HEADER_ROW = 7


def node(tag, parent=None, **attrs):
    element = ET.Element(tag, {k: str(v) for k, v in attrs.items()})
    if parent is not None:
        parent.append(element)
    return element


def add_pivot(sheet, records, headers, groups, table_name=SOURCE_TABLE, pivot_name='ClaroResumenCC'):
    definition = node('pivotTableDefinition', name=pivot_name, cacheId=1,
                      dataCaption='Valores', grandTotalCaption='Total general',
                      rowHeaderCaption='Etiquetas de fila', createdVersion=3,
                      updatedVersion=3, minRefreshableVersion=3)
    node('location', definition, ref=f'G{PIVOT_HEADER_ROW}:H{len(groups) + PIVOT_HEADER_ROW + 1}',
         firstHeaderRow=1, firstDataRow=1, firstDataCol=1)
    fields = node('pivotFields', definition, count=len(headers))
    node('pivotField', fields, showAll=0)
    node('pivotField', fields, dataField=1, showAll=0)
    cc_field = node('pivotField', fields, axis='axisRow', showAll=0, sortType='ascending')
    items = node('items', cc_field, count=len(groups) + 1)
    for index in range(len(groups)):
        node('item', items, x=index)
    node('item', items, t='default')
    if len(headers) > 3:
        node('pivotField', fields, showAll=0)
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
    cache = node('pivotCacheDefinition', recordCount=len(records), refreshOnLoad=1,
                 missingItemsLimit=0, enableRefresh=1, saveData=1,
                 createdVersion=3, refreshedVersion=3, minRefreshableVersion=3)
    node('worksheetSource', node('cacheSource', cache, type='worksheet'), name=table_name)
    fields = node('cacheFields', cache, count=len(headers))
    node('sharedItems', node('cacheField', fields, name=headers[0], numFmtId=49))
    node('sharedItems', node('cacheField', fields, name=headers[1], numFmtId=4),
         containsSemiMixedTypes=0, containsString=0, containsNumber=1)
    shared = node('sharedItems', node('cacheField', fields, name=headers[2], numFmtId=49), count=len(groups))
    for cc in groups:
        node('s', shared, v=cc)
    if len(headers) > 3:
        node('sharedItems', node('cacheField', fields, name=headers[3]))
    indices = {cc: i for i, cc in enumerate(groups)}
    cache_records = node('pivotCacheRecords', count=len(records))
    for item in records:
        record = node('r', cache_records)
        node('s', record, v=item['numero'])
        node('n', record, v=item['amount'])
        node('x', record, v=indices[item['cc']])
        if len(headers) > 3:
            node('s', record, v=item['observacion'])
    pivot = TableDefinition.from_tree(definition)
    pivot.cache = CacheDefinition.from_tree(cache)
    pivot.cache.records = RecordList.from_tree(cache_records)
    sheet.add_pivot(pivot)


def text(sheet, row, col, value):
    cell = sheet.cell(row, col, str(value or ''))
    cell.data_type = 's'
    return cell


def generate_claro_excel(report, original_bytes):
    from claro_build import generate_claro_excel as generate
    return generate(report, original_bytes)
