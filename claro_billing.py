"""Cargos del Excel original, validación PDF y prorrateo en dos niveles."""
import io
import re
from collections import defaultdict
from decimal import Decimal

import openpyxl
from pypdf import PdfReader

from claro_engine import clean, phone, money, amounts, parse_metadata, ADJUSTMENT_RE, PHONE_RE, CENT, ZERO, json_values
from claro_workbook import generate_claro_excel
from processing_diagnostics import PhaseTimings


def read_billing(data):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError('No se pudo leer el Excel de Claro en formato XLSX.') from exc
    try:
        if 'Factura Detallada' not in wb.sheetnames:
            raise ValueError('El Excel de Claro debe contener la hoja Factura Detallada.')
        rows = iter(wb['Factura Detallada'].values)
        preamble = []
        for row in rows:
            h = [clean(v) for v in row]
            ni = next((i for i, v in enumerate(h) if v in ('linea', 'numero', 'telefono')), None)
            ai = next((i for i, v in enumerate(h) if 'total cargos por linea' in v and 'sin igv' in v), None)
            if ni is not None and ai is not None:
                pi = next((i for i, v in enumerate(h) if v == 'plan'), None)
                oi = next((i for i, v in enumerate(h) if v == 'otros cargos y abonos'), None)
                break
            preamble.append(' '.join(str(v) for v in row if v is not None))
        else:
            raise ValueError('Faltan Línea y Total Cargos por Línea (Sin IGV) en Factura Detallada.')
        records, general = [], []
        for row in rows:
            if clean(row[0]) in ('total', 'totales', 'subtotal'):
                continue
            raw = row[ni]
            number = phone(raw)
            label = str(raw or '').strip()
            is_general = bool(ADJUSTMENT_RE.search(clean(label)))
            if not number and not is_general:
                continue
            if row[ai] is None:
                raise ValueError(f'Importe sin IGV vacío para {label}.')
            item = {'numero': number or 'Varios', 'amount': money(row[ai]),
                    'plan': str(row[pi] or '') if pi is not None else '',
                    'source_row': label}
            if number:
                records.append(item)
            else:
                item['otros_cargos'] = money(row[oi]) if oi is not None and row[oi] is not None else None
                general.append(item)
        if not records:
            raise ValueError('Factura Detallada no contiene teléfonos con importes válidos.')
        return records, general, '\n'.join(preamble)
    finally:
        wb.close()


def read_inventory(data, company):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError('No se pudo leer el Inventario General.') from exc
    try:
        sheet = next((s for s in wb if clean(s.title) == company.lower()), None)
        if sheet is None:
            raise ValueError(f'El inventario no contiene la hoja {company}.')
        rows = iter(sheet.values)
        for row in rows:
            if len(row) >= 8 and clean(row[2]) == 'numero' and clean(row[7]) == 'cc dpto':
                headers = [clean(v) for v in row]
                user_col = next((i for i, v in enumerate(headers) if v == 'usuario'), None)
                break
        else:
            raise ValueError(f'En {company}, se requiere Numero en C y CC Dpto en H.')
        mapping, conflicts = {}, set()
        for row in rows:
            n = phone(row[2])
            if not n:
                continue
            entry = {'cc': str(row[7] or '').strip(), 'usuario': str(row[user_col] or '') if user_col is not None else ''}
            if n in mapping and mapping[n]['cc'] != entry['cc']:
                conflicts.add(n)
            else:
                mapping[n] = entry
        for n in conflicts:
            mapping[n]['cc'] = ''
        return mapping
    finally:
        wb.close()


def read_pdf(data):
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted and not reader.decrypt(''):
            raise ValueError('El PDF está protegido con contraseña.')
        pages = []
        for index, page in enumerate(reader.pages):
            plain = page.extract_text() or ''
            if index == 0 or re.search(r'otros cargos|ajustes|abonos|descuentos', clean(plain)):
                plain = page.extract_text(extraction_mode='layout', layout_mode_space_vertically=False) or plain
            pages.append(plain)
        text = '\f'.join(pages)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('No se pudo leer el PDF de Claro.') from exc
    meta = parse_metadata(text)
    general, summary = [], []
    active = False
    found_detail = False
    pending = ''
    for line in text.splitlines():
        line = line.strip()
        norm = clean(line)
        if re.match(r'^otros cargos(?: y abonos)? no afectos', norm):
            active = False
            pending = ''
            continue
        values = amounts(line)
        if re.fullmatch(r'otros cargos(?: y abonos)?|ajustes(?: y abonos)?|abonos|devoluciones|descuentos', norm):
            active, pending = True, ''
            found_detail = True
            continue
        if re.match(r'^otros cargos y abonos\s', norm) and values:
            summary.append(values[-1][1])
            continue
        if norm.startswith('descripcion'):
            continue
        if re.match(r'^(?:total|pagina|recibo|detalle|consolidado|consumo)\b', norm):
            active, pending = False, ''
            continue
        if not active:
            continue
        if not values:
            if norm and not re.search(r'igv|i\.g\.v|importe|prodimin|prodac', norm):
                pending = ' '.join(filter(None, [pending, line]))
            continue
        description = ' '.join(filter(None, [pending, line[:values[-1][0].start()].strip()]))
        pending = ''
        if not PHONE_RE.search(description):
            general.append({'description': description, 'amount': values[-1][1]})
    if not general and summary and not found_detail:
        general = [{'description': 'Otros Cargos y Abonos (resumen del PDF)', 'amount': summary[0]}]
    return meta, general


def apportion(total, weights):
    """Redondeo comercial por fila y residuo exacto en la última fila."""
    if not weights:
        if total:
            raise ValueError('No hay centros de costo reales para distribuir Varios.')
        return []
    denominator = sum(weights, ZERO)
    if not denominator:
        if total:
            raise ValueError('No se puede prorratear Varios sobre un total de líneas igual a cero.')
        return [ZERO] * len(weights)
    result = [money(total * weight / denominator) for weight in weights[:-1]]
    return result + [total - sum(result, ZERO)]


def process_claro_files(pdf_bytes, inventory_bytes, fact_bytes=None, company=None):
    if not pdf_bytes or not inventory_bytes or not fact_bytes:
        raise ValueError('Debe cargar los tres archivos: Excel de Claro, Inventario General y recibo PDF.')
    if company not in ('PRODAC', 'PRODIMIN'):
        raise ValueError('Seleccione PRODAC o PRODIMIN antes de procesar.')
    timings = PhaseTimings()
    meta, pdf_general = timings.run('pdf_read_ms', read_pdf, pdf_bytes)
    if meta['company'] != company:
        raise ValueError(f'La empresa seleccionada es {company}, pero el PDF corresponde a {meta["company"]}.')
    records, general, preamble = timings.run('excel_read_ms', read_billing, fact_bytes)
    declared = set(v.upper() for v in re.findall(r'\b(prodac|prodimin)\b', clean(preamble)))
    if declared and declared != {company}:
        raise ValueError('La empresa del Excel de Claro no coincide con la empresa seleccionada y el PDF.')
    pdf_adjustment = sum((r['amount'] for r in pdf_general), ZERO)
    adjustment = sum((r['amount'] for r in general), ZERO)
    if abs(pdf_adjustment - adjustment) > CENT:
        raise ValueError('Los ajustes o Varios del PDF no coinciden con Factura Detallada. Revise la fila Varios y su importe sin IGV; la hoja original se conserva sin alteraciones.')
    for item in general:
        if item['otros_cargos'] is not None and abs(item['otros_cargos'] - item['amount']) > CENT:
            raise ValueError('La fila Varios tiene importes diferentes en Total sin IGV y Otros Cargos y Abonos.')
    timings.mark('reconciliation_ms')
    inventory = timings.run('excel_read_ms', read_inventory, inventory_bytes, company)
    groups = defaultdict(lambda: {'amount': ZERO, 'indices': []})
    matched, missing = set(), set()
    for index, item in enumerate(records):
        inv = inventory.get(item['numero'], {})
        item['cc'] = inv.get('cc') or 'SIN CC - REVISAR'
        item['usuario'] = inv.get('usuario') or 'No asignado'
        item['observacion'] = '' if inv.get('cc') else 'Sin centro de costo en inventario'
        (matched if inv.get('cc') else missing).add(item['numero'])
        groups[item['cc']]['amount'] += item['amount']
        groups[item['cc']]['indices'].append(index)
    timings.mark('inventory_match_ms')
    eligible = [cc for cc in sorted(groups) if cc != 'SIN CC - REVISAR' and groups[cc]['amount'] != ZERO]
    allocations = dict(zip(eligible, apportion(adjustment, [groups[cc]['amount'] for cc in eligible])))
    for cc, group in groups.items():
        allocated = allocations.get(cc, ZERO)
        portions = apportion(allocated, [records[i]['amount'] for i in group['indices']])
        for index, portion in zip(group['indices'], portions):
            item = records[index]
            item['prorrateo'] = portion
            item['final_amount'] = item['amount'] + portion
            item['costo'] = float(item['final_amount'])
    line_total = sum((r['amount'] for r in records), ZERO)
    total = sum((r['final_amount'] for r in records), ZERO)
    difference = total - meta['cargos_del_mes']
    warnings = ['Existen líneas sin centro de costo'] if missing else []
    if missing and adjustment:
        warnings.append('Varios se distribuye entre los CC identificados; las líneas sin CC conservan su monto original.')
    summary = [{'cc': cc, 'antes': group['amount'], 'prorrateo': allocations.get(cc, ZERO),
                'total': group['amount'] + allocations.get(cc, ZERO), 'lineas': len(group['indices'])}
               for cc, group in sorted(groups.items())]
    report = {**meta, 'numeros_unicos': len(matched | missing), 'registros_cargos': len(records),
              'registros_procesados': len(records) + len(general), 'lineas_con_cc': len(matched),
              'lineas_sin_cc': len(missing), 'suma_lineas': line_total, 'ajustes_abonos': adjustment,
              'total_calculado': total, 'diferencia': difference, 'cuadra': abs(difference) <= CENT,
              'estado': 'CUADRA' if abs(difference) <= CENT else 'NO CUADRA', 'warnings': warnings,
              'details': records, 'summary': summary, 'varios': general, 'pdf_adjustments': pdf_general}
    report['download_filename'] = f"Conciliacion_Claro_{company}_{meta['periodo'].replace(' ', '_')}.xlsx"
    timings.mark('reconciliation_ms')
    report['excel_bytes'] = timings.run('workbook_ms', generate_claro_excel, report, fact_bytes)
    report['timings'] = timings.snapshot()
    return json_values(report)
