"""Conciliación Movistar independiente: PDF oficial, inventario y Excel opcionales."""
import io
import re
import unicodedata
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import msoffcrypto
import openpyxl
import pypdf

from movistar_workbook import generate_movistar_excel
from processing_diagnostics import PhaseTimings

ZERO = Decimal('0.00')
TOLERANCE = Decimal('0.01')
AMOUNT = r'[+−-]?\s*\d[\d,]*\.\d{2}'
CURRENCY = r'S\s*/\.?\s*'
END_AMOUNT = re.compile(rf'(?:{CURRENCY})?({AMOUNT})\s*$', re.I)


def normalize(value):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFD', str(value or ''))
                           if not unicodedata.combining(c)).lower().split())


def clean_phone(value):
    if value is None:
        return ''
    if isinstance(value, (float, Decimal)) and value == int(value):
        value = int(value)
    return re.sub(r'\D', '', str(value))[-9:]


def to_decimal(value):
    if value is None or value == '':
        raise ValueError('No se puede interpretar un monto vacío en los archivos de Movistar.')
    clean = re.sub(r'S\s*/\.?|\$|\s', '', str(value), flags=re.I).replace(',', '').replace('−', '-')
    try:
        result = Decimal(clean)
        if not result.is_finite():
            raise InvalidOperation
        return result.quantize(TOLERANCE, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        raise ValueError('No se puede interpretar un monto en los archivos de Movistar.') from None


def decrypt_excel(file_bytes, password=''):
    """Los documentos descifrados y la contraseña sólo viven en memoria."""
    try:
        office = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        if not office.is_encrypted():
            return io.BytesIO(file_bytes)
        if not password:
            raise ValueError('Ingrese la contraseña mensual para validar los Excel protegidos.')
        office.load_key(password=password)
        stream = io.BytesIO()
        office.decrypt(stream)
        stream.seek(0)
        return stream
    except ValueError:
        raise
    except Exception:
        raise ValueError('No se pudo abrir el Excel opcional. Revise el archivo y su contraseña mensual.') from None


def _is_page_header(line):
    text = normalize(line)
    return (not text or text in ('empresas movil', 'importe s/', 'importe s/.')
            or text.startswith(('integratel peru', 'telefonica del peru'))
            or bool(re.fullmatch(r'pagina\s+\d+\s*/\s*\d+', text)))


def _discount_kind(description):
    text = normalize(description)
    duration = re.search(r'(?<!\d)(12|3|18)\s*m(?:eses|es)?\b', text)
    if duration and duration[1] == '12':
        return 'perm' if re.search(r'\bperm\w*\b', text) else '12m'
    if duration and duration[1] == '3':
        return '3m'
    return '18m' if duration and duration[1] == '18' else 'incorporado'


def extract_pdf_data(pdf_bytes):
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        raw = '\n'.join(page.extract_text() or '' for page in reader.pages)
    except Exception:
        raise ValueError('No se pudo leer el PDF. Cargue un recibo Movistar válido con texto seleccionable.') from None
    if not raw.strip():
        raise ValueError('El PDF está vacío o no tiene texto seleccionable.')
    if not any(word in normalize(raw) for word in ('movistar', 'integratel')):
        raise ValueError('El PDF no corresponde a un recibo de Movistar.')

    # Unir páginas antes de reconocer Anexos, sin posiciones ni cantidades fijas.
    text = '\n'.join(line.strip() for line in raw.splitlines() if not _is_page_header(line))
    headers = list(re.finditer(r'^Anexo\s+(\d+)\s*:\s*([^\n]*)', text, re.I | re.M))
    if not headers:
        raise ValueError('No se encontraron Anexos de líneas telefónicas en el PDF.')
    summary = text[:headers[0].start()]
    summary_normal = '\n'.join(normalize(line) for line in summary.splitlines())

    def amount(label, required=False, display_label=None):
        match = re.search(rf'^{label}\s*:?\s*{CURRENCY}({AMOUNT})\s*$', summary_normal, re.I | re.M)
        if match:
            return to_decimal(match[1])
        if required:
            raise ValueError(f'No se pudo identificar el importe de {display_label or label} en el resumen del PDF.')
        return None

    period = re.search(r'ciclo de facturacion\s*[-–:]\s*([a-z]+\s+\d{4})', normalize(summary))
    if not period:
        period = re.search(r'(?:recibo|periodo(?: de facturacion)?)\s*:?\s*([a-z]+\s+\d{4})', normalize(summary))
    if not period:
        raise ValueError('No se pudo identificar el mes y año del recibo Movistar.')
    receipt = re.search(r'n[°ºo]?\s*recibo\s*:\s*([\w-]+)', summary, re.I)
    if not receipt:
        receipt = re.search(r'resumen de recibo\s*-\s*n[°ºo]?\s*([\w-]+)', summary, re.I)
    subtotal = amount(r'subtotal(?: sin igv)?', True, 'Subtotal sin IGV')
    igv = amount(r'igv(?:\s*\(\s*18\s*%\s*\))?', True, 'IGV')
    total_afecto = amount('total afecto', True)
    inafectos_pdf = amount('cargos adicionales inafectos', True)
    total_facturado = amount('total facturado')
    total_pagar = amount(r'total a pagar(?: del mes)?')
    if total_facturado is None and total_pagar is None:
        raise ValueError('No se pudo extraer el Total facturado o Total a pagar del PDF.')
    redondeo = amount('redondeo')
    if redondeo is None:
        components = {}
        for match in re.finditer(rf'redondeo del mes (actual|anterior)\s*{CURRENCY}({AMOUNT})', normalize(text), re.I):
            components[match[1]] = to_decimal(match[2])
        if not components:
            raise ValueError('No se pudo identificar el redondeo neto del recibo PDF.')
        redondeo = sum(components.values(), ZERO)

    breakdown = {f'{prefix}_{kind}': (0 if prefix == 'count' else ZERO)
                 for kind in ('12m', '3m', 'perm', '18m') for prefix in ('count', 'monto')}
    parsed, duplicates, seen = [], [], set()
    section_re = re.compile(r'^(Cargos Mensuales|Descuentos y Bonificaciones Afectos|Cargos Adicionales Afectos)\s*:\s*(.*)', re.I)
    stop_re = re.compile(r'^(?:detalle\b|redondeo\b|cargos adicionales inafectos\b|descuentos y bonificaciones inafectos\b)', re.I)
    for i, header in enumerate(headers):
        phone = clean_phone(header[2])
        if len(phone) != 9:
            raise ValueError(f'No se puede identificar el teléfono del Anexo {header[1]}.')
        if phone in seen:
            duplicates.append(phone)
        seen.add(phone)
        body = text[header.end():headers[i + 1].start() if i + 1 < len(headers) else len(text)]
        concepts, section, pending = [], None, ''
        for line in body.splitlines():
            if _is_page_header(line):
                continue
            if stop_re.match(normalize(line)):
                if pending:
                    raise ValueError(f'No se puede interpretar un monto del Anexo {header[1]}.')
                section = None
                continue
            heading = section_re.match(line)
            if heading:
                if pending:
                    raise ValueError(f'No se puede interpretar un monto del Anexo {header[1]}.')
                section = normalize(heading[1])
                line = heading[2].strip()
            if section is None or not line.strip():
                continue
            pending = f'{pending} {line}'.strip()
            match = END_AMOUNT.search(pending)
            if not match:
                continue
            concept = pending[:match.start()].strip()
            if not concept:
                raise ValueError(f'Falta la descripción de un cargo del Anexo {header[1]}.')
            concepts.append((section, concept, to_decimal(match[1])))
            pending = ''
        if pending:
            raise ValueError(f'No se puede interpretar un monto del Anexo {header[1]}.')
        monthly = [(c, value) for s, c, value in concepts if s == 'cargos mensuales']
        if not monthly:
            raise ValueError(f'No se encontraron cargos mensuales del Anexo {header[1]}.')
        cargos = sum((v for _, v in monthly), ZERO)
        additional = sum((v for s, _, v in concepts if s == 'cargos adicionales afectos'), ZERO)
        desc, incorporated = ZERO, ZERO
        for section, concept, value in concepts:
            if section != 'descuentos y bonificaciones afectos':
                continue
            kind, value = _discount_kind(concept), abs(value)
            if kind in ('12m', '3m', 'perm'):
                desc += value
            else:
                incorporated += value
            if kind != 'incorporado':
                breakdown[f'count_{kind}'] += 1
                breakdown[f'monto_{kind}'] += value
        monto = cargos + additional - incorporated
        parsed.append({'anx': int(header[1]), 'phone': phone, 'plan': monthly[0][0],
                       'monto': monto, 'desc': desc, 'monto_f': monto - desc,
                       'cargos_mensuales': cargos, 'cargos_adic': additional,
                       'desc_18m': incorporated})

    # Sólo el bloque detallado, evitando las cuotas repetidas en el resumen inicial.
    inafectos = []
    for block in re.finditer(r'^Cargos Adicionales Inafectos\s*:\s*', text, re.M | re.I):
        body = text[block.end():]
        stop = re.search(r'^(?:Redondeo\s*:|Nota de Cr[eé]dito|Conceptos facturables|Lugares de pago|Anexo\s+\d+\s*:)', body, re.I | re.M)
        if stop:
            body = body[:stop.start()]
        item_headers = list(re.finditer(r'^(?:Cuota\s+\d+\s+de\s+\d+\s+)?financiamiento\s+(?:Equipo\b|(?:de\s+)?venta\b)', body, re.I | re.M))
        for index, start in enumerate(item_headers):
            end = item_headers[index + 1].start() if index + 1 < len(item_headers) else len(body)
            item = body[start.start():end]
            match = re.search(rf'{CURRENCY}({AMOUNT})(?!\d)', item, re.I)
            ph = re.search(r'TELEF\s*:\s*([\d+\s-]+)', item, re.I)
            phone = clean_phone(ph[1]) if ph else ''
            if len(phone) != 9:
                raise ValueError('No se puede identificar el teléfono de una cuota de financiamiento del PDF.')
            if not match:
                raise ValueError('No se puede interpretar el monto de una cuota de financiamiento del PDF.')
            description = ' '.join(item[:match.start()].split())
            normalized_desc = normalize(description)
            inafectos.append({'phone': phone, 'description': description, 'monto': to_decimal(match[1]),
                               'is_equipo': 'equipo' in normalized_desc,
                               'is_venta': 'venta' in normalized_desc})
    breakdown['total_desc_count'] = sum(breakdown[f'count_{k}'] for k in ('12m', '3m', 'perm'))
    breakdown['total_desc_monto'] = sum((breakdown[f'monto_{k}'] for k in ('12m', '3m', 'perm')), ZERO)
    declared = re.search(r'^Anexo\s*:\s*(\d+)\s*$', summary, re.M | re.I)
    warnings = []
    if declared and int(declared[1]) != len(parsed):
        warnings.append(f'El recibo declara {declared[1]} Anexos y se extrajeron {len(parsed)}. Faltan registros o hay duplicados.')
    if duplicates:
        warnings.append('Existen números duplicados en el PDF: ' + ', '.join(sorted(set(duplicates))))
    return {'periodo': period[1].title(), 'recibo_numero': receipt[1] if receipt else '',
            'cargos_mensuales_resumen': amount('cargos mensuales'),
            'descuentos_afectos_resumen': amount('descuentos y bonificaciones afectos'),
            'cargos_adic_resumen': amount('cargos adicionales afectos'),
            'descuentos_inafectos': amount('descuentos y bonificaciones inafectos') or ZERO,
            'subtotal': subtotal, 'igv': igv, 'total_afecto': total_afecto,
            'cargos_inafectos': inafectos_pdf, 'redondeo': redondeo,
            'total_facturado': total_facturado, 'total_pagar': total_pagar,
            'pdf_total': total_facturado if total_facturado is not None else total_pagar,
            'lines': parsed, 'inafectos': inafectos, 'duplicate_phones': duplicates,
            'integrity_warnings': warnings, 'discount_breakdown': breakdown,
            'inafectos_breakdown': {'count_equipos': sum(x['is_equipo'] for x in inafectos),
                                   'count_ventas': sum(x['is_venta'] for x in inafectos),
                                   'total_count': len(inafectos),
                                   'total_inafectos': sum((x['monto'] for x in inafectos), ZERO)}}


def load_inventory(inv_bytes, with_warnings=False):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(inv_bytes), data_only=True, read_only=True)
    except Exception:
        raise ValueError('No se pudo leer el Inventario Celulares General. Cargue un archivo .xlsx válido.') from None
    try:
        if 'PRODAC' not in wb.sheetnames:
            raise ValueError('No se encuentra la hoja PRODAC en el Inventario Celulares General.')
        rows = iter(wb['PRODAC'].iter_rows(values_only=True))
        columns = {normalize(value): i for i, value in enumerate(next(rows, ()))}
        if 'numero' not in columns or 'cc dpto' not in columns:
            raise ValueError('Faltan los campos Numero o CC Dpto en la hoja PRODAC del Inventario General.')
        inv, duplicates = {}, set()
        for row in rows:
            phone = clean_phone(row[columns['numero']])
            if len(phone) != 9:
                continue
            cc_value = row[columns['cc dpto']]
            cc = str(cc_value).strip() if cc_value is not None else ''
            if phone in inv:
                duplicates.add(phone)
                if inv[phone] != cc:
                    inv[phone] = ''
            else:
                inv[phone] = cc
        if not inv:
            raise ValueError('No se encontraron números válidos en la hoja PRODAC del Inventario General.')
        warnings = ['Números duplicados en el inventario; revise su CC: ' + ', '.join(sorted(duplicates))] if duplicates else []
        return (inv, warnings) if with_warnings else inv
    finally:
        wb.close()


def _plan_key(plan):
    plan = re.sub(r'\([^)]*\bal\b[^)]*\)', '', normalize(plan))
    return re.sub(r'[^a-z0-9]', '', plan)


def validate_optional_files(fact_bytes, traf_bytes, cargos_bytes, password, parsed_lines, inafectos):
    warnings = []
    pdf_by_phone = {line['phone']: line for line in parsed_lines}
    for label, data in [('Fact', fact_bytes), ('Traf', traf_bytes), ('Cargos', cargos_bytes)]:
        if not data:
            continue
        try:
            wb = openpyxl.load_workbook(decrypt_excel(data, password), data_only=True, read_only=True)
        except Exception:
            warnings.append(f'{label}: No se pudo abrir el archivo; revise el Excel y la contraseña mensual en Validación opcional.')
            continue
        try:
            # El portal puede declarar A1:A1 aunque el XML contenga todas las filas.
            wb.active.reset_dimensions()
            rows = list(wb.active.iter_rows(values_only=True))
            if label in ('Fact', 'Traf'):
                # Columnas del formato del portal, sin fijar las filas de cabecera.
                entries = defaultdict(list)
                for row in rows:
                    phone = clean_phone(row[1]) if len(row) > 1 else ''
                    if len(row) > 4 and row[2] and not str(row[2]).strip().isdigit() and not row[4]:
                        continue  # Fila de cuenta financiera, no una línea/anexo.
                    if len(phone) == 9:
                        entries[phone].append(row)
                for phone, line in pdf_by_phone.items():
                    if phone not in entries:
                        warnings.append(f'{label}: Número {phone} presente en PDF y ausente en {label}.')
                        continue
                    if label == 'Fact':
                        total = sum((to_decimal(row[19] or 0) + to_decimal(row[28] or 0)
                                     for row in entries[phone]), ZERO)
                        if abs(total - line['monto']) > TOLERANCE:
                            warnings.append(f'Fact: Monto diferente para {phone}: Excel S/ {total:,.2f}; PDF MONTO S/ {line["monto"]:,.2f}.')
                    else:
                        plan = entries[phone][0][4] if len(entries[phone][0]) > 4 else ''
                        if _plan_key(plan) != _plan_key(line['plan']):
                            warnings.append(f'Traf: Plan diferente entre PDF y Traf para {phone}. PDF: {line["plan"]}; Traf: {plan or "vacío"}.')
                for phone in sorted(entries.keys() - pdf_by_phone.keys()):
                    warnings.append(f'{label}: Número {phone} presente en {label} y ausente en PDF.')
                if not entries:
                    warnings.append(f'{label}: No se encontraron teléfonos válidos para comparar.')
            else:
                entries = Counter()
                for row in rows:
                    for col, cell in enumerate(row):
                        if not cell or not re.search(r'TELEF\s*:', str(cell), re.I):
                            continue
                        match = re.search(r'TELEF\s*:\s*([\d+\s-]+)', str(cell), re.I)
                        phone = clean_phone(match[1]) if match else ''
                        if len(phone) != 9 or col + 1 >= len(row) or row[col + 1] is None:
                            continue
                        entries[(phone, to_decimal(row[col + 1]))] += 1
                total_excel = sum((amount * count for (_, amount), count in entries.items()), ZERO)
                total_pdf = sum((x['monto'] for x in inafectos), ZERO)
                warnings.append(f'Cargos: Excel S/ {total_excel:,.2f}; PDF S/ {total_pdf:,.2f}; diferencia S/ {total_pdf - total_excel:,.2f}.')
                pdf_entries = Counter((x['phone'], x['monto']) for x in inafectos)
                for (phone, amount), count in (pdf_entries - entries).items():
                    warnings.append(f'Cargos: {count} cuota(s) de {phone} por S/ {amount:,.2f} presente(s) en PDF y ausente(s) en Cargos.')
                for (phone, amount), count in (entries - pdf_entries).items():
                    warnings.append(f'Cargos: {count} cuota(s) de {phone} por S/ {amount:,.2f} presente(s) en Cargos y ausente(s) en PDF.')
        except (ValueError, IndexError):
            warnings.append(f'{label}: No se pudo interpretar la estructura o un monto del Excel opcional.')
        finally:
            wb.close()
    return warnings


def process_movistar_files(*args, **kwargs):
    """Dos entradas obligatorias; conserva compatibilidad con la antigua llamada API."""
    keys = ['pdf_bytes', 'inventory_bytes', 'fact_bytes', 'traf_bytes', 'cargos_bytes', 'password']
    if len(args) >= 4 and not (isinstance(args[0], bytes) and args[0].startswith(b'%PDF')):
        keys = ['fact_bytes', 'traf_bytes', 'cargos_bytes', 'pdf_bytes', 'inventory_bytes', 'password']
    values = dict(zip(keys, args))
    values.update(kwargs)
    pdf_bytes, inventory_bytes = values.get('pdf_bytes'), values.get('inventory_bytes')
    if not pdf_bytes:
        raise ValueError('Debe cargar el recibo mensual de Movistar en PDF.')
    if not inventory_bytes:
        raise ValueError('Debe cargar el archivo Inventario Celulares General.')
    timings = PhaseTimings()
    pdf = timings.run('pdf_read_ms', extract_pdf_data, pdf_bytes)
    inv, inventory_warnings = timings.run('excel_read_ms', load_inventory, inventory_bytes, with_warnings=True)
    lines, inafectos = pdf['lines'], pdf['inafectos']
    sin_cc, cc_counts, cc_sums = [], Counter(), defaultdict(lambda: ZERO)
    for line in lines:
        cc = inv.get(line['phone']) or 'SIN CC'
        line['cc'] = cc
        if cc == 'SIN CC':
            sin_cc.append({'phone': line['phone'], 'plan': line['plan'], 'monto_f': float(line['monto_f']),
                           'obs': 'Sin CC en el Inventario General (hoja PRODAC); revisar asignación.'})
        cc_counts[cc] += 1
        cc_sums[cc] += line['monto_f']
    timings.mark('inventory_match_ms')
    totals = {key: sum((l[key] for l in lines), ZERO) for key in ('monto', 'desc', 'monto_f')}
    inaf_total = sum((i['monto'] for i in inafectos), ZERO)
    total_afecto = totals['monto_f'] + pdf['igv']
    calculated = total_afecto + inaf_total + pdf['redondeo']
    difference = calculated - pdf['pdf_total']
    sub_difference = totals['monto_f'] - pdf['subtotal']
    inaf_difference = inaf_total - pdf['cargos_inafectos']
    warnings = list(pdf['integrity_warnings'])
    if abs(sub_difference) > TOLERANCE:
        warnings.append(f'El subtotal calculado no coincide con el PDF: diferencia S/ {sub_difference:,.2f}.')
    if abs(inaf_difference) > TOLERANCE:
        warnings.append(f'Los cargos inafectos extraídos no coinciden con el PDF: diferencia S/ {inaf_difference:,.2f}.')
    if abs(total_afecto - pdf['total_afecto']) > TOLERANCE:
        warnings.append('El Total afecto calculado no coincide con el resumen del PDF.')
    if pdf['descuentos_inafectos']:
        warnings.append('El recibo contiene descuentos inafectos; requiere revisión de la conciliación.')
    if abs(difference) > TOLERANCE:
        warnings.append(f'El total calculado no coincide con el Total facturado del PDF: diferencia S/ {difference:,.2f}.')
    cuadra = not warnings
    display_warnings = warnings + inventory_warnings
    if sin_cc:
        display_warnings.append(f'{len(sin_cc)} líneas sin centro de costo; sus importes permanecen incluidos en todos los totales.')
    timings.mark('reconciliation_ms')
    optional = timings.run('excel_read_ms', validate_optional_files, values.get('fact_bytes'), values.get('traf_bytes'),
                                       values.get('cargos_bytes'), values.get('password') or '', lines, inafectos)
    result = {'periodo': pdf['periodo'], 'recibo_numero': pdf['recibo_numero'],
              'line_count': len(lines), 'cargo_mensual_sin_igv': float(totals['monto']),
              'descuento_total': float(totals['desc']), 'cargo_final_sin_igv': float(totals['monto_f']),
              'subtotal_pdf': float(pdf['subtotal']), 'igv': float(pdf['igv']),
              'total_afecto': float(total_afecto), 'cargos_inafectos': float(inaf_total),
              'cargos_inafectos_pdf': float(pdf['cargos_inafectos']), 'redondeo': float(pdf['redondeo']),
              'total_facturado': float(pdf['total_facturado']) if pdf['total_facturado'] is not None else None,
              'total_pagar': float(pdf['total_pagar']) if pdf['total_pagar'] is not None else None,
              'descuentos_inafectos': float(pdf['descuentos_inafectos']),
              'pdf_total': float(pdf['pdf_total']), 'total_calculado': float(calculated),
              'diferencia': float(difference), 'diferencia_subtotal': float(sub_difference),
              'diferencia_inafectos': float(inaf_difference), 'cuadra': cuadra,
              'estado': 'CUADRA' if cuadra else 'NO CUADRA',
              'discount_count': pdf['discount_breakdown']['total_desc_count'],
              'discount_breakdown': {k: float(v) if isinstance(v, Decimal) else v
                                     for k, v in pdf['discount_breakdown'].items()},
              'inafectos_breakdown': {k: float(v) if isinstance(v, Decimal) else v
                                      for k, v in pdf['inafectos_breakdown'].items()},
              'sin_cc_count': len(sin_cc), 'sin_cc_lines': sin_cc,
              'resumen_cc': [{'cc': cc, 'lineas': cc_counts[cc], 'monto': float(amount),
                              'pct': round(float(amount / totals['monto_f'] * 100), 2) if totals['monto_f'] else 0}
                             for cc, amount in sorted(cc_sums.items())],
              'lines_sample': [{k: float(v) if isinstance(v, Decimal) else v for k, v in l.items()} for l in lines[:50]],
              'warnings': display_warnings, 'optional_warnings': optional,
              'resumen_pdf': {key: float(value) if isinstance(value, Decimal) else value
                              for key, value in pdf.items() if key in (
                                  'cargos_mensuales_resumen', 'descuentos_afectos_resumen',
                                  'cargos_adic_resumen', 'subtotal', 'igv', 'total_afecto',
                                  'descuentos_inafectos', 'cargos_inafectos', 'redondeo',
                                  'total_facturado', 'total_pagar')},
              }
    timings.mark('reconciliation_ms')
    result['excel_bytes'] = timings.run('workbook_ms', generate_movistar_excel, lines, inafectos, inv, pdf)
    result['timings'] = timings.snapshot()
    result.update(cargo_mensual_final_sin_igv=result['cargo_final_sin_igv'], total_pdf=result['pdf_total'],
                  cargo_mensual_con_igv=result['total_afecto'], equipos_pdf=result['cargos_inafectos_pdf'],
                  equipos=result['cargos_inafectos'], cc_summary=result['resumen_cc'],
                  disc_count_12m=pdf['discount_breakdown']['count_12m'])
    return result
