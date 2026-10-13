"""
movistar_engine.py
Motor independiente para la conciliacion mensual de Movistar - PRODAC.
Extrae descuentos del PDF (sin lista maestra separada).
Version 2.0 - Procesamiento dinamico completo.
"""

import io
import re
import os
import json
import uuid
import datetime
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
import msoffcrypto
import pypdf


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def clean_phone(val):
    """Normaliza un numero telefonico: extrae los ultimos 9 digitos."""
    if val is None:
        return ""
    digits = re.sub(r'\D', '', str(val).strip())
    if len(digits) >= 9:
        return digits[-9:]
    return digits


def to_decimal(val):
    """Convierte un valor a Decimal de forma segura."""
    try:
        return Decimal(str(val or 0)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal('0.00')


# ---------------------------------------------------------------------------
# Descifrado de Excel protegido
# ---------------------------------------------------------------------------

def decrypt_excel(file_bytes, password):
    """
    Descifra un archivo Excel protegido en memoria.
    Lanza ValueError con mensaje claro si la password es incorrecta.
    """
    if not password or not str(password).strip():
        raise ValueError("Debe ingresar la contrasena de los archivos del mes.")

    try:
        office = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        office.load_key(password=str(password).strip())
        dec = io.BytesIO()
        office.decrypt(dec)
        dec.seek(0)
        return dec
    except Exception as e:
        msg = str(e).lower()
        # Intentar si el archivo ya no esta cifrado
        try:
            openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            return io.BytesIO(file_bytes)
        except Exception:
            raise ValueError(
                "La contrasena ingresada no permite abrir los archivos de Movistar."
            )


# ---------------------------------------------------------------------------
# Extraccion del PDF
# ---------------------------------------------------------------------------

def extract_pdf_data(pdf_bytes):
    """
    Extrae del recibo PDF de Movistar:
    - Periodo de facturacion
    - Totales del resumen (subtotal, IGV, equipos, redondeo, total)
    - Descuentos de 12 meses por numero de anexo
    """
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    if not reader.pages:
        raise ValueError("El PDF del recibo esta vacio.")

    # ---------- Texto completo ----------
    all_text = "\n".join(
        (page.extract_text() or "") for page in reader.pages
    )

    # ---------- Periodo ----------
    periodo = "Periodo no detectado"
    pm = re.search(
        r'Ciclo de facturaci[oó]n\s*[-\u2013]\s*([A-Za-záéíóúÁÉÍÓÚñÑ]+\s+\d{4})',
        all_text, re.IGNORECASE
    )
    if not pm:
        pm = re.search(r'Recibo\s+([A-Za-záéíóúÁÉÍÓÚñÑ]+\s+\d{4})', all_text, re.IGNORECASE)
    if not pm:
        pm = re.search(r'Recibo\s+([A-Za-záéíóúÁÉÍÓÚñÑ]+)', all_text, re.IGNORECASE)
    if pm:
        periodo = pm.group(1).strip()

    def find_amount(pattern):
        m = re.search(pattern, all_text, re.IGNORECASE)
        if m:
            return to_decimal(m.group(1).replace(',', ''))
        return None

    subtotal     = find_amount(r'Subtotal\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')
    igv          = find_amount(r'IGV\s*\(\s*18%\s*\)\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')
    total_afecto = find_amount(r'Total Afecto\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')
    equipos_pdf  = find_amount(r'Cargos Adicionales Inafectos\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')
    redondeo     = find_amount(r'Redondeo\s+S/?[/ ]*(-?[0-9,]+\.[0-9]{2})')
    total_fact   = find_amount(r'Total Facturado\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')
    total_pagar  = find_amount(r'Total a pagar\s+S/?[/ ]*([0-9,]+\.[0-9]{2})')

    pdf_total = total_fact if total_fact is not None else total_pagar
    if pdf_total is None:
        raise ValueError(
            "No se pudo extraer el Total del recibo PDF. "
            "Verifique que sea un recibo Movistar original."
        )

    # ---------- Descuentos de 12 meses por Anexo ----------
    # Partir el texto completo en bloques por "Anexo N: PHONE"
    disc_map = {}   # phone -> Decimal monto_descuento
    disc_count_12m = 0

    parts = re.split(r'(?=Anexo\s+\d+\s*:\s*\d+)', all_text)

    DISC_12M_PATTERN = re.compile(
        r'(?:Dscto[^\n]*?50%[^\n]*?12\s*m(?:eses)?|'
        r'Descuento[^\n]*?50%[^\n]*?12\s*m(?:eses)?)',
        re.IGNORECASE
    )
    # Exclusion: 18 meses
    DISC_18M_PATTERN = re.compile(r'18\s*m(?:eses)?', re.IGNORECASE)

    for part in parts:
        anx_m = re.match(r'Anexo\s+(\d+)\s*:\s*(\d+)', part)
        if not anx_m:
            continue
        phone = clean_phone(anx_m.group(2))
        if not phone:
            continue

        # Buscar cada linea de descuento de 12m en este bloque
        phone_disc = Decimal('0.00')
        for line in part.split('\n'):
            if DISC_12M_PATTERN.search(line) and not DISC_18M_PATTERN.search(line):
                # Extraer importe negativo en la misma linea o siguiente
                amt_m = re.search(r'S/\s*-\s*([\d,]+\.?\d*)', line)
                if amt_m:
                    phone_disc += to_decimal(amt_m.group(1).replace(',', ''))

        if phone_disc > Decimal('0.00'):
            # Si el mismo telefono aparece en varios Anexos, acumular
            disc_map[phone] = disc_map.get(phone, Decimal('0.00')) + phone_disc
            disc_count_12m += 1

    return {
        'periodo': periodo,
        'subtotal': subtotal,
        'igv': igv,
        'total_afecto': total_afecto,
        'equipos_pdf': equipos_pdf or Decimal('0.00'),
        'redondeo': redondeo if redondeo is not None else Decimal('0.00'),
        'pdf_total': pdf_total,
        'disc_map': disc_map,          # {phone: Decimal}
        'disc_count_12m': disc_count_12m,
    }


# ---------------------------------------------------------------------------
# Carga del inventario
# ---------------------------------------------------------------------------

def load_inventory(inv_bytes):
    """
    Carga el Inventario Celulares General desde bytes.
    Busca la hoja PRODAC y mapea Numero -> CC Dpto.
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(inv_bytes), data_only=True)
    except Exception as e:
        raise ValueError(f"No se pudo leer el Inventario Celulares General: {e}")

    sheet_name = 'PRODAC' if 'PRODAC' in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    # Detectar columnas de numero y CC Dpto
    col_num = 3   # default C
    col_cc  = 8   # default H
    for c in range(1, min(25, ws.max_column + 1)):
        h = str(ws.cell(1, c).value or '').strip().upper()
        if h in ('NUMERO', 'NÚMERO', 'N\xdaMERO'):
            col_num = c
        elif 'CC DPTO' in h or 'CC_DPTO' in h or 'CENTRO DE COSTO' in h:
            col_cc = c

    inv_map = {}
    for r in range(2, ws.max_row + 1):
        phone = clean_phone(ws.cell(r, col_num).value)
        cc    = str(ws.cell(r, col_cc).value or '').strip()
        if phone:
            inv_map[phone] = cc

    if not inv_map:
        raise ValueError(
            "No se encontraron lineas en el Inventario Celulares General. "
            "Verifique que la hoja PRODAC tenga datos."
        )
    return inv_map


# ---------------------------------------------------------------------------
# Motor principal
# ---------------------------------------------------------------------------

def process_movistar_files(
    fact_bytes, traf_bytes, cargos_bytes, pdf_bytes,
    inventory_bytes, password
):
    """
    Flujo principal de conciliacion mensual Movistar.
    Recibe todos los archivos como bytes.
    Retorna dict con resultados y excel_bytes.
    """

    # 1. Descifrar los tres Excel protegidos
    fact_dec   = decrypt_excel(fact_bytes,   password)
    traf_dec   = decrypt_excel(traf_bytes,   password)
    cargos_dec = decrypt_excel(cargos_bytes, password)

    fact_wb   = openpyxl.load_workbook(fact_dec,   data_only=True)
    traf_wb   = openpyxl.load_workbook(traf_dec,   data_only=True)
    cargos_wb = openpyxl.load_workbook(cargos_dec, data_only=True)

    fact_ws   = fact_wb.active
    traf_ws   = traf_wb.active
    cargos_ws = cargos_wb.active

    # 2. Extraer datos del PDF (periodo + totales + descuentos)
    pdf_info = extract_pdf_data(pdf_bytes)

    # 3. Cargar inventario
    inv_map = load_inventory(inventory_bytes)

    # 4. Procesar Fact: detectar fila de encabezado y extraer lineas
    #    Buscar la fila donde la col B sea "TELEFONO" o donde empiecen los datos
    data_start_row = 12  # valor por defecto
    for r in range(1, 20):
        b_val = str(fact_ws.cell(r, 2).value or '').upper().strip()
        if 'TELEFONO' in b_val:
            data_start_row = r + 4  # saltar filas de sub-encabezados
            break
        # Tambien aceptar cuando la columna B tiene un numero de 9 digitos
        phone_test = clean_phone(fact_ws.cell(r, 2).value)
        if phone_test and len(phone_test) >= 8:
            data_start_row = r
            break

    fact_lines = []
    seen_phones = set()
    dup_phones  = []

    for r in range(data_start_row, fact_ws.max_row + 1):
        phone_raw = fact_ws.cell(r, 2).value
        phone = clean_phone(phone_raw)
        if not phone or len(phone) < 8:
            continue
        # Descartar si es el mismo numero de cuenta financiera (sin empezar con 9)
        if not phone.startswith('9') and not phone.startswith('94') and not phone.startswith('93') and not phone.startswith('96') and not phone.startswith('97') and not phone.startswith('98'):
            # Puede ser un numero valido que no empieza con 9 (ej. 940xxx)
            # Solo descartar si coincide con ID de cuenta de factura
            account_cell = str(fact_ws.cell(3, 3).value or '')
            if clean_phone(account_cell) == phone and len(account_cell) > 5:
                continue

        if phone in seen_phones:
            dup_phones.append(phone)
        seen_phones.add(phone)

        t_val  = to_decimal(fact_ws.cell(r, 20).value)
        ac_val = to_decimal(fact_ws.cell(r, 29).value)
        monto  = t_val + ac_val

        fact_lines.append({
            'phone': phone,
            'monto': monto,
        })

    if not fact_lines:
        raise ValueError(
            "No se encontraron lineas telefonicas validas en el archivo Fact. "
            "Verifique que sea el archivo correcto de Movistar."
        )

    # 5. Procesar Traf: telefono -> plan
    traf_map = {}
    for r in range(1, traf_ws.max_row + 1):
        p = clean_phone(traf_ws.cell(r, 2).value)
        plan = str(traf_ws.cell(r, 5).value or '').strip()
        if p and plan and plan.upper() not in ('PLAN', 'PLANES', 'E', ''):
            traf_map[p] = plan

    # 6. Procesar Cargos: extraer cuotas de equipos
    equipment_rows = []
    for r in range(1, cargos_ws.max_row + 1):
        for c in range(1, min(cargos_ws.max_column + 1, 10)):
            desc = str(cargos_ws.cell(r, c).value or '').strip()
            if 'TELEF:' in desc.upper() and ('CUOTA' in desc.upper() or 'FINANCIAMIENTO' in desc.upper() or 'EQUIPO' in desc.upper()):
                # Extraer numero de telefono
                tel_m = re.search(r'TELEF:\s*(\d+)', desc, re.IGNORECASE)
                if not tel_m:
                    continue
                phone = clean_phone(tel_m.group(1))
                # Buscar monto en las columnas siguientes
                monto_eq = Decimal('0.00')
                for mc in range(c + 1, min(c + 5, cargos_ws.max_column + 1)):
                    mv = cargos_ws.cell(r, mc).value
                    if mv is not None:
                        try:
                            monto_eq = to_decimal(mv)
                            break
                        except Exception:
                            pass
                # Si el monto esta en formato S/xxx en la misma celda
                if monto_eq == Decimal('0.00'):
                    amt_m = re.search(r'S/\s*([\d,]+\.?\d*)', desc)
                    if amt_m:
                        monto_eq = to_decimal(amt_m.group(1).replace(',', ''))

                if monto_eq > Decimal('0.00'):
                    equipment_rows.append({
                        'phone': phone,
                        'desc': desc,
                        'monto': monto_eq,
                    })
                break  # una vez encontrado en esta fila, no buscar mas columnas

    # 7. Cruzar datos y calcular por linea
    disc_map   = pdf_info['disc_map']
    processed  = []
    total_monto  = Decimal('0')
    total_desc   = Decimal('0')
    total_monto_f = Decimal('0')
    sin_cc_lines  = []
    sin_plan_lines = []
    cc_warnings   = []

    for line in fact_lines:
        phone = line['phone']
        monto = line['monto']
        plan  = traf_map.get(phone, "")

        if not plan:
            sin_plan_lines.append(phone)

        # Descuento de 12 meses desde el PDF
        desc = disc_map.get(phone, Decimal('0.00'))

        monto_f = monto - desc

        # Centro de costo del inventario
        cc = inv_map.get(phone, "")
        if not cc:
            sin_cc_lines.append({
                'phone': phone,
                'plan': plan,
                'monto_f': float(monto_f),
            })
        elif len(cc) > 20 or any(kw in cc.lower() for kw in ['incl', 'igv', 'iva']):
            cc_warnings.append(
                f"Linea {phone} tiene CC con texto particular: '{cc}'"
            )

        total_monto   += monto
        total_desc    += desc
        total_monto_f += monto_f

        processed.append({
            'phone':   phone,
            'plan':    plan,
            'monto':   monto,
            'desc':    desc if desc > 0 else None,
            'monto_f': monto_f,
            'cc':      cc,
        })

    # Resumen por CC
    cc_stats = {}
    for item in processed:
        cc_name = item['cc'] or 'Sin CC'
        if cc_name not in cc_stats:
            cc_stats[cc_name] = {'count': 0, 'monto': Decimal('0')}
        cc_stats[cc_name]['count']  += 1
        cc_stats[cc_name]['monto']  += item['monto_f']

    cc_summary = []
    for cc_name, st in cc_stats.items():
        m = st['monto']
        pct = float(
            (m / total_monto_f * 100).quantize(Decimal('0.1'))
        ) if total_monto_f > 0 else 0.0
        cc_summary.append({
            'cc': cc_name, 'count': st['count'],
            'monto': float(m), 'pct': pct,
        })
    cc_summary.sort(key=lambda x: x['monto'], reverse=True)

    # 8. Totales generales y validacion contra el PDF
    total_equipos  = sum(r['monto'] for r in equipment_rows)
    calc_igv       = (total_monto_f * Decimal('0.18')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    calc_afecto    = total_monto_f + calc_igv
    pdf_redondeo   = pdf_info['redondeo']
    pdf_total      = pdf_info['pdf_total']

    total_calculado = calc_afecto + total_equipos + pdf_redondeo
    diferencia      = total_calculado - pdf_total
    cuadra          = abs(diferencia) <= Decimal('0.01')

    causas = []
    if not cuadra:
        if pdf_info['equipos_pdf'] and abs(total_equipos - pdf_info['equipos_pdf']) > Decimal('0.01'):
            causas.append(
                f"Equipos calculado S/{total_equipos:.2f} vs PDF S/{pdf_info['equipos_pdf']:.2f}"
            )
        if pdf_info['subtotal'] and abs(total_monto_f - pdf_info['subtotal']) > Decimal('0.01'):
            causas.append(
                f"Subtotal afecto calculado S/{total_monto_f:.2f} vs PDF S/{pdf_info['subtotal']:.2f}"
            )
        if not causas:
            causas.append("Diferencia general entre el calculo y el total del recibo.")

    # 9. Generar Excel
    excel_wb = generate_excel(
        processed, equipment_rows, pdf_info['subtotal'], total_monto_f
    )
    out = io.BytesIO()
    excel_wb.save(out)
    out.seek(0)

    del fact_wb, traf_wb, cargos_wb

    return {
        'periodo':                    pdf_info['periodo'],
        'line_count':                 len(processed),
        'anexos_count':               len(fact_lines),
        'disc_count_12m':             pdf_info['disc_count_12m'],
        'discount_count':             sum(1 for p in processed if p['desc']),
        'cargo_mensual_sin_igv':      float(total_monto),
        'descuento_total':            float(total_desc),
        'cargo_mensual_final_sin_igv':float(total_monto_f),
        'igv':                        float(calc_igv),
        'cargo_mensual_con_igv':      float(calc_afecto),
        'equipos':                    float(total_equipos),
        'redondeo':                   float(pdf_redondeo),
        'total_calculado':            float(total_calculado),
        'total_pdf':                  float(pdf_total),
        'diferencia':                 float(diferencia),
        'cuadra':                     cuadra,
        'causas_no_cuadra':           causas,
        'sin_cc_lines':               sin_cc_lines,
        'sin_cc_count':               len(sin_cc_lines),
        'sin_plan_lines':             sin_plan_lines,
        'sin_plan_count':             len(sin_plan_lines),
        'dup_phones':                 dup_phones,
        'cc_warnings':                cc_warnings,
        'cc_summary':                 cc_summary,
        'excel_bytes':                out.getvalue(),
    }


# ---------------------------------------------------------------------------
# Generacion del Excel de salida
# ---------------------------------------------------------------------------

def generate_excel(processed_lines, equipment_rows, pdf_subtotal, total_monto_f):
    """
    Genera el Excel conciliado con tres hojas en el orden correcto:
    1. CARGOS ADICIONALES
    2. CARGO MENSUAL
    3. Hoja1
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    FMT_CURR   = '#,##0.00'
    FMT_NUM    = '@'  # texto para telefonos
    F_ARIAL    = Font(name='Arial', size=10)
    F_HEADER   = Font(name='Arial', size=10, bold=True)
    FILL_NONE  = None

    def hdr(ws, row, col, value):
        c = ws.cell(row, col, value)
        c.font = F_HEADER
        return c

    def dat(ws, row, col, value, is_curr=False, is_phone=False):
        c = ws.cell(row, col, value)
        c.font = F_ARIAL
        if is_curr:
            c.number_format = FMT_CURR
        if is_phone:
            c.number_format = FMT_NUM
            c.value = str(value) if value is not None else ''
        return c

    # ================================================================
    # HOJA 1: CARGOS ADICIONALES
    # ================================================================
    ws1 = wb.create_sheet('CARGOS ADICIONALES')
    ws1.column_dimensions['A'].width = 5
    ws1.column_dimensions['B'].width = 16
    ws1.column_dimensions['C'].width = 85
    ws1.column_dimensions['D'].width = 16
    ws1.column_dimensions['E'].width = 14
    ws1.freeze_panes = 'B3'

    hdr(ws1, 2, 2, 'NUMERO')
    hdr(ws1, 2, 3, 'DESCRIPCION')
    hdr(ws1, 2, 4, 'MONTO')

    row = 3
    # Los datos del archivo Cargos ya contienen las cuotas de equipos.
    # La fila del descuento general (NN) la ponemos primero si existe.
    for eq in equipment_rows:
        dat(ws1, row, 2, str(eq['phone']), is_phone=True)
        dat(ws1, row, 3, eq['desc'])
        dat(ws1, row, 4, float(eq['monto']), is_curr=True)
        row += 1

    last_data = row - 1
    tot_cell = ws1.cell(row, 4, f'=SUM(D3:D{last_data})')
    tot_cell.font = F_HEADER
    tot_cell.number_format = FMT_CURR

    # ================================================================
    # HOJA 2: CARGO MENSUAL
    # ================================================================
    ws2 = wb.create_sheet('CARGO MENSUAL')
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 52
    ws2.column_dimensions['D'].width = 14
    ws2.column_dimensions['E'].width = 14
    ws2.column_dimensions['F'].width = 14
    ws2.freeze_panes = 'B3'

    hdr(ws2, 2, 2, 'NUMERO')
    hdr(ws2, 2, 3, 'PLAN')
    hdr(ws2, 2, 4, 'MONTO')
    hdr(ws2, 2, 5, 'DESC')
    hdr(ws2, 2, 6, 'MONTO F')

    m_row = 3
    idx   = 1
    for item in processed_lines:
        ws2.cell(m_row, 1, idx).font = F_ARIAL
        dat(ws2, m_row, 2, str(item['phone']), is_phone=True)
        dat(ws2, m_row, 3, item['plan'])
        dat(ws2, m_row, 4, float(item['monto']), is_curr=True)

        # DESC: solo poner valor si hay descuento; sino celda vacia
        if item['desc'] is not None and item['desc'] > 0:
            c_desc = ws2.cell(m_row, 5, float(item['desc']))
            c_desc.font = F_ARIAL
            c_desc.number_format = FMT_CURR
        else:
            ws2.cell(m_row, 5, None).number_format = FMT_CURR

        # MONTO F: formula
        c_mf = ws2.cell(m_row, 6, f'=D{m_row}-E{m_row}')
        c_mf.font = F_ARIAL
        c_mf.number_format = FMT_CURR

        m_row += 1
        idx   += 1

    last_m = m_row - 1

    # Fila total
    def total_cell(ws, row, col, formula):
        c = ws.cell(row, col, formula)
        c.font = F_HEADER
        c.number_format = FMT_CURR
        return c

    total_cell(ws2, m_row, 4, f'=SUM(D3:D{last_m})')
    total_cell(ws2, m_row, 5, f'=SUM(E3:E{last_m})')
    total_cell(ws2, m_row, 6, f'=SUM(F3:F{last_m})')

    # Filas de verificacion: subtotal del PDF y diferencia
    row_pdf = m_row + 1
    row_dif = m_row + 2
    if pdf_subtotal is not None:
        c_pdf = ws2.cell(row_pdf, 4, float(pdf_subtotal))
        c_pdf.font = F_HEADER
        c_pdf.number_format = FMT_CURR
        c_dif = ws2.cell(row_dif, 4, f'=F{m_row}-D{row_pdf}')
        c_dif.font = F_HEADER
        c_dif.number_format = FMT_CURR

    # ================================================================
    # HOJA 3: Hoja1
    # ================================================================
    ws3 = wb.create_sheet('Hoja1')
    ws3.column_dimensions['A'].width = 16
    ws3.column_dimensions['B'].width = 14
    ws3.column_dimensions['C'].width = 24
    ws3.column_dimensions['G'].width = 24
    ws3.column_dimensions['H'].width = 16
    ws3.column_dimensions['I'].width = 14
    ws3.column_dimensions['J'].width = 16
    ws3.freeze_panes = 'A2'

    # Encabezados en fila 1
    hdr(ws3, 1, 1, 'NUMERO')
    hdr(ws3, 1, 2, 'MONTO')
    hdr(ws3, 1, 3, 'CC')

    # Ordenar por CC luego por telefono
    sorted_lines = sorted(
        processed_lines,
        key=lambda x: (x['cc'] or 'ZZZZ', x['phone'])
    )

    h_row   = 2
    cc_totals = {}   # cc -> Decimal

    for item in sorted_lines:
        dat(ws3, h_row, 1, str(item['phone']), is_phone=True)
        c_m = ws3.cell(h_row, 2, float(item['monto_f']))
        c_m.font = F_ARIAL
        c_m.number_format = FMT_CURR
        dat(ws3, h_row, 3, item['cc'] or '')

        cc_name = item['cc'] or 'Sin CC'
        cc_totals[cc_name] = cc_totals.get(cc_name, Decimal('0')) + item['monto_f']

        h_row += 1

    last_h = h_row - 1
    tot_h = ws3.cell(h_row, 2, f'=SUM(B2:B{last_h})')
    tot_h.font = F_HEADER
    tot_h.number_format = FMT_CURR

    # Tabla de resumen por CC (columnas G, H, I, J) - encabezado en fila 4
    hdr(ws3, 4, 7, 'Etiquetas de fila')
    hdr(ws3, 4, 8, 'Suma de MONTO')
    hdr(ws3, 4, 9, 'IGV')
    hdr(ws3, 4, 10, 'TOTAL')

    s_row = 5
    for cc_name, cc_sum in cc_totals.items():
        ws3.cell(s_row, 7, cc_name).font = F_ARIAL
        c_sm = ws3.cell(s_row, 8, float(cc_sum))
        c_sm.font = F_ARIAL
        c_sm.number_format = FMT_CURR
        c_igv = ws3.cell(s_row, 9, f'=H{s_row}*0.18')
        c_igv.font = F_ARIAL
        c_igv.number_format = FMT_CURR
        c_tot = ws3.cell(s_row, 10, f'=H{s_row}+I{s_row}')
        c_tot.font = F_ARIAL
        c_tot.number_format = FMT_CURR
        s_row += 1

    last_s = s_row - 1
    ws3.cell(s_row, 7, 'Total general').font = F_HEADER
    c_tg_h = ws3.cell(s_row, 8, f'=SUM(H5:H{last_s})')
    c_tg_h.font = F_HEADER
    c_tg_h.number_format = FMT_CURR
    c_tg_j = ws3.cell(s_row, 10, f'=SUM(J5:J{last_s})')
    c_tg_j.font = F_HEADER
    c_tg_j.number_format = FMT_CURR

    return wb
