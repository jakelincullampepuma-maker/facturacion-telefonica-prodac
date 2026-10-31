"""
entel_engine.py
Motor dinamico e independiente para la conciliacion mensual de Entel - PRODAC / PRODIMIN.
Caracteristicas:
- Deteccion automatica de hojas, columnas y filas de datos del Excel del proveedor.
- Extraccion del concepto PLAN del PDF (independiente de equipos y total a pagar).
- Cruce con Inventario Celulares General (hojas PRODAC, PRODIMIN) filtrando operador Entel.
- Asignacion de CC Dpto como centro de costo oficial.
- Comparacion contra Hoja1 previa (si existe) para alertar cambios de CC.
- Generacion de hoja 'Resultado Entel' en posicion 0 conservando el 100% del Excel original.
- Tabla resumen por CC Dpto con formulas SUMIF y COUNTIF que se actualizan automaticamente.
- 100% dinamico sin valores ni telefonos ni importes hardcodeados.
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
import pypdf


# ---------------------------------------------------------------------------
# Utilidades de normalizacion
# ---------------------------------------------------------------------------

def clean_phone(val):
    """
    Normaliza un numero telefonico extrayendo los ultimos 9 digitos
    o el numero entero si tiene 9 digitos.
    """
    if val is None:
        return ""
    # Si viene como float/int, convertir a str sin decimales
    if isinstance(val, (int, float)):
        val_str = f"{int(val)}"
    else:
        val_str = str(val).strip()

    digits = re.sub(r'\D', '', val_str)
    if len(digits) >= 9:
        return digits[-9:]
    return digits


def to_decimal(val):
    """Convierte un valor a Decimal con 2 decimales redondeado comercialmente."""
    if val is None:
        return Decimal('0.00')
    if isinstance(val, (int, float)):
        return Decimal(str(val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    # Si es string, limpiar caracteres como 'S/', '$', comas, espacios
    s = str(val).strip().replace('S/', '').replace('$', '').replace(' ', '')
    # Manejar formato peruano/europeo de miles y comas
    if ',' in s and '.' in s:
        # e.g. 1,234.56 -> quitar coma
        s = s.replace(',', '')
    elif ',' in s and '.' not in s:
        # e.g. 1478,17 -> cambiar a punto
        s = s.replace(',', '.')

    try:
        return Decimal(s).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal('0.00')


# ---------------------------------------------------------------------------
# Extraccion del PDF de Entel
# ---------------------------------------------------------------------------

def extract_pdf_entel(pdf_bytes):
    """
    Extrae del recibo oficial de Entel:
    - Empresa facturada
    - Recibo No
    - Periodo de facturacion (emision, inicio y fin)
    - Cuenta y RUC
    - Importe del concepto PLAN (nunca Total a pagar)
    - Importe de Equipos
    - Total a pagar de la factura
    """
    if not pdf_bytes or len(pdf_bytes.strip()) == 0:
        raise ValueError("El PDF del recibo de Entel esta vacio.")

    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        if not reader.pages:
            raise ValueError("El PDF del recibo de Entel esta vacio.")
    except Exception as e:
        raise ValueError(f"Error al leer el archivo PDF de Entel: {e}")

    p1_text = reader.pages[0].extract_text() or ""
    all_text = "\n".join((page.extract_text() or "") for page in reader.pages)

    # 1. Empresa
    empresa = "PRODUCTOS DE ACERO CASSADO S.A. - PRODAC"
    emp_m = re.search(r'(PRODUCTOS DE ACERO CASSADO[^\n]*|PRODIMIN[^\n]*)', p1_text, re.IGNORECASE)
    if emp_m:
        empresa = emp_m.group(1).strip()

    # 2. Recibo No
    recibo_nro = "No detectado"
    # Formato Entel: Serie-Correlativo
    rec_m = re.search(r'([A-Z0-9]{3,5}-\d{6,})', p1_text)
    if rec_m:
        recibo_nro = rec_m.group(1).strip()
    else:
        # Buscar por texto Recibo
        rec_m2 = re.search(r'Recibo\s*N[°ºo\.]*\s*[:\s]*\n*([A-Za-z0-9\-]+)', p1_text, re.IGNORECASE)
        if rec_m2:
            recibo_nro = rec_m2.group(1).strip()

    # 3. Fechas y Periodo
    dates_found = re.findall(r'\b\d{1,2}/[A-Za-z]{3}/\d{4}\b', p1_text)
    emision = dates_found[0] if len(dates_found) > 0 else ""
    inicio_periodo = dates_found[1] if len(dates_found) > 1 else ""
    fin_periodo = dates_found[2] if len(dates_found) > 2 else ""

    periodo = ""
    if inicio_periodo and fin_periodo:
        periodo = f"{inicio_periodo} - {fin_periodo}"
    elif emision:
        periodo = f"Emision: {emision}"
    else:
        periodo = "Periodo no especificado"

    # 4. Numero de Cuenta
    cuenta = ""
    acc_m = re.search(r'\b(\d+\.\d+\.\d+\.\d+\.\d+)\b', p1_text)
    if acc_m:
        cuenta = acc_m.group(1).strip()

    # 5. RUC
    ruc = ""
    ruc_m = re.search(r'\b(20\d{9})\b', p1_text)
    if ruc_m:
        ruc = ruc_m.group(1).strip()

    # 6. Importes de la fila de resumen de conceptos:
    # Encabezados: Plan | Prorrateo | Cargos Adicionales | Descuentos | Equipos | Saldo Anterior
    plan_val = None
    equipos_val = None

    # Patron 1: fila completa de 6 montos con S/
    m6 = re.search(
        r'S/\s*([\d,]+\.\d{2})\s+S/\s*([\d,]+\.\d{2})\s+S/\s*([\d,]+\.\d{2})\s+S/\s*([\d,]+\.\d{2})\s+S/\s*([\d,]+\.\d{2})\s+S/\s*([\d,]+\.\d{2})',
        p1_text
    )
    if m6:
        plan_val = to_decimal(m6.group(1))
        equipos_val = to_decimal(m6.group(5))

    # Patron 2: fila resumen en la tabla de la pagina 2
    if plan_val is None:
        tot_m = re.search(
            r'Total\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})\s*\n\s*([\d,]+\.\d{2})',
            all_text
        )
        if tot_m:
            plan_val = to_decimal(tot_m.group(1))
            equipos_val = to_decimal(tot_m.group(5))

    # Patron 3: Desglose tributario en Pagina 1
    if plan_val is None:
        desglose_m = re.search(
            r'Total\s*\(Incl\.\s*IGV\)[\s\S]*?Montos ya gravados[\s\S]*?(\d[\d,]*\.\d{2})[\s\S]*?(\d[\d,]*\.\d{2})',
            p1_text
        )
        if desglose_m:
            plan_val = to_decimal(desglose_m.group(1))
            equipos_val = to_decimal(desglose_m.group(2))

    # 7. Total a pagar
    total_pagar = None
    tot_pagar_m = re.search(r'Total\s*a\s*pagar[\s\S]*?S/\s*([\d,]+\.\d{2})', p1_text, re.IGNORECASE)
    if tot_pagar_m:
        total_pagar = to_decimal(tot_pagar_m.group(1))
    else:
        # Buscar el ultimo monto de la tabla de totales
        tot_end_m = re.search(r'Total a pagar[\s\S]*?(\d[\d,]*\.\d{2})\s*$', p1_text, re.MULTILINE)
        if tot_end_m:
            total_pagar = to_decimal(tot_end_m.group(1))

    if plan_val is None:
        raise ValueError(
            "No se pudo extraer el concepto PLAN del PDF de Entel. "
            "Verifique que el PDF sea un recibo mensual original de Entel."
        )

    return {
        'empresa': empresa,
        'recibo_nro': recibo_nro,
        'periodo': periodo,
        'emision': emision,
        'cuenta': cuenta,
        'ruc': ruc,
        'plan': plan_val,
        'equipos': equipos_val if equipos_val is not None else Decimal('0.00'),
        'total_pagar': total_pagar if total_pagar is not None else (plan_val + (equipos_val or Decimal('0.00'))),
    }


# ---------------------------------------------------------------------------
# Carga e indexacion del Inventario Celulares General
# ---------------------------------------------------------------------------

def load_inventory_entel(inv_bytes):
    """
    Carga el Inventario Celulares General desde bytes.
    Inspecciona hojas como PRODAC y PRODIMIN.
    Identifica dinamicamente encabezados (Operador, Numero, Usuario, Area, Depto, CC Area, CC Dpto).
    Filtra UNICAMENTE lineas correspondientes a Entel para evitar cruzamientos incorrectos.
    Registra tambien numeros de otros operadores para generar advertencias claras.
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(inv_bytes), data_only=True)
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo Inventario Celulares General: {e}")

    entel_map = {}       # phone -> { 'empresa', 'usuario', 'area', 'departamento', 'cc_area', 'cc_dpto', 'row', 'raw_phone' }
    other_op_map = {}    # phone -> operator_name
    dup_inv_entel = []

    # Buscar en hojas candidatas (PRODAC, PRODIMIN o todas las hojas relevantes)
    target_sheets = [s for s in wb.sheetnames if s.upper() in ('PRODAC', 'PRODIMIN')]
    if not target_sheets:
        target_sheets = [wb.sheetnames[0]]

    for sname in target_sheets:
        ws = wb[sname]
        if ws.max_row < 2:
            continue

        # Detectar columnas por encabezado en las primeras 5 filas
        col_op = None
        col_num = None
        col_user = None
        col_area = None
        col_dpto = None
        col_cc_area = None
        col_cc_dpto = None
        header_row = 1

        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, min(25, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or '').strip().upper()
                if not val:
                    continue
                if 'OPERADOR' in val:
                    col_op = c
                    header_row = r
                elif val in ('NUMERO', 'NÚMERO', 'LINEA', 'LÍNEA', 'TELEFONO', 'TELÉFONO') or 'NUMERO' in val:
                    col_num = c
                    header_row = r
                elif any(k in val for k in ('USUARIO', 'RESPONSABLE', 'NOMBRE')):
                    col_user = c
                elif val in ('AREA', 'ÁREA') or 'AREA' in val:
                    if 'CC' not in val:
                        col_area = c
                elif any(k in val for k in ('DEPARTAMENTO', 'DPTO', 'DEPTO')):
                    if 'CC' not in val:
                        col_dpto = c
                elif 'CC AREA' in val or 'CC_AREA' in val:
                    col_cc_area = c
                elif 'CC DPTO' in val or 'CC_DPTO' in val or 'CC DEPTO' in val:
                    col_cc_dpto = c

            if col_num and col_op and (col_cc_dpto or col_user):
                break

        # Fallbacks si no se detectaron por encabezado exacto
        col_op = col_op or 2       # default B
        col_num = col_num or 3     # default C
        col_area = col_area or 4   # default D
        col_dpto = col_dpto or 5   # default E
        col_user = col_user or 6   # default F
        col_cc_area = col_cc_area or 7  # default G
        col_cc_dpto = col_cc_dpto or 8  # default H

        for r in range(header_row + 1, ws.max_row + 1):
            num_raw = ws.cell(r, col_num).value
            phone = clean_phone(num_raw)
            if not phone or len(phone) < 8:
                continue

            op_raw = str(ws.cell(r, col_op).value or '').strip()
            is_entel = 'entel' in op_raw.lower()

            if is_entel:
                record = {
                    'empresa': sname,
                    'operador': op_raw,
                    'usuario': str(ws.cell(r, col_user).value or '').strip(),
                    'area': str(ws.cell(r, col_area).value or '').strip(),
                    'departamento': str(ws.cell(r, col_dpto).value or '').strip(),
                    'cc_area': str(ws.cell(r, col_cc_area).value or '').strip(),
                    'cc_dpto': str(ws.cell(r, col_cc_dpto).value or '').strip(),
                    'row_inv': r,
                    'raw_phone': num_raw,
                }
                if phone in entel_map:
                    dup_inv_entel.append(phone)
                else:
                    entel_map[phone] = record
            else:
                if op_raw and phone not in other_op_map:
                    other_op_map[phone] = op_raw

    if not entel_map:
        raise ValueError(
            "No se encontraron lineas de Entel en el Inventario Celulares General. "
            "Verifique que las hojas PRODAC/PRODIMIN contengan registros con Operadora 'Entel'."
        )

    return {
        'entel_map': entel_map,
        'other_op_map': other_op_map,
        'dup_inv_entel': dup_inv_entel,
    }


# ---------------------------------------------------------------------------
# Inspeccion y deteccion dinamica del Excel de Entel
# ---------------------------------------------------------------------------

def inspect_entel_excel(excel_bytes):
    """
    Inspecciona todas las hojas del Excel de facturacion de Entel.
    Detecta automaticamente:
    - La hoja de detalle de lineas (excluyendo Resumen, Hoja1, Otros Cargos, etc.)
    - Las columnas de Numero Entel, Plan Tarifario y Cargo Fijo.
    - La fila inicial de datos y cantidad de lineas detectadas.
    Permite confirmacion de columnas antes de procesar.
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    candidate_sheets = []

    best_sheet = None
    best_score = -1
    best_config = None

    for sname in wb.sheetnames:
        sname_lower = sname.lower()
        # Ignorar hojas conocidas de resumen o anexos no-lineas
        if sname_lower in ('hoja1', 'sheet1', 'resultado entel', 'resumen', 'ceco'):
            continue

        ws = wb[sname]
        if ws.max_row < 2:
            continue

        col_phone = None
        col_plan = None
        col_cargo = None
        header_row = 1

        # Escaneo de filas superiores en busca de encabezados
        for r in range(1, min(10, ws.max_row + 1)):
            for c in range(1, min(30, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or '').strip().lower()
                if not val:
                    continue
                if any(k in val for k in ['número entel', 'numero entel', 'n° entel', 'no entel', 'nº entel']) or (val in ['numero', 'número', 'linea', 'línea', 'móvil', 'movil'] and col_phone is None):
                    col_phone = c
                    header_row = r
                elif any(k in val for k in ['plan tarifario', 'nombre plan', 'descripción plan', 'tarifa']) or (val == 'plan' and col_plan is None):
                    col_plan = c
                elif any(k in val for k in ['cargo fijo', 'cargos fijos', 'renta mensual', 'monto f', 'cargo_fijo']):
                    col_cargo = c

            if col_phone and col_cargo:
                break

        # Si col_cargo no se encontro por 'cargo fijo', buscar columna numerica con importes de planes
        col_phone_cand = col_phone or 3
        col_plan_cand = col_plan or 4
        col_cargo_cand = col_cargo or 5

        # Contar cuantas filas tienen numeros celulares validos de 9 digitos
        valid_phones = 0
        data_start_row = None
        data_end_row = None
        preview_rows = []

        for r in range((header_row or 1) + 1, ws.max_row + 1):
            raw_val = ws.cell(r, col_phone_cand).value
            phone = clean_phone(raw_val)
            if phone and len(phone) == 9 and phone.startswith('9'):
                valid_phones += 1
                if data_start_row is None:
                    data_start_row = r
                data_end_row = r
                if len(preview_rows) < 5:
                    preview_rows.append({
                        'row': r,
                        'phone': phone,
                        'plan': str(ws.cell(r, col_plan_cand).value or '').strip(),
                        'cargo': float(to_decimal(ws.cell(r, col_cargo_cand).value)),
                    })

        sheet_info = {
            'sheet_name': sname,
            'header_row': header_row,
            'col_phone': col_phone_cand,
            'col_plan': col_plan_cand,
            'col_cargo': col_cargo_cand,
            'col_phone_letter': get_column_letter(col_phone_cand),
            'col_plan_letter': get_column_letter(col_plan_cand),
            'col_cargo_letter': get_column_letter(col_cargo_cand),
            'line_count': valid_phones,
            'data_start_row': data_start_row or (header_row + 1),
            'data_end_row': data_end_row,
            'preview': preview_rows,
        }
        candidate_sheets.append(sheet_info)

        if valid_phones > best_score:
            best_score = valid_phones
            best_sheet = sname
            best_config = sheet_info

    if not best_config or best_score == 0:
        # Intento fallback en cualquier hoja
        for sname in wb.sheetnames:
            ws = wb[sname]
            for c in range(1, min(10, ws.max_column + 1)):
                phones_count = 0
                for r in range(1, min(50, ws.max_row + 1)):
                    p = clean_phone(ws.cell(r, c).value)
                    if len(p) == 9 and p.startswith('9'):
                        phones_count += 1
                if phones_count >= 5:
                    best_config = {
                        'sheet_name': sname,
                        'header_row': 1,
                        'col_phone': c,
                        'col_plan': c + 1,
                        'col_cargo': c + 2,
                        'col_phone_letter': get_column_letter(c),
                        'col_plan_letter': get_column_letter(c + 1),
                        'col_cargo_letter': get_column_letter(c + 2),
                        'line_count': phones_count,
                        'data_start_row': 2,
                        'data_end_row': ws.max_row,
                        'preview': [],
                    }
                    break
            if best_config:
                break

    if not best_config:
        raise ValueError(
            "No se pudo identificar automaticamente la hoja de detalle de lineas en el Excel de Entel. "
            "Asegurese de cargar el archivo mensual original proporcionado por Entel."
        )

    return {
        'all_sheets': wb.sheetnames,
        'has_hoja1': 'Hoja1' in wb.sheetnames,
        'best_config': best_config,
        'candidate_sheets': candidate_sheets,
    }


# ---------------------------------------------------------------------------
# Procesamiento de Hoja1 previa (si existe)
# ---------------------------------------------------------------------------

def load_previous_hoja1(wb):
    """
    Si existe Hoja1 en el libro (usada en meses anteriores como resultado manual),
    extrae el mapeo Numero -> Centro de Costo (CC) para detectar discrepancias
    con el inventario vigente sin sobreescribir la hoja original.
    """
    if 'Hoja1' not in wb.sheetnames:
        return {}

    ws = wb['Hoja1']
    hoja1_map = {}

    col_num = 1
    col_cc = 3
    header_row = 1

    for r in range(1, min(5, ws.max_row + 1)):
        for c in range(1, min(10, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '').upper().strip()
            if 'NUMERO' in v or 'NUM' in v:
                col_num = c
                header_row = r
            elif v == 'CC' or 'CENTRO' in v:
                col_cc = c

    for r in range(header_row + 1, ws.max_row + 1):
        num_raw = ws.cell(r, col_num).value
        phone = clean_phone(num_raw)
        cc_raw = ws.cell(r, col_cc).value
        if phone:
            hoja1_map[phone] = str(cc_raw).strip() if cc_raw is not None else ""

    return hoja1_map


# ---------------------------------------------------------------------------
# Generacion y estilizacion de la hoja 'Resultado Entel'
# ---------------------------------------------------------------------------

def write_resultado_entel_sheet(wb, processed_lines, cc_summary, total_excel):
    """
    Crea la hoja 'Resultado Entel' en la posicion 0 (primera hoja).
    Las demas hojas permanecen 100% intactas en sus posiciones originales.
    Genera:
    1. Tabla principal de Resultados (Columnas A a K).
    2. Fila de Total con formula SUM.
    3. Tabla de Resumen por Centro de Costo (Columnas M a P) con formulas COUNTIF, SUMIF y %.
    4. Estilos visuales corporativos PRODAC / Entel.
    """
    sheet_title = "Resultado Entel"
    if sheet_title in wb.sheetnames:
        del wb[sheet_title]

    ws = wb.create_sheet(sheet_title, 0)
    ws.views.sheetView[0].showGridLines = True

    # Paleta de colores
    COLOR_PRIMARY = "0F2744"      # Azul marino corporativo PRODAC
    COLOR_ENTEL   = "EA580C"      # Naranja representativo Entel
    COLOR_BG_TH   = "1E3A8A"      # Azul oscuro para encabezados
    COLOR_BG_ALT  = "F8FAFC"      # Fila alterna muy suave
    COLOR_BORDER  = "CBD5E1"      # Borde sutil gris azulado
    COLOR_ACCENT  = "0284C7"      # Azul cielo para resumen

    font_title   = Font(name="Calibri", size=14, bold=True, color="0F2744")
    font_sub     = Font(name="Calibri", size=10, italic=True, color="64748B")
    font_header  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data    = Font(name="Calibri", size=10, color="0F172A")
    font_total   = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_warn    = Font(name="Calibri", size=10, color="B91C1C", bold=True)
    font_obs     = Font(name="Calibri", size=9, color="475569")

    fill_header_main = PatternFill(start_color=COLOR_BG_TH, end_color=COLOR_BG_TH, fill_type="solid")
    fill_header_res  = PatternFill(start_color="0369A1", end_color="0369A1", fill_type="solid")
    fill_alt         = PatternFill(start_color=COLOR_BG_ALT, end_color=COLOR_BG_ALT, fill_type="solid")
    fill_total       = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_warn_cell   = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color=COLOR_BORDER),
        right=Side(style='thin', color=COLOR_BORDER),
        top=Side(style='thin', color=COLOR_BORDER),
        bottom=Side(style='thin', color=COLOR_BORDER),
    )
    total_border = Border(
        top=Side(style='thin', color="0F172A"),
        bottom=Side(style='double', color="0F172A"),
    )

    fmt_curr = '"S/ "#,##0.00'
    fmt_pct  = '0.0%'
    fmt_int  = '#,##0'

    # Titulo de la hoja
    ws['A1'] = "CONCILIACION MENSUAL DE LINEAS - ENTEL"
    ws['A1'].font = font_title
    ws['A2'] = f"Generado dinamicamente: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')} | Cruce oficial contra Inventario Celulares General"
    ws['A2'].font = font_sub

    # Encabezados de la Tabla Principal (Fila 4)
    HEADERS_MAIN = [
        ("A", "Numero telefonico", 16, Alignment(horizontal="center")),
        ("B", "Plan tarifario",    26, Alignment(horizontal="left")),
        ("C", "Cargo fijo con IGV",18, Alignment(horizontal="right")),
        ("D", "Empresa",           12, Alignment(horizontal="center")),
        ("E", "Usuario",           28, Alignment(horizontal="left")),
        ("F", "Area",              24, Alignment(horizontal="left")),
        ("G", "Departamento",      26, Alignment(horizontal="left")),
        ("H", "Centro de costo",   16, Alignment(horizontal="center")),
        ("I", "Codigo PRS / Area", 16, Alignment(horizontal="center")),
        ("J", "Estado conciliacion", 18, Alignment(horizontal="center")),
        ("K", "Observaciones",     42, Alignment(horizontal="left")),
    ]

    row_h = 4
    for col_letter, title, width, align in HEADERS_MAIN:
        c = ws[f"{col_letter}{row_h}"]
        c.value = title
        c.font = font_header
        c.fill = fill_header_main
        c.alignment = Alignment(horizontal=align.horizontal, vertical="center", wrap_text=True)
        c.border = thin_border
        ws.column_dimensions[col_letter].width = width

    # Escribir filas de datos
    cur_row = 5
    for idx, item in enumerate(processed_lines):
        r = cur_row
        is_even = (idx % 2 == 1)
        row_fill = fill_alt if is_even else None

        # A: Numero (formato texto para no truncar ceros)
        cA = ws[f"A{r}"]
        cA.value = str(item['phone'])
        cA.number_format = '@'
        cA.alignment = Alignment(horizontal="center", vertical="center")
        cA.font = font_data

        # B: Plan
        cB = ws[f"B{r}"]
        cB.value = item['plan']
        cB.alignment = Alignment(horizontal="left", vertical="center")
        cB.font = font_data

        # C: Cargo Fijo con IGV (numero)
        cC = ws[f"C{r}"]
        cC.value = float(item['cargo_fijo'])
        cC.number_format = fmt_curr
        cC.alignment = Alignment(horizontal="right", vertical="center")
        cC.font = font_data

        # D: Empresa
        cD = ws[f"D{r}"]
        cD.value = item['empresa'] or "Sin info"
        cD.alignment = Alignment(horizontal="center", vertical="center")
        cD.font = font_data

        # E: Usuario
        cE = ws[f"E{r}"]
        cE.value = item['usuario'] or "Sin informacion"
        cE.alignment = Alignment(horizontal="left", vertical="center")
        cE.font = font_data

        # F: Area
        cF = ws[f"F{r}"]
        cF.value = item['area'] or "Sin info"
        cF.alignment = Alignment(horizontal="left", vertical="center")
        cF.font = font_data

        # G: Departamento
        cG = ws[f"G{r}"]
        cG.value = item['departamento'] or "Sin info"
        cG.alignment = Alignment(horizontal="left", vertical="center")
        cG.font = font_data

        # H: Centro de costo (CC Dpto)
        cH = ws[f"H{r}"]
        cH.value = item['cc_dpto'] or "Sin CC"
        cH.alignment = Alignment(horizontal="center", vertical="center")
        cH.font = font_data if item['cc_dpto'] else font_warn

        # I: Codigo PRS / CC Area (si existe y es distinto)
        cI = ws[f"I{r}"]
        cI.value = item['prs_code'] or "-"
        cI.alignment = Alignment(horizontal="center", vertical="center")
        cI.font = font_data

        # J: Estado de conciliacion
        cJ = ws[f"J{r}"]
        cJ.value = item['estado']
        cJ.alignment = Alignment(horizontal="center", vertical="center")
        if item['estado'] == 'Conciliado':
            cJ.font = font_data
        else:
            cJ.font = font_warn

        # K: Observaciones
        cK = ws[f"K{r}"]
        cK.value = item['observacion'] or ""
        cK.alignment = Alignment(horizontal="left", vertical="center")
        cK.font = font_obs

        # Aplicar bordes y relleno
        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
            cell = ws[f"{col_l}{r}"]
            cell.border = thin_border
            if row_fill and not (col_l == 'H' and not item['cc_dpto']):
                cell.fill = row_fill
            if col_l == 'H' and not item['cc_dpto']:
                cell.fill = fill_warn_cell

        cur_row += 1

    last_data_row = cur_row - 1

    # Fila de Totales de la Tabla Principal
    tot_row = cur_row
    ws[f"A{tot_row}"] = "Total general"
    ws[f"A{tot_row}"].font = font_total
    ws[f"A{tot_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{tot_row}"].fill = fill_total
    ws[f"A{tot_row}"].border = total_border

    for col_l in ["B", "D", "E", "F", "G", "H", "I", "J", "K"]:
        c = ws[f"{col_l}{tot_row}"]
        c.fill = fill_total
        c.border = total_border

    c_tot = ws[f"C{tot_row}"]
    c_tot.value = f"=SUM(C5:C{last_data_row})"
    c_tot.font = font_total
    c_tot.number_format = fmt_curr
    c_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_tot.fill = fill_total
    c_tot.border = total_border

    # =========================================================================
    # TABLA RESUMEN POR CENTRO DE COSTO (Columnas M a P)
    # Requisito: formulas SUMIF y COUNTIF para que sea verificable y dinamica
    # =========================================================================
    ws.column_dimensions["L"].width = 4  # Columna separadora

    HEADERS_RESUMEN = [
        ("M", "Centro de costo",      18, Alignment(horizontal="center")),
        ("N", "Cantidad de lineas",   16, Alignment(horizontal="right")),
        ("O", "Total Cargo Fijo",     18, Alignment(horizontal="right")),
        ("P", "% del Total",          14, Alignment(horizontal="right")),
    ]

    for col_l, title, width, align in HEADERS_RESUMEN:
        c = ws[f"{col_l}{row_h}"]
        c.value = title
        c.font = font_header
        c.fill = fill_header_res
        c.alignment = Alignment(horizontal=align.horizontal, vertical="center", wrap_text=True)
        c.border = thin_border
        ws.column_dimensions[col_l].width = width

    res_row = 5
    for idx, cc_item in enumerate(cc_summary):
        r = res_row
        cc_name = cc_item['cc']

        # M: Centro de Costo
        cM = ws[f"M{r}"]
        cM.value = cc_name
        cM.font = font_data
        cM.alignment = Alignment(horizontal="center", vertical="center")
        cM.border = thin_border

        # N: Cantidad de lineas con COUNTIF
        cN = ws[f"N{r}"]
        cN.value = f'=COUNTIF($H$5:$H${last_data_row}, M{r})'
        cN.font = font_data
        cN.number_format = fmt_int
        cN.alignment = Alignment(horizontal="right", vertical="center")
        cN.border = thin_border

        # O: Total Cargo Fijo con SUMIF
        cO = ws[f"O{r}"]
        cO.value = f'=SUMIF($H$5:$H${last_data_row}, M{r}, $C$5:$C${last_data_row})'
        cO.font = font_data
        cO.number_format = fmt_curr
        cO.alignment = Alignment(horizontal="right", vertical="center")
        cO.border = thin_border

        # P: Porcentaje con formula referenciada a la celda total
        cP = ws[f"P{r}"]
        tot_res_row = 5 + len(cc_summary)
        cP.value = f'=O{r}/$O${tot_res_row}'
        cP.font = font_data
        cP.number_format = fmt_pct
        cP.alignment = Alignment(horizontal="right", vertical="center")
        cP.border = thin_border

        res_row += 1

    last_res_row = res_row - 1
    tot_res_row = res_row

    # Fila Total de la Tabla Resumen
    ws[f"M{tot_res_row}"] = "Total general"
    ws[f"M{tot_res_row}"].font = font_total
    ws[f"M{tot_res_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"M{tot_res_row}"].fill = fill_total
    ws[f"M{tot_res_row}"].border = total_border

    ws[f"N{tot_res_row}"] = f"=SUM(N5:N{last_res_row})"
    ws[f"N{tot_res_row}"].font = font_total
    ws[f"N{tot_res_row}"].number_format = fmt_int
    ws[f"N{tot_res_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"N{tot_res_row}"].fill = fill_total
    ws[f"N{tot_res_row}"].border = total_border

    ws[f"O{tot_res_row}"] = f"=SUM(O5:O{last_res_row})"
    ws[f"O{tot_res_row}"].font = font_total
    ws[f"O{tot_res_row}"].number_format = fmt_curr
    ws[f"O{tot_res_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"O{tot_res_row}"].fill = fill_total
    ws[f"O{tot_res_row}"].border = total_border

    ws[f"P{tot_res_row}"] = f"=SUM(P5:P{last_res_row})"
    ws[f"P{tot_res_row}"].font = font_total
    ws[f"P{tot_res_row}"].number_format = fmt_pct
    ws[f"P{tot_res_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"P{tot_res_row}"].fill = fill_total
    ws[f"P{tot_res_row}"].border = total_border

    # Congelar paneles en encabezados (fila 5 hacia abajo, visible A y B)
    ws.freeze_panes = "C5"

    # Agregar autofiltro a la tabla principal
    ws.auto_filter.ref = f"A4:K{last_data_row}"


# ---------------------------------------------------------------------------
# Motor principal de Conciliacion Entel
# ---------------------------------------------------------------------------

def process_entel_files(excel_bytes, pdf_bytes, inventory_bytes, custom_config=None):
    """
    Flujo integral de conciliacion mensual para Entel:
    1. Carga del Excel original del proveedor (preservando formato original).
    2. Carga del Inventario Celulares General (filtrando operador Entel).
    3. Extraccion de datos del PDF oficial (concepto PLAN).
    4. Cruce de datos por numero telefonico normalizado de 9 digitos.
    5. Asignacion de CC Dpto como centro de costo.
    6. Comparacion contra Hoja1 previa (si existe) alertando diferencias.
    7. Validacion de sumas (Total Excel vs Total PLAN PDF, tolerancia S/ 0.01).
    8. Creacion de la hoja 'Resultado Entel' en la posicion 0 del Excel original.
    9. Generacion del Excel final en memoria sin alterar las hojas existentes.
    """
    # 1. Cargar el libro original de Entel para lectura y modificacion
    try:
        # data_only=False para conservar formulas y formatos originales de las otras hojas
        wb_out = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
        # wb_read para leer valores calculados si habia formulas
        wb_read = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    except Exception as e:
        raise ValueError(f"No se pudo abrir el archivo Excel de facturacion de Entel: {e}")

    # 2. Extraer datos del PDF (periodo, recibo, concepto PLAN, equipos, total a pagar)
    pdf_info = extract_pdf_entel(pdf_bytes)
    plan_pdf = pdf_info['plan']

    # 3. Cargar el inventario (filtrando por operador Entel)
    inv_data = load_inventory_entel(inventory_bytes)
    entel_map = inv_data['entel_map']
    other_op_map = inv_data['other_op_map']

    # 4. Detectar o aplicar configuracion de hoja y columnas de Entel
    if custom_config and custom_config.get('sheet_name'):
        sheet_name = custom_config['sheet_name']
        col_phone  = int(custom_config.get('col_phone', 3))
        col_plan   = int(custom_config.get('col_plan', 4))
        col_cargo  = int(custom_config.get('col_cargo', 5))
        start_row  = int(custom_config.get('data_start_row', 5))
    else:
        inspection = inspect_entel_excel(excel_bytes)
        cfg = inspection['best_config']
        sheet_name = cfg['sheet_name']
        col_phone  = cfg['col_phone']
        col_plan   = cfg['col_plan']
        col_cargo  = cfg['col_cargo']
        start_row  = cfg['data_start_row']

    if sheet_name not in wb_read.sheetnames:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el Excel de Entel.")

    ws_lines = wb_read[sheet_name]

    # 5. Cargar Hoja1 previa si existe en el Excel del proveedor
    hoja1_map = load_previous_hoja1(wb_read)

    # 6. Extraer y procesar lineas telefonicas
    processed_lines = []
    seen_phones = set()
    dup_phones_file = []
    total_cargo_excel = Decimal('0.00')

    sin_cc_count = 0
    no_inv_count = 0
    conciliado_count = 0
    cc_discrepancies = []
    general_observations = []

    for r in range(start_row, ws_lines.max_row + 1):
        raw_phone = ws_lines.cell(r, col_phone).value
        phone = clean_phone(raw_phone)
        if not phone or len(phone) < 8:
            continue

        # Validar si no es un numero celular (debe tener 9 digitos o empezar con 9)
        if len(phone) != 9:
            continue

        if phone in seen_phones:
            dup_phones_file.append(phone)
        seen_phones.add(phone)

        # Plan tarifario
        plan_val = str(ws_lines.cell(r, col_plan).value or '').strip()

        # Cargo fijo con IGV
        cargo_raw = ws_lines.cell(r, col_cargo).value
        cargo_fijo = to_decimal(cargo_raw)
        total_cargo_excel += cargo_fijo

        # Cruce con Inventario
        inv_item = entel_map.get(phone)
        observaciones_linea = []
        estado_linea = "Conciliado"

        empresa_val = ""
        usuario_val = ""
        area_val = ""
        dpto_val = ""
        cc_dpto_val = ""
        cc_area_val = ""
        prs_code_val = ""

        if inv_item:
            empresa_val  = inv_item['empresa']
            usuario_val  = inv_item['usuario']
            area_val     = inv_item['area']
            dpto_val     = inv_item['departamento']
            cc_dpto_val  = inv_item['cc_dpto']
            cc_area_val  = inv_item['cc_area']

            # Codigo PRS: si CC Area difiere de CC Dpto, guardarlo en columna I
            if cc_area_val and cc_area_val != cc_dpto_val:
                prs_code_val = cc_area_val

            if not cc_dpto_val:
                estado_linea = "Sin CC"
                observaciones_linea.append("Sin Centro de Costo en inventario")
                sin_cc_count += 1
            else:
                conciliado_count += 1
        else:
            # No encontrado en lineas de Entel
            no_inv_count += 1
            estado_linea = "No encontrado"
            if phone in other_op_map:
                op_otro = other_op_map[phone]
                observaciones_linea.append(f"Registrado en inventario bajo otro operador ({op_otro})")
            else:
                observaciones_linea.append("No figura en el Inventario de Celulares")

        # Comparar con Hoja1 anterior si existe
        if phone in hoja1_map:
            hoja1_cc = hoja1_map[phone]
            if cc_dpto_val and hoja1_cc and cc_dpto_val != hoja1_cc:
                diff_msg = f"CC difiere de Hoja1 previa (Inv: {cc_dpto_val} vs Hoja1: {hoja1_cc})"
                observaciones_linea.append(diff_msg)
                cc_discrepancies.append({
                    'phone': phone,
                    'inv_cc': cc_dpto_val,
                    'hoja1_cc': hoja1_cc,
                    'usuario': usuario_val,
                })
                # No invalidar la conciliacion si el inventario actual tiene CC, pero reflejarlo
                if estado_linea == "Conciliado":
                    estado_linea = "Conciliado (CC cambiado)"

        # Duplicados en archivo
        if phone in dup_phones_file:
            observaciones_linea.append("Numero duplicado en la factura")
            estado_linea = "Duplicado"

        processed_lines.append({
            'phone': phone,
            'plan': plan_val,
            'cargo_fijo': cargo_fijo,
            'empresa': empresa_val,
            'usuario': usuario_val,
            'area': area_val,
            'departamento': dpto_val,
            'cc_dpto': cc_dpto_val,
            'prs_code': prs_code_val,
            'estado': estado_linea,
            'observacion': "; ".join(observaciones_linea),
        })

    if not processed_lines:
        raise ValueError(
            f"No se pudieron extraer lineas telefonicas de la hoja '{sheet_name}'. "
            "Verifique la configuracion de columnas del Excel de Entel."
        )

    # 7. Resumen agrupado por CC Dpto
    cc_stats = {}
    for item in processed_lines:
        cc_name = item['cc_dpto'] or "Sin Centro de Costo"
        if cc_name not in cc_stats:
            cc_stats[cc_name] = {'count': 0, 'monto': Decimal('0.00')}
        cc_stats[cc_name]['count'] += 1
        cc_stats[cc_name]['monto'] += item['cargo_fijo']

    cc_summary = []
    for cc_name, st in cc_stats.items():
        monto_cc = st['monto']
        pct = float((monto_cc / total_cargo_excel * 100).quantize(Decimal('0.1'))) if total_cargo_excel > 0 else 0.0
        cc_summary.append({
            'cc': cc_name,
            'count': st['count'],
            'monto': float(monto_cc),
            'pct': pct,
        })
    # Ordenar por monto descendente, dejando 'Sin Centro de Costo' al final si existe
    cc_summary.sort(key=lambda x: (x['cc'] == 'Sin Centro de Costo', -x['monto']))

    # 8. Conciliacion monetaria contra el PDF (Tolerancia S/ 0.01)
    diferencia = (total_cargo_excel - plan_pdf).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    cuadra_monetaria = abs(diferencia) <= Decimal('0.01')

    # Validacion integral: diferenciacon entre conciliacion monetaria y validacion de datos
    es_valido_todo = cuadra_monetaria and (sin_cc_count == 0) and (no_inv_count == 0) and (len(dup_phones_file) == 0)

    if cuadra_monetaria:
        estado_general = "CONCILIADO"
        mensaje_general = "CONCILIADO — Los importes coinciden."
    else:
        estado_general = "REVISAR"
        mensaje_general = f"REVISAR — Existe una diferencia de S/ {diferencia:+.2f} entre Excel y PDF."

    if cc_discrepancies:
        general_observations.append(
            f"Se detectaron {len(cc_discrepancies)} linea(s) con centro de costo diferente respecto a la Hoja1 previa."
        )
    if sin_cc_count > 0:
        general_observations.append(
            f"{sin_cc_count} linea(s) no tienen centro de costo asignado en el inventario."
        )
    if no_inv_count > 0:
        general_observations.append(
            f"{no_inv_count} linea(s) no se encontraron en el inventario de Entel."
        )
    if dup_phones_file:
        general_observations.append(
            f"Se encontraron {len(dup_phones_file)} numero(s) duplicados en la factura."
        )

    # 9. Escribir hoja 'Resultado Entel' en la posicion 0
    write_resultado_entel_sheet(wb_out, processed_lines, cc_summary, total_cargo_excel)

    # 10. Guardar libro resultante en memoria
    excel_out = io.BytesIO()
    wb_out.save(excel_out)
    excel_out.seek(0)
    final_excel_bytes = excel_out.getvalue()

    # Preparar lineas de observacion para descarga CSV
    obs_lines_csv = [
        p for p in processed_lines if p['observacion'] or p['estado'] != 'Conciliado'
    ]

    return {
        'periodo': pdf_info['periodo'],
        'recibo_nro': pdf_info['recibo_nro'],
        'empresa': pdf_info['empresa'],
        'cuenta': pdf_info['cuenta'],
        'ruc': pdf_info['ruc'],
        'sheet_used': sheet_name,
        'line_count': len(processed_lines),
        'conciliado_count': conciliado_count,
        'sin_cc_count': sin_cc_count,
        'no_inv_count': no_inv_count,
        'dup_count': len(dup_phones_file),
        'total_excel': float(total_cargo_excel),
        'total_pdf_plan': float(plan_pdf),
        'total_pdf_equipos': float(pdf_info['equipos']),
        'total_pdf_factura': float(pdf_info['total_pagar']),
        'diferencia': float(diferencia),
        'cuadra': cuadra_monetaria,
        'es_valido_todo': es_valido_todo,
        'estado_conciliacion': estado_general,
        'mensaje_conciliacion': mensaje_general,
        'cc_discrepancies': cc_discrepancies,
        'general_observations': general_observations,
        'lines': [
            {
                'phone': p['phone'],
                'plan': p['plan'],
                'cargo_fijo': float(p['cargo_fijo']),
                'empresa': p['empresa'],
                'usuario': p['usuario'],
                'area': p['area'],
                'departamento': p['departamento'],
                'cc_dpto': p['cc_dpto'],
                'prs_code': p['prs_code'],
                'estado': p['estado'],
                'observacion': p['observacion'],
            }
            for p in processed_lines
        ],
        'cc_summary': cc_summary,
        'obs_lines_count': len(obs_lines_csv),
        'obs_lines': [
            {
                'phone': p['phone'],
                'plan': p['plan'],
                'cargo_fijo': float(p['cargo_fijo']),
                'usuario': p['usuario'],
                'cc_dpto': p['cc_dpto'],
                'estado': p['estado'],
                'observacion': p['observacion'],
            }
            for p in obs_lines_csv
        ],
        'excel_bytes': final_excel_bytes,
    }
