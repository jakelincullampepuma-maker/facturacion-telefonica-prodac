"""Lectura y controles PDF de Claro; el Excel original aporta los cargos."""
import datetime as dt
import io
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import openpyxl
from pypdf import PdfReader

from claro_workbook import generate_claro_excel

CENT = Decimal('0.01')
ZERO = Decimal('0.00')
MONTHS = ('Enero Febrero Marzo Abril Mayo Junio Julio Agosto Septiembre Octubre '
          'Noviembre Diciembre').split()
MONTH_ABBR = ('ENE FEB MAR ABR MAY JUN JUL AGO SEP OCT NOV DIC').lower().split()
MONEY_RE = re.compile(r'(?<![\d.,])(?:[-−–]\s*)?(?:S\s*/\.?\s*)?'
                      r'\(?\d[\d.,]*[.,]\d{2}\)?(?![\d%]|\s*%)')
PHONE_RE = re.compile(r'(?<!\d)(?:\+?51[ -]?)?(9(?:[ -]?\d){8})(?!\d)')
ADJUSTMENT_RE = re.compile(r'\b(?:varios|ajustes?|abonos?|devoluciones?|'
                           r'descuentos?|otros cargos|dev\.|bonificaci\w*)', re.I)


def clean(value):
    return ''.join(c for c in unicodedata.normalize('NFD', str(value or ''))
                   if unicodedata.category(c) != 'Mn').strip().lower()


def phone(value):
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    digits = re.sub(r'\D', '', str(value or ''))
    return digits[-9:] if len(digits) >= 9 else ''


def money(value):
    if isinstance(value, (int, float, Decimal)):
        result = Decimal(str(value))
    else:
        text = str(value).strip().replace('−', '-').replace('–', '-')
        negative = '-' in text or '(' in text
        text = re.sub(r'[^\d.,]', '', text)
        if not text:
            raise ValueError('Importe vacío o inválido.')
        # El último separador con dos decimales es el separador decimal.
        if re.search(r'[.,]\d{2}$', text):
            text = re.sub(r'[.,]', '', text[:-3]) + '.' + text[-2:]
        else:
            text = text.replace(',', '')
        try:
            result = Decimal(text) * (-1 if negative else 1)
        except InvalidOperation as exc:
            raise ValueError('Importe inválido.') from exc
    if not result.is_finite():
        raise ValueError('Importe no finito.')
    return result.quantize(CENT, rounding=ROUND_HALF_UP)


def amounts(text):
    return [(m, money(m.group())) for m in MONEY_RE.finditer(text)]


def parse_period(text):
    normalized = clean(text)
    month_pattern = '|'.join(clean(m) for m in MONTHS)
    # Priorizar la etiqueta del periodo sobre fechas de vencimiento/emisión.
    label = re.search(r'(?:periodo(?:\s+(?:de\s+)?facturad\w*|\s+de\s+facturacion)?|'
                      r'mes\s+de\s+facturacion)\s*[:\-]?\s*([^\n]+)', normalized)
    if not label:
        raise ValueError('No se pudo identificar el periodo facturado en el PDF.')
    value = label.group(1).strip()
    named = re.search(r'\b(' + month_pattern + r')\s*(?:de|del|/|-)?\s*(20\d{2})\b', value)
    if named:
        month = [clean(m) for m in MONTHS].index(named.group(1)) + 1
        year = int(named.group(2))
    else:
        abbreviated = re.findall(r'\b(\d{1,2})/([a-z]{3})/(\d{2}|20\d{2})\b', value)
        dates = re.findall(r'\b(\d{1,2})[/.-](\d{1,2})[/.-](20\d{2})\b', value)
        short = re.search(r'\b(0?[1-9]|1[0-2])[/.-](20\d{2})\b', value)
        if abbreviated:
            day, abbreviation, year_text = abbreviated[-1]
            if abbreviation not in MONTH_ABBR:
                raise ValueError('Mes del periodo no reconocido.')
            month, year = MONTH_ABBR.index(abbreviation) + 1, int(year_text)
            if len(year_text) == 2:
                year += 2000
            dt.date(year, month, int(day))
        elif dates:
            # El último día del rango determina el mes facturado.
            day, month, year = map(int, dates[-1])
            dt.date(year, month, day)
        elif short:
            month, year = map(int, short.groups())
        else:
            raise ValueError('El periodo del PDF no contiene un mes y año reconocibles.')
    return {'periodo': f'{MONTHS[month - 1]} {year}', 'mes': month, 'anio': year,
            'periodo_facturado': value}


def parse_metadata(text):
    normalized = clean(text)
    companies = set(re.findall(r'\b(prodac|prodimin)\b', normalized))
    if re.search(r'productos\s+de\s+acero\s+cassado', normalized):
        companies.add('prodac')
    if re.search(r'productos\s+(?:de\s+minerales|minerales|mineros)', normalized):
        companies.add('prodimin')
    if len(companies) != 1:
        raise ValueError('El PDF debe identificar inequívocamente a PRODAC o PRODIMIN.')
    company = companies.pop().upper()
    result = {'company': company, **parse_period(text)}
    lines = text.splitlines()
    owner = next((i for i, line in enumerate(lines)
                  if company.lower() in clean(line) or
                  re.search(r'productos\s+(?:de\s+acero|de\s+minerales|minerales|mineros)', clean(line))), 0)
    result['razon_social'] = lines[owner].strip()
    # El RUC del cliente se busca cerca de su razón social, nunca por constante.
    vicinity = '\n'.join(lines[max(0, owner - 2):owner + 7])
    ruc = re.search(r'R\.?\s*U\.?\s*C\.?\s*[:\-]?\s*(\d{11})', vicinity, re.I)
    result['ruc'] = ruc.group(1) if ruc else ''
    receipt = re.search(r'(?:recibo|factura)(?:\s+electronico)?\s*(?:n[°ºo.]?\s*)?[:#\s]*'
                        r'([A-Z0-9]{3,5}\s*[-–]\s*\d{5,12})', normalized, re.I)
    result['numero_recibo'] = re.sub(r'\s+', '', receipt.group(1)).upper() if receipt else ''
    emission = re.search(r'fecha\s*(?:de\s*)?emision\s*:?\s*([^\n]+)', normalized)
    result['fecha_emision'] = emission.group(1).strip() if emission else ''
    count = re.search(r'(?:cantidad|numero|total|n[°º])\s*(?:de\s*)?lineas\s*:?\s*(\d+)', normalized)
    result['cantidad_lineas_pdf'] = int(count.group(1)) if count else None

    def labeled_amount(pattern, required=False):
        match = re.search(pattern + r'\s*:?\s*(?:\([^\n)]*\))?\s*([^\n]*)', normalized)
        if match:
            found = amounts(match.group(1))
            if not found:
                found = amounts(normalized[match.end():].lstrip('\n').split('\n')[0])
            if found:
                return found[0][1]
        if required:
            raise ValueError('No se pudo leer Cargos del Mes en el PDF; no se sustituirá por otro total.')
        return None

    result['cargos_del_mes'] = labeled_amount(r'\bcargos\s+del\s+mes\b', True)
    result['igv'] = labeled_amount(r'\bi\.?g\.?v\.?\s*(?:18\s*%)?')
    result['total_recibo'] = labeled_amount(r'\b(?:importe\s+total|total\s+(?:del\s+recibo|recibo\s+actual|a\s+pagar|facturado))\b')
    return result


def process_claro_files(pdf_bytes, inventory_bytes, fact_bytes=None, company=None):
    from claro_billing import process_claro_files as process
    return process(pdf_bytes, inventory_bytes, fact_bytes, company)

def json_values(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, dict):
        return {k: json_values(v) for k, v in value.items()}
    if isinstance(value, list):
        return [json_values(v) for v in value]
    return value
