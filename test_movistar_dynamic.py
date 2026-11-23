"""Regresiones financieras con documentos sintéticos y meses distintos."""
import io
import unittest
import zipfile
from decimal import Decimal
from unittest.mock import patch
from xml.etree import ElementTree as ET

import openpyxl
import pypdf
import movistar_engine as engine
from test_support import movistar_fixtures, simple_pdf, workbook_bytes


def inventory(rows=None, sheet='PRODAC', headers=None):
    wb = openpyxl.Workbook()
    wb.active.title = sheet
    wb.active.append(headers or ['Numero', 'CC Dpto'])
    for row in rows or [['951234567', 'CC-A'], ['952345678', 'CC-B']]:
        wb.active.append(row)
    return workbook_bytes(wb)


def pdf_text(**changes):
    text = '''Movistar
Recibo Enero 2027
Anexo: 2
Subtotal S/ 35.00
IGV (18%) S/ 6.30
Total Afecto S/ 41.30
Cargos Adicionales Inafectos S/ 12.00
Redondeo S/ -0.02
Total Facturado S/ 53.28
Total a pagar S/ 53.28
Anexo 1: +51 951-234-567
Cargos Mensuales: Plan variable S/ 10.00
Descuentos y Bonificaciones Afectos: Descuento 50% al plan x 3m S/-5.00
Anexo 2: 952345678
Cargos Mensuales: Otro plan S/ 30.00
Cargos Adicionales Inafectos: Cuota 1 de 4 financiamiento de venta Nro. (TELEF: +51 951-234-567) S/ 12.00'''
    for old, new in changes.items():
        text = text.replace(old, new)
    return text


class MovistarDynamicTests(unittest.TestCase):
    def process(self, text=None, inv=None, **kwargs):
        return engine.process_movistar_files(simple_pdf(text or pdf_text()), inv or inventory(), **kwargs)

    def test_new_month_negative_rounding_missing_cc_and_fresh_pivot(self):
        result = self.process(inv=inventory([['+51 951-234-567', 'CC-NUEVO']]))
        self.assertEqual(result['periodo'], 'Enero 2027')
        self.assertTrue(result['cuadra'])
        self.assertEqual(result['sin_cc_count'], 1)
        self.assertEqual(result['redondeo'], -0.02)
        wb = openpyxl.load_workbook(io.BytesIO(result['excel_bytes']))
        h1 = wb['Hoja1']
        pivot = h1._pivots[0]
        self.assertEqual(pivot.cache.recordCount, 2)
        self.assertEqual(pivot.cache.records.count, 2)
        self.assertEqual(pivot.cache.cacheSource.worksheetSource.name, 'MovistarDatos')
        self.assertIsNone(pivot.cache.cacheSource.worksheetSource.ref)
        self.assertEqual(h1.tables['MovistarDatos'].ref, 'A1:C4')
        self.assertEqual(h1.tables['MovistarDatos'].totalsRowCount, 1)
        self.assertTrue(pivot.cache.refreshOnLoad)
        self.assertEqual(pivot.dataFields[0].subtotal, 'sum')
        self.assertEqual(h1['G5'].value, 'CC-NUEVO')
        self.assertEqual(h1['G6'].value, 'SIN CC')
        values = openpyxl.load_workbook(io.BytesIO(result['excel_bytes']), data_only=True)
        self.assertEqual(values['CARGO MENSUAL']['F2'].value, 5)
        self.assertEqual(values['Hoja1']['H7'].value, 35)
        self.assertEqual(wb['Hoja1']['A2'].data_type, 's')
        self.assertEqual(wb['CARGOS ADICIONALES']['A2'].value, '951234567')
        self.assertIsNone(h1['I4'].value)
        with zipfile.ZipFile(io.BytesIO(result['excel_bytes'])) as archive:
            records = ET.fromstring(archive.read('xl/pivotCache/pivotCacheRecords1.xml'))
            self.assertEqual(len(records), 2)

    def test_classifies_durations_and_normalizes_descriptions(self):
        cases = {'DÉSCUENTO  de 50% al plan x 12m': '12m',
                 'Dscto x Perm B2B 50% al plan x 12 m': 'perm',
                 'descuento al plan 3 meses': '3m', 'descuento x18m': '18m',
                 'Ajuste IP de S/3 x18m': '18m', 'descuento x13m': 'incorporado',
                 'bonificacion extraordinaria': 'incorporado'}
        for concept, kind in cases.items():
            with self.subTest(concept=concept):
                self.assertEqual(engine._discount_kind(concept), kind)
        result = self.process(pdf_text(**{'Descuento 50% al plan x 3m': 'Ajuste IP x18m'}))
        self.assertEqual(result['descuento_total'], 0)
        self.assertEqual(result['cargo_mensual_sin_igv'], 35)
        self.assertEqual(result['cargo_final_sin_igv'], 35)

    def test_anexo_and_concept_continue_on_next_page(self):
        first, second = pdf_text().split('Otro plan S/ 30.00')
        writer = pypdf.PdfWriter()
        for part in (first + 'Otro plan', 'Empresas Movil\nImporte S/\nS/ 30.00' + second):
            writer.append(pypdf.PdfReader(io.BytesIO(simple_pdf(part))))
        out = io.BytesIO()
        writer.write(out)
        result = engine.process_movistar_files(out.getvalue(), inventory())
        self.assertEqual(result['line_count'], 2)
        self.assertTrue(result['cuadra'])
        self.assertEqual(result['lines_sample'][1]['plan'], 'Otro plan')

    def test_missing_quota_never_reports_cuadra_using_summary_total(self):
        result = self.process(pdf_text().split('Cargos Adicionales Inafectos:')[0])
        self.assertFalse(result['cuadra'])
        self.assertEqual(result['cargos_inafectos'], 0)
        self.assertEqual(result['cargos_inafectos_pdf'], 12)
        self.assertEqual(result['diferencia'], -12)
        self.assertTrue(any('inafectos' in w for w in result['warnings']))

    def test_financing_description_without_cuota_prefix(self):
        result = self.process(pdf_text(**{'Cuota 1 de 4 financiamiento': 'financiamiento'}))
        self.assertTrue(result['cuadra'])
        self.assertEqual(result['inafectos_breakdown']['count_ventas'], 1)

    def test_duplicate_pdf_and_missing_declared_anexo_are_visible(self):
        for replacement in ({'Anexo: 2': 'Anexo: 3'}, {'952345678': '951234567'}):
            with self.subTest(replacement=replacement):
                result = self.process(pdf_text(**replacement))
                self.assertFalse(result['cuadra'])
                self.assertTrue(result['warnings'])
                self.assertEqual(result['line_count'], 2)

    def test_inventory_requires_real_fields_and_ambiguous_cc_is_empty(self):
        for inv, message in [(inventory(sheet='OTRA'), 'PRODAC'),
                             (inventory(headers=['Numero', 'Otro']), 'CC Dpto')]:
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                engine.load_inventory(inv)
        inv = inventory([['951234567', 'CC-A'], ['51 951-234-567', 'CC-B']])
        result = self.process(inv=inv)
        self.assertEqual(result['sin_cc_count'], 2)
        self.assertTrue(any('duplicados' in w for w in result['warnings']))

    def test_invalid_phone_amount_pdf_and_missing_month_fail_clearly(self):
        for old, new, message in [('951-234-567', '123', 'teléfono'),
                                  ('Plan variable S/ 10.00', 'Plan variable S/ error', 'monto'),
                                  ('Movistar', 'Otro operador', 'Movistar'),
                                  ('Enero 2027', 'sin fecha', 'mes y año')]:
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                self.process(pdf_text(**{old: new}))
        with self.assertRaisesRegex(ValueError, 'monto'):
            engine.to_decimal('error')
        self.assertEqual(engine.clean_phone(51951234567.0), '951234567')

    def test_rounding_components_are_netted_and_not_counted_twice(self):
        text = pdf_text().replace('Redondeo S/ -0.02',
                                  'Redondeo del mes Actual S/-0.05\nRedondeo del mes Anterior S/0.03')
        result = self.process(text)
        self.assertTrue(result['cuadra'])
        self.assertEqual(result['redondeo'], -0.02)

    def test_optional_mismatches_and_bad_password_do_not_change_pdf_totals(self):
        fixtures = movistar_fixtures()
        result = self.process(fact_bytes=fixtures['fact_file'], traf_bytes=fixtures['traf_file'],
                              cargos_bytes=fixtures['cargos_file'], password='test-only')
        self.assertTrue(result['cuadra'])
        self.assertTrue(any('Plan diferente' in w for w in result['optional_warnings']))
        with patch.object(engine, 'decrypt_excel', side_effect=ValueError('secret-value')):
            result = self.process(fact_bytes=b'broken', password='secret-value')
        self.assertTrue(result['cuadra'])
        self.assertNotIn('secret-value', str(result['optional_warnings']))


if __name__ == '__main__':
    unittest.main(verbosity=2)
