"""Pies CFF mal decodificados y conciliación completa entre páginas.

Los documentos y expectativas privados se configuran en un manifiesto local
ignorado por Git; las regresiones sintéticas se ejecutan también en CI.
"""
from contextlib import ExitStack
from decimal import Decimal
import io
import json
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import openpyxl
import movistar_engine as engine
from test_movistar_dynamic import inventory, pdf_text


FOOTERS = [
    'Integratel Perú S.A.A.', 'Integratel Perœ S.A.A.', 'Integratel Per',
    'Integratel Perú - R.U.C. 20100017491', 'R.U.C. 20100017491',
    'Dirección: Jr. Domingo Martínez 445', 'Página 2 / 20', 'Pagina 2 de 20',
    'PÆgina 2 / 20', 'PÃ¡gina 2 / 20', 'P\ufffdgina 2 / 20',
    'Empresas Móvil', 'Empresas Movil', 'Importe S/', 'Importe S/.',
]


class MovistarPageBreakTests(unittest.TestCase):
    def test_recognizes_separate_and_combined_footer_lines(self):
        for line in FOOTERS:
            with self.subTest(line=line):
                self.assertTrue(engine._is_page_header(line))
        for line in ['Plan variable S/ 10.00', 'Descuento 50% al plan x 3m S/-5.00',
                     'Cargos Mensuales: Otro plan S/ 30.00', 'Anexo 2: 952345678']:
            self.assertFalse(engine._is_page_header(line))

    def process_pages(self, pages):
        extracted = SimpleNamespace(pages=[SimpleNamespace(extract_text=lambda text=text: text)
                                          for text in pages])
        with patch.object(engine.pypdf, 'PdfReader', return_value=extracted):
            result = engine.process_movistar_files(b'%PDF-synthetic-extraction', inventory())
        self.assertTrue(result['cuadra'])
        self.assertEqual(result['line_count'], 2)
        self.assertEqual(result['diferencia'], 0)
        self.assertEqual(result['total_calculado'], 53.28)
        self.assertEqual(result['lines_sample'][0]['desc'], 5)
        self.assertEqual(result['lines_sample'][1]['plan'], 'Otro plan')
        wb = openpyxl.load_workbook(io.BytesIO(result['excel_bytes']))
        self.assertEqual(wb.sheetnames, ['CARGOS ADICIONALES', 'CARGO MENSUAL', 'Hoja1'])
        self.assertEqual(wb['CARGO MENSUAL']['F2'].value, '=D2-E2')
        self.assertTrue(wb['Hoja1']._pivots)
        wb.close()
        return result

    def test_footer_after_completed_charge_does_not_start_pending_charge(self):
        first, second = pdf_text().split('Anexo 2:')
        for footer in FOOTERS:
            with self.subTest(footer=footer):
                result = self.process_pages([first+'\n'+footer, 'Anexo 2:'+second])
                self.assertEqual(result['lines_sample'][0]['plan'], 'Plan variable')

    def test_unfinished_charge_and_discount_survive_page_headers(self):
        footer = '\n'.join(FOOTERS)
        for end in ['Otro plan', 'Descuento 50% al plan x 3m']:
            first, second = pdf_text().split(end)
            result = self.process_pages([first+end+'\n'+footer,
                                         'Empresas Móvil\nImporte S/\n'+second.lstrip()])
            self.assertEqual(result['lines_sample'][0]['plan'], 'Plan variable')

    def test_section_remainder_is_checked_before_appending_pending(self):
        text = pdf_text().replace('Cargos Mensuales: Otro plan',
                                 'Cargos Mensuales: Empresas Móvil\nCargos Mensuales: Otro plan')
        self.process_pages([text])


class MovistarRealPageBreakTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        manifest = Path(os.environ.get('MOVISTAR_REAL_CASES',
                        Path(__file__).resolve().parent/'data/movistar-page-breaks/cases.json'))
        if not manifest.is_file():
            raise unittest.SkipTest('Manifiesto privado de octubre/noviembre no configurado.')
        cls.cases = json.loads(manifest.read_text(encoding='utf-8'))

    def check_month(self, month):
        case = self.cases.get(month)
        if not case or not all(Path(case[key]).is_file() for key in ('pdf', 'inventory')):
            self.skipTest(f'Documentos privados de {month} no disponibles.')
        for fallback in (False, True):
            with self.subTest(month=month, without_fonttools=fallback), ExitStack() as stack:
                if fallback:
                    stack.enter_context(patch.dict(sys.modules, {'fontTools.cffLib': None}))
                result = engine.process_movistar_files(Path(case['pdf']).read_bytes(),
                                                       Path(case['inventory']).read_bytes())
                self.assertEqual(result['line_count'], case['line_count'])
                line = next(item for item in result['lines_sample'] if item['anx'] == case['anexo'])
                self.assertEqual(line['phone'], case['phone'])
                self.assertEqual(Decimal(str(line['cargos_mensuales'])), Decimal(case['charge']))
                # Los descuentos incorporados siguen en MONTO, no se duplican en DESC.
                discount = Decimal(str(line['desc'])) + Decimal(str(line['desc_18m']))
                self.assertEqual(-discount, Decimal(case['discount']))
                self.assertTrue(result['cuadra'])
                self.assertEqual(result['diferencia'], 0)
                self.assertEqual(result['total_calculado'], result['pdf_total'])
                wb = openpyxl.load_workbook(io.BytesIO(result['excel_bytes']))
                cached = openpyxl.load_workbook(io.BytesIO(result['excel_bytes']), data_only=True)
                self.assertEqual(wb.sheetnames, ['CARGOS ADICIONALES', 'CARGO MENSUAL', 'Hoja1'])
                row = next(row for row in wb['CARGO MENSUAL'].iter_rows(min_row=2)
                           if row[1].value == case['phone'])
                self.assertEqual(row[5].value, f'=D{row[0].row}-E{row[0].row}')
                self.assertAlmostEqual(cached['CARGO MENSUAL'].cell(row[0].row, 6).value, line['monto_f'])
                self.assertEqual(wb['Hoja1']._pivots[0].cache.recordCount, case['line_count'])
                wb.close()
                cached.close()

    def test_november_complete_reconciliation_and_excel(self):
        self.check_month('november')

    def test_october_complete_reconciliation_and_excel(self):
        self.check_month('october')


if __name__ == '__main__':
    unittest.main(verbosity=2)
