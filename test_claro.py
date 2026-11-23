"""Tres entradas, conservación del proveedor y prorrateo decimal de Claro."""
import io
import json
import os
from pathlib import Path
import threading
import unittest
import urllib.error
import urllib.request
import zipfile
from decimal import Decimal
from xml.etree import ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill
import claro_engine as claro
from claro_billing import apportion, read_billing
from test_support import simple_pdf, workbook_bytes, multipart


def fixtures(company='PRODIMIN', adjustment=True, missing=False, duplicate=False, total=None):
    inv = openpyxl.Workbook()
    inv.remove(inv.active)
    for label in ['PRODAC', 'PRODIMIN']:
        sheet = inv.create_sheet(label)
        sheet.append(['ID', 'Usuario', 'Numero', '', '', '', '', 'CC Dpto'])
        sheet.append([1, 'Usuario A', '+51 951 234 567', '', '', '', '', 'CC-A' if label == company else 'CC-OTRA'])
        if not missing:
            sheet.append([2, 'Usuario B', 952345678, '', '', '', '', 'CC-B'])
    fact = openpyxl.Workbook()
    sheet = fact.active
    sheet.title = 'Factura Detallada'
    sheet.append([company])
    sheet.merge_cells('A1:D1')
    sheet['A1'].font = Font(name='Arial', bold=True, color='FF0000')
    sheet['A1'].fill = PatternFill('solid', fgColor='FFFF00')
    sheet.append(['Línea', 'Plan', 'Total Cargos por Línea (Sin IGV)', 'Otros Cargos y Abonos'])
    sheet.append(['TOTAL', '', 999, 0])
    sheet.append(['951234567', 'Plan A', 12.50, 0])
    sheet.append(['952345678', 'Plan B', 7.50, 0])
    if duplicate:
        sheet.append(['951234567', 'Servicio adicional', 1.25, 0])
    if adjustment:
        sheet.append(['Varios', 'Varios', -1.01, -1.01])
    sheet.column_dimensions['A'].width = 25
    sheet.freeze_panes = 'C4'
    old = fact.create_sheet('Hoja1')
    old.append(['Resultado anterior que se recalcula'])
    base = Decimal('20.00') + (Decimal('1.25') if duplicate else 0)
    amount = base - (Decimal('1.01') if adjustment else 0)
    if total is not None:
        amount = Decimal(str(total))
    text = f'Claro {company}\nRUC: 20123456789\nRecibo F001-00001234\nPeriodo: Marzo 2027\nCargos del Mes {amount:.2f}\nIGV 3.42\nTotal a pagar 22.41'
    if adjustment:
        text += '\nOTROS CARGOS Y ABONOS\nDescripcion Importe\nDev. por interrupcion -1.01\nTotal de Otros Cargos y Abonos -1.01'
    text += '\nOTROS CARGOS NO AFECTOS A I.G.V.\nInteres legal -0.55'
    return {'pdf_file': simple_pdf(text), 'inventory_file': workbook_bytes(inv), 'fact_file': workbook_bytes(fact)}


def source_sheet(data):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        sheet = next(x for x in ET.fromstring(z.read('xl/workbook.xml')).find('{'+ns+'}sheets') if x.get('name') == 'Factura Detallada')
        rid = sheet.get('{'+rel+'}id')
        target = next(x.get('Target') for x in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')) if x.get('Id') == rid)
        path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        return z.read(path)


class ClaroTests(unittest.TestCase):
    def process(self, company='PRODIMIN', **kwargs):
        f = fixtures(company=company, **kwargs)
        return claro.process_claro_files(f['pdf_file'], f['inventory_file'], f['fact_file'], company)

    def test_three_files_required(self):
        f = fixtures()
        with self.assertRaisesRegex(ValueError, 'tres archivos'):
            claro.process_claro_files(f['pdf_file'], f['inventory_file'])
        with self.assertRaisesRegex(ValueError, 'Seleccione'):
            claro.process_claro_files(f['pdf_file'], f['inventory_file'], f['fact_file'])

    def test_general_adjustment_is_distributed_once(self):
        r = self.process()
        self.assertEqual(r['suma_lineas'], 20)
        self.assertEqual(r['ajustes_abonos'], -1.01)
        self.assertEqual(r['total_calculado'], 18.99)
        self.assertEqual(r['diferencia'], 0)
        self.assertEqual(r['estado'], 'CUADRA')
        self.assertEqual(sum(Decimal(str(x['prorrateo'])) for x in r['summary']), Decimal('-1.01'))
        self.assertEqual([x['prorrateo'] for x in r['details']], [-.63, -.38])
        self.assertEqual([x['cc'] for x in r['details']], ['CC-A', 'CC-B'])
        self.assertEqual(len(r['details']), 2)
        self.assertTrue(all(x['numero'] != 'Varios' for x in r['details']))

    def test_no_adjustment_and_duplicates(self):
        r = self.process(adjustment=False, duplicate=True)
        self.assertEqual(r['numeros_unicos'], 2)
        self.assertEqual(r['registros_cargos'], 3)
        self.assertEqual(r['total_calculado'], 21.25)
        self.assertEqual(r['ajustes_abonos'], 0)
        self.assertTrue(r['cuadra'])

    def test_missing_cc_is_preserved(self):
        r = self.process(adjustment=False, missing=True)
        self.assertEqual(r['lineas_sin_cc'], 1)
        self.assertEqual(r['details'][1]['cc'], 'SIN CC - REVISAR')
        self.assertTrue(r['cuadra'])
        r = self.process(missing=True)
        self.assertEqual(r['details'][1]['final_amount'], 7.5)
        self.assertEqual(r['total_calculado'], 18.99)

    def test_selected_company_validated_with_pdf(self):
        f = fixtures()
        with self.assertRaisesRegex(ValueError, 'seleccionada'):
            claro.process_claro_files(f['pdf_file'], f['inventory_file'], f['fact_file'], 'PRODAC')
        self.assertEqual(self.process(company='PRODAC')['company'], 'PRODAC')

    def test_varios_must_exist_in_original(self):
        f = fixtures()
        no_adjustment = fixtures(adjustment=False)['fact_file']
        with self.assertRaisesRegex(ValueError, 'no coinciden'):
            claro.process_claro_files(f['pdf_file'], f['inventory_file'], no_adjustment, 'PRODIMIN')

    def test_original_preserved_and_exact_sheet_order(self):
        f = fixtures()
        r = claro.process_claro_files(f['pdf_file'], f['inventory_file'], f['fact_file'], 'PRODIMIN')
        self.assertEqual(source_sheet(f['fact_file']), source_sheet(r['excel_bytes']))
        wb = openpyxl.load_workbook(io.BytesIO(r['excel_bytes']))
        self.assertEqual(wb.sheetnames, ['Factura Detallada', 'Hoja2', 'Hoja1'])
        self.assertEqual(wb['Factura Detallada']['A1'].font.color.rgb, '00FF0000')
        self.assertEqual(wb['Factura Detallada'].freeze_panes, 'C4')
        self.assertEqual(wb['Factura Detallada']['C6'].value, -1.01)

    def test_two_pivots_have_independent_structured_caches(self):
        r = self.process()
        wb = openpyxl.load_workbook(io.BytesIO(r['excel_bytes']))
        for name, table, count in [('Hoja2','ClaroTrabajo',3), ('Hoja1','ClaroFinal',2)]:
            sheet = wb[name]
            pivot = sheet._pivots[0]
            self.assertEqual(pivot.cache.cacheSource.worksheetSource.name, table)
            self.assertTrue(pivot.cache.refreshOnLoad)
            self.assertEqual(pivot.cache.recordCount, count)
            self.assertEqual(pivot.cache.records.count, count)
            self.assertEqual(sheet.tables[table].ref, f'A1:C{count+1}')
            self.assertEqual(sheet['G7'].value, 'Etiquetas de fila')
            self.assertEqual(pivot.location.ref, f'G7:H{pivot.cache.cacheFields[2].sharedItems.count + 8}')
            self.assertIsNone(sheet.tables[table].autoFilter)
            self.assertFalse(sheet.tables[table].tableStyleInfo.name)
            for cells in sheet.iter_rows(min_col=1, max_col=3, max_row=count+1):
                for cell in cells:
                    self.assertFalse(cell.font.bold)
                    self.assertEqual(cell.font.color.rgb, '00000000')
                    self.assertIsNone(cell.fill.patternType)
                    self.assertFalse(any(edge and edge.style for edge in
                                         [cell.border.left, cell.border.right, cell.border.top, cell.border.bottom]))
            self.assertEqual(sheet['G7'].fill.fgColor.rgb, '00C6E7F4')
        self.assertNotEqual(wb['Hoja2']._pivots[0].cacheId, wb['Hoja1']._pivots[0].cacheId)
        self.assertEqual(wb['Hoja2']['I7'].value, 'prorrateo')
        self.assertEqual(wb['Hoja2']['J7'].value, 'total ajustado')
        self.assertFalse(wb['Hoja2']['I7'].font.bold)
        self.assertIsNone(wb['Hoja2']['I7'].fill.patternType)
        self.assertNotIn('[Red]', wb['Hoja2']['I8'].number_format)
        self.assertIsNone(wb['Hoja1']['I7'].value)
        values = openpyxl.load_workbook(io.BytesIO(r['excel_bytes']), data_only=True)
        self.assertAlmostEqual(values['Hoja2']['I11'].value, -1.01)
        self.assertAlmostEqual(values['Hoja2']['J11'].value, 18.99)
        self.assertAlmostEqual(values['Hoja1']['H10'].value, 18.99)

    def test_rounding_residue_on_last_cc_and_last_line(self):
        self.assertEqual(apportion(Decimal('-0.01'), [Decimal('1')]*3), [Decimal('0'),Decimal('0'),Decimal('-0.01')])
        self.assertEqual(sum(apportion(Decimal('-1.01'), [Decimal('1')]*7)), Decimal('-1.01'))
        r = self.process(duplicate=True)
        for group in r['summary']:
            rows = [x for x in r['details'] if x['cc'] == group['cc']]
            self.assertEqual(sum(Decimal(str(x['prorrateo'])) for x in rows), Decimal(str(group['prorrateo'])))

    def test_cent_tolerance_and_invalid_files(self):
        self.assertTrue(self.process(total=19)['cuadra'])
        self.assertFalse(self.process(total=19.01)['cuadra'])
        f = fixtures()
        with self.assertRaises(ValueError):
            claro.process_claro_files(b'corrupto', f['inventory_file'], f['fact_file'], 'PRODIMIN')

    def test_rerun_preserves_original_and_replaces_derived_sheets(self):
        f = fixtures()
        one = claro.process_claro_files(f['pdf_file'], f['inventory_file'], f['fact_file'], 'PRODIMIN')
        two = claro.process_claro_files(f['pdf_file'], f['inventory_file'], one['excel_bytes'], 'PRODIMIN')
        self.assertEqual(source_sheet(f['fact_file']), source_sheet(two['excel_bytes']))
        self.assertEqual(openpyxl.load_workbook(io.BytesIO(two['excel_bytes'])).sheetnames, ['Factura Detallada','Hoja2','Hoja1'])
        self.assertEqual(two['total_calculado'], one['total_calculado'])

    def test_dynamic_values_not_in_production(self):
        for name in ['claro_engine.py','claro_billing.py','claro_build.py','claro_package.py','js/app.js']:
            code = Path(name).read_text(encoding='utf-8')
            for value in ['5348.44', '5341.07', '5762.44', 'Pmin - Jul']:
                self.assertNotIn(value, code)


class ClaroApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import server
        cls.httpd = server.ThreadedServer(('127.0.0.1',0), server.AppCelularesHTTPHandler)
        threading.Thread(target=cls.httpd.serve_forever,daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.httpd.server_address[1]}'
    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown(); cls.httpd.server_close()
    def test_three_file_api_and_download(self):
        body, kind = multipart(fixtures(), {'company':'PRODIMIN'})
        with urllib.request.urlopen(urllib.request.Request(self.base+'/api/claro/process',body,{'Content-Type':kind})) as response:
            r = json.load(response)
        self.assertTrue(r['cuadra'])
        with urllib.request.urlopen(self.base+'/api/claro/download-excel?token='+r['download_token']) as response:
            self.assertEqual(openpyxl.load_workbook(io.BytesIO(response.read())).sheetnames, ['Factura Detallada','Hoja2','Hoja1'])
    def test_missing_excel_rejected(self):
        files = fixtures(); del files['fact_file']
        body, kind = multipart(files, {'company':'PRODIMIN'})
        with self.assertRaises(urllib.error.HTTPError) as error:
            urllib.request.urlopen(urllib.request.Request(self.base+'/api/claro/process',body,{'Content-Type':kind}))
        self.assertEqual(error.exception.code,400)
        error.exception.close()


class ClaroRealAcceptance(unittest.TestCase):
    def real(self, month):
        paths = [os.environ.get('CLARO_'+month+'_PDF'), os.environ.get('CLARO_INVENTORY'), os.environ.get('CLARO_'+month+'_REFERENCE')]
        if not all(paths):
            self.skipTest('Faltan los tres documentos reales de '+month)
        original = Path(paths[2]).read_bytes()
        r = claro.process_claro_files(Path(paths[0]).read_bytes(),Path(paths[1]).read_bytes(),original,'PRODIMIN')
        self.assertEqual(source_sheet(original),source_sheet(r['excel_bytes']))
        out = Path('data/claro-results'); out.mkdir(parents=True,exist_ok=True)
        (out/r['download_filename']).write_bytes(r.pop('excel_bytes'))
        (out/(month+'-restored-validation.json')).write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
        return r
    def test_july_real(self):
        r = self.real('JULY')
        for key,value in {'numeros_unicos':143,'registros_cargos':143,'suma_lineas':5348.44,'ajustes_abonos':-7.37,'total_calculado':5341.07,'diferencia':0,'lineas_sin_cc':0}.items():
            self.assertEqual(r[key],value,key)
        self.assertEqual(r['estado'],'CUADRA')
    def test_october_real(self):
        r = self.real('OCTOBER')
        self.assertEqual(r['numeros_unicos'],150)
        self.assertEqual(r['registros_cargos'],151)
        self.assertEqual(r['total_calculado'],5762.44)
        self.assertEqual(r['diferencia'],0)
        self.assertEqual(r['estado'],'CUADRA')
    def test_october_excel_records(self):
        path = os.environ.get('CLARO_OCTOBER_REFERENCE')
        if not path:
            self.skipTest('Excel real de octubre no configurado.')
        rows, general, _ = read_billing(Path(path).read_bytes())
        self.assertEqual(len(rows),151)
        self.assertEqual(len({r['numero'] for r in rows}),150)
        from collections import Counter
        repeated = [count for count in Counter(r['numero'] for r in rows).values() if count > 1]
        self.assertEqual(repeated,[2])
        self.assertEqual(sum((r['amount'] for r in rows+general),Decimal('0')),Decimal('5762.44'))

if __name__ == '__main__': unittest.main(verbosity=2)
