"""Prueba segura local o pública: genera todos los documentos sintéticamente.

APP_PUBLIC_TEST_BASE_URL habilita la URL remota explícita. Nunca lee documentos
privados, rutas de Downloads ni contraseñas. Las evidencias quedan en data/.
"""
import io
import json
import os
from pathlib import Path
import sys
import threading
import unittest
import urllib.request

import openpyxl

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / '.test-tools'))
from playwright.sync_api import sync_playwright
import server
from build_views import VIEWS
from test_claro import fixtures as claro_fixtures
from test_support import movistar_fixtures, simple_pdf, workbook_bytes


def synthetic_entel():
    bill = openpyxl.Workbook()
    bill.active.title = 'Detalle sintético'
    bill.active.append(['Numero', 'Plan', 'Cargo Fijo'])
    bill.active.append(['951234567', 'Plan A', 12.50])
    bill.active.append(['952345678', 'Plan B', 7.50])
    inv = openpyxl.Workbook()
    inv.active.title = 'PRODAC'
    inv.active.append(['Numero', 'Operador', 'CC Dpto', 'Usuario'])
    inv.active.append(['951234567', 'Entel', 'CC-A', 'Usuario de prueba A'])
    inv.active.append(['952345678', 'Entel', 'CC-B', 'Usuario de prueba B'])
    pdf = simple_pdf('Entel - documento sintetico\n01/Nov/2027 01/Oct/2027 31/Oct/2027\nPLAN Prorrateo Cargos Descuentos Equipos Saldo\nS/20.00 S/0.00 S/0.00 S/0.00 S/0.00 S/20.00\nTotal a pagar S/20.00')
    return {'excel': workbook_bytes(bill), 'inventory': workbook_bytes(inv), 'pdf': pdf}


class SafeReleaseBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        public = os.environ.get('APP_PUBLIC_TEST_BASE_URL', '').rstrip('/')
        cls.httpd = None
        if public:
            cls.base = public
        else:
            cls.httpd = server.ThreadedServer(('127.0.0.1', 0), server.AppCelularesHTTPHandler)
            threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()
            cls.base = f'http://127.0.0.1:{cls.httpd.server_address[1]}'
        cls.out = ROOT / 'data' / ('release-public' if public else 'release-local')
        cls.out.mkdir(parents=True, exist_ok=True)
        cls.report = {'base_url': cls.base, 'synthetic_inputs_only': True, 'operators': {}}
        cls.pw = sync_playwright().start()
        chrome = Path('C:/Program Files/Google/Chrome/Application/chrome.exe')
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None, headless=True)

    @classmethod
    def tearDownClass(cls):
        (cls.out/'verification.json').write_text(json.dumps(cls.report, ensure_ascii=False, indent=2), encoding='utf-8')
        cls.browser.close()
        cls.pw.stop()
        if cls.httpd:
            cls.httpd.shutdown()
            cls.httpd.server_close()

    def test_health(self):
        with urllib.request.urlopen(self.base+'/health', timeout=120) as response:
            health = json.load(response)
        self.assertEqual(health['status'], 'ok')
        self.assertEqual(set(health['modules']), {'claro', 'movistar', 'entel'})
        expected = os.environ.get('EXPECTED_RELEASE_COMMIT')
        if expected:
            self.assertEqual(health['commit'], expected)
        self.report['health'] = health

    def test_clean_routes_redirects_and_navigation(self):
        with self.browser.new_context() as context:
            query = '?source=smoke&x=a%2Fb&x=dos+valores'
            aliases = {'/': '/claro', '/index.html': '/claro'}
            for operator in VIEWS:
                aliases[f'/{operator}.html'] = f'/{operator}'
                aliases[f'/{operator}/'] = f'/{operator}'
            for source, destination in aliases.items():
                for suffix in ['', query]:
                    response = context.request.get(self.base+source+suffix, max_redirects=0)
                    self.assertEqual(response.status, 301)
                    self.assertEqual(response.headers['location'], destination+suffix)
                    self.assertEqual(response.body(), b'')
            page = context.new_page()
            page.set_default_timeout(120000)
            errors, failed_resources = [], []
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.on('response', lambda r: failed_resources.append(r.url) if r.url.startswith(self.base) and r.status >= 400 else None)
            page.goto(self.base+'/?source=home')
            self.assertEqual(page.url, self.base+'/claro?source=home')
            for operator in VIEWS:
                page.locator(f'.operator-link[href="/{operator}"]').click()
                self.assertEqual(page.url, self.base+f'/{operator}')
                self.assertEqual(page.locator('.operator-link').evaluate_all('nodes=>nodes.map(n=>n.getAttribute("href"))'), ['/claro', '/movistar', '/entel'])
                self.assertEqual(page.locator('.sidebar-logo').get_attribute('href'), '/claro')
                self.assertTrue(page.locator('link[rel="stylesheet"]').evaluate_all('nodes=>nodes.every(n=>Boolean(n.sheet))'))
                self.assertTrue(page.evaluate('Boolean(window.UploadUI && window.ModuleUI && window.AppSidebar)'))
                self.assertTrue(page.locator('#'+VIEWS[operator]['result']).is_hidden())
            self.assertEqual(errors, [])
            self.assertEqual(failed_resources, [])
            self.report['routes'] = {'permanent_redirects': aliases, 'queries_preserved': True, 'navigation': list(VIEWS), 'resource_errors': failed_resources, 'page_errors': errors}

    def navigation(self, page, download=False):
        for name in ['dashboard', 'upload', 'results', 'sin-cc', 'export']:
            if page.locator('#sidebar-menu-button').is_visible():
                page.locator('#sidebar-menu-button').click()
            if name == 'export' and download:
                with page.expect_download():
                    page.locator('#nav-'+name).click()
            else:
                page.locator('#nav-'+name).click()

    def screenshots(self, page, operator, phase):
        for label, width, height in [('desktop', 1366, 1000), ('mobile', 390, 844)]:
            page.set_viewport_size({'width': width, 'height': height})
            page.evaluate("scrollTo({top:0,behavior:'instant'})")
            self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
            page.screenshot(path=str(self.out/f'{operator}-{phase}-{label}.png'), full_page=True, animations='disabled')
            if label == 'mobile':
                self.navigation(page, download=phase == 'results')
        page.set_viewport_size({'width': 1366, 'height': 1000})

    def case(self, operator, company='PRODIMIN'):
        if operator == 'claro':
            f = claro_fixtures(company)
            files = {'claro': f['fact_file'], 'inv': f['inventory_file'], 'pdf': f['pdf_file']}
            expected = {'numeros_unicos': 2, 'total_calculado': 18.99, 'diferencia': 0}
            sheets = ['Factura Detallada', 'Hoja2', 'Hoja1']
            sheet, pivot_total = 'Hoja1', 18.99
        elif operator == 'movistar':
            f = movistar_fixtures()
            files = {'pdf': f['pdf_file'], 'inventory': f['inventory_file']}
            expected = {'line_count': 2, 'total_calculado': 53.30, 'diferencia': 0}
            sheets = ['CARGOS ADICIONALES', 'CARGO MENSUAL', 'Hoja1']
            sheet, pivot_total = 'Hoja1', 35
        else:
            files = synthetic_entel()
            expected = {'line_count': 2, 'total_excel': 20, 'total_pdf_plan': 20, 'diferencia': 0}
            sheets = ['Resultado Entel', 'Detalle sintético']
            sheet, pivot_total = 'Resultado Entel', 20
        with self.browser.new_context(accept_downloads=True, viewport={'width': 1366, 'height': 1000}) as context:
            page = context.new_page()
            page.set_default_timeout(120000)
            errors, console_errors, posts = [], [], []
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.on('console', lambda msg: console_errors.append(msg.text.split('?')[0]) if msg.type == 'error' else None)
            page.on('request', lambda r: posts.append(r.url.split('?')[0]) if r.method == 'POST' else None)
            page.add_init_script('''(() => {
              window.documentReads = 0;
              for (const name of ['arrayBuffer','text','stream']) {
                const original = Blob.prototype[name];
                Blob.prototype[name] = function(...args) { window.documentReads++; return original.apply(this,args); };
              }
            })();''')
            page.goto(self.base+f'/{operator}')
            self.assertEqual(page.url, self.base+f'/{operator}')
            page.evaluate('document.fonts.ready')
            results = page.locator('#'+VIEWS[operator]['result'])
            self.assertTrue(results.is_hidden())
            self.assertTrue(page.locator('#btn-export-excel').is_disabled())
            self.assertEqual(page.locator('.sidebar-link').all_text_contents(), ['Dashboard', 'Cargar archivos', 'Resultados', 'Sin CC - Revisar', 'Descargar reporte'])
            if operator == 'claro':
                page.locator('#tab-'+company.lower()).click()
            self.navigation(page)
            self.screenshots(page, operator, 'initial')
            def select_files():
                for field, payload in files.items():
                    pdf = field == 'pdf'
                    page.locator('#input-'+field+'-file').set_input_files({'name': 'Sintetico.'+('pdf' if pdf else 'xlsx'), 'mimeType': 'application/pdf' if pdf else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'buffer': payload})
                    self.assertTrue(results.is_hidden())
                    self.assertTrue(page.locator('input[type=file]').evaluate_all('nodes => nodes.every(n => !n.disabled)'))
            select_files()
            self.assertEqual(page.evaluate('window.documentReads'), 0)
            self.assertEqual(posts, [])
            self.assertFalse(page.locator('#btn-process-files').is_disabled())
            diagnostics = page.evaluate('UploadDiagnostics.snapshot()')
            self.assertEqual(diagnostics['runs'], [])
            self.assertTrue(all(n == 1 for n in diagnostics['listeners'].values()))
            self.assertLess(max(s['selection_ms'] for s in diagnostics['selections']), 100)
            # Clear and reload must retain exactly one listener and no eager reads.
            page.locator('#btn-clear-files').click()
            select_files()
            self.assertEqual(page.evaluate('window.documentReads'), 0)
            if operator == 'entel':
                page.evaluate('''() => {const run = EntelProcessor.process; EntelProcessor.process = async (...args) => {const result = await run(...args); window.syntheticResult = result; return result;};}''')
                page.locator('#btn-process-files').click()
                page.wait_for_function('Boolean(window.syntheticResult)')
                result = page.evaluate('(({excel_blob,...r})=>r)(window.syntheticResult)')
                self.assertEqual(posts, [])
            else:
                with page.expect_response(lambda r: r.url.endswith(f'/api/{operator}/process')) as response:
                    page.locator('#btn-process-files').click()
                self.assertEqual(response.value.status, 200)
                result = response.value.json()
                self.assertEqual(len(posts), 1)
            page.wait_for_function("!document.getElementById('btn-export-excel').disabled")
            self.assertTrue(results.is_visible())
            self.assertEqual(page.locator('#banner-pill-badge').inner_text(), 'CUADRA')
            self.assertEqual(page.locator('.conciliation-banner-card:visible').count(), 1)
            for key, value in expected.items():
                self.assertAlmostEqual(result[key], value, places=2)
            self.screenshots(page, operator, 'results')
            self.navigation(page, download=True)
            with page.expect_download() as pending:
                page.locator('#btn-export-excel').click()
            output = self.out/f'{operator}-{company}.xlsx'
            pending.value.save_as(output)
            wb = openpyxl.load_workbook(output, data_only=True)
            self.assertEqual(wb.sheetnames, sheets)
            self.assertTrue(wb[sheet]._pivots)
            total_row = next(row for row in wb[sheet].iter_rows() if row[6].value == 'Total general')
            self.assertAlmostEqual(total_row[7].value, pivot_total, places=2)
            wb.close()
            if operator == 'claro':
                page.locator('#tab-'+('prodac' if company == 'PRODIMIN' else 'prodimin')).click()
                self.assertTrue(results.is_hidden())
            page.locator('#btn-clear-files').click()
            self.assertTrue(results.is_hidden())
            self.assertTrue(page.locator('#btn-export-excel').is_disabled())
            self.assertEqual(page.locator('#period-display-text').inner_text(), 'Periodo: pendiente de detectar')
            self.assertTrue(page.locator('input[type=file]').evaluate_all('nodes=>nodes.every(n=>!n.disabled && !n.files.length)'))
            self.assertEqual(errors, [])
            self.assertEqual(console_errors, [])
            self.report['operators'][operator+'-'+company] = {'values': expected, 'sheets': sheets, 'pivot_total': pivot_total, 'selection_ms': [s['selection_ms'] for s in diagnostics['selections']], 'page_errors': errors, 'console_errors': console_errors, 'passed': True}

    def test_claro_prodac(self): self.case('claro', 'PRODAC')
    def test_claro_prodimin(self): self.case('claro', 'PRODIMIN')
    def test_movistar(self): self.case('movistar')
    def test_entel(self): self.case('entel')


if __name__ == '__main__':
    unittest.main(verbosity=2)
