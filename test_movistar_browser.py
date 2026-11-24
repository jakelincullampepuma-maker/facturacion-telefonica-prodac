"""Carga, procesamiento y descarga de Movistar en un navegador real."""
import io
import json
import os
from pathlib import Path
import sys
import threading
import unittest
import openpyxl

ROOT = Path(__file__).resolve().parent
if (ROOT / ".test-tools").is_dir():
    sys.path.insert(0, str(ROOT / ".test-tools"))
from playwright.sync_api import sync_playwright
import server
from test_support import movistar_fixtures

OUT = ROOT / "data/browser-results"
PRIVATE_KEYS = ["MOVISTAR_FACT_PATH", "MOVISTAR_TRAF_PATH", "MOVISTAR_CARGOS_PATH", "MOVISTAR_INVENTORY_PATH", "MOVISTAR_PDF_PATH"]


class MovistarBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = os.environ.get("APP_TEST_BASE_URL", "").rstrip("/")
        cls.httpd = None
        if not cls.base:
            cls.httpd = server.ThreadedServer(("127.0.0.1", 0), server.AppCelularesHTTPHandler)
            cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
            cls.thread.start()
            cls.base = f"http://127.0.0.1:{cls.httpd.server_address[1]}"
        OUT.mkdir(parents=True, exist_ok=True)
        cls.pw = sync_playwright().start()
        chrome = Path("C:/Program Files/Google/Chrome/Application/chrome.exe")
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None, headless=True)
        cls.context = cls.browser.new_context(accept_downloads=True)

    @classmethod
    def tearDownClass(cls):
        cls.context.close()
        cls.browser.close()
        cls.pw.stop()
        if cls.httpd:
            cls.httpd.shutdown()
            cls.httpd.server_close()

    def process(self, files, password, name):
        page = self.context.new_page()
        page.set_default_timeout(240000)
        page.goto(self.base + "/movistar")
        if any(k in files for k in ('fact_file', 'traf_file', 'cargos_file')):
            page.locator('#details-optional-val').evaluate('(e) => e.open = true')
        for key, path in files.items():
            input_name = {"fact_file": "fact", "traf_file": "traf", "cargos_file": "cargos",
                          "inventory_file": "inventory", "pdf_file": "pdf"}[key]
            page.locator("#input-" + input_name + "-file").set_input_files(str(path))
        if password:
            page.locator("#input-movistar-password").fill(password)
        with page.expect_response(lambda r: r.url.endswith("/api/movistar/process"), timeout=240000) as pending:
            page.locator("#btn-process-files").click()
        response = pending.value
        self.assertEqual(response.status, 200)
        result = response.json()
        self.assertTrue(result["success"])
        page.wait_for_function("!document.getElementById('btn-export-excel').disabled")
        self.assertEqual(page.locator("#input-movistar-password").input_value(), "")
        self.assertEqual(page.locator("#kpi-val-lineas").inner_text(), str(result["line_count"]))
        self.assertEqual(page.locator("#banner-pill-badge").inner_text(), "REVISAR" if result.get('sin_cc_count') or result.get('warnings') or result.get('optional_warnings') else "CUADRA")
        self.assertTrue(result["cuadra"])
        with page.expect_download() as pending_download:
            page.locator("#btn-export-excel").click()
        path = OUT / name
        pending_download.value.save_as(path)
        wb = openpyxl.load_workbook(path)
        self.assertEqual(wb.sheetnames, ["CARGOS ADICIONALES", "CARGO MENSUAL", "Hoja1"])
        self.assertEqual(wb["CARGO MENSUAL"]["F2"].value, "=D2-E2")
        page.close()
        return result

    def test_synthetic_upload_process_download(self):
        paths = {}
        for key, contents in movistar_fixtures().items():
            path = OUT / ("movistar-synthetic-" + key + (".pdf" if key == "pdf_file" else ".xlsx"))
            path.write_bytes(contents)
            paths[key] = path
        result = self.process(paths, "synthetic-test-only", "Movistar-control-result.xlsx")
        self.assertEqual(result["line_count"], 2)
        self.assertEqual(result["discount_count"], 1)
        self.assertEqual(result["diferencia"], 0)
        self.assertAlmostEqual(result["total_pdf"], 53.30)

    @unittest.skipUnless(all((ROOT / 'data/movistar_nov_test' / name).is_file() for name in ('Movistar-Nov.pdf', 'Inventario Celulares General.xlsx')), 'Referencias privadas de noviembre no disponibles.')
    def test_november_pdf_inventory_download(self):
        directory = ROOT / 'data/movistar_nov_test'
        result = self.process({'pdf_file': directory / 'Movistar-Nov.pdf',
                               'inventory_file': directory / 'Inventario Celulares General.xlsx'},
                              '', 'Movistar_Conciliado_Noviembre_2026.xlsx')
        self.assertEqual(result['line_count'], 325)
        for key, value in {'cargo_mensual_sin_igv': 12597.72, 'descuento_total': 1046.68,
                           'cargo_final_sin_igv': 11551.04, 'subtotal_pdf': 11551.04,
                           'igv': 2079.19, 'cargos_inafectos': 25452.11, 'redondeo': 0.06,
                           'total_calculado': 39082.40, 'pdf_total': 39082.40, 'diferencia': 0}.items():
            self.assertAlmostEqual(result[key], value, places=2)
        path = OUT / 'Movistar_Conciliado_Noviembre_2026.xlsx'
        wb = openpyxl.load_workbook(path)
        values = openpyxl.load_workbook(path, data_only=True)
        self.assertEqual(len(wb['Hoja1']._pivots), 1)
        self.assertEqual(wb['Hoja1']._pivots[0].cache.recordCount, 325)
        self.assertEqual(wb['Hoja1']._pivots[0].cache.records.count, 325)
        self.assertEqual(wb['Hoja1'].tables['MovistarDatos'].autoFilter.ref, 'A1:C326')
        self.assertEqual(wb['Hoja1']._pivots[0].cache.cacheSource.worksheetSource.name, 'MovistarDatos')
        self.assertEqual(wb['CARGO MENSUAL']['F326'].value, '=D326-E326')
        self.assertEqual(values['CARGOS ADICIONALES']['C94'].value, 25452.11)
        self.assertEqual(values['CARGO MENSUAL']['F327'].value, 11551.04)
        self.assertEqual(values['Hoja1']['B327'].value, 11551.04)
        metrics = {k: result[k] for k in ('periodo', 'line_count', 'cargo_mensual_sin_igv',
                   'descuento_total', 'cargo_final_sin_igv', 'subtotal_pdf', 'igv', 'total_afecto',
                   'cargos_inafectos', 'redondeo', 'total_calculado', 'pdf_total', 'diferencia',
                   'estado', 'sin_cc_count', 'discount_breakdown', 'inafectos_breakdown')}
        (OUT / 'movistar-november-validation.json').write_text(json.dumps(metrics, indent=2), encoding='utf-8')

    def test_layout_and_invalidated_results(self):
        paths = {}
        for key, contents in movistar_fixtures().items():
            if key not in ('pdf_file', 'inventory_file'):
                continue
            path = OUT / ('movistar-layout-' + key + ('.pdf' if key == 'pdf_file' else '.xlsx'))
            path.write_bytes(contents)
            paths[key] = path
        page = self.context.new_page()
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto(self.base + '/movistar')
        self.assertFalse(page.locator('#details-optional-val').evaluate('e => e.open'))
        self.assertTrue(page.locator('#btn-process-files').is_disabled())
        for key, name in [('pdf_file', 'pdf'), ('inventory_file', 'inventory')]:
            page.locator(f'#input-{name}-file').set_input_files(str(paths[key]))
        page.locator('#btn-process-files').click()
        page.wait_for_function("!document.getElementById('btn-export-excel').disabled")
        for width, height in [(1366, 768), (1280, 720), (1024, 600)]:
            page.set_viewport_size({'width': width, 'height': height})
            self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
        page.screenshot(path=str(OUT / 'movistar-results.png'), full_page=True)
        page.locator('#input-pdf-file').set_input_files([])
        page.locator('#input-pdf-file').set_input_files(str(paths['pdf_file']))
        self.assertTrue(page.locator('#btn-export-excel').is_disabled())
        self.assertTrue(page.locator('#section-resultados-movistar').is_hidden())
        page.locator('#btn-clear-files').click()
        self.assertEqual(page.locator('#input-inventory-file').input_value(), '')
        self.assertFalse(errors)
        page.close()

    @unittest.skipUnless(not os.environ.get("APP_TEST_BASE_URL") and os.environ.get("MOVISTAR_TEST_PASSWORD") and all(os.environ.get(k) for k in PRIVATE_KEYS),
                         "Prueba privada Movistar disponible sólo en servidor local con referencias configuradas.")
    def test_private_september_upload_process_download(self):
        files = dict(zip(["fact_file", "traf_file", "cargos_file", "inventory_file", "pdf_file"], [os.environ[k] for k in PRIVATE_KEYS]))
        result = self.process(files, os.environ["MOVISTAR_TEST_PASSWORD"], "Movistar-private-result.xlsx")
        self.assertEqual(result["line_count"], 318)
        self.assertEqual(result["discount_count"], 64)
        self.assertEqual(result["sin_cc_count"], 0)
        self.assertEqual(result["diferencia"], 0)
        self.assertAlmostEqual(result["total_pdf"], 31502.60)
        # Sólo métricas de validación en el archivo local, sin números ni usuarios.
        metrics = {k: result[k] for k in ["line_count", "discount_count", "sin_cc_count", "diferencia", "total_pdf"]}
        (OUT / "movistar-validation.json").write_text(json.dumps(metrics), encoding="utf-8")


if __name__ == "__main__":
    unittest.main(verbosity=2, testRunner=unittest.TextTestRunner(stream=sys.stdout, verbosity=2))
