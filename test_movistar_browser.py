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
        page.goto(self.base + "/movistar.html")
        for key, path in files.items():
            input_name = {"fact_file": "fact", "traf_file": "traf", "cargos_file": "cargos",
                          "inventory_file": "inventory", "pdf_file": "pdf"}[key]
            page.locator("#input-" + input_name + "-file").set_input_files(str(path))
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
        self.assertEqual(page.locator("#banner-pill-badge").inner_text(), "CUADRA")
        self.assertTrue(result["cuadra"])
        with page.expect_download() as pending_download:
            page.locator("#btn-export-excel").click()
        path = OUT / name
        pending_download.value.save_as(path)
        wb = openpyxl.load_workbook(path)
        self.assertEqual(wb.sheetnames, ["CARGOS ADICIONALES", "CARGO MENSUAL", "Hoja1"])
        self.assertEqual(wb["CARGO MENSUAL"]["F3"].value, "=D3-E3")
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
