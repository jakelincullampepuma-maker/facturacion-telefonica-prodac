"""Pruebas portables de servidor, privacidad, caché y API Movistar."""
import datetime
import io
import json
import os
import threading
import unittest
import urllib.error
import urllib.request
from unittest.mock import patch
import openpyxl
import server
from test_support import movistar_fixtures, multipart


class TestReleaseServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = server.ThreadedServer(("127.0.0.1", 0), server.AppCelularesHTTPHandler)
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f"http://127.0.0.1:{cls.httpd.server_address[1]}"

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=5)

    def test_health_and_public_pages(self):
        with patch.dict(os.environ, {"RENDER_GIT_COMMIT": "test-commit"}):
            with urllib.request.urlopen(self.base + "/health") as response:
                data = json.load(response)
            self.assertEqual(data["status"], "ok")
            self.assertEqual(data["commit"], "test-commit")
        for path in ["/", "/claro.html", "/movistar.html", "/entel.html", "/js/entelWorkbook.js", "/css/styles.css"]:
            with urllib.request.urlopen(self.base + path) as response:
                self.assertEqual(response.status, 200)
        req = urllib.request.Request(self.base + "/health", method="HEAD")
        with urllib.request.urlopen(req) as response:
            self.assertEqual(response.status, 200)
            self.assertEqual(response.read(), b"")

    def test_private_paths_are_not_served(self):
        for path in ["/.git/config", "/.env", "/.env.test", "/data/", "/data/inventario_movistar.xlsx",
                     "/server.py", "/requirements.txt", "/test_movistar_acceptance.py", "/js/",
                     "/js/%2e%2e/server.py", "/css/../../.git/config", "/.test-tools/"]:
            for method in ["GET", "HEAD"]:
                with self.subTest(path=path, method=method), self.assertRaises(urllib.error.HTTPError) as ctx:
                    urllib.request.urlopen(urllib.request.Request(self.base + path, method=method))
                self.assertEqual(ctx.exception.code, 404)

    def test_movistar_upload_process_download(self):
        body, content_type = multipart(movistar_fixtures(), {"password": "synthetic-test-only"})
        req = urllib.request.Request(self.base + "/api/movistar/process", data=body, headers={"Content-Type": content_type})
        with urllib.request.urlopen(req) as response:
            result = json.load(response)
        self.assertTrue(result["success"])
        self.assertEqual(result["line_count"], 2)
        self.assertEqual(result["discount_count"], 1)
        self.assertEqual(result["descuento_total"], 5)
        self.assertAlmostEqual(result["total_calculado"], 53.30)
        self.assertEqual(result["diferencia"], 0)
        self.assertTrue(result["cuadra"])
        with urllib.request.urlopen(self.base + "/api/movistar/download-excel?token=" + result["download_token"]) as response:
            wb = openpyxl.load_workbook(io.BytesIO(response.read()))
        self.assertEqual(wb.sheetnames, ["CARGOS ADICIONALES", "CARGO MENSUAL", "Hoja1"])
        self.assertEqual(wb["CARGO MENSUAL"]["F3"].value, "=D3-E3")

    def test_expired_downloads_and_upload_limit(self):
        with server.CACHE_LOCK:
            server.EXCEL_CACHE["expired-test"] = {"bytes": b"x", "created": datetime.datetime.now() - datetime.timedelta(hours=1)}
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(self.base + "/api/movistar/download-excel?token=expired-test")
        self.assertEqual(ctx.exception.code, 404)
        body, content_type = multipart({}, {"password": "synthetic-test-only"})
        with patch.object(server, "MAX_UPLOAD_BYTES", 1):
            with self.assertRaises(urllib.error.HTTPError) as ctx:
                urllib.request.urlopen(urllib.request.Request(self.base + "/api/movistar/process", data=body, headers={"Content-Type": content_type}))
        self.assertEqual(ctx.exception.code, 413)


if __name__ == "__main__":
    unittest.main(verbosity=2)
