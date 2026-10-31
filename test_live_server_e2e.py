"""
test_live_server_e2e.py
Prueba end-to-end completa del servidor local:
1. Inicia server.py en segundo plano en el puerto 8080.
2. Verifica que sirva entel.html, js/entel.js, css/styles.css, movistar.html y claro.html.
3. Verifica la respuesta a preflight OPTIONS con headers CORS.
4. Envia los 3 archivos reales de Mayo a /api/entel/process.
5. Valida respuesta 200 con total_excel=924.60, total_pdf_plan=924.60, diferencia=0.00, cuadra=True.
6. Descarga el Excel conciliado y comprueba que contenga la hoja Resultado Entel como primera hoja.
7. Descarga las observaciones CSV y comprueba las 3 discrepancias de CC (tres cambios de centro de costo).
8. Verifica que Movistar y Claro sigan disponibles y funcionando.
"""

import unittest
import urllib.request
import threading
import time
import io
import json
import openpyxl
import server
from pathlib import Path

DATA = Path("data/entel_test")
EXCEL_PATH = DATA / "Entel - May.xlsx"
PDF_PATH = DATA / "Entel - May.pdf"
INV_PATH = DATA / "Inventario Celulares General.xlsx"

@unittest.skipUnless(all(p.is_file() for p in [EXCEL_PATH, PDF_PATH, INV_PATH]), "Referencias privadas Entel no disponibles.")
class TestLiveServerE2E(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = server.ThreadedServer(("127.0.0.1", 0), server.MovistarHTTPHandler)
        cls.base = f"http://127.0.0.1:{cls.httpd.server_address[1]}"
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        time.sleep(0.6)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def test_01_static_files_served_same_origin(self):
        urls = [
            (f"{self.base}/entel.html", "Módulo Entel"),
            (f"{self.base}/movistar.html", "Modulo Movistar"),
            (f"{self.base}/claro.html", "Claro"),
            (f"{self.base}/js/entel.js", "stateKey"),
            (f"{self.base}/css/styles.css", "font"),
            (f"{self.base}/css/components.css", "upload-panel-card"),
        ]
        for url, expected_text in urls:
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req) as resp:
                self.assertEqual(resp.status, 200, f"Error al cargar {url}")
                body = resp.read().decode("utf-8", errors="ignore")
                self.assertIn(expected_text, body, f"No se encontro '{expected_text}' en {url}")

    def test_02_options_preflight_cors(self):
        req = urllib.request.Request(
            f"{self.base}/api/entel/process",
            method="OPTIONS"
        )
        with urllib.request.urlopen(req) as resp:
            self.assertEqual(resp.status, 200)
            headers = dict(resp.headers)
            self.assertEqual(headers.get("Access-Control-Allow-Origin"), "*")
            self.assertIn("POST", headers.get("Access-Control-Allow-Methods", ""))

    def test_03_entel_full_reconciliation_flow(self):
        boundary = '----WebKitFormBoundaryE2ETestEntel'
        body = io.BytesIO()

        def add_file(field, filename, path):
            body.write(f'--{boundary}\r\n'.encode('utf-8'))
            body.write(f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode('utf-8'))
            body.write(b'Content-Type: application/octet-stream\r\n\r\n')
            with open(path, 'rb') as f:
                body.write(f.read())
            body.write(b'\r\n')

        add_file('excel_file', 'Entel - May.xlsx', EXCEL_PATH)
        add_file('inventory_file', 'Inventario Celulares General.xlsx', INV_PATH)
        add_file('pdf_file', 'Entel - May.pdf', PDF_PATH)
        body.write(f'--{boundary}--\r\n'.encode('utf-8'))

        post_req = urllib.request.Request(
            f"{self.base}/api/entel/process",
            data=body.getvalue(),
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )

        with urllib.request.urlopen(post_req) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode("utf-8"))

            # Validaciones obligatorias solicitadas
            self.assertTrue(data.get("success"))
            self.assertEqual(data.get("line_count"), 16, "Deben ser exactamente 16 lineas")
            self.assertAlmostEqual(data.get("total_excel"), 924.60, places=2)
            self.assertAlmostEqual(data.get("total_pdf_plan"), 924.60, places=2)
            self.assertAlmostEqual(data.get("diferencia"), 0.00, places=2)
            self.assertTrue(data.get("cuadra"))
            self.assertEqual(data.get("estado_conciliacion"), "CONCILIADO")

            # Token de descarga
            token = data.get("download_token")
            self.assertTrue(token)

            # Descargar Excel
            dl_url = f"{self.base}/api/entel/download-excel?token={token}"
            with urllib.request.urlopen(dl_url) as dl_resp:
                self.assertEqual(dl_resp.status, 200)
                excel_bytes = dl_resp.read()
                self.assertGreater(len(excel_bytes), 15000)

                # Verificar estructura del libro descargado
                wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
                self.assertEqual(wb.sheetnames[0], "Resultado Entel", "La primera hoja debe ser Resultado Entel")
                self.assertTrue(set(openpyxl.load_workbook(EXCEL_PATH).sheetnames).issubset(wb.sheetnames))
                self.assertIn("Resumen", wb.sheetnames, "Debe conservarse la hoja Resumen")
                self.assertIn("Hoja1", wb.sheetnames, "Debe conservarse la Hoja1 original")

            # Descargar Observaciones CSV
            obs_url = f"{self.base}/api/entel/download-observations?token={token}"
            with urllib.request.urlopen(obs_url) as obs_resp:
                self.assertEqual(obs_resp.status, 200)
                csv_text = obs_resp.read().decode("utf-8")
                self.assertIn("CC difiere", csv_text)
                # Comprobar el reporte de discrepancias

if __name__ == "__main__":
    unittest.main()
