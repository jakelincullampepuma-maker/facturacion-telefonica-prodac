"""
test_server_movistar_api.py
Prueba de integración del API HTTP de Movistar en servidor en vivo.
"""

import urllib.request
import threading
import time
import io
import json
import unittest
import os
import openpyxl
import server
from pathlib import Path

DATA_NOV = Path("data/movistar_nov_test")
PDF_PATH = DATA_NOV / "Movistar-Nov.pdf"
INV_PATH = DATA_NOV / "Inventario Celulares General.xlsx"
FACT_PATH = DATA_NOV / "Fact_Nov.xlsx"
TRAF_PATH = DATA_NOV / "Traf_Nov.xlsx"
CARGOS_PATH = DATA_NOV / "Cargos_Nov.xlsx"


@unittest.skipUnless(PDF_PATH.is_file() and INV_PATH.is_file(), 'Referencias privadas de noviembre no disponibles.')
class TestServerMovistarAPI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server_address = ('127.0.0.1', 0)
        cls.httpd = server.ThreadedServer(cls.server_address, server.MovistarHTTPHandler)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def test_process_mandatory_files_only(self):
        """
        Prueba que procesa únicamente los 2 archivos obligatorios (PDF e Inventario).
        """
        boundary = '----WebKitFormBoundaryMovistarTest123'
        body = io.BytesIO()

        def add_file(field, filename, path):
            body.write(f'--{boundary}\r\n'.encode('utf-8'))
            body.write(f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode('utf-8'))
            body.write(b'Content-Type: application/octet-stream\r\n\r\n')
            with open(path, 'rb') as f:
                body.write(f.read())
            body.write(b'\r\n')

        add_file('pdf_file', 'Movistar-Nov.pdf', PDF_PATH)
        add_file('inventory_file', 'Inventario Celulares General.xlsx', INV_PATH)
        body.write(f'--{boundary}--\r\n'.encode('utf-8'))

        req = urllib.request.Request(
            f'http://127.0.0.1:{self.port}/api/movistar/process',
            data=body.getvalue(),
            headers={'Content-Type': f'multipart/form-data; boundary={boundary}'}
        )

        with urllib.request.urlopen(req) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get('success'))
            self.assertEqual(data.get('line_count'), 325)
            self.assertAlmostEqual(data.get('cargo_mensual_sin_igv'), 12597.72, places=2)
            self.assertAlmostEqual(data.get('descuento_total'), 1046.68, places=2)
            self.assertAlmostEqual(data.get('cargo_final_sin_igv'), 11551.04, places=2)
            self.assertAlmostEqual(data.get('subtotal_pdf'), 11551.04, places=2)
            self.assertAlmostEqual(data.get('igv'), 2079.19, places=2)
            self.assertAlmostEqual(data.get('total_afecto'), 13630.23, places=2)
            self.assertAlmostEqual(data.get('cargos_inafectos'), 25452.11, places=2)
            self.assertAlmostEqual(data.get('redondeo'), 0.06, places=2)
            self.assertAlmostEqual(data.get('total_calculado'), 39082.40, places=2)
            self.assertAlmostEqual(data.get('pdf_total'), 39082.40, places=2)
            self.assertAlmostEqual(data.get('diferencia'), 0.00, places=2)
            self.assertEqual(data.get('estado'), 'CUADRA')
            self.assertEqual(data.get('download_filename'), 'Movistar_Conciliado_Noviembre_2026.xlsx')

            token = data.get('download_token')
            self.assertTrue(token)

        # 2. Probar descarga de Excel
        dl_url = f'http://127.0.0.1:{self.port}/api/movistar/download-excel?token={token}'
        with urllib.request.urlopen(dl_url) as dl_resp:
            self.assertEqual(dl_resp.status, 200)
            excel_bytes = dl_resp.read()
            self.assertGreater(len(excel_bytes), 20000)

            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
            self.assertEqual(wb.sheetnames, ['CARGOS ADICIONALES', 'CARGO MENSUAL', 'Hoja1'])
            # Verificar formula final CARGOS ADICIONALES
            self.assertEqual(wb['CARGOS ADICIONALES'].cell(94, 3).value, '=SUM(C2:C93)')
            # Verificar formula final CARGO MENSUAL
            self.assertEqual(wb['CARGO MENSUAL'].cell(329, 4).value, '=F327-D328')

        # 3. Probar descarga de observaciones CSV
        obs_url = f'http://127.0.0.1:{self.port}/api/movistar/download-observations?token={token}'
        with urllib.request.urlopen(obs_url) as obs_resp:
            self.assertEqual(obs_resp.status, 200)
            csv_str = obs_resp.read().decode('utf-8')
            self.assertIn("NUMERO", csv_str)
            self.assertTrue(data["sin_cc_lines"])
            for line in data["sin_cc_lines"]:
                self.assertIn(line["phone"], csv_str)

    def test_process_with_optional_files(self):
        """
        Prueba con archivos opcionales Fact, Traf, Cargos y contraseña.
        """
        if not os.environ.get('MOVISTAR_TEST_PASSWORD'):
            self.skipTest('Contraseña de prueba no configurada.')
        boundary = '----WebKitFormBoundaryMovistarTestOptional'
        body = io.BytesIO()

        def add_file(field, filename, path):
            body.write(f'--{boundary}\r\n'.encode('utf-8'))
            body.write(f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode('utf-8'))
            body.write(b'Content-Type: application/octet-stream\r\n\r\n')
            with open(path, 'rb') as f:
                body.write(f.read())
            body.write(b'\r\n')

        def add_field(name, val):
            body.write(f'--{boundary}\r\n'.encode('utf-8'))
            body.write(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode('utf-8'))
            body.write(val.encode('utf-8'))
            body.write(b'\r\n')

        add_file('pdf_file', 'Movistar-Nov.pdf', PDF_PATH)
        add_file('inventory_file', 'Inventario Celulares General.xlsx', INV_PATH)
        add_file('fact_file', 'Fact_Nov.xlsx', FACT_PATH)
        add_file('traf_file', 'Traf_Nov.xlsx', TRAF_PATH)
        add_file('cargos_file', 'Cargos_Nov.xlsx', CARGOS_PATH)
        add_field('password', os.environ['MOVISTAR_TEST_PASSWORD'])
        body.write(f'--{boundary}--\r\n'.encode('utf-8'))

        req = urllib.request.Request(
            f'http://127.0.0.1:{self.port}/api/movistar/process',
            data=body.getvalue(),
            headers={'Content-Type': f'multipart/form-data; boundary={boundary}'}
        )

        with urllib.request.urlopen(req) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get('success'))
            self.assertEqual(data.get('estado'), 'CUADRA')
            warns = data.get('optional_warnings', [])
            self.assertTrue(len(warns) > 0)


if __name__ == '__main__':
    unittest.main()
