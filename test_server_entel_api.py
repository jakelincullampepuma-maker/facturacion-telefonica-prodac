import urllib.request
import threading
import time
import io
import json
import unittest
import server
from pathlib import Path
DATA = Path("data/entel_test")

@unittest.skipUnless(all((DATA / n).is_file() for n in ["Entel - May.xlsx", "Entel - May.pdf", "Inventario Celulares General.xlsx"]), "Referencias privadas Entel no disponibles.")

class TestServerEntelAPI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.port = 0
        cls.server_address = ('127.0.0.1', cls.port)
        cls.httpd = server.ThreadedServer(cls.server_address, server.MovistarHTTPHandler)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def test_process_and_downloads(self):
        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        body = io.BytesIO()

        def add_file(field, filename, path):
            body.write(f'--{boundary}\r\n'.encode('utf-8'))
            body.write(f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode('utf-8'))
            body.write(b'Content-Type: application/octet-stream\r\n\r\n')
            with open(path, 'rb') as f:
                body.write(f.read())
            body.write(b'\r\n')

        add_file('excel_file', 'Entel - May.xlsx', DATA / 'Entel - May.xlsx')
        add_file('inventory_file', 'Inventario Celulares General.xlsx', DATA / 'Inventario Celulares General.xlsx')
        add_file('pdf_file', 'Entel - May.pdf', DATA / 'Entel - May.pdf')
        body.write(f'--{boundary}--\r\n'.encode('utf-8'))

        req = urllib.request.Request(
            f'http://127.0.0.1:{self.port}/api/entel/process',
            data=body.getvalue(),
            headers={'Content-Type': f'multipart/form-data; boundary={boundary}'}
        )

        with urllib.request.urlopen(req) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode('utf-8'))
            self.assertTrue(data.get('success'))
            self.assertEqual(data.get('line_count'), 16)
            self.assertAlmostEqual(data.get('total_excel'), 924.60, places=2)
            self.assertAlmostEqual(data.get('total_pdf_plan'), 924.60, places=2)
            self.assertAlmostEqual(data.get('diferencia'), 0.00, places=2)
            token = data.get('download_token')
            self.assertTrue(token)

        # Test download excel
        dl_url = f'http://127.0.0.1:{self.port}/api/entel/download-excel?token={token}'
        with urllib.request.urlopen(dl_url) as dl_resp:
            self.assertEqual(dl_resp.status, 200)
            excel_bytes = dl_resp.read()
            self.assertGreater(len(excel_bytes), 10000)

        # Test download observations CSV
        obs_url = f'http://127.0.0.1:{self.port}/api/entel/download-observations?token={token}'
        with urllib.request.urlopen(obs_url) as obs_resp:
            self.assertEqual(obs_resp.status, 200)
            csv_str = obs_resp.read().decode('utf-8')
            self.assertIn("NUMERO", csv_str)
            self.assertIn("CC difiere", csv_str)

if __name__ == '__main__':
    unittest.main()
