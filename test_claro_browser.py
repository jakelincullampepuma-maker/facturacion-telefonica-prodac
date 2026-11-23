"""Pantalla original de Claro: tres tarjetas, selectores y descarga preservada."""
import json
import os
from pathlib import Path
import sys
import threading
import unittest
import openpyxl
ROOT = Path(__file__).resolve().parent
if (ROOT/'.test-tools').is_dir(): sys.path.insert(0,str(ROOT/'.test-tools'))
from playwright.sync_api import sync_playwright
import server
from test_claro import fixtures, source_sheet

class ClaroBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd=server.ThreadedServer(('127.0.0.1',0),server.AppCelularesHTTPHandler)
        threading.Thread(target=cls.httpd.serve_forever,daemon=True).start()
        cls.base=f'http://127.0.0.1:{cls.httpd.server_address[1]}'
        cls.pw=sync_playwright().start()
        chrome=Path('C:/Program Files/Google/Chrome/Application/chrome.exe')
        cls.browser=cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None,headless=True)
        cls.out=ROOT/'data/claro-results'; cls.out.mkdir(parents=True,exist_ok=True)
    @classmethod
    def tearDownClass(cls):
        cls.browser.close(); cls.pw.stop(); cls.httpd.shutdown(); cls.httpd.server_close()
    def load(self,page,files):
        for key,field in [('pdf_file','pdf'),('inventory_file','inv'),('fact_file','claro')]:
            page.locator('#input-'+field+'-file').set_input_files({'name':key+('.pdf' if field=='pdf' else '.xlsx'),'mimeType':'application/octet-stream','buffer':files[key]})
    def assert_results_hidden(self,page):
        self.assertTrue(page.locator('#claroResultsSection').is_hidden())
        self.assertTrue(page.locator('#conciliation-banner').is_hidden())
        self.assertEqual(page.locator('main .card-kpi-concil:visible, main table:visible').count(),0)
        self.assertTrue(page.locator('#claroMissingCcSection').is_hidden())
        self.assertTrue(page.locator('#btn-export-excel').is_disabled())
        self.assertEqual(page.locator('#period-display-text').inner_text(),'Periodo: pendiente de detectar')
    def assert_files_cleared(self,page):
        for field in ['claro','inv','pdf']:
            self.assertEqual(page.locator('#input-'+field+'-file').evaluate('(el)=>el.files.length'),0)
            self.assertEqual(page.locator('#name-'+field+'-file').inner_text(),'Seleccionar archivo')
        self.assertTrue(page.locator('#btn-process-files').is_disabled())
    def test_initial_results_hidden_even_before_javascript(self):
        with self.browser.new_context(java_script_enabled=False) as context:
            # Esta comprobación de HTML/CSS local no depende de las fuentes remotas.
            context.route('https://fonts.googleapis.com/**', lambda route: route.abort())
            context.route('https://fonts.gstatic.com/**', lambda route: route.abort())
            for path in ['/claro.html','/']:
                page=context.new_page();page.goto(self.base+path)
                self.assert_results_hidden(page)
                self.assertEqual(page.locator('h1').inner_text(),'Módulo Claro')
                self.assertTrue(page.locator('#section-carga-archivos').is_visible())
                self.assertTrue(page.locator('#claroResultsSection').evaluate('(el)=>el.hidden'))
                page.close()
    def test_three_cards_tabs_and_preserved_download(self):
        with self.browser.new_context(accept_downloads=True,viewport={'width':1440,'height':1000}) as context:
            page=context.new_page(); errors=[]; page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(self.base+'/claro.html')
            self.assertTrue(page.locator('#tab-prodac').is_visible())
            self.assertTrue(page.locator('#tab-prodimin').is_visible())
            self.assertEqual(page.locator('.upload-file-card').count(),3)
            positions=page.locator('.upload-file-card').evaluate_all('(els)=>els.map(e=>Math.round(e.getBoundingClientRect().top))')
            self.assertEqual(len(set(positions)),1)
            self.assertEqual(page.locator('.card-step-num').all_text_contents(),['01','02','03'])
            self.assertNotIn('Validación opcional',page.locator('main').inner_text())
            self.assert_results_hidden(page)
            page.locator('#tab-prodimin').click()
            files=fixtures(duplicate=True)
            page.locator('#input-pdf-file').set_input_files({'name':'recibo.pdf','mimeType':'application/pdf','buffer':files['pdf_file']})
            page.locator('#input-inv-file').set_input_files({'name':'inventario.xlsx','mimeType':'application/octet-stream','buffer':files['inventory_file']})
            self.assertTrue(page.locator('#btn-process-files').is_disabled())
            self.load(page,files)
            self.assert_results_hidden(page)
            page.locator('#btn-process-files').click()
            page.wait_for_function("document.getElementById('banner-pill-badge').textContent === 'CUADRA'")
            self.assertTrue(page.locator('#claroResultsSection').is_visible())
            self.assertTrue(page.locator('#claroMissingCcSection').is_visible())
            self.assertEqual(page.locator('#kpi-val-lineas').inner_text(),'2')
            self.assertIn('3 registros',page.locator('#kpi-sub-lineas').inner_text())
            self.assertEqual(page.locator('#tbody-full-detail-list tr').count(),3)
            self.assertEqual(page.locator('#tbody-resumen-cc tr').count(),2)
            with page.expect_download() as download: page.locator('#btn-export-excel').click()
            target=self.out/'browser-restored.xlsx'; download.value.save_as(target)
            self.assertEqual(source_sheet(files['fact_file']),source_sheet(target.read_bytes()))
            self.assertEqual(openpyxl.load_workbook(target).sheetnames,['Factura Detallada','Hoja2','Hoja1'])
            page.screenshot(path=str(self.out/'claro-restored.png'),full_page=True)
            page.locator('#btn-ver-todos-detalle').click()
            self.assertTrue(page.locator('#full-detail-section').is_visible())
            page.locator('#tab-prodac').click()
            self.assert_results_hidden(page)
            page.locator('#tab-prodimin').click()
            self.assert_results_hidden(page)
            page.locator('#btn-process-files').click()
            page.locator('#claroResultsSection').wait_for(state='visible')
            page.locator('#btn-remove-claro').click()
            self.assert_results_hidden(page)
            self.assertTrue(page.locator('#btn-process-files').is_disabled())
            self.load(page,files)
            page.locator('#btn-process-files').click()
            page.locator('#claroResultsSection').wait_for(state='visible')
            page.locator('#btn-clear-files').click()
            self.assert_results_hidden(page)
            self.assert_files_cleared(page)
            self.assertTrue(page.locator('#process-message').is_hidden())
            self.assertEqual(errors,[])
    def test_company_mismatch_and_home_entry(self):
        with self.browser.new_context() as context:
            page=context.new_page(); page.goto(self.base+'/')
            self.load(page,fixtures())
            page.locator('#btn-process-files').click()
            page.locator('#process-message').wait_for(state='visible')
            self.assertIn('seleccionada',page.locator('#process-message').inner_text())
            self.assert_results_hidden(page)
    def test_processing_error_and_stale_responses_stay_hidden(self):
        with self.browser.new_context() as context:
            page=context.new_page();page.goto(self.base+'/claro.html');page.locator('#tab-prodimin').click()
            files=fixtures();self.load(page,files)
            with page.expect_response('**/api/claro/process') as response:
                page.locator('#btn-process-files').click()
            report=response.value.json()
            page.locator('#claroResultsSection').wait_for(state='visible')
            pending=[]
            page.route('**/api/claro/process',lambda route:pending.append(route))
            page.locator('#btn-process-files').click()
            self.assertEqual(page.locator('#process-button-label').inner_text(),'Procesando…')
            self.assertEqual(page.locator('#btn-process-files').get_attribute('aria-busy'),'true')
            self.assert_results_hidden(page)
            for field in ['claro','inv','pdf']:
                self.assertTrue(page.locator('#input-'+field+'-file').is_disabled())
            self.assertTrue(page.locator('#btn-clear-files').is_disabled())
            pending.pop().fulfill(status=500,content_type='application/json',body=json.dumps({'success':False,'error':'Error de prueba al leer el PDF.'}))
            page.locator('#process-message').wait_for(state='visible')
            self.assertEqual(page.locator('#process-message').inner_text(),'Error de prueba al leer el PDF.')
            self.assert_results_hidden(page)
            page.wait_for_function("!document.getElementById('btn-process-files').disabled")
            for field in ['claro','inv','pdf']:
                self.assertTrue(page.locator('#input-'+field+'-file').is_enabled())
            page.locator('#btn-process-files').click()
            self.assert_results_hidden(page)
            pending.pop().fulfill(status=200,content_type='application/json',body=json.dumps(report))
            page.locator('#claroResultsSection').wait_for(state='visible')
            page.wait_for_function("!document.getElementById('btn-clear-files').disabled")
            page.locator('#btn-clear-files').click()
            self.assert_results_hidden(page)
            self.assertTrue(page.locator('#process-message').is_hidden())
            self.assert_files_cleared(page)
    def test_observations_and_financial_difference_states(self):
        cases=[('missing-cc',fixtures(missing=True),'Hay observaciones por revisar','REVISAR',1),
               ('difference',fixtures(total=19.01),'La conciliación no cuadra','NO CUADRA',0)]
        for name,files,title,status,missing in cases:
            with self.subTest(case=name),self.browser.new_context() as context:
                page=context.new_page();page.goto(self.base+'/claro.html');page.locator('#tab-prodimin').click()
                self.load(page,files);page.locator('#btn-process-files').click()
                page.locator('#claroResultsSection').wait_for(state='visible')
                self.assertEqual(page.locator('#banner-title-text').inner_text(),title)
                self.assertEqual(page.locator('#banner-pill-badge').inner_text(),status)
                self.assertEqual(page.locator('main .card-kpi-concil:visible').count(),6)
                self.assertEqual(page.locator('#tbody-sin-cc tr').count(),missing)
                self.assertTrue(page.locator('#claroMissingCcSection').is_visible())
                self.assertFalse(page.locator('#btn-export-excel').is_disabled())
    @unittest.skipUnless(all(os.environ.get(k) for k in ['CLARO_JULY_PDF','CLARO_INVENTORY','CLARO_JULY_REFERENCE']),'Documentos reales de julio no configurados.')
    def test_real_july_three_files(self):
        with self.browser.new_context(accept_downloads=True,viewport={'width':1440,'height':1000}) as context:
            page=context.new_page();page.goto(self.base+'/claro.html')
            self.assert_results_hidden(page)
            page.screenshot(path=str(self.out/'claro-initial.png'),full_page=True)
            page.locator('#tab-prodimin').click()
            for field,key in [('pdf','CLARO_JULY_PDF'),('inv','CLARO_INVENTORY'),('claro','CLARO_JULY_REFERENCE')]:
                page.locator('#input-'+field+'-file').set_input_files(os.environ[key])
            with page.expect_response('**/api/claro/process',timeout=120000) as response: page.locator('#btn-process-files').click()
            result=response.value.json()
            self.assertTrue(result['success'])
            self.assertEqual(result['total_calculado'],5341.07)
            self.assertEqual(result['lineas_sin_cc'],0)
            page.wait_for_function("document.getElementById('banner-pill-badge').textContent === 'CUADRA'")
            self.assertTrue(page.locator('#claroResultsSection').is_visible())
            self.assertEqual(page.locator('#banner-title-text').inner_text(),'La conciliación cuadra')
            self.assertEqual(page.locator('#kpi-val-lineas').inner_text(),'143')
            with page.expect_download() as download: page.locator('#btn-export-excel').click()
            target=self.out/('restored-'+download.value.suggested_filename);download.value.save_as(target)
            self.assertEqual(source_sheet(Path(os.environ['CLARO_JULY_REFERENCE']).read_bytes()),source_sheet(target.read_bytes()))
            page.screenshot(path=str(self.out/'claro-restored-july.png'),full_page=True)
            page.locator('#btn-clear-files').click()
            self.assert_results_hidden(page)
            self.assert_files_cleared(page)
            page.screenshot(path=str(self.out/'claro-cleared.png'),full_page=True)

if __name__=='__main__':unittest.main(verbosity=2)
