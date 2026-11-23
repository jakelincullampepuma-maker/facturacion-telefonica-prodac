"""Comprueba los tres paneles compartidos al mismo zoom (100%)."""
import json
import hashlib
import io
import os
from pathlib import Path
import sys
import threading
import unittest

ROOT = Path(__file__).resolve().parent
if (ROOT / ".test-tools").is_dir():
    sys.path.insert(0, str(ROOT / ".test-tools"))
from playwright.sync_api import sync_playwright
import server
import openpyxl
from test_claro import fixtures as claro_fixtures
from test_support import movistar_fixtures, simple_pdf, workbook_bytes


class OperatorLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = os.environ.get("APP_TEST_BASE_URL", "").rstrip("/")
        cls.httpd = None
        if not cls.base:
            cls.httpd = server.ThreadedServer(("127.0.0.1", 0), server.AppCelularesHTTPHandler)
            threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()
            cls.base = f"http://127.0.0.1:{cls.httpd.server_address[1]}"
        cls.pw = sync_playwright().start()
        chrome = Path("C:/Program Files/Google/Chrome/Application/chrome.exe")
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None, headless=True)
        cls.out = ROOT / "data/sidebar-comparison"
        cls.out.mkdir(parents=True, exist_ok=True)
        cls.reports = []

    @classmethod
    def tearDownClass(cls):
        (cls.out / "after.json").write_text(json.dumps(cls.reports, indent=2), encoding="utf-8")
        cls.browser.close()
        cls.pw.stop()
        if cls.httpd:
            cls.httpd.shutdown()
            cls.httpd.server_close()

    def page(self, context, module):
        page = context.new_page()
        page.goto(self.base + "/" + module + ".html")
        page.evaluate("document.fonts.ready")
        return page

    def nav_click(self, page, action):
        if page.viewport_size['width'] <= 760:
            page.locator('#sidebar-menu-button').click()
        page.locator('#nav-'+action).click()

    def metrics(self, page):
        return page.evaluate("""() => {
          const sidebar=document.querySelector('.sidebar'), nav=document.querySelector('.sidebar-nav');
          const measure=e=>{
            const s=getComputedStyle(e), r=e.getBoundingClientRect();
            return {x:r.x,width:r.width,top:r.top,bottom:r.bottom,height:r.height,
              fontSize:s.fontSize,fontFamily:s.fontFamily,padding:s.padding,margin:s.margin,gap:s.gap};
          };
          return {sidebar:{...measure(sidebar),scrollHeight:sidebar.scrollHeight,clientHeight:sidebar.clientHeight,scrollWidth:sidebar.scrollWidth,clientWidth:sidebar.clientWidth},
            header:measure(document.querySelector('.sidebar-header')),nav:measure(nav),
            footer:measure(document.querySelector('.sidebar-footer')),main:measure(document.querySelector('.main-wrapper')),
            button:measure(nav.querySelector('.nav-item')),
            badge:measure(nav.querySelector('a[href="entel.html"] .op-badge-soon')),
            options:[...nav.querySelectorAll('button,a')].map(measure),
            labels:[...nav.querySelectorAll('.sidebar-link')].map(e=>e.textContent.trim()),
            activeOperator:nav.querySelector('.operator-link.active').dataset.operator,
            icons:[...nav.querySelectorAll('.sidebar-link svg,.operator-icon')].map(measure),
            dpr:devicePixelRatio,zoom:getComputedStyle(document.documentElement).zoom};
        }""")

    def test_laptop_dimensions_match_and_all_options_fit(self):
        for width, height in [(1366, 768), (1280, 720), (1024, 600), (1024, 576), (640, 800)]:
            with self.subTest(viewport=(width, height)):
                context = self.browser.new_context(viewport={"width": width, "height": height}, device_scale_factor=1)
                metrics = {}
                normalized = []
                for module in ["claro", "movistar", "entel"]:
                    page = self.page(context, module)
                    if width <= 760:
                        self.assertEqual(page.locator('.main-wrapper').bounding_box()['x'],0)
                        page.locator('#sidebar-menu-button').click()
                    data = metrics[module] = self.metrics(page)
                    self.reports.append({"module": module, "viewport": [width, height], **data})
                    self.assertEqual(data["dpr"], 1)
                    self.assertEqual(float(data["zoom"]), 1)
                    self.assertEqual(data["sidebar"]["width"], 240)
                    self.assertEqual(data["sidebar"]["height"], height)
                    self.assertEqual(data["sidebar"]["x"], 0)
                    self.assertEqual(data["main"]["x"], 0 if width <= 760 else 240)
                    self.assertEqual(data['labels'], ['Dashboard','Cargar archivos','Resultados','Sin CC - Revisar','Descargar reporte'])
                    self.assertEqual(data['activeOperator'], module)
                    self.assertEqual(len(data['options']), 8)
                    self.assertEqual(page.locator('#nav-resumen, #nav-obs').count(), 0)
                    if module == 'entel':
                        self.assertEqual(page.locator('h1').inner_text(), 'Módulo Entel')
                        self.assertEqual(page.locator('h1 span, h1 svg').count(), 0)
                    self.assertLessEqual(data["sidebar"]["scrollHeight"], data["sidebar"]["clientHeight"] + 1)
                    self.assertLessEqual(data["sidebar"]["scrollWidth"], data["sidebar"]["clientWidth"])
                    for option in data["options"]:
                        self.assertGreaterEqual(option["top"], 0)
                        self.assertLessEqual(option["bottom"], height)
                    self.assertLessEqual(data["footer"]["bottom"], height + 1)
                    page.screenshot(path=str(self.out / f'{module}-{width}x{height}.png'), animations='disabled')
                    panel = page.locator('.sidebar')
                    panel.screenshot(path=str(self.out / f'panel-{module}-{width}x{height}.png'), animations='disabled')
                    page.add_style_tag(content='.sidebar *{transition:none!important;}.sidebar .operator-link,.sidebar .operator-link.active,.sidebar .operator-link .op-badge-soon,.sidebar .operator-link .operator-icon{background:transparent!important;border-color:transparent!important;color:#8fa2b7!important;}')
                    normalized.append(hashlib.sha256(panel.screenshot(path=str(self.out / f'normalized-{module}-{width}x{height}.png'), animations='disabled')).hexdigest())
                self.assertEqual(len(set(normalized)), 1, 'Los paneles difieren incluso al neutralizar únicamente sus colores activos.')
                for module in ['movistar','entel']:
                    for element in ['header','button','badge','footer','nav','main']:
                        for key in ['x','width','fontSize','fontFamily','padding','margin','gap']:
                            self.assertEqual(metrics[module][element][key], metrics['claro'][element][key], module+'.'+element+'.'+key)
                    self.assertEqual(metrics[module]['options'],metrics['claro']['options'])
                    self.assertEqual(metrics[module]['icons'],metrics['claro']['icons'])
                context.close()
        (self.out/'comparacion.html').write_text('''<!doctype html><html lang="es"><meta charset="utf-8"><title>Comparación de paneles</title><style>body{font:16px system-ui;margin:24px;background:#f4f7fb}section{display:flex;gap:24px}.panel img{width:240px}#panels.overlay{position:relative;height:768px}#panels.overlay .panel{position:absolute;top:0;left:0;opacity:.34}#panels.overlay h2{visibility:hidden}</style><h1>Panel compartido: Claro, Movistar y Entel</h1><p>Ancho: 240 px. Inicio horizontal del contenido: 240 px. Posición, iconos, texto y botones idénticos; sólo cambia el color del operador activo.</p><label><input type="checkbox" onchange="document.getElementById('panels').classList.toggle('overlay',this.checked)"> Superponer capturas</label><section id="panels">'''+''.join(f'<div class="panel"><h2>{name.title()}</h2><img src="panel-{name}-1366x768.png" alt="Panel de {name}"></div>' for name in ['claro','movistar','entel'])+'</section></html>',encoding='utf-8')

    def test_small_height_keeps_scroll_and_right_content_is_independent(self):
        context = self.browser.new_context(viewport={"width": 1280, "height": 400}, device_scale_factor=1)
        for module in ["claro", "movistar", "entel"]:
            with self.subTest(module=module):
                page = self.page(context, module)
                data = self.metrics(page)
                self.assertGreater(data["sidebar"]["scrollHeight"], data["sidebar"]["clientHeight"])
                # En una ventana baja, se puede llegar a la última opción.
                page.locator('.sidebar a[href="entel.html"]').scroll_into_view_if_needed()
                self.assertGreater(page.locator(".sidebar").evaluate("e=>e.scrollTop"), 0)
                rect = page.locator('.sidebar a[href="entel.html"]').bounding_box()
                self.assertLessEqual(rect["y"] + rect["height"], 400)
                # El contenido largo a la derecha debe seguir desplazándose.
                page.evaluate("""() => {
                  document.querySelector('.sidebar').scrollTop=0;
                  const longContent=document.createElement('div');
                  longContent.style.height='2000px';
                  document.querySelector('main').append(longContent);
                  window.scrollTo({top:300,behavior:'instant'});
                }""")
                self.assertGreater(page.evaluate("window.scrollY"), 0)
                self.assertEqual(page.locator(".sidebar").evaluate("e=>e.scrollTop"), 0)
                self.assertEqual(page.locator(".sidebar").bounding_box()["y"], 0)
                page.close()
        context.close()

    def navigation_files(self, module):
        if module == 'claro':
            files = claro_fixtures(company='PRODAC', missing=True, adjustment=False)
            return [('claro', files['fact_file']), ('inv', files['inventory_file']), ('pdf', files['pdf_file'])]
        if module == 'movistar':
            files = movistar_fixtures()
            inventory = openpyxl.load_workbook(io.BytesIO(files['inventory_file']))
            inventory.active['B3'] = None
            return [('pdf',files['pdf_file']),('inventory',workbook_bytes(inventory))]
        billing = openpyxl.Workbook(); billing.active.title = 'Detalle'
        billing.active.append(['Número','Plan','Cargo Fijo'])
        billing.active.append(['951234567','Plan A',12.5]); billing.active.append(['952345678','Plan B',7.5])
        inventory = openpyxl.Workbook(); inventory.active.title = 'PRODAC'
        inventory.active.append(['Numero','Operador','CC Dpto','Usuario'])
        inventory.active.append(['951234567','Entel','CC-A','Usuario A'])
        inventory.active.append(['952345678','Entel',None,'Usuario B'])
        pdf = simple_pdf('Entel\nPLAN Prorrateo Cargos Descuentos Equipos Saldo\nS/ 20.00 S/ 0.00 S/ 0.00 S/ 0.00 S/ 0.00 S/ 0.00\nTotal a pagar S/ 20.00')
        return [('excel',workbook_bytes(billing)),('inventory',workbook_bytes(inventory)),('pdf',pdf)]

    def test_all_five_actions_before_and_after_processing(self):
        for width in [1366, 640]:
          for module,results,missing in [('claro','claroResultsSection','claroMissingCcSection'),
                                      ('movistar','section-resultados-movistar','section-sin-cc-movistar'),
                                      ('entel','section-resultados-entel','section-sin-cc-entel')]:
            with self.subTest(module=module,width=width),self.browser.new_context(accept_downloads=True,viewport={'width':width,'height':768}) as context:
                page=self.page(context,module);errors=[];downloads=[]
                page.on('pageerror',lambda error:errors.append(str(error)))
                page.on('download',lambda download:downloads.append(download))
                for action in ['dashboard','upload','results','sin-cc','export']:
                    self.nav_click(page,action)
                self.assertTrue(page.locator('#'+results).is_hidden())
                self.assertEqual(downloads,[])
                for field,content in self.navigation_files(module):
                    page.locator('#input-'+field+'-file').set_input_files({'name':field+('.pdf' if field=='pdf' else '.xlsx'),'mimeType':'application/octet-stream','buffer':content})
                page.locator('#btn-process-files').click()
                page.wait_for_function("!document.getElementById('btn-export-excel').disabled")
                self.assertTrue(page.locator('#'+missing).is_visible())
                page.evaluate('''()=>{window.sidebarScrolls=[];const original=Element.prototype.scrollIntoView;Element.prototype.scrollIntoView=function(...args){sidebarScrolls.push(this.id);return original.apply(this,args);};window.sidebarTopCalls=0;const scroll=window.scrollTo;window.scrollTo=function(...args){sidebarTopCalls++;return scroll.apply(this,args);};}''')
                self.nav_click(page,'dashboard');self.assertEqual(page.evaluate('sidebarTopCalls'),1)
                for action,target in [('upload','section-carga-archivos'),('results',results),('sin-cc',missing)]:
                    self.nav_click(page,action)
                    self.assertEqual(page.evaluate('sidebarScrolls[sidebarScrolls.length-1]'),target)
                with page.expect_download() as download:self.nav_click(page,'export')
                output=self.out/(module+f'-navigation-{width}.xlsx');download.value.save_as(output)
                book=openpyxl.load_workbook(output)
                if module=='claro':self.assertEqual(book.sheetnames,['Factura Detallada','Hoja2','Hoja1'])
                self.assertEqual(len(downloads),1)
                page.locator('#btn-clear-files').click()
                for action in ['results','sin-cc','export']:
                    self.nav_click(page,action)
                self.assertTrue(page.locator('#'+results).is_hidden())
                self.assertEqual(len(downloads),1)
                self.assertEqual(errors,[])


if __name__ == "__main__":
    unittest.main(verbosity=2)
