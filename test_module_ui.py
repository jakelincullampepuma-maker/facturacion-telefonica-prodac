"""Contrato visual común y evidencias con los documentos reales disponibles."""
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import threading
import unittest
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / '.test-tools'))
from playwright.sync_api import sync_playwright
import server
from build_views import VIEWS, render

OUT = ROOT / 'data/unified-ui'
DOWNLOADS = Path.home() / 'Downloads'
CASES = {
    'claro': {'files': {'claro':Path(os.environ.get('CLARO_JULY_REFERENCE', DOWNLOADS/'dgv/Pmin - Jul.xlsx')),
                       'inv':Path(os.environ.get('CLARO_INVENTORY', DOWNLOADS/'Inventario Celulares General.xlsx')),
                       'pdf':Path(os.environ.get('CLARO_JULY_PDF', DOWNLOADS/'dgv/Pmin - Jul.PDF'))},
              'expected':{'numeros_unicos':143,'suma_lineas':5348.44,'ajustes_abonos':-7.37,'total_calculado':5341.07,'cargos_del_mes':5341.07,'diferencia':0},
              'kpis':{'lineas':'numeros_unicos','total-lineas':'suma_lineas','ajustes':'ajustes_abonos','total':'total_calculado','cargos-mes':'cargos_del_mes','diferencia':'diferencia'}},
    'movistar': {'files':{'pdf':ROOT/'data/movistar_nov_test/Movistar-Nov.pdf','inventory':ROOT/'data/movistar_nov_test/Inventario Celulares General.xlsx'},
                 'expected':{'line_count':325,'cargo_mensual_sin_igv':12597.72,'descuento_total':1046.68,'cargo_final_sin_igv':11551.04,'igv':2079.19,'total_afecto':13630.23,'cargos_inafectos':25452.11,'redondeo':.06,'pdf_total':39082.40,'total_calculado':39082.40,'diferencia':0},
                 'kpis':{'lineas':'line_count','monto':'cargo_mensual_sin_igv','desc':'descuento_total','monto-f':'cargo_final_sin_igv','igv':'igv','total-afecto':'total_afecto','inafectos':'cargos_inafectos','redondeo':'redondeo','total-pdf':'pdf_total','total-calculado':'total_calculado','diferencia':'diferencia'}},
    'entel': {'files':{'excel':ROOT/'data/entel_test/Entel - May.xlsx','inventory':ROOT/'data/entel_test/Inventario Celulares General.xlsx','pdf':ROOT/'data/entel_test/Entel - May.pdf'},
              'expected':{'line_count':16,'conciliado_count':16,'total_excel':924.60,'total_pdf_plan':924.60,'diferencia':0},
              'kpis':{'lineas':'line_count','conciliadas':'conciliado_count','observaciones':'obs_lines_count','excel-plan':'total_excel','pdf-plan':'total_pdf_plan','equipos':'total_pdf_equipos','total-factura':'total_pdf_factura','diferencia':'diferencia'}},
}


class ModuleUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        OUT.mkdir(parents=True, exist_ok=True)
        cls.httpd = server.ThreadedServer(('127.0.0.1',0),server.AppCelularesHTTPHandler)
        threading.Thread(target=cls.httpd.serve_forever,daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.httpd.server_address[1]}'
        cls.pw = sync_playwright().start()
        chrome=Path('C:/Program Files/Google/Chrome/Application/chrome.exe')
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None,headless=True)
        cls.report = {'geometry':[], 'real':{}}

    @classmethod
    def tearDownClass(cls):
        before = OUT/'business-before.json'
        if before.is_file():
            hashes = json.loads(before.read_text(encoding='utf-8-sig'))
            # El servidor cambia de rutas; los motores y lectores conservan su contrato.
            cls.report['business_unchanged'] = all(hashlib.sha256((ROOT/path).read_bytes()).hexdigest().upper()==digest for path,digest in hashes.items() if path != 'server.py')
        (OUT/'verification.json').write_text(json.dumps(cls.report,ensure_ascii=False,indent=2),encoding='utf-8')
        sections = []
        for size in ['desktop','laptop','small']:
            for phase in ['initial','results']:
                cards = ''.join(f'<article><h3>{op.title()}</h3><a href="{op}-{phase}-{size}.png"><img src="{op}-{phase}-{size}.png" alt="{op}: {phase}, {size}"></a></article>' for op in VIEWS if (OUT/f'{op}-{phase}-{size}.png').is_file())
                sections.append(f'<h2>{size} · {phase}</h2><section>{cards}</section>')
        (OUT/'comparacion.html').write_text('<!doctype html><html lang="es"><meta charset="utf-8"><title>Comparación de los tres módulos</title><style>body{font:16px system-ui;background:#f4f7fb;color:#0f172a;margin:24px}section{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:20px}article{min-width:0}img{width:100%;border:1px solid #e2e8f0;border-radius:8px}a{color:#0284c7}@media(max-width:700px){section{grid-template-columns:1fr}}</style><h1>Claro · Movistar · Entel</h1><p>Capturas completas con documentos reales. Abra cada imagen para ver el tamaño original.</p><p><a href="verification.json">Resultados y comparación de los Excel</a> · <a href="full-tests.log">Suite completa</a></p>'+''.join(sections)+'</html>',encoding='utf-8')
        cls.browser.close();cls.pw.stop();cls.httpd.shutdown();cls.httpd.server_close()

    def test_generated_pages_use_one_template(self):
        for op in VIEWS:
            self.assertEqual((ROOT/f'{op}.html').read_text(encoding='utf-8'),render(op))
        self.assertEqual((ROOT/'index.html').read_text(encoding='utf-8'),render('claro'))

    @unittest.skipUnless((OUT/'business-before.json').is_file(),'Registro local anterior no disponible.')
    def test_business_sources_unchanged(self):
        hashes=json.loads((OUT/'business-before.json').read_text(encoding='utf-8-sig'))
        for path,digest in hashes.items():
            if path == 'server.py':
                continue  # El enrutamiento HTTP se verifica en test_release_server.
            self.assertEqual(hashlib.sha256((ROOT/path).read_bytes()).hexdigest().upper(),digest,path)

    def test_initial_geometry_and_presentation_states(self):
        with self.browser.new_context() as context:
            for width,height in [(1366,1000),(1024,800),(390,844)]:
                metrics = []
                for op in VIEWS:
                    page=context.new_page();page.set_viewport_size({'width':width,'height':height});page.goto(self.base+f'/{op}')
                    page.evaluate('document.fonts.ready')
                    self.assertTrue(page.locator('#'+VIEWS[op]['result']).is_hidden())
                    self.assertEqual(page.locator('main .card-kpi-concil:visible,main table:visible').count(),0)
                    self.assertEqual(page.locator('.header-title-group p').inner_text(),'Conciliación mensual por centro de costo')
                    self.assertEqual(page.locator('.upload-panel-title h2').inner_text(),'Carga de archivos – '+VIEWS[op]['name'])
                    self.assertEqual(page.locator('.upload-buttons-bar button').all_text_contents(),['Procesar archivos','Limpiar','Descargar Excel conciliado'])
                    self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
                    m=page.evaluate('''()=>{const rect=e=>{let r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height]};return {header:rect(document.querySelector('.module-header')),upload:rect(document.querySelector('.upload-panel-card')),mainX:document.querySelector('main').getBoundingClientRect().x,cardHeights:[...document.querySelectorAll('.upload-file-card')].filter(e=>e.getClientRects().length).map(e=>e.getBoundingClientRect().height)};}''')
                    self.report['geometry'].append({'operator':op,'width':width,**m});metrics.append(m)
                    self.assertTrue(all(h==164 for h in m['cardHeights']))
                    for ok,obs,title,label,color in [(True,0,'La conciliación cuadra','CUADRA','rgb(16, 185, 129)'),(True,2,'Hay observaciones por revisar','REVISAR','rgb(245, 158, 11)'),(False,0,'La conciliación no cuadra','NO CUADRA','rgb(239, 68, 68)')]:
                        page.evaluate('(args)=>ModuleUI.status(args)',{'cuadra':ok,'observations':obs,'difference':'S/ 1.00'})
                        self.assertEqual(page.locator('#banner-title-text').text_content(),title)
                        self.assertEqual(page.locator('#banner-pill-badge').text_content(),label)
                        self.assertEqual(page.locator('#banner-status-circle').evaluate('e=>getComputedStyle(e).backgroundColor'),color)
                    self.assertEqual(page.evaluate("ModuleUI.period('01/Oct/2026 - 31/Oct/2026')"),'Octubre 2026')
                    self.assertEqual(page.evaluate("ModuleUI.period('Noviembre 2026')"),'Noviembre 2026')
                    page.close()
                for m in metrics[1:]:
                    self.assertEqual(m['header'],metrics[0]['header'])
                    self.assertEqual(m['upload'][:3],metrics[0]['upload'][:3])
                    self.assertEqual(m['mainX'],metrics[0]['mainX'])

    def screenshots(self,page,op,phase):
        for size,width,height in [('desktop',1366,1000),('laptop',1024,800),('small',390,844)]:
            page.set_viewport_size({'width':width,'height':height})
            page.evaluate("window.scrollTo({top:0,behavior:'instant'})")
            self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
            if phase=='results':
                columns=page.locator('.conciliation-kpis-grid').evaluate('e=>getComputedStyle(e).gridTemplateColumns.split(" ").length')
                self.assertEqual(columns,4 if width>1200 else 2 if width>520 else 1)
                for card in page.locator('.card-kpi-concil').all():
                    self.assertEqual(card.bounding_box()['height'],174)
                    self.assertTrue(card.evaluate('e=>e.scrollHeight<=e.clientHeight'))
                for button in page.locator('.upload-buttons-bar > button').all():
                    self.assertTrue(button.evaluate('e=>{let r=e.getBoundingClientRect(),p=e.parentElement.getBoundingClientRect();return r.left>=p.left&&r.right<=p.right+.5}'))
            page.screenshot(path=str(OUT/f'{op}-{phase}-{size}.png'),full_page=True,animations='disabled')
        page.set_viewport_size({'width':1366,'height':1000})

    def real_case(self,op):
        case=CASES[op]
        with self.browser.new_context(accept_downloads=True,viewport={'width':1366,'height':1000}) as context:
            page=context.new_page();page.set_default_timeout(120000);errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
            page.goto(self.base+f'/{op}');page.evaluate('document.fonts.ready')
            if op=='claro':page.locator('#tab-prodimin').click()
            self.screenshots(page,op,'initial')
            for field,path in case['files'].items():page.locator('#input-'+field+'-file').set_input_files(str(path))
            if op=='entel':
                page.evaluate('''()=>{const run=EntelProcessor.process;EntelProcessor.process=async(...args)=>{const r=await run(...args);window.uiResult=r;return r;};}''')
                page.locator('#btn-process-files').click();page.wait_for_function('!!window.uiResult')
                result=page.evaluate('(({excel_blob,...r})=>r)(window.uiResult)')
            else:
                with page.expect_response(lambda r:r.url.endswith(f'/api/{op}/process')) as response:page.locator('#btn-process-files').click()
                self.assertEqual(response.value.status,200);result=response.value.json()
            page.wait_for_function("!document.getElementById('btn-export-excel').disabled")
            for key,expected in case['expected'].items():self.assertAlmostEqual(result[key],expected,places=2,msg=op+'.'+key)
            for key,source in case['kpis'].items():
                number=float(re.sub(r'[^\d.\-]','',page.locator('#kpi-val-'+key).inner_text()))
                self.assertAlmostEqual(number,result[source],places=2,msg=op+'.'+key)
            self.assertEqual(page.locator('.conciliation-banner-card:visible').count(),1)
            if op!='claro':self.assertEqual(page.locator('#'+op+'-alert-container').inner_text(),'')
            self.assertEqual(page.locator('.two-columns-bottom-grid > .table-widget-card').count(),2)
            self.assertEqual(page.locator('.two-columns-bottom-grid .table-widget-title').all_text_contents(),['Resumen por CC Dpto.','Sin CC - Revisar'])
            self.screenshots(page,op,'results')
            with page.expect_download() as download:page.locator('#btn-export-excel').click()
            output=OUT/(op.title()+'.xlsx');download.value.save_as(output)
            baseline=OUT/'baseline'/(op.title()+'.xlsx')
            comparison={}
            if baseline.is_file():
                with zipfile.ZipFile(baseline) as old,zipfile.ZipFile(output) as new:
                    self.assertEqual(set(old.namelist()),set(new.namelist()))
                    def same(name):
                        left,right=old.read(name),new.read(name)
                        if left==right:return True
                        if name.endswith(('.xml','.rels')):
                            return ET.canonicalize(left,rewrite_prefixes=True)==ET.canonicalize(right,rewrite_prefixes=True)
                        return False
                    changed=[name for name in old.namelist() if name!='docProps/core.xml' and not same(name)]
                    self.assertEqual(changed,[],f'{op}: contenido del Excel modificado')
                    comparison={'same_package_entries':True,'all_xml_and_resources_identical':True,'xml_comparison':'XML canónico: admite prefijos de espacios de nombres equivalentes','excluded':'docProps/core.xml (fecha de creación)'}
            self.report['real'][op]={'values':{k:result[k] for k in case['expected']},'period_internal':result['periodo'],'period_display':page.locator('#period-display-text').inner_text(),'status':page.locator('#banner-pill-badge').inner_text(),'workbook_comparison':comparison}
            page.locator('#btn-clear-files').click();self.assertTrue(page.locator('#'+VIEWS[op]['result']).is_hidden());self.assertEqual(errors,[])

    @unittest.skipUnless(all(p.is_file() for p in CASES['claro']['files'].values()),'Referencias reales de Claro no disponibles.')
    def test_real_claro(self):self.real_case('claro')

    @unittest.skipUnless(all(p.is_file() for p in CASES['movistar']['files'].values()),'Referencias reales de Movistar no disponibles.')
    def test_real_movistar(self):self.real_case('movistar')

    @unittest.skipUnless(all(p.is_file() for p in CASES['entel']['files'].values()),'Referencias reales de Entel no disponibles.')
    def test_real_entel(self):self.real_case('entel')


if __name__=='__main__':unittest.main(verbosity=2)
