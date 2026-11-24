"""Selección sin lecturas, selectores consecutivos y diagnóstico por etapa."""
import json
import os
from pathlib import Path
import sys
import threading
import time
import unittest

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / '.test-tools'))
from playwright.sync_api import sync_playwright
import server
from test_claro import fixtures as claro_files
from test_support import movistar_fixtures

PROBE = r"""(() => {
  const probe = window.uploadProbe = {changes: [], listeners: {}, reads: [], fileReads: {}, posts: []};
  const add = EventTarget.prototype.addEventListener;
  EventTarget.prototype.addEventListener = function(type, listener, options) {
    if (type === 'change' && this instanceof HTMLInputElement && this.type === 'file') {
      const id = this.id;
      probe.listeners[id] = (probe.listeners[id] || 0) + 1;
      const original = listener;
      listener = function(event) {
        const start = performance.now();
        try { return original.call(this, event); }
        finally { probe.changes.push({id, ms: performance.now() - start}); }
      };
    }
    return add.call(this, type, listener, options);
  };
  for (const method of ['arrayBuffer', 'text', 'stream']) {
    const original = Blob.prototype[method];
    if (original) Blob.prototype[method] = function(...args) {
      if (this instanceof File && method === 'arrayBuffer') probe.fileReads[this.name] = (probe.fileReads[this.name] || 0) + 1;
      probe.reads.push(method); return original.apply(this, args);
    };
  }
  for (const method of ['readAsArrayBuffer', 'readAsText', 'readAsDataURL', 'readAsBinaryString']) {
    const original = FileReader.prototype[method];
    if (original) FileReader.prototype[method] = function(...args) {
      probe.reads.push(method); return original.apply(this, args);
    };
  }
  const fetchOriginal = window.fetch;
  window.fetch = function(url, options) {
    if (options?.method?.toUpperCase() === 'POST') probe.posts.push(String(url));
    return fetchOriginal.apply(this, arguments);
  };
})();"""


class UploadPerformanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = server.ThreadedServer(('127.0.0.1', 0), server.AppCelularesHTTPHandler)
        threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.httpd.server_address[1]}'
        cls.pw = sync_playwright().start()
        chrome = Path('C:/Program Files/Google/Chrome/Application/chrome.exe')
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.is_file() else None, headless=True)
        cls.out = ROOT / 'data/upload-performance'
        cls.out.mkdir(parents=True, exist_ok=True)
        cls.measurements = {}

    @classmethod
    def tearDownClass(cls):
        label = 'before' if os.environ.get('UPLOAD_BASELINE') else 'after'
        (cls.out / (label + '.json')).write_text(json.dumps(cls.measurements, indent=2), encoding='utf-8')
        cls.browser.close(); cls.pw.stop(); cls.httpd.shutdown(); cls.httpd.server_close()

    def inputs(self, module):
        if module == 'claro':
            files = claro_files(company='PRODAC')
            return [(field, files[key]) for field, key in [('claro','fact_file'), ('inv','inventory_file'), ('pdf','pdf_file')]]
        if module == 'movistar':
            files = movistar_fixtures()
            return [(field, files[key]) for field, key in [('pdf','pdf_file'), ('inventory','inventory_file'), ('fact','fact_file'), ('traf','traf_file'), ('cargos','cargos_file')]]
        folder = ROOT / 'data/entel_test'
        if not all((folder/name).is_file() for name in ['Entel - May.xlsx', 'Inventario Celulares General.xlsx', 'Entel - May.pdf']):
            from test_release_public import synthetic_entel
            return list(synthetic_entel().items())
        return [('excel', (folder/'Entel - May.xlsx').read_bytes()),
                ('inventory', (folder/'Inventario Celulares General.xlsx').read_bytes()),
                ('pdf', (folder/'Entel - May.pdf').read_bytes())]

    def test_selection_has_no_content_processing_or_duplicate_listeners(self):
        for module in ['claro', 'movistar', 'entel']:
            with self.subTest(module=module), self.browser.new_context(viewport={'width':1440,'height':1000}) as context:
                context.add_init_script(PROBE)
                page = context.new_page(); page.goto(self.base+'/'+module)
                if module == 'entel': page.wait_for_function('window.EntelProcessor && window.UploadUI')
                page.evaluate("""() => {
                  for (const [object, name] of [[window.XLSX,'read'],[window.pdfjsLib,'getDocument']]) {
                    if (!object) continue;
                    const original=object[name];
                    object[name]=function(...args){uploadProbe.reads.push(name);return original.apply(this,args);};
                  }
                }""")
                fields = self.inputs(module)
                chooser_ms = []
                for cycle in range(2):
                    for field, content in reversed(fields):
                        if module == 'movistar' and field in ['fact','traf','cargos']:
                            page.locator('#details-optional-val').evaluate('(el)=>el.open=true')
                        start = time.perf_counter()
                        with page.expect_file_chooser(timeout=2000) as chooser:
                            page.locator('#dropzone-'+field).click()
                        chooser_ms.append(round((time.perf_counter()-start)*1000, 2))
                        chooser.value.set_files({'name':field+('.pdf' if field=='pdf' else '.xlsx'), 'mimeType':'application/octet-stream', 'buffer':content})
                        self.assertTrue(all(page.locator('#input-'+name+'-file').is_enabled() for name, _ in fields))
                    if module == 'movistar': page.locator('#input-movistar-password').fill('synthetic-test')
                    self.assertTrue(page.locator('#btn-process-files').is_enabled())
                    probe = page.evaluate('uploadProbe')
                    self.assertEqual(probe['reads'], [])
                    self.assertEqual(probe['posts'], [])
                    self.assertTrue(all(count == 1 for count in probe['listeners'].values()))
                    self.assertEqual(len(probe['changes']), len(fields)*(cycle+1))
                    page.locator('#btn-clear-files').click()
                probe = page.evaluate('uploadProbe')
                self.measurements.setdefault(module, {}).update({'selection_events':probe['changes'], 'listeners':probe['listeners'],
                                             'content_reads_before_button':probe['reads'], 'posts_before_button':probe['posts'],
                                             'next_chooser_ms':chooser_ms})
                self.assertLess(max(event['ms'] for event in probe['changes']), 50)
                if not os.environ.get('UPLOAD_BASELINE'):
                    self.assertTrue(page.locator('#btn-process-files').is_disabled())

    def load_required(self, page, module):
        for field, content in self.inputs(module):
            if module == 'movistar' and field in ['fact', 'traf', 'cargos']: continue
            page.locator('#input-'+field+'-file').set_input_files({'name':field+('.pdf' if field=='pdf' else '.xlsx'),
                'mimeType':'application/octet-stream', 'buffer':content})

    def test_busy_error_releases_controls_without_duplicate_processing(self):
        for module in ['claro','movistar','entel']:
            with self.subTest(module=module), self.browser.new_context() as context:
                context.add_init_script(PROBE)
                page=context.new_page();page.goto(self.base+'/'+module)
                self.load_required(page,module)
                pending=[]
                if module=='entel':
                    page.evaluate("()=>{EntelProcessor.process=()=>{window.processCalls=(window.processCalls||0)+1;return new Promise((resolve,reject)=>window.rejectProcess=reject);};}")
                else: page.route('**/api/'+module+'/process',lambda route:pending.append(route))
                page.locator('#btn-process-files').click()
                page.wait_for_function("typeof window.rejectProcess === 'function' || uploadProbe.posts.length === 1")
                self.assertIn('Procesando',page.locator('#btn-process-files').inner_text())
                for field,_ in self.inputs(module): self.assertTrue(page.locator('#input-'+field+'-file').is_disabled())
                page.locator('#btn-process-files').evaluate('(el)=>{el.click();el.click();}')
                if module=='entel':
                    self.assertEqual(page.evaluate('processCalls'),1)
                    page.evaluate("rejectProcess(new Error('Error de prueba'))")
                else:
                    self.assertEqual(page.evaluate('uploadProbe.posts.length'),1)
                    pending.pop().fulfill(status=500,content_type='application/json',body=json.dumps({'success':False,'error':'Error de prueba'}))
                page.wait_for_function("UploadDiagnostics.runs[0].status === 'error'")
                for field,_ in self.inputs(module): self.assertTrue(page.locator('#input-'+field+'-file').is_enabled())
                self.assertTrue(page.locator('#btn-process-files').is_enabled())
                page.locator('#btn-clear-files').click()
                self.assertTrue(page.locator('#btn-process-files').is_disabled())
                self.assertEqual(len(page.evaluate('UploadDiagnostics.runs')),1)

    def test_invalid_extension_and_size_do_not_read_files_or_scroll(self):
        for module in ['claro','movistar','entel']:
            with self.subTest(module=module), self.browser.new_context() as context:
                context.add_init_script(PROBE)
                page=context.new_page();page.goto(self.base+'/'+module)
                self.load_required(page,module)
                for name,size in [('invalid.txt',100),('oversize.pdf',100*1024*1024+1)]:
                    before=page.evaluate('scrollY')
                    page.evaluate("""({name,size})=>{
                      const input=document.getElementById('input-pdf-file');
                      const file=new File(['x'],name);Object.defineProperty(file,'size',{value:size});
                      Object.defineProperty(input,'files',{configurable:true,value:[file]});
                      input.dispatchEvent(new Event('change',{bubbles:true}));
                    }""",{'name':name,'size':size})
                    self.assertTrue(page.locator('#btn-process-files').is_disabled())
                    self.assertEqual(page.evaluate('scrollY'),before)
                    self.assertEqual(page.evaluate('uploadProbe.reads'),[])
                    self.assertEqual(page.evaluate('uploadProbe.posts'),[])

    def test_processing_diagnostics_and_entel_worker(self):
        for module in ['claro','movistar','entel']:
            with self.subTest(module=module),self.browser.new_context() as context:
                context.add_init_script(PROBE)
                page=context.new_page();page.set_default_timeout(120000);page.goto(self.base+'/'+module)
                self.load_required(page,module)
                if module=='entel':
                    page.evaluate("""()=>{
                      const original=EntelProcessor.process;
                      EntelProcessor.process=async(...args)=>{const result=await original(...args);window.entelWorkerUsed=result.excel_worker;return result;};
                    }""")
                page.evaluate('window.uiTicks=0;window.uiTimer=setInterval(()=>uiTicks++,10)')
                page.locator('#btn-process-files').click()
                page.wait_for_function("UploadDiagnostics.runs.length && UploadDiagnostics.runs[0].status !== 'processing'")
                page.evaluate('clearInterval(uiTimer)')
                run=page.evaluate('UploadDiagnostics.snapshot().runs[0]')
                self.measurements.setdefault(module,{}).update({'processing':run,'ui_ticks':page.evaluate('uiTicks')})
                self.assertEqual(run['status'],'success')
                for phase in ['pdf_read_ms','excel_read_ms','inventory_match_ms','reconciliation_ms','workbook_ms']:
                    self.assertGreater(run['timings'][phase],0,phase)
                self.assertGreater(page.evaluate('uiTicks'),1)
                if module=='entel':
                    self.assertTrue(page.evaluate('entelWorkerUsed'))
                    self.assertEqual(page.evaluate('uploadProbe.fileReads'),{'excel.xlsx':1,'inventory.xlsx':1,'pdf.pdf':1})
                    self.assertEqual(page.evaluate('uploadProbe.posts'),[])
                else: self.assertEqual(len(page.evaluate('uploadProbe.posts')),1)


if __name__ == '__main__': unittest.main(verbosity=2)
