"""Comprueba geometría y desplazamiento de ambos menús al mismo zoom (100%)."""
import json
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
        cls.out = ROOT / "data/layout-entel"
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
            dpr:devicePixelRatio,zoom:getComputedStyle(document.documentElement).zoom};
        }""")

    def test_laptop_dimensions_match_and_all_options_fit(self):
        for width, height in [(1366, 768), (1280, 720), (1024, 600), (1024, 576)]:
            with self.subTest(viewport=(width, height)):
                context = self.browser.new_context(viewport={"width": width, "height": height}, device_scale_factor=1)
                metrics = {}
                for module in ["movistar", "entel"]:
                    page = self.page(context, module)
                    data = metrics[module] = self.metrics(page)
                    self.reports.append({"module": module, "viewport": [width, height], **data})
                    self.assertEqual(data["dpr"], 1)
                    self.assertEqual(float(data["zoom"]), 1)
                    self.assertEqual(data["sidebar"]["width"], 240)
                    self.assertLessEqual(data["sidebar"]["scrollHeight"], data["sidebar"]["clientHeight"] + 1)
                    self.assertLessEqual(data["sidebar"]["scrollWidth"], data["sidebar"]["clientWidth"])
                    for option in data["options"]:
                        self.assertGreaterEqual(option["top"], 0)
                        self.assertLessEqual(option["bottom"], height)
                    self.assertLessEqual(data["footer"]["bottom"], height + 1)
                    if width == 1366:
                        page.screenshot(path=str(self.out / (module + "-after.png")))
                for element in ["header", "button", "badge"]:
                    for key in ["width", "height", "fontSize", "fontFamily", "padding", "margin", "gap"]:
                        self.assertEqual(metrics["entel"][element][key], metrics["movistar"][element][key], element + "." + key)
                for key in ["x", "width", "padding", "margin", "fontSize"]:
                    self.assertEqual(metrics["entel"]["main"][key], metrics["movistar"]["main"][key], "main." + key)
                context.close()

    def test_small_height_keeps_scroll_and_right_content_is_independent(self):
        context = self.browser.new_context(viewport={"width": 1280, "height": 400}, device_scale_factor=1)
        for module in ["movistar", "entel"]:
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
