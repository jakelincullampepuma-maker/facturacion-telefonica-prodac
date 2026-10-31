"""Pruebas de navegador Entel, sin backend. Requiere playwright para desarrollo.
Ejecutar: python test_entel_browser.py
Los archivos reales permanecen en data/entel_test (ignorado por Git).
"""
import contextlib
import functools
import hashlib
import http.server
import io
import json
from pathlib import Path
import sys
import threading
import unittest
import zipfile
import openpyxl

ROOT = Path(__file__).resolve().parent
if (ROOT / ".test-tools").exists():
    sys.path.insert(0, str(ROOT / ".test-tools"))
from playwright.sync_api import sync_playwright

DATA = ROOT / "data/entel_test"
OUT = ROOT / "data/browser-results"
PYTHON_ENGINE = None


def simple_pdf(text):
    # PDF mínimo con texto seleccionable: prueba de lógica, no documento real.
    stream = ("BT /F1 12 Tf 40 750 Td (" + text + ") Tj ET").encode("ascii")
    objects = [b"<< /Type /Catalog /Pages 2 0 R >>",
               b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
               b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
               b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
               b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream"]
    data = b"%PDF-1.4\n"
    offsets = [0]
    for i, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data += str(i).encode() + b" 0 obj\n" + obj + b"\nendobj\n"
    xref = len(data)
    data += b"xref\n0 6\n0000000000 65535 f \n"
    data += b"".join(f"{o:010} 00000 n \n".encode() for o in offsets[1:])
    data += f"trailer << /Size 6 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode()
    return data


class BrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        OUT.mkdir(parents=True, exist_ok=True)
        handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(ROOT))
        cls.httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f"http://127.0.0.1:{cls.httpd.server_port}"
        cls.pw = sync_playwright().start()
        chrome = Path("C:/Program Files/Google/Chrome/Application/chrome.exe")
        cls.browser = cls.pw.chromium.launch(executable_path=str(chrome) if chrome.exists() else None, headless=True)
        cls.context = cls.browser.new_context(accept_downloads=True)
        cls.protected = {p: hashlib.sha256((ROOT / p).read_bytes()).hexdigest() for p in [
            "claro.html", "index.html", "movistar.html", "js/app.js", "js/excelParser.js",
            "js/pdfParser.js", "js/matcher.js", "js/excelExporter.js", "js/movistar.js", "server.py", "entel_engine.py"]}
        (OUT / "baseline-hashes.json").write_text(json.dumps(cls.protected, indent=2))

    @classmethod
    def tearDownClass(cls):
        cls.context.close()
        cls.browser.close()
        cls.pw.stop()
        cls.httpd.shutdown()
        cls.httpd.server_close()
        for p, digest in cls.protected.items():
            assert hashlib.sha256((ROOT / p).read_bytes()).hexdigest() == digest, p + " fue modificado durante las pruebas"

    def entel_page(self, local=False):
        page = self.context.new_page()
        page.set_default_timeout(60000)
        page.goto((ROOT / "entel.html").as_uri() if local else self.base + "/entel.html")
        page.wait_for_function("window.EntelProcessor && window.EntelWorkbook && window.XLSX && window.pdfjsLib && window.JSZip")
        # Mantiene el resultado observable sólo dentro del navegador de prueba.
        page.evaluate("""() => {
          const run = EntelProcessor.process;
          EntelProcessor.process = async (...args) => {
            try { const result = await run(...args); window.testResult = result; return result; } finally { window.testCompleted = (window.testCompleted || 0) + 1; }
          };
        }""")
        return page

    def load_and_process(self, page, excel=None, inventory=None, pdf=None):
        for selector in ["#input-excel-file", "#input-inventory-file", "#input-pdf-file"]:
            page.locator(selector).set_input_files([])
        completed = page.evaluate("window.testCompleted || 0")
        page.locator("#input-excel-file").set_input_files(str(excel or DATA / "Entel - May.xlsx"))
        page.locator("#input-inventory-file").set_input_files(str(inventory or DATA / "Inventario Celulares General.xlsx"))
        page.locator("#input-pdf-file").set_input_files(str(pdf or DATA / "Entel - May.pdf"))
        page.locator("#btn-process-files").click()
        page.wait_for_function("(before) => (window.testCompleted || 0) > before", arg=completed)
        page.wait_for_function("!document.getElementById('btn-process-files').disabled")
        return page.locator("#entel-alert-container").inner_text()

    def assert_originals(self, output, original):
        wb = openpyxl.load_workbook(output)
        source = openpyxl.load_workbook(original)
        self.assertEqual(wb.sheetnames, ["Resultado Entel"] + source.sheetnames)
        with zipfile.ZipFile(output) as outzip, zipfile.ZipFile(original) as srczip:
            allowed = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml", "xl/styles.xml"}
            for name in srczip.namelist():
                self.assertIn(name, outzip.namelist())
                if name not in allowed:
                    self.assertEqual(srczip.read(name), outzip.read(name), "Se alteró la entrada original " + name)
        self.assert_compact_result(output, self.result)

    def assert_compact_result(self, output, result):
        wb = openpyxl.load_workbook(output)
        ws = wb["Resultado Entel"]
        cached = openpyxl.load_workbook(output, data_only=True)["Resultado Entel"]
        count = result["line_count"]
        self.assertEqual([ws.cell(1, c).value for c in range(1, 4)], ["NUMERO", "MONTO", "CC"])
        self.assertEqual(ws.max_column, 8)
        self.assertTrue(all(ws.cell(r, c).value is None for r in range(1, ws.max_row + 1) for c in [4, 5, 6]))
        for row, line in enumerate(result["lines"], 2):
            self.assertEqual(ws.cell(row, 1).value, line["phone"])
            self.assertEqual(ws.cell(row, 1).data_type, "s")
            self.assertAlmostEqual(ws.cell(row, 2).value, line["cargo_fijo"])
            self.assertEqual(ws.cell(row, 2).number_format, "#,##0.00")
            self.assertEqual(ws.cell(row, 3).value, line["cc_dpto"] or "SIN CC")
            self.assertEqual(ws.cell(row, 3).number_format, "@")
        self.assertEqual(ws.cell(count + 2, 1).value, "Total general")
        self.assertAlmostEqual(cached.cell(count + 2, 2).value, result["total_excel"])
        self.assertEqual(len(ws.tables), 1)
        table = list(ws.tables.values())[0]
        self.assertEqual(table.ref, "A1:C" + str(count + 2))
        self.assertEqual(table.autoFilter.ref, "A1:C" + str(count + 1))
        self.assertEqual(table.totalsRowCount, 1)
        self.assertEqual(ws.cell(count + 2, 2).value, "=SUM(" + table.name + "[MONTO])")
        self.assertEqual(ws["G4"].value, "Etiquetas de fila")
        self.assertEqual(ws["H4"].value, "Suma de MONTO")
        expected = {}
        for line in result["lines"]:
            cc = line["cc_dpto"] or "SIN CC"
            expected[cc] = expected.get(cc, 0) + round(line["cargo_fijo"] * 100)
        self.assertEqual([ws.cell(r, 7).value for r in range(5, len(expected) + 5)], sorted(expected))
        for r in range(5, len(expected) + 5):
            self.assertAlmostEqual(ws.cell(r, 8).value, expected[ws.cell(r, 7).value] / 100)
        rt = len(expected) + 5
        self.assertEqual(ws.cell(rt, 7).value, "Total general")
        self.assertAlmostEqual(ws.cell(rt, 8).value, result["total_excel"])
        self.assertTrue(ws.cell(rt, 8).font.bold)
        self.assertEqual(ws["G4"].fill.fgColor.rgb, "FFC0E6F5")
        self.assertEqual(ws.cell(rt, 8).fill.fgColor.rgb, "FFC0E6F5")
        # Una definición nativa, con caché actualizable y origen en la tabla Excel.
        self.assertEqual(len(ws._pivots), 1)
        pivot = ws._pivots[0]
        self.assertEqual(pivot.location.ref, "G4:H" + str(rt))
        self.assertEqual(pivot.rowFields[0].x, 2)
        self.assertEqual(pivot.dataFields[0].fld, 1)
        self.assertEqual(pivot.dataFields[0].subtotal, "sum")
        self.assertEqual(pivot.cache.cacheSource.worksheetSource.name, table.name)
        self.assertEqual(pivot.cache.recordCount, count)
        self.assertEqual(len(pivot.cache.records.r), count)

    @unittest.skipUnless(all((DATA / n).exists() for n in ["Entel - May.xlsx", "Entel - May.pdf", "Inventario Celulares General.xlsx"]), "Referencias privadas de mayo no disponibles.")
    def test_01_may_static_and_file_without_backend(self):
        import entel_engine
        expected = entel_engine.process_entel_files((DATA / "Entel - May.xlsx").read_bytes(), (DATA / "Entel - May.pdf").read_bytes(), (DATA / "Inventario Celulares General.xlsx").read_bytes())
        for local in [False, True]:
            with self.subTest(file_protocol=local):
                page = self.entel_page(local)
                api_calls = []
                page.on("request", lambda r: api_calls.append(r.url) if "/api/" in r.url else None)
                errors = []
                page.on("pageerror", lambda e: errors.append(str(e)))
                message = self.load_and_process(page)
                self.assertIn("Los importes de PLAN coinciden", message)
                self.result = page.evaluate("(({excel_blob, ...r}) => r)(window.testResult)")
                for key in ["line_count", "total_excel", "total_pdf_plan", "total_pdf_equipos", "total_pdf_factura", "diferencia", "cuadra", "cc_discrepancies", "cc_summary"]:
                    self.assertEqual(self.result[key], expected[key], key)
                self.assertEqual(self.result["line_count"], 16)
                self.assertEqual(self.result["total_excel"], 924.60)
                self.assertEqual(self.result["diferencia"], 0)
                self.assertEqual(len(self.result["cc_discrepancies"]), 3)
                self.assertEqual(len(api_calls), 0)
                self.assertEqual(errors, [])
                self.assertEqual(page.locator("#tbody-lineas-entel tr").count(), 16)
                page.screenshot(path=str(OUT / ("entel-file.png" if local else "entel-static.png")), full_page=True)
                with page.expect_download() as d:
                    page.locator("#btn-export-excel").click()
                output = OUT / ("Entel-file.xlsx" if local else "Entel-static.xlsx")
                d.value.save_as(output)
                self.assert_originals(output, DATA / "Entel - May.xlsx")
                with page.expect_download() as d:
                    page.locator("#btn-download-obs").click()
                d.value.save_as(OUT / "Observaciones Entel.csv")
                csv = (OUT / "Observaciones Entel.csv").read_text(encoding="utf-8-sig")
                for item in self.result["cc_discrepancies"]:
                    self.assertIn(item["phone"], csv)
                page.locator("#input-search-lines").fill(self.result["lines"][0]["phone"])
                self.assertEqual(page.locator("#tbody-lineas-entel tr").count(), 1)
                page.locator("#btn-remove-inventory").click()
                self.assertTrue(page.locator("#btn-export-excel").is_disabled())
                page.close()
        (OUT / "may-result.json").write_text(json.dumps(self.result, ensure_ascii=False, indent=2), encoding="utf-8")

    def test_02_dynamic_structure_and_validation(self):
        # Distintos teléfonos, montos, nombres de hoja, columnas y filas iniciales.
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Detalle nuevo"
        ws.append(["Reporte de otro periodo"])
        ws.append([])
        ws.append(["Cargo Fijo", "Plan Tarifario", "Numero Entel"])
        ws.append([11.25, "Plan variable A", "951234567"])
        ws.append([19.75, "Plan variable B", "952345678"])
        inv = openpyxl.Workbook()
        inv.active.title = "Datos actualizados"
        inv.active.append(["CC Dpto", "Usuario", "Numero", "Operador", "CC Area", "Departamento", "Area"])
        inv.active.append(["CC-NUEVO-A", "Usuario A", "951234567", "Entel", "AREA-A", "Dpto A", "Área A"])
        inv.active.append(["CC-NUEVO-B", "Usuario B", "952345678", "Entel", "AREA-B", "Dpto B", "Área B"])
        x, i, p = OUT / "dynamic.xlsx", OUT / "inventory.xlsx", OUT / "dynamic.pdf"
        wb.save(x); inv.save(i)
        p.write_bytes(simple_pdf("Plan Prorrateo Cargos Descuentos Equipos Saldo S/ 31.00 S/ 0.00 S/ 0.00 S/ 0.00 S/ 0.00 S/ 0.00 Total a pagar S/ 31.00"))
        page = self.entel_page()
        self.assertIn("coinciden", self.load_and_process(page, x, i, p))
        r = page.evaluate("(({excel_blob, ...r}) => r)(window.testResult)")
        self.assertEqual(r["line_count"], 2)
        self.assertEqual(r["total_excel"], 31)
        self.assertEqual(r["diferencia"], 0)
        self.assertTrue(r["es_valido_todo"])
        with page.expect_download() as d:
            page.locator("#btn-export-excel").click()
        dynamic_out = OUT / "Entel-dynamic.xlsx"; d.value.save_as(dynamic_out)
        self.assert_compact_result(dynamic_out, r)
        # Sin CC, operador incorrecto, duplicados y diferencia no se silencian.
        ws.append([3.50, "Otra línea", "951234567"])
        wb.save(x)
        inv.active.cell(2, 1).value = None
        inv.active.cell(3, 4).value = "Claro"
        inv.save(i)
        self.load_and_process(page, x, i, p)
        r = page.evaluate("(({excel_blob, ...r}) => r)(window.testResult)")
        self.assertEqual(r["dup_count"], 1)
        self.assertEqual(r["no_inv_count"], 1)
        self.assertEqual(r["sin_cc_count"], 2)
        self.assertEqual(r["diferencia"], 3.5)
        self.assertFalse(r["cuadra"])
        self.assertFalse(r["es_valido_todo"])
        with page.expect_download() as d:
            page.locator("#btn-export-excel").click()
        missing_out = OUT / "Entel-sin-cc.xlsx"; d.value.save_as(missing_out)
        self.assert_compact_result(missing_out, r)
        # Importe corrupto debe bloquear, nunca convertirse en cero.
        ws.cell(4, 1).value = "importe inválido"; wb.save(x)
        self.assertIn("no válido", self.load_and_process(page, x, i, p))
        self.assertTrue(page.locator("#btn-export-excel").is_disabled())
        # PDF legible sin PLAN no puede dar una conciliación falsa.
        wb.remove(ws); ws = wb.create_sheet("Detalle alternativo")
        ws.append(["Numero", "Cargo Fijo"]); ws.append(["951234567", 11.25]); wb.save(x)
        p.write_bytes(simple_pdf("Total a pagar S/ 11.25"))
        self.assertIn("concepto PLAN", self.load_and_process(page, x, i, p))
        self.assertTrue(page.locator("#btn-export-excel").is_disabled())
        page.close()


    @unittest.skipUnless(all((DATA / n).exists() for n in ["Entel - May.xlsx", "Entel - May.pdf", "Inventario Celulares General.xlsx"]), "Referencias privadas de mayo no disponibles.")
    def test_025_reprocess_and_corrupt_inputs(self):
        page = self.entel_page()
        self.load_and_process(page)
        with page.expect_download() as d:
            page.locator("#btn-export-excel").click()
        first = OUT / "roundtrip-first.xlsx"; d.value.save_as(first)
        self.load_and_process(page, excel=first)
        self.assertEqual(page.evaluate("window.testResult.total_excel"), 924.60)
        with page.expect_download() as d:
            page.locator("#btn-export-excel").click()
        second = OUT / "roundtrip-second.xlsx"; d.value.save_as(second)
        book = openpyxl.load_workbook(second)
        self.assertEqual(book.sheetnames[0], "Resultado Entel")
        self.assertIn("Resultado Entel anterior 1", book.sheetnames)
        self.assertEqual(len(book.sheetnames), len(openpyxl.load_workbook(first).sheetnames) + 1)
        with zipfile.ZipFile(first) as src, zipfile.ZipFile(second) as out:
            for name in src.namelist():
                if name.startswith("xl/worksheets/"):
                    self.assertEqual(src.read(name), out.read(name))
        bad = OUT / "corrupt.xlsx"; bad.write_bytes(b"archivo incorrecto")
        self.assertIn("No se", self.load_and_process(page, excel=bad))
        self.assertTrue(page.locator("#btn-export-excel").is_disabled())
        bad_pdf = OUT / "corrupt.pdf"; bad_pdf.write_bytes(b"no es un PDF")
        self.assertIn("No se pudieron", self.load_and_process(page, pdf=bad_pdf))
        self.assertTrue(page.locator("#btn-export-excel").is_disabled())
        page.close()

    def test_03_claro_unchanged_synthetic_flow(self):
        # No hay archivos reales de Claro en el repositorio: prueba controlada.
        wb = openpyxl.Workbook()
        wb.active.title = "Factura Detallada"
        wb.active.append(["Línea", "Total Cargos por Línea (Sin IGV)"])
        wb.active.append(["953456789", 17.5])
        x, i, p = OUT / "claro-control.xlsx", OUT / "claro-inventory.xlsx", OUT / "claro-control.pdf"
        wb.save(x)
        inv = openpyxl.Workbook(); inv.active.title = "PRODAC"
        inv.active.append(["Numero", "CC Dpto", "Usuario"])
        inv.active.append(["953456789", "CC-CONTROL", "Usuario prueba"]); inv.save(i)
        p.write_bytes(simple_pdf("Cargos del Mes S/ 17.50"))
        page = self.context.new_page()
        page.set_default_timeout(60000)
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.on("dialog", lambda d: d.accept())
        page.goto(self.base + "/claro.html")
        for input_id, path in [("input-claro-file", x), ("input-inv-file", i), ("input-pdf-file", p)]:
            page.locator("#" + input_id).set_input_files(str(path))
        page.locator("#btn-process-files").click()
        page.wait_for_function("document.getElementById('kpi-val-lineas').textContent === '1'")
        self.assertEqual(page.locator("#kpi-val-lineas").inner_text(), "1")
        self.assertEqual(page.locator("#banner-pill-badge").inner_text(), "CUADRA")
        with page.expect_download() as d:
            page.locator("#btn-export-excel").click()
        out = OUT / "Claro-control-result.xlsx"; d.value.save_as(out)
        result = openpyxl.load_workbook(out)
        self.assertEqual(result.sheetnames, ["Detalle", "Resumen CC", "Conciliación"])
        self.assertEqual(result["Conciliación"]["B7"].value, 0)
        page.locator('a[href="entel.html"]').click()
        page.wait_for_function("window.EntelProcessor && window.JSZip")
        self.assertIn("Módulo Entel", page.title())
        self.assertEqual(errors, [])
        page.close()


if __name__ == "__main__":
    unittest.main(verbosity=2, testRunner=unittest.TextTestRunner(stream=sys.stdout, verbosity=2))
