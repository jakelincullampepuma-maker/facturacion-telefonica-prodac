"""
test_movistar_acceptance.py
Pruebas obligatorias y de aceptación del Módulo Movistar.
1. Verifica que ningún monto fijo, cantidad de líneas, periodo o contraseña esté escrito directamente en el código de producción.
2. Ejecuta la prueba obligatoria de aceptación completa con los archivos de referencia de Noviembre.
3. Comprueba estructura, orden exacto de las 3 hojas, totales, fórmulas, tabla dinámica y estado CUADRA.
4. Comprueba validación opcional con Fact, Traf y Cargos protegidos.
5. Comprueba manejo de errores y validaciones de archivos faltantes.
"""

import os
import io
import re
import json
import unittest
import openpyxl
import zipfile
from decimal import Decimal
from pathlib import Path

import movistar_engine

# Rutas de prueba de Noviembre
NOV_DIR = Path(__file__).resolve().parent / "data" / "movistar_nov_test"
NOV_PDF_PATH = NOV_DIR / "Movistar-Nov.pdf"
NOV_INV_PATH = NOV_DIR / "Inventario Celulares General.xlsx"
NOV_FACT_PATH = NOV_DIR / "Fact_Nov.xlsx"
NOV_TRAF_PATH = NOV_DIR / "Traf_Nov.xlsx"
NOV_CARGOS_PATH = NOV_DIR / "Cargos_Nov.xlsx"
NOV_PASSWORD = os.environ.get('MOVISTAR_TEST_PASSWORD', '')


class TestMovistarDynamicCode(unittest.TestCase):
    """
    Verifica que no existan constantes numéricas fijas de pruebas escritas en el código de producción.
    """
    FORBIDDEN_VALUES = [
        # Valores de septiembre
        "11538.81", "11,538.81",
        "1566.65", "1,566.65",
        "9972.16", "9,972.16",
        "19735.40", "19,735.40",
        "1794.99", "1,794.99",
        "11767.15", "11,767.15",
        "31502.60", "31,502.60",
        # Valores de noviembre
        "12597.72", "12,597.72",
        "1046.68", "1,046.68",
        "11551.04", "11,551.04",
        "25452.11", "25,452.11",
        "39082.40", "39,082.40",
    ]

    FILES_TO_CHECK = [
        "movistar_engine.py",
        "movistar_workbook.py",
        "server.py",
        "js/movistar.js"
    ]

    def test_no_hardcoded_values(self):
        for fname in self.FILES_TO_CHECK:
            if not os.path.exists(fname):
                continue
            with open(fname, "r", encoding="utf-8") as f:
                content = f.read()
            for val in self.FORBIDDEN_VALUES:
                self.assertNotIn(
                    val, content,
                    f"El valor '{val}' está escrito directamente en {fname}."
                )


@unittest.skipUnless(NOV_PDF_PATH.is_file() and NOV_INV_PATH.is_file(), 'Referencias privadas de noviembre no disponibles.')
class TestMovistarNovemberAcceptance(unittest.TestCase):
    """
    Prueba obligatoria de aceptación con los archivos reales de Noviembre.
    """
    def setUp(self):
        self.assertTrue(NOV_PDF_PATH.is_file(), f"Falta archivo de prueba: {NOV_PDF_PATH}")
        self.assertTrue(NOV_INV_PATH.is_file(), f"Falta archivo de prueba: {NOV_INV_PATH}")
        with open(NOV_PDF_PATH, "rb") as f:
            self.pdf_bytes = f.read()
        with open(NOV_INV_PATH, "rb") as f:
            self.inv_bytes = f.read()

    def test_complete_november_reconciliation(self):
        """
        Prueba completa requerida para Noviembre:
        325 líneas, S/ 12,597.72, S/ 1,046.68, S/ 11,551.04, S/ 2,079.19,
        S/ 25,452.11, S/ 0.06, Total S/ 39,082.40, Diferencia S/ 0.00, CUADRA.
        """
        result = movistar_engine.process_movistar_files(
            pdf_bytes=self.pdf_bytes,
            inventory_bytes=self.inv_bytes
        )

        # 1. Periodo y recibo
        self.assertIn("Noviembre", result["periodo"])
        self.assertIn("2026", result["periodo"])
        from pypdf import PdfReader
        source_text = PdfReader(io.BytesIO(self.pdf_bytes)).pages[0].extract_text()
        self.assertTrue(result["recibo_numero"])
        self.assertIn(result["recibo_numero"], source_text)

        # 2. Cantidad de líneas
        self.assertEqual(result["line_count"], 325, "Deben ser exactamente 325 líneas")

        # 3. MONTO total antes de descuentos en DESC
        self.assertAlmostEqual(result["cargo_mensual_sin_igv"], 12597.72, places=2)

        # 4. Descuento total en columna DESC
        self.assertAlmostEqual(result["descuento_total"], 1046.68, places=2)

        # 5. MONTO F (Subtotal sin IGV)
        self.assertAlmostEqual(result["cargo_final_sin_igv"], 11551.04, places=2)
        self.assertAlmostEqual(result["subtotal_pdf"], 11551.04, places=2)

        # 6. IGV (18%)
        self.assertAlmostEqual(result["igv"], 2079.19, places=2)

        # 7. Total afecto
        self.assertAlmostEqual(result["total_afecto"], 13630.23, places=2)

        # 8. Cargos adicionales inafectos
        self.assertAlmostEqual(result["cargos_inafectos"], 25452.11, places=2)

        # 9. Redondeo
        self.assertAlmostEqual(result["redondeo"], 0.06, places=2)

        # 10. Total calculado y Total PDF
        self.assertAlmostEqual(result["total_calculado"], 39082.40, places=2)
        self.assertAlmostEqual(result["pdf_total"], 39082.40, places=2)

        # 11. Diferencia
        self.assertAlmostEqual(result["diferencia"], 0.00, places=2)

        # 12. Estado CUADRA
        self.assertTrue(result["cuadra"])
        self.assertEqual(result["estado"], "CUADRA")

        # 13. Desglose de descuentos
        db = result["discount_breakdown"]
        self.assertEqual(db["count_3m"], 7, "Deben ser 7 descuentos de 3 meses")
        self.assertEqual(db["count_perm"], 7, "Deben ser 7 descuentos permanentes")
        self.assertEqual(db["count_18m"], 149, "Deben ser 149 descuentos de 18 meses incorporados en MONTO")
        self.assertAlmostEqual(db["total_desc_monto"], 1046.68, places=2)

        # 14. Desglose de inafectos (87 equipos + 5 ventas = 92)
        ib = result["inafectos_breakdown"]
        self.assertEqual(ib["count_equipos"], 87, "Deben ser 87 financiamientos de equipos")
        self.assertEqual(ib["count_ventas"], 5, "Deben ser 5 financiamientos de venta")
        self.assertEqual(ib["total_count"], 92, "Deben ser 92 registros inafectos en total")
        self.assertAlmostEqual(ib["total_inafectos"], 25452.11, places=2)

        # 15. Líneas sin CC
        self.assertEqual(result["sin_cc_count"], 7, "Deben identificarse 7 líneas sin CC en el inventario")

        # 16. Validación de la estructura del Excel generado
        excel_io = io.BytesIO(result["excel_bytes"])
        wb = openpyxl.load_workbook(excel_io, data_only=False)

        # Tres hojas exactas en orden
        expected_sheets = ["CARGOS ADICIONALES", "CARGO MENSUAL", "Hoja1"]
        self.assertEqual(wb.sheetnames, expected_sheets)

        # Comprobar hoja CARGOS ADICIONALES
        ws_ca = wb["CARGOS ADICIONALES"]
        self.assertEqual([c.value for c in ws_ca[1]], ["NUMERO", "DESCRIPCION", "MONTO"])
        self.assertEqual(ws_ca.cell(94, 3).value, "=SUM(C2:C93)")
        self.assertEqual(ws_ca.auto_filter.ref, 'A1:C93')

        # Comprobar hoja CARGO MENSUAL
        ws_cm = wb["CARGO MENSUAL"]
        self.assertEqual([c.value for c in ws_cm[1]], ["N°", "NUMERO", "PLAN", "MONTO", "DESC", "MONTO F"])
        self.assertEqual(ws_cm.cell(2, 6).value, "=D2-E2")
        self.assertEqual(ws_cm.cell(327, 4).value, "=SUM(D2:D326)")
        self.assertEqual(ws_cm.cell(327, 5).value, "=SUM(E2:E326)")
        self.assertEqual(ws_cm.cell(327, 6).value, "=SUM(F2:F326)")
        self.assertAlmostEqual(float(ws_cm.cell(328, 4).value), 11551.04, places=2)
        self.assertEqual(ws_cm.cell(329, 4).value, "=F327-D328")
        self.assertEqual(ws_cm.auto_filter.ref, 'A1:F326')

        # Comprobar hoja Hoja1
        ws_h1 = wb["Hoja1"]
        self.assertEqual(ws_h1.cell(1, 1).value, "NUMERO")
        self.assertEqual(ws_h1.cell(1, 2).value, "MONTO")
        self.assertEqual(ws_h1.cell(1, 3).value, "CC")
        self.assertEqual(ws_h1.cell(327, 2).value, "=SUM(MovistarDatos[MONTO])")
        source = ws_h1.tables['MovistarDatos']
        self.assertEqual(source.ref, 'A1:C327')
        self.assertEqual(source.autoFilter.ref, 'A1:C326')
        self.assertEqual(source.totalsRowCount, 1)
        self.assertEqual([col.name for col in source.tableColumns], ['NUMERO', 'MONTO', 'CC'])
        self.assertEqual(ws_h1.cell(4, 7).value, "Etiquetas de fila")
        self.assertEqual(ws_h1.cell(4, 8).value, "Suma de MONTO")

        # Comprobar preservación y auto-refresco de la Tabla Dinámica
        with zipfile.ZipFile(excel_io, 'r') as z:
            namelist = z.namelist()
            self.assertIn("xl/pivotCache/pivotCacheDefinition1.xml", namelist)
            self.assertIn("xl/pivotTables/pivotTable1.xml", namelist)
            cache_xml = z.read("xl/pivotCache/pivotCacheDefinition1.xml").decode("utf-8")
            self.assertIn('refreshOnLoad="1"', cache_xml)
            self.assertIn('worksheetSource name="MovistarDatos"', cache_xml)
            self.assertNotIn('worksheetSource ref=', cache_xml)
            self.assertIn('xl/tables/table1.xml', namelist)

    def test_optional_validation_files(self):
        """
        Prueba la sección de validación opcional con Fact, Traf y Cargos protegidos.
        """
        if not (NOV_PASSWORD and NOV_FACT_PATH.is_file() and NOV_TRAF_PATH.is_file() and NOV_CARGOS_PATH.is_file()):
            self.skipTest("Archivos opcionales o contraseña mensual de prueba no disponibles")

        with open(NOV_FACT_PATH, "rb") as f: fb = f.read()
        with open(NOV_TRAF_PATH, "rb") as f: tb = f.read()
        with open(NOV_CARGOS_PATH, "rb") as f: cb = f.read()

        result = movistar_engine.process_movistar_files(
            pdf_bytes=self.pdf_bytes,
            inventory_bytes=self.inv_bytes,
            fact_bytes=fb,
            traf_bytes=tb,
            cargos_bytes=cb,
            password=NOV_PASSWORD
        )

        self.assertTrue(result["cuadra"])
        warnings = result.get("optional_warnings", [])
        self.assertTrue(len(warnings) > 0, "Debe generar notas de validación opcional")
        # Debe informar la diferencia de S/ 5.00 en Cargos
        cargos_warn = [w for w in warnings if "Cargos:" in w]
        self.assertTrue(len(cargos_warn) > 0)
        self.assertIn("5.00", cargos_warn[0])

    def test_missing_files_error_handling(self):
        """
        Verifica que la omisión de archivos obligatorios genere un error claro.
        """
        with self.assertRaises(ValueError) as ctx:
            movistar_engine.process_movistar_files(pdf_bytes=None, inventory_bytes=self.inv_bytes)
        self.assertIn("PDF", str(ctx.exception))

        with self.assertRaises(ValueError) as ctx:
            movistar_engine.process_movistar_files(pdf_bytes=self.pdf_bytes, inventory_bytes=None)
        self.assertIn("Inventario", str(ctx.exception))


if __name__ == '__main__':
    unittest.main()
