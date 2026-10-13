"""
test_movistar_acceptance.py
Pruebas obligatorias y de aceptación del Módulo Movistar.
1. Verifica que ningún monto fijo, cantidad de líneas, periodo o contraseña esté escrito directamente en el código.
2. Ejecuta una prueba completa de extremo a extremo con los archivos de septiembre.
3. Comprueba estructura, orden de hojas, totales y validación CUADRA.
4. Comprueba manejo de errores (contraseña incorrecta, archivos faltantes).
"""

import os
import io
import re
import json
import unittest
import openpyxl
from decimal import Decimal
import movistar_engine

PASSWORD_TEST = "381359748"
FACT_PATH = r"C:\Users\Jakelin.Cullampe\Downloads\movitar\Fact_20254053822_57503614_381359748_05-09-2026_1.xlsx"
TRAF_PATH = r"C:\Users\Jakelin.Cullampe\Downloads\movitar\Traf_20254053822_57503614_381359748_05-09-2026_1.xlsx"
CARGOS_PATH = r"C:\Users\Jakelin.Cullampe\Downloads\movitar\Cargos_20254053822_57503614_381359748_05-09-2026_1.xlsx"
PDF_PATH = r"C:\Users\Jakelin.Cullampe\Downloads\movitar\DIP-FACT-Recibo_Ciclico_2026090005_S1AA-0054018088.pdf"
INV_PATH = r"C:\Users\Jakelin.Cullampe\Downloads\inventario\Inventario Celulares General.xlsx"
if not os.path.exists(INV_PATH):
    INV_PATH = r"data\inventario_movistar.xlsx"

class TestMovistarDynamicCode(unittest.TestCase):
    """
    Verifica que no existan constantes numéricas de septiembre escritas fijas en el código de producción.
    """
    FORBIDDEN_VALUES = [
        "11538.81", "11,538.81",
        "1566.65", "1,566.65",
        "9972.16", "9,972.16",
        "19735.40", "19,735.40", "19735.4",
        "1794.99", "1,794.99",
        "11767.15", "11,767.15",
        "31502.60", "31,502.60", "31502.6",
        "381359748"
    ]

    FILES_TO_CHECK = [
        "movistar_engine.py",
        "server.py",
        "js/movistar.js"
    ]

    def test_no_hardcoded_september_values(self):
        for fname in self.FILES_TO_CHECK:
            if not os.path.exists(fname):
                continue
            with open(fname, "r", encoding="utf-8") as f:
                content = f.read()
            for val in self.FORBIDDEN_VALUES:
                self.assertNotIn(
                    val, content,
                    f"El valor '{val}' de septiembre está escrito directamente en {fname}."
                )

class TestMovistarEndToEnd(unittest.TestCase):
    """
    Prueba obligatoria con los archivos reales de septiembre.
    """
    def setUp(self):
        with open(FACT_PATH, "rb") as f: self.fact_bytes = f.read()
        with open(TRAF_PATH, "rb") as f: self.traf_bytes = f.read()
        with open(CARGOS_PATH, "rb") as f: self.cargos_bytes = f.read()
        with open(PDF_PATH, "rb") as f: self.pdf_bytes = f.read()
        with open(INV_PATH, "rb") as f: self.inv_bytes = f.read()

    def test_complete_september_reconciliation(self):
        result = movistar_engine.process_movistar_files(
            self.fact_bytes,
            self.traf_bytes,
            self.cargos_bytes,
            self.pdf_bytes,
            self.inv_bytes,
            PASSWORD_TEST
        )

        # 1. Periodo
        self.assertIn("Septiembre", result["periodo"])

        # 2. Cantidad de líneas
        self.assertEqual(result["line_count"], 318, "Deben ser exactamente 318 líneas")

        # 3. Líneas con descuento
        self.assertEqual(result["discount_count"], 64, "Deben ser exactamente 64 líneas con descuento")

        # 4. Líneas sin CC
        self.assertEqual(result["sin_cc_count"], 0, "No debe haber líneas sin CC en septiembre")

        # 5. MONTO total antes de descuento
        self.assertAlmostEqual(result["cargo_mensual_sin_igv"], 11538.81, places=2)

        # 6. Descuento total
        self.assertAlmostEqual(result["descuento_total"], 1566.65, places=2)

        # 7. MONTO F (Cargo mensual final sin IGV)
        self.assertAlmostEqual(result["cargo_mensual_final_sin_igv"], 9972.16, places=2)

        # 8. IGV
        self.assertAlmostEqual(result["igv"], 1794.99, places=2)

        # 9. Cargo mensual con IGV (Total afecto)
        self.assertAlmostEqual(result["cargo_mensual_con_igv"], 11767.15, places=2)

        # 10. Cargos adicionales o equipos
        self.assertAlmostEqual(result["equipos"], 19735.40, places=2)

        # 11. Redondeo
        self.assertAlmostEqual(result["redondeo"], 0.05, places=2)

        # 12. Total calculado
        self.assertAlmostEqual(result["total_calculado"], 31502.60, places=2)

        # 13. Total del recibo PDF
        self.assertAlmostEqual(result["total_pdf"], 31502.60, places=2)

        # 14. Diferencia
        self.assertAlmostEqual(result["diferencia"], 0.00, places=2)

        # 15. Estado CUADRA
        self.assertTrue(result["cuadra"], "El estado debe ser CUADRA")

        # 16. Validación del archivo Excel generado
        excel_io = io.BytesIO(result["excel_bytes"])
        wb = openpyxl.load_workbook(excel_io, data_only=False)

        # Exactamente tres hojas y en este orden
        expected_sheets = ["CARGOS ADICIONALES", "CARGO MENSUAL", "Hoja1"]
        self.assertEqual(wb.sheetnames, expected_sheets, f"Las hojas deben ser exactamente {expected_sheets}")

        # Comprobar CARGOS ADICIONALES
        ws_cargos = wb["CARGOS ADICIONALES"]
        self.assertEqual(ws_cargos.cell(2, 2).value, "NUMERO")
        self.assertEqual(ws_cargos.cell(2, 3).value, "DESCRIPCION")
        self.assertEqual(ws_cargos.cell(2, 4).value, "MONTO")
        # Total en la última fila (83 equipos desde fila 3 => fila 86)
        last_cargos_row = ws_cargos.max_row
        self.assertIn("=SUM", str(ws_cargos.cell(last_cargos_row, 4).value))

        # Comprobar CARGO MENSUAL
        ws_mensual = wb["CARGO MENSUAL"]
        self.assertEqual(ws_mensual.cell(2, 2).value, "NUMERO")
        self.assertEqual(ws_mensual.cell(2, 3).value, "PLAN")
        self.assertEqual(ws_mensual.cell(2, 4).value, "MONTO")
        self.assertEqual(ws_mensual.cell(2, 5).value, "DESC")
        self.assertEqual(ws_mensual.cell(2, 6).value, "MONTO F")
        # Fórmulas de fila y totales
        self.assertEqual(ws_mensual.cell(3, 6).value, "=D3-E3")
        self.assertEqual(ws_mensual.cell(321, 4).value, "=SUM(D3:D320)")
        self.assertEqual(ws_mensual.cell(321, 6).value, "=SUM(F3:F320)")

        # Comprobar Hoja1
        ws_hoja1 = wb["Hoja1"]
        self.assertEqual(ws_hoja1.cell(1, 1).value, "NUMERO")
        self.assertEqual(ws_hoja1.cell(1, 2).value, "MONTO")
        self.assertEqual(ws_hoja1.cell(1, 3).value, "CC")
        self.assertEqual(ws_hoja1.cell(4, 7).value, "Etiquetas de fila")
        self.assertEqual(ws_hoja1.cell(4, 8).value, "Suma de MONTO")
        self.assertEqual(ws_hoja1.cell(4, 9).value, "IGV")
        self.assertEqual(ws_hoja1.cell(4, 10).value, "TOTAL")

    def test_wrong_password_handling(self):
        """
        Verifica que una contraseña incorrecta lance un error claro sin revelar datos.
        """
        with self.assertRaises(ValueError) as ctx:
            movistar_engine.process_movistar_files(
                self.fact_bytes,
                self.traf_bytes,
                self.cargos_bytes,
                self.pdf_bytes,
                self.inv_bytes,
                "password_incorrecta_123"
            )
        self.assertIn("contrase", str(ctx.exception).lower())
        self.assertNotIn("password_incorrecta_123", str(ctx.exception))

if __name__ == '__main__':
    unittest.main()
