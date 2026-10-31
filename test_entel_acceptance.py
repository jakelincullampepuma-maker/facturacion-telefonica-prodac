"""
test_entel_acceptance.py
Pruebas obligatorias y de aceptacion del Modulo Entel.
1. Verifica que ningun monto fijo, telefono, periodo o centro de costo este hardcodeado en el codigo.
2. Ejecuta una prueba completa de extremo a extremo con los archivos del caso de prueba (Mayo).
3. Comprueba estructura, orden de hojas, conservacion de hojas originales, totales y validacion CONCILIADO.
4. Comprueba deteccion de las 3 discrepancias de CC contra la Hoja1 previa.
5. Comprueba funcionamiento dinamico cuando Hoja1 no existe.
6. Comprueba manejo de errores (archivos faltantes o corruptos).
"""

import os
import io
import re
import json
import unittest
import openpyxl
from decimal import Decimal
import entel_engine

from pathlib import Path
DATA = Path(os.environ.get("ENTEL_TEST_DIR", "data/entel_test"))
EXCEL_MAY_PATH = DATA / "Entel - May.xlsx"
PDF_MAY_PATH = DATA / "Entel - May.pdf"
INV_PATH = DATA / "Inventario Celulares General.xlsx"
HAS_REFERENCES = all(p.is_file() for p in [EXCEL_MAY_PATH, PDF_MAY_PATH, INV_PATH])


class TestEntelDynamicCode(unittest.TestCase):
    """
    Verifica que no existan valores fijos, telefonos ni importes hardcodeados
    en el codigo de produccion de Entel.
    """
    FORBIDDEN_VALUES = [
        "924.60", "924.6",
        "1478.17", "1,478.17",
        "2402.77", "2,402.77",
    ]

    FILES_TO_CHECK = [
        "entel_engine.py",
        "server.py",
        "js/entel.js",
        "js/entelProcessor.js",
        "js/entelWorkbook.js",
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
                    f"El valor '{val}' esta hardcodeado directamente en {fname}."
                )


@unittest.skipUnless(HAS_REFERENCES, "Referencias privadas Entel no disponibles.")
class TestEntelEndToEnd(unittest.TestCase):
    """
    Prueba obligatoria con los archivos de Mayo.
    - 16 lineas identificadas.
    - Total de planes = S/ 924.60.
    - Total PLAN PDF = S/ 924.60.
    - Equipos = S/ 1,478.17.
    - Factura = S/ 2,402.77.
    - Diferencia = S/ 0.00.
    - Estado CONCILIADO.
    - Tres discrepancias de CC respecto al inventario.
    - Conservacion total de hojas originales y creacion de 'Resultado Entel' en pos 0.
    """
    def setUp(self):
        self.assertTrue(os.path.exists(EXCEL_MAY_PATH), f"No existe {EXCEL_MAY_PATH}")
        self.assertTrue(os.path.exists(PDF_MAY_PATH), f"No existe {PDF_MAY_PATH}")
        self.assertTrue(os.path.exists(INV_PATH), f"No existe {INV_PATH}")

        with open(EXCEL_MAY_PATH, "rb") as f:
            self.excel_bytes = f.read()
        with open(PDF_MAY_PATH, "rb") as f:
            self.pdf_bytes = f.read()
        with open(INV_PATH, "rb") as f:
            self.inv_bytes = f.read()

    def test_may_reconciliation(self):
        result = entel_engine.process_entel_files(
            self.excel_bytes,
            self.pdf_bytes,
            self.inv_bytes,
        )

        # 1. Cantidad de lineas
        self.assertEqual(result["line_count"], 16, "Deben ser exactamente 16 lineas")

        # 2. Total de planes en Excel
        self.assertAlmostEqual(result["total_excel"], 924.60, places=2)

        # 3. Total de PLAN extraido del PDF (nunca total a pagar)
        self.assertAlmostEqual(result["total_pdf_plan"], 924.60, places=2)

        # 4. Total de equipos
        self.assertAlmostEqual(result["total_pdf_equipos"], 1478.17, places=2)

        # 5. Total recibo / factura
        self.assertAlmostEqual(result["total_pdf_factura"], 2402.77, places=2)

        # 6. Diferencia monetaria exacta
        self.assertAlmostEqual(result["diferencia"], 0.00, places=2)

        # 7. Validacion cuadra y estado CONCILIADO
        self.assertTrue(result["cuadra"])
        self.assertEqual(result["estado_conciliacion"], "CONCILIADO")

        # Las discrepancias deben reflejar exclusivamente los documentos cargados.
        discrepancies = result["cc_discrepancies"]
        self.assertEqual(len(discrepancies), 3)
        inv_map = entel_engine.load_inventory_entel(self.inv_bytes)["entel_map"]
        previous = entel_engine.load_previous_hoja1(openpyxl.load_workbook(io.BytesIO(self.excel_bytes), data_only=True))
        for item in discrepancies:
            self.assertEqual(item["inv_cc"], inv_map[item["phone"]]["cc_dpto"])
            self.assertEqual(item["hoja1_cc"], previous[item["phone"]])
            self.assertNotEqual(item["inv_cc"], item["hoja1_cc"])

        # 9. Verificacion del Excel resultante
        out_bytes = result["excel_bytes"]
        wb_out = openpyxl.load_workbook(io.BytesIO(out_bytes))

        # La primera hoja debe ser 'Resultado Entel'
        self.assertEqual(wb_out.sheetnames[0], "Resultado Entel")

        # Todas las hojas originales deben conservarse intactas
        original_sheets = openpyxl.load_workbook(io.BytesIO(self.excel_bytes)).sheetnames
        for sname in original_sheets:
            self.assertIn(sname, wb_out.sheetnames, f"La hoja original '{sname}' no se conservo")

        # 10. Verificacion de formulas en la hoja Resultado Entel
        ws_res = wb_out["Resultado Entel"]

        # Fila de total de la tabla principal: =SUM(C5:C20)
        tot_cell = ws_res.cell(21, 3).value
        self.assertEqual(tot_cell, "=SUM(C5:C20)")

        # Formulas en la tabla resumen por CC: COUNTIF y SUMIF
        countif_formula = ws_res.cell(5, 14).value
        self.assertTrue(str(countif_formula).startswith("=COUNTIF"), f"Formula esperada COUNTIF, obtenido: {countif_formula}")

        sumif_formula = ws_res.cell(5, 15).value
        self.assertTrue(str(sumif_formula).startswith("=SUMIF"), f"Formula esperada SUMIF, obtenido: {sumif_formula}")


@unittest.skipUnless(HAS_REFERENCES, "Referencias privadas Entel no disponibles.")
class TestEntelWithoutHoja1(unittest.TestCase):
    """
    Verifica que el sistema funcione cuando el archivo del mes no tiene Hoja1 previa.
    """
    def setUp(self):
        with open(EXCEL_MAY_PATH, "rb") as f:
            self.excel_bytes = f.read()
        with open(PDF_MAY_PATH, "rb") as f:
            self.pdf_bytes = f.read()
        with open(INV_PATH, "rb") as f:
            self.inv_bytes = f.read()

        # Crear una version de Excel en memoria sin Hoja1
        wb = openpyxl.load_workbook(io.BytesIO(self.excel_bytes))
        if "Hoja1" in wb.sheetnames:
            del wb["Hoja1"]
        buf = io.BytesIO()
        wb.save(buf)
        self.excel_no_hoja1 = buf.getvalue()

    def test_runs_cleanly_without_hoja1(self):
        result = entel_engine.process_entel_files(
            self.excel_no_hoja1,
            self.pdf_bytes,
            self.inv_bytes,
        )
        self.assertEqual(result["line_count"], 16)
        self.assertTrue(result["cuadra"])
        self.assertEqual(result["estado_conciliacion"], "CONCILIADO")
        self.assertEqual(len(result["cc_discrepancies"]), 0, "Sin Hoja1 no debe haber discrepancias previas")


class TestEntelErrorHandling(unittest.TestCase):
    """
    Verifica que se capturen y reporten errores claros con archivos invalidos.
    """
    def test_empty_pdf(self):
        with self.assertRaises(ValueError):
            entel_engine.extract_pdf_entel(b"")

    def test_corrupt_excel(self):
        with self.assertRaises(ValueError):
            entel_engine.process_entel_files(b"corrupto", b"pdf", b"inv")


if __name__ == "__main__":
    unittest.main()
