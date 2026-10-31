"""Datos sintéticos de prueba; no contiene documentos ni credenciales corporativas."""
import io
import openpyxl


def workbook_bytes(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def simple_pdf(text):
    objects = []
    commands = ["BT /F1 12 Tf 40 750 Td"]
    for index, line in enumerate(text.splitlines()):
        if index:
            commands.append("0 -18 Td")
        escaped = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        commands.append("(" + escaped + ") Tj")
    commands.append("ET")
    stream = "\n".join(commands).encode("ascii")
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


def movistar_fixtures():
    fact = openpyxl.Workbook()
    for r, phone, amount in [(12, "951234567", 10), (13, "952345678", 30)]:
        fact.active.cell(r, 2, phone)
        fact.active.cell(r, 20, amount)
        fact.active.cell(r, 29, 0)
    traf = openpyxl.Workbook()
    for r, phone in [(1, "951234567"), (2, "952345678")]:
        traf.active.cell(r, 2, phone)
        traf.active.cell(r, 5, "Plan de prueba")
    cargos = openpyxl.Workbook()
    cargos.active.append(["CUOTA EQUIPO TELEF: 951234567", 12])
    inventory = openpyxl.Workbook()
    inventory.active.title = "PRODAC"
    inventory.active.append(["Numero", "CC Dpto", "Usuario"])
    inventory.active.append(["951234567", "CC-A", "Usuario de prueba A"])
    inventory.active.append(["952345678", "CC-B", "Usuario de prueba B"])
    pdf = simple_pdf("Recibo Septiembre 2026\nSubtotal S/ 35.00\nIGV (18%) S/ 6.30\nTotal Afecto S/ 41.30\nCargos Adicionales Inafectos S/ 12.00\nRedondeo S/ 0.00\nTotal Facturado S/ 53.30\nAnexo 1: 951234567\nDescuento 50% al plan 12 meses S/ - 5.00")
    return {key: data for key, data in [
        ("fact_file", workbook_bytes(fact)), ("traf_file", workbook_bytes(traf)),
        ("cargos_file", workbook_bytes(cargos)), ("inventory_file", workbook_bytes(inventory)),
        ("pdf_file", pdf)]}


def multipart(files, fields=None):
    boundary = "AppCelularesTestBoundary"
    parts = []
    for key, value in (fields or {}).items():
        parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"\r\n\r\n{value}\r\n'.encode())
    for key, payload in files.items():
        extension = "pdf" if key == "pdf_file" else "xlsx"
        parts.extend([f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"; filename="{key}.{extension}"\r\nContent-Type: application/octet-stream\r\n\r\n'.encode(), payload, b"\r\n"])
    parts.append(f"--{boundary}--\r\n".encode())
    return b"".join(parts), "multipart/form-data; boundary=" + boundary
