import http.server
import socketserver
import os
import json
import uuid
import datetime
import urllib.parse
import email.parser
import email.policy
import traceback

import movistar_engine

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(DIRECTORY, "data")

# Cache en memoria para descargas temporales
EXCEL_CACHE = {}


def parse_multipart(body_bytes, content_type_header):
    """Parsea multipart/form-data usando email.parser."""
    headers = f"Content-Type: {content_type_header}\r\nMIME-Version: 1.0\r\n\r\n".encode("latin1")
    msg     = email.parser.BytesParser(policy=email.policy.default).parsebytes(headers + body_bytes)
    fields  = {}
    files   = {}
    for part in msg.iter_parts():
        name     = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload  = part.get_payload(decode=True)
        if filename is not None:
            files[name] = {"filename": filename, "content": payload}
        elif name:
            fields[name] = payload.decode("utf-8", errors="ignore") if payload else ""
    return fields, files


class MovistarHTTPHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        # Solo imprimir requests de API, no de archivos estaticos
        if "/api/" in args[0] if args else False:
            super().log_message(format, *args)

    # ------------------------------------------------------------------
    # GET
    # ------------------------------------------------------------------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path   = parsed.path
        query  = urllib.parse.parse_qs(parsed.query)

        if path == "/api/movistar/download-excel":
            token = query.get("token", [""])[0]
            if token in EXCEL_CACHE:
                cached     = EXCEL_CACHE[token]
                excel_bytes = cached["bytes"]
                filename   = cached["filename"]
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                self.send_header(
                    "Content-Disposition",
                    f'attachment; filename="{filename}"'
                )
                self.send_header("Content-Length", str(len(excel_bytes)))
                self.end_headers()
                self.wfile.write(excel_bytes)
            else:
                self.send_json(404, {"error": "El archivo ya no esta disponible. Procese nuevamente."})
            return

        elif path == "/api/movistar/download-observations":
            token = query.get("token", [""])[0]
            if token in EXCEL_CACHE:
                cached  = EXCEL_CACHE[token]
                sin_cc  = cached.get("sin_cc_lines", [])
                lines   = ["\ufeffNUMERO,PLAN,MONTO F,OBSERVACION"]
                for item in sin_cc:
                    p  = item.get("phone", "")
                    pl = item.get("plan", "").replace(",", " ")
                    m  = f"{item.get('monto_f', 0):.2f}"
                    lines.append(f'"{p}","{pl}",{m},"Sin Centro de Costo en inventario"')
                csv_bytes = "\r\n".join(lines).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header(
                    "Content-Disposition",
                    'attachment; filename="Observaciones_Sin_CC_Movistar.csv"'
                )
                self.send_header("Content-Length", str(len(csv_bytes)))
                self.end_headers()
                self.wfile.write(csv_bytes)
            else:
                self.send_json(404, {"error": "El reporte ya no esta disponible."})
            return

        super().do_GET()

    # ------------------------------------------------------------------
    # POST
    # ------------------------------------------------------------------
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path   = parsed.path

        content_length = int(self.headers.get("Content-Length", 0))
        content_type   = self.headers.get("Content-Type", "")

        if not content_type.startswith("multipart/form-data"):
            self.send_json(400, {"error": "Se requiere multipart/form-data."})
            return

        body_bytes      = self.rfile.read(content_length)
        fields, files   = parse_multipart(body_bytes, content_type)

        try:
            if path == "/api/movistar/process":
                # --- Validar password ---
                password = fields.get("password", "").strip()
                if not password:
                    self.send_json(400, {"error": "Debe ingresar la contrasena de los archivos del mes."})
                    return

                # --- Validar 5 archivos obligatorios ---
                required = [
                    ("fact_file",      "Archivo Fact (Fact_*.xlsx)"),
                    ("traf_file",      "Archivo Traf (Traf_*.xlsx)"),
                    ("cargos_file",    "Archivo Cargos (Cargos_*.xlsx)"),
                    ("inventory_file", "Inventario Celulares General"),
                    ("pdf_file",       "Recibo PDF mensual de Movistar"),
                ]
                for key, label in required:
                    if key not in files or not files[key]["content"]:
                        self.send_json(400, {"error": f"Falta cargar: {label}."})
                        return

                result = movistar_engine.process_movistar_files(
                    fact_bytes      = files["fact_file"]["content"],
                    traf_bytes      = files["traf_file"]["content"],
                    cargos_bytes    = files["cargos_file"]["content"],
                    pdf_bytes       = files["pdf_file"]["content"],
                    inventory_bytes = files["inventory_file"]["content"],
                    password        = password,
                )

                # Guardar Excel en cache
                token        = str(uuid.uuid4())
                periodo_slug = re.sub(r'[^a-zA-Z0-9_]', '_', result["periodo"])
                filename     = f"Movistar_Conciliado_{periodo_slug}.xlsx"

                EXCEL_CACHE[token] = {
                    "bytes":        result["excel_bytes"],
                    "filename":     filename,
                    "sin_cc_lines": result["sin_cc_lines"],
                    "created":      datetime.datetime.now(),
                }

                del result["excel_bytes"]
                result["download_token"]    = token
                result["download_filename"] = filename
                result["success"]           = True

                self.send_json(200, result)

            else:
                self.send_json(404, {"error": "Ruta no encontrada."})

        except ValueError as ve:
            print(f"[Aviso Movistar] {ve}")
            self.send_json(400, {"error": str(ve)})
        except Exception:
            print("[Error Servidor Movistar]")
            traceback.print_exc()
            self.send_json(500, {
                "error": "Error interno al procesar los archivos. Revise la consola del servidor."
            })


import re  # necesario para periodo_slug en do_POST

if __name__ == "__main__":
    os.chdir(DIRECTORY)
    os.makedirs(DATA_DIR, exist_ok=True)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), MovistarHTTPHandler) as httpd:
        print("=" * 55)
        print(f" Servidor Movistar iniciado en: http://localhost:{PORT}")
        print(f" Directorio: {DIRECTORY}")
        print(" Presione Ctrl+C para detener.")
        print("=" * 55)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")
