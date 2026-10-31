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
import re
import threading
from pathlib import Path

import movistar_engine
import entel_engine

PORT = int(os.environ.get("PORT", 8080))
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(DIRECTORY, "data")

# Cache en memoria para descargas temporales
EXCEL_CACHE = {}
CACHE_LOCK = threading.RLock()
CACHE_TTL_SECONDS = 30 * 60
MAX_CACHE_BYTES = 64 * 1024 * 1024
MAX_UPLOAD_BYTES = 100 * 1024 * 1024


def prune_cache():
    """Limita la vida y el tamaño de las descargas temporales en RAM."""
    now = datetime.datetime.now()
    with CACHE_LOCK:
        for token, item in list(EXCEL_CACHE.items()):
            if (now - item["created"]).total_seconds() > CACHE_TTL_SECONDS:
                EXCEL_CACHE.pop(token, None)
        size = sum(len(item["bytes"]) for item in EXCEL_CACHE.values())
        for token in sorted(EXCEL_CACHE, key=lambda key: EXCEL_CACHE[key]["created"]):
            if size <= MAX_CACHE_BYTES:
                break
            size -= len(EXCEL_CACHE[token]["bytes"])
            EXCEL_CACHE.pop(token, None)



def get_cached_download(token):
    with CACHE_LOCK:
        return EXCEL_CACHE.get(token)


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
        self.send_header("Content-Length", "0")
        self.end_headers()

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def log_message(self, format, *args):
        # No registrar tokens de descarga, contraseñas ni contenido subido.
        path = urllib.parse.urlparse(self.path).path
        if path.startswith("/api/"):
            super().log_message("%s %s", self.command, path)

    def send_head(self):
        # SimpleHTTPRequestHandler expondría .git, .env, data y pruebas.
        # Sólo se sirven los recursos públicos de la aplicación.
        path = urllib.parse.unquote(urllib.parse.urlparse(self.path).path)
        public_pages = {"/", "/index.html", "/claro.html", "/movistar.html", "/entel.html"}
        allowed = path in public_pages or re.fullmatch(r"/(?:js|css)/[A-Za-z0-9_-]+\.(?:js|css)", path)
        target = Path(self.translate_path(self.path)).resolve()
        root = Path(DIRECTORY).resolve()
        if not allowed or not target.is_relative_to(root):
            self.send_error(404, "Recurso no disponible.")
            return None
        return super().send_head()

    def list_directory(self, path):
        self.send_error(404, "Recurso no disponible.")
        return None

    def do_HEAD(self):
        if urllib.parse.urlparse(self.path).path == "/health":
            self.do_GET()
        else:
            super().do_HEAD()

    # ------------------------------------------------------------------
    # GET
    # ------------------------------------------------------------------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path   = parsed.path
        query  = urllib.parse.parse_qs(parsed.query)

        if path == "/health":
            self.send_json(200, {
                "status": "ok",
                "commit": os.environ.get("RENDER_GIT_COMMIT", "local"),
                "modules": ["claro", "movistar", "entel"],
            })
            return
        prune_cache()

        if path == "/api/movistar/download-excel":
            token = query.get("token", [""])[0]
            cached = get_cached_download(token)
            if cached:
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
            cached = get_cached_download(token)
            if cached:
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

        elif path == "/api/entel/download-excel":
            token = query.get("token", [""])[0]
            cached = get_cached_download(token)
            if cached:
                excel_bytes = cached["bytes"]
                filename    = cached["filename"]
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

        elif path == "/api/entel/download-observations":
            token = query.get("token", [""])[0]
            cached = get_cached_download(token)
            if cached:
                obs = cached.get("obs_lines", [])
                lines = ["\ufeffNUMERO,PLAN,CARGO_FIJO,USUARIO,CC_DPTO,ESTADO,OBSERVACION"]
                for item in obs:
                    p = item.get("phone", "")
                    pl = str(item.get("plan", "")).replace(",", " ").replace('"', '')
                    m = f"{item.get('cargo_fijo', 0):.2f}"
                    u = str(item.get("usuario", "")).replace('"', '')
                    cc = str(item.get("cc_dpto", "")).replace('"', '')
                    st = str(item.get("estado", "")).replace('"', '')
                    ob = str(item.get("observacion", "")).replace('"', '')
                    lines.append(f'"{p}","{pl}",{m},"{u}","{cc}","{st}","{ob}"')
                csv_bytes = "\r\n".join(lines).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header(
                    "Content-Disposition",
                    'attachment; filename="Observaciones_Entel.csv"'
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

        try:
            content_length = int(self.headers.get("Content-Length", 0))
        except ValueError:
            content_length = -1
        if content_length <= 0 or content_length > MAX_UPLOAD_BYTES:
            self.close_connection = True
            self.send_json(413 if content_length > MAX_UPLOAD_BYTES else 400,
                           {"error": "Carga vacía, inválida o superior al límite de 100 MB."})
            return
        content_type = self.headers.get("Content-Type", "")
        prune_cache()

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

                with CACHE_LOCK:
                    EXCEL_CACHE[token] = {
                        "bytes":        result["excel_bytes"],
                        "filename":     filename,
                        "sin_cc_lines": result["sin_cc_lines"],
                        "created":      datetime.datetime.now(),
                    }

                prune_cache()
                del result["excel_bytes"]
                result["download_token"]    = token
                result["download_filename"] = filename
                result["success"]           = True

                self.send_json(200, result)

            elif path == "/api/entel/inspect":
                if "excel_file" not in files or not files["excel_file"]["content"]:
                    self.send_json(400, {"error": "Debe cargar el archivo Excel de facturacion de Entel."})
                    return
                inspection = entel_engine.inspect_entel_excel(files["excel_file"]["content"])
                self.send_json(200, inspection)

            elif path == "/api/entel/process":
                required = [
                    ("excel_file",     "Excel de facturacion Entel"),
                    ("inventory_file", "Inventario Celulares General"),
                    ("pdf_file",       "Recibo PDF mensual de Entel"),
                ]
                for key, label in required:
                    if key not in files or not files[key]["content"]:
                        self.send_json(400, {"error": f"Falta cargar: {label}."})
                        return

                custom_cfg = None
                if "custom_config" in fields and fields["custom_config"]:
                    try:
                        custom_cfg = json.loads(fields["custom_config"])
                    except Exception:
                        pass

                result = entel_engine.process_entel_files(
                    excel_bytes     = files["excel_file"]["content"],
                    pdf_bytes       = files["pdf_file"]["content"],
                    inventory_bytes = files["inventory_file"]["content"],
                    custom_config   = custom_cfg,
                )

                token = str(uuid.uuid4())
                periodo_clean = re.sub(r'[^a-zA-Z0-9_]', '_', result.get("periodo", "mes"))
                filename = f"Entel_Conciliado_{periodo_clean}.xlsx"

                with CACHE_LOCK:
                    EXCEL_CACHE[token] = {
                        "bytes":     result["excel_bytes"],
                        "filename":  filename,
                        "obs_lines": result["obs_lines"],
                        "created":   datetime.datetime.now(),
                    }

                prune_cache()
                del result["excel_bytes"]
                result["download_token"]    = token
                result["download_filename"] = filename
                result["success"]           = True

                self.send_json(200, result)

            else:
                self.send_json(404, {"error": "Ruta no encontrada."})

        except ValueError as ve:
            print("[Validación] Archivo o configuración no válidos.")
            self.send_json(400, {"error": str(ve)})
        except Exception:
            print("[Error Servidor Movistar]")
            traceback.print_exc()
            self.send_json(500, {
                "error": "Error interno al procesar los archivos. Revise la consola del servidor."
            })



class ThreadedServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    daemon_threads = True
    allow_reuse_address = True

AppCelularesHTTPHandler = MovistarHTTPHandler

if __name__ == "__main__":
    os.chdir(DIRECTORY)
    os.makedirs(DATA_DIR, exist_ok=True)
    with ThreadedServer(("0.0.0.0", PORT), MovistarHTTPHandler) as httpd:
        print("=" * 60)
        print(f" Servidor iniciado en:   http://localhost:{PORT}")
        print(f"   -> Módulo Entel:     http://localhost:{PORT}/entel.html")
        print(f"   -> Módulo Movistar:  http://localhost:{PORT}/movistar.html")
        print(f"   -> Módulo Claro:     http://localhost:{PORT}/claro.html")
        print(f" Directorio base:       {DIRECTORY}")
        print(" Presione Ctrl+C para detener.")
        print("=" * 60)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")
