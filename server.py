# -*- coding: utf-8 -*-

"""
SISTEMA SMAP 2026 - SERVIDOR WEB

Compatible con:
- Render
- Windows / ejecución local
- GitHub
- Actualización mediante API /api/sync

El servidor:
1. Sirve index.html y los archivos del tablero.
2. Lee data.json como fuente de datos web.
3. Expone los datos mediante /api/data.
4. Permite sincronización mediante POST /api/sync.
5. Utiliza automáticamente el puerto PORT proporcionado por Render.
"""

from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime
import json
import os
import threading


# ============================================================
# CONFIGURACIÓN
# ============================================================

ROOT = Path(__file__).resolve().parent

DATA_FILE = ROOT / "data.json"
DEMO_FILE = ROOT / "demo_data.json"
CONFIG_FILE = ROOT / "config.json"

# Render proporciona PORT automáticamente.
# Para ejecución local utilizamos 8876.
PORT = int(os.environ.get("PORT", "8876"))

# Render necesita escuchar en todas las interfaces.
HOST = "0.0.0.0"

# Token para sincronización.
SYNC_TOKEN = os.environ.get(
    "SMAP_SYNC_TOKEN",
    "CAMBIAR_TOKEN_SMAPPRO"
)


# ============================================================
# MEMORIA
# ============================================================

LOCK = threading.Lock()

CACHE = {
    "data": [],
    "loaded_at": None,
    "source": None,
    "error": None
}


# ============================================================
# UTILIDADES
# ============================================================

def now():
    return datetime.now().isoformat(timespec="seconds")


def load_json_file(path):
    """
    Lee un JSON y devuelve una lista.
    """

    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, dict):
        raw = raw.get("data", [])

    if not isinstance(raw, list):
        return []

    return raw


def load_data_file():
    """
    Carga data.json.
    Si no existe, utiliza demo_data.json.
    """

    try:

        if DATA_FILE.exists():

            path = DATA_FILE
            source = "data.json"

        elif DEMO_FILE.exists():

            path = DEMO_FILE
            source = "demo_data.json"

        else:

            raise FileNotFoundError(
                "No existe data.json ni demo_data.json"
            )

        data = load_json_file(path)

        with LOCK:

            CACHE["data"] = data
            CACHE["loaded_at"] = now()
            CACHE["source"] = source
            CACHE["error"] = None

        print(
            f"[SMAP] Datos cargados: {len(data)} registros "
            f"desde {source}",
            flush=True
        )

        return data

    except Exception as e:

        with LOCK:
            CACHE["error"] = str(e)

        print(
            f"[SMAP] ERROR cargando datos: {e}",
            flush=True
        )

        return []


def save_data(data):
    """
    Guarda la información sincronizada en data.json.
    """

    temporary = ROOT / "data.tmp.json"

    payload = {
        "data": data,
        "updated_at": now()
    }

    with open(
        temporary,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            payload,
            f,
            ensure_ascii=False,
            separators=(",", ":")
        )

    temporary.replace(DATA_FILE)

    with LOCK:

        CACHE["data"] = data
        CACHE["loaded_at"] = payload["updated_at"]
        CACHE["source"] = "data.json"
        CACHE["error"] = None

    print(
        f"[SMAP] Sincronización completada: "
        f"{len(data)} registros",
        flush=True
    )


# ============================================================
# HANDLER WEB
# ============================================================

class SMAPHandler(SimpleHTTPRequestHandler):

    # --------------------------------------------------------
    # JSON
    # --------------------------------------------------------

    def send_json(self, obj, status=200):

        raw = json.dumps(
            obj,
            ensure_ascii=False,
            separators=(",", ":")
        ).encode("utf-8")

        self.send_response(status)

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8"
        )

        self.send_header(
            "Content-Length",
            str(len(raw))
        )

        self.send_header(
            "Cache-Control",
            "no-store, no-cache, must-revalidate"
        )

        self.end_headers()

        self.wfile.write(raw)

    # --------------------------------------------------------
    # GET
    # --------------------------------------------------------

    def do_GET(self):

        path = urlparse(self.path).path

        # ----------------------------------------------------
        # DATOS
        # ----------------------------------------------------

        if path == "/api/data":

            with LOCK:

                response = {
                    "data": CACHE["data"],
                    "meta": {
                        "loaded_at": CACHE["loaded_at"],
                        "source": CACHE["source"],
                        "rows": len(CACHE["data"]),
                        "error": CACHE["error"]
                    }
                }

            self.send_json(response)

            return

        # ----------------------------------------------------
        # METADATA
        # ----------------------------------------------------

        if path == "/api/meta":

            with LOCK:

                response = {
                    "loaded_at": CACHE["loaded_at"],
                    "source": CACHE["source"],
                    "rows": len(CACHE["data"]),
                    "error": CACHE["error"]
                }

            self.send_json(response)

            return

        # ----------------------------------------------------
        # HEALTH CHECK
        # ----------------------------------------------------

        if path == "/api/health":

            with LOCK:

                response = {
                    "ok": True,
                    "status": "online",
                    "rows": len(CACHE["data"]),
                    "time": now()
                }

            self.send_json(response)

            return

        # ----------------------------------------------------
        # REFRESH
        # ----------------------------------------------------

        if path == "/api/refresh":

            data = load_data_file()

            self.send_json({
                "ok": True,
                "rows": len(data),
                "loaded_at": CACHE["loaded_at"]
            })

            return

        # ----------------------------------------------------
        # ARCHIVOS WEB
        # ----------------------------------------------------

        return super().do_GET()

    # --------------------------------------------------------
    # POST
    # --------------------------------------------------------

    def do_POST(self):

        path = urlparse(self.path).path

        # ----------------------------------------------------
        # SINCRONIZACIÓN
        # ----------------------------------------------------

        if path != "/api/sync":

            self.send_json(
                {
                    "ok": False,
                    "error": "Ruta no encontrada"
                },
                404
            )

            return

        # ----------------------------------------------------
        # VALIDAR TOKEN
        # ----------------------------------------------------

        token = self.headers.get(
            "X-SMAP-TOKEN",
            ""
        )

        if not SYNC_TOKEN or token != SYNC_TOKEN:

            self.send_json(
                {
                    "ok": False,
                    "error": "No autorizado"
                },
                401
            )

            return

        # ----------------------------------------------------
        # LEER INFORMACIÓN
        # ----------------------------------------------------

        try:

            content_length = int(
                self.headers.get(
                    "Content-Length",
                    "0"
                )
            )

            body = self.rfile.read(
                content_length
            )

            payload = json.loads(
                body.decode("utf-8")
            )

            if isinstance(payload, dict):

                data = payload.get(
                    "data",
                    []
                )

            else:

                data = payload

            if not isinstance(data, list):

                raise ValueError(
                    "El campo data debe ser una lista."
                )

            # Guardar datos
            save_data(data)

            self.send_json({
                "ok": True,
                "rows": len(data),
                "updated_at": CACHE["loaded_at"]
            })

        except Exception as e:

            print(
                f"[SMAP] ERROR en sincronización: {e}",
                flush=True
            )

            self.send_json(
                {
                    "ok": False,
                    "error": str(e)
                },
                400
            )


# ============================================================
# SERVIDOR
# ============================================================

def start_server():

    # Servir archivos desde la carpeta del proyecto.
    os.chdir(ROOT)

    # Cargar información inicial.
    load_data_file()

    print("")
    print("==============================================", flush=True)
    print("       SISTEMA SMAP 2026 - WEB", flush=True)
    print("==============================================", flush=True)
    print(
        f"HOST: {HOST}",
        flush=True
    )
    print(
        f"PORT: {PORT}",
        flush=True
    )
    print(
        f"REGISTROS: {len(CACHE['data'])}",
        flush=True
    )
    print(
        f"CARPETA: {ROOT}",
        flush=True
    )
    print("==============================================", flush=True)
    print(
        "SERVIDOR INICIANDO...",
        flush=True
    )

    # Crear servidor.
    server = ThreadingHTTPServer(
        (HOST, PORT),
        SMAPHandler
    )

    # Permitir cierre limpio de los hilos.
    server.daemon_threads = True

    print(
        f"SISTEMA SMAP ESCUCHANDO EN "
        f"{HOST}:{PORT}",
        flush=True
    )

    print(
        "SERVIDOR ACTIVO.",
        flush=True
    )

    # Mantener proceso vivo.
    try:

        server.serve_forever(
            poll_interval=0.5
        )

    except KeyboardInterrupt:

        print(
            "Servidor detenido manualmente.",
            flush=True
        )

    finally:

        server.server_close()


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":

    start_server()
