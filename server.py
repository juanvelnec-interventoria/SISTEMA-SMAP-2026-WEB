# -*- coding: utf-8 -*-
"""SISTEMA SMAP 2026 - servidor web/local.

- Sirve el tablero index.html.
- Usa data.json como fuente web inicial.
- Permite recibir una actualización desde el PC administrador mediante POST /api/sync.
- Si existe un Excel local y openpyxl está instalado, también puede leerlo en local.
"""
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime, date
import json, math, os, threading

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

ROOT = Path(__file__).resolve().parent
CFG = ROOT / 'config.json'
DATA_FILE = ROOT / 'data.json'
DEMO_FILE = ROOT / 'demo_data.json'

try:
    config = json.loads(CFG.read_text(encoding='utf-8'))
except Exception:
    config = {}

CONFIG_EXCEL = Path(config.get('excel_path', r'C:\Users\USER\Desktop\INTERVENTORIA VELNEC\SMAP\SMAP_VELNEC_2026_JCA.xlsx'))
SHEET = config.get('sheet_name', 'Tabla SMAP-CIV')
SYNC_TOKEN = os.environ.get('SMAP_SYNC_TOKEN', config.get('sync_token', 'CAMBIAR_TOKEN_SMAPPRO'))
PORT = int(os.environ.get('PORT', '10000'))
HOST = os.environ.get('HOST', '0.0.0.0')

LOCK = threading.Lock()
CACHE = {'data': [], 'loaded_at': None, 'source': None, 'error': None}


def norm(s):
    return ' '.join(str(s or '').replace('\n', ' ').split()).strip().lower()


def blank(v):
    return v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() == '' or str(v).strip().lower() in {'nan', '#n/a', 'none'}


def serial(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def load_json_file(path):
    raw = json.loads(path.read_text(encoding='utf-8'))
    if isinstance(raw, dict):
        raw = raw.get('data', [])
    return raw if isinstance(raw, list) else []


def load_data_file():
    path = DATA_FILE if DATA_FILE.exists() else DEMO_FILE
    try:
        data = load_json_file(path)
        with LOCK:
            CACHE.update(data=data, loaded_at=datetime.now().isoformat(timespec='seconds'), source=path.name, error=None)
        return data
    except Exception as e:
        with LOCK:
            CACHE['error'] = str(e)
        return []


def resolve_excel_path():
    if CONFIG_EXCEL.exists():
        return CONFIG_EXCEL
    home = Path.home()
    roots = [
        home / 'Desktop' / 'INTERVENTORIA VELNEC' / 'SMAP',
        home / 'OneDrive' / 'Desktop' / 'INTERVENTORIA VELNEC' / 'SMAP',
        Path(r'C:\Users\USER\Desktop\INTERVENTORIA VELNEC\SMAP'),
        Path(r'C:\Users\USER\OneDrive\Desktop\INTERVENTORIA VELNEC\SMAP'),
    ]
    candidates = []
    for root in roots:
        try:
            if root.exists():
                candidates += list(root.glob('SMAP_VELNEC_2026_JCA*.xlsx'))
        except Exception:
            pass
    return max(candidates, key=lambda x: x.stat().st_mtime) if candidates else CONFIG_EXCEL


def read_excel(path=None):
    if load_workbook is None:
        raise RuntimeError('openpyxl no está instalado.')
    excel = Path(path) if path else resolve_excel_path()
    if not excel.exists():
        raise FileNotFoundError(f'No se encontró el Excel maestro: {excel}')
    wb = load_workbook(excel, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        raise KeyError(f'No existe la hoja {SHEET!r}. Hojas: {wb.sheetnames}')
    ws = wb[SHEET]
    it = ws.iter_rows(min_row=2, values_only=True)
    headers = list(next(it))
    H = [norm(x) for x in headers]

    def idx(*names):
        ns = [norm(x) for x in names]
        for n in ns:
            for i, h in enumerate(H):
                if h == n:
                    return i
        return -1

    I = {
        'consecutivo': idx('Consecutivo'), 'mes': idx('Mes'), 'anio': idx('Año'), 'solicitante': idx('Solicitante'),
        'direccion': idx('Dirección'), 'lat': idx('Coordenadas Latitud'), 'lon': idx('Coordenadas Longitud'),
        'barrio': idx('Barrio/Vereda'), 'comuna': idx('Comuna'), 'nombreComuna': idx('Nombre comuna'),
        'falla': idx('N° Falla'), 'tipoRecorrido': idx('TIPO RECORRIDO'), 'oficio': idx('Oficio de envio'),
        'fechaEnvio': idx('Fecha de envio CIV'), 'estado': idx('Estado de SMAP'), 'noSS': idx('No. de SS'),
        'tipoSS': idx('Tipo de SS'), 'fechaSS': idx('Fecha recibo del numero de SS'), 'diasSS': idx('Días Trascurridos entrega numero de SS'),
        'cumpleSS': idx('Cumple ANS entrega de numero de SS'), 'nivel': idx('Tipo de Nivel'),
        'fechaEntrega': idx('Fecha entrega de Respuesta'), 'diasRespuesta': idx('Días Trascurridos entrega de Respuesta'),
        'vencimiento': idx('VENCIMIENTO ENTREGA RESPUESTA'), 'cumpleRespuesta': idx('Cumple ANS entrega de Respuesta'),
        'fechaRespuesta': idx('Fecha de Respuesta'), 'radicadoResp': idx('Radicado Respuesta EPM2'), 'observaciones': idx('Observaciones'),
        'aval': idx('Radicado AVAL VELNEC'), 'fechaAval': idx('Fecha de AVAL'), 'aprobDistrito': idx('RAD.APROBACION DISTRITO'),
        'fechaAprob': idx('Fecha de APROB DISTRITO'), 'ssRevision': idx('FECHA RESPUEST SS REVISAR_EJECUCION'),
        'diasRevision': idx('Días Trascurridos entrega de SS REVISAR'), 'vencRevision': idx('VENCIMIENTO'),
        'cumpleRevision': idx('Cumple ANS entrega de SS REVISAR'), 'radEjec': idx('Radicado EJECUCION VELNEC'),
        'fechaEjec': idx('Fecha de RADI EJECUCION'), 'iniciales': idx('Iniciales'), 'interventor': idx('INTERVENTOR'),
        'tecnico': idx('Tecnicos'), 'fallaSmap': idx('FALLAS smap')
    }
    exact_interventor = next((i for i, h in enumerate(H) if h == 'interventor'), -1)
    if exact_interventor >= 0:
        I['interventor'] = exact_interventor
    elif len(headers) > 5:
        I['interventor'] = 5

    data = []
    for r in it:
        if not r or I['consecutivo'] < 0:
            continue
        con = r[I['consecutivo']] if I['consecutivo'] < len(r) else None
        if blank(con):
            continue
        d = {k: (serial(r[j]) if j >= 0 and j < len(r) else None) for k, j in I.items()}
        d['pendienteActualizacion'] = not bool(d.get('anio') or d.get('mes') or d.get('direccion'))
        data.append(d)
    wb.close()
    return data


def save_data(data):
    tmp = DATA_FILE.with_suffix('.tmp')
    payload = {'data': data, 'updated_at': datetime.now().isoformat(timespec='seconds')}
    tmp.write_text(json.dumps(payload, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    tmp.replace(DATA_FILE)
    with LOCK:
        CACHE.update(data=data, loaded_at=payload['updated_at'], source='data.json', error=None)


class Handler(SimpleHTTPRequestHandler):
    def _send_json(self, obj, status=200):
        raw = json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(raw)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        p = urlparse(self.path).path
        if p == '/api/data':
            with LOCK:
                data = CACHE['data']
                meta = {'loaded_at': CACHE['loaded_at'], 'source': CACHE['source'], 'rows': len(data), 'error': CACHE['error']}
            self._send_json({'data': data, 'meta': meta})
            return
        if p == '/api/meta':
            with LOCK:
                self._send_json({'loaded_at': CACHE['loaded_at'], 'source': CACHE['source'], 'rows': len(CACHE['data']), 'error': CACHE['error']})
            return
        if p == '/api/health':
            self._send_json({'ok': True, 'rows': len(CACHE['data'])})
            return
        if p == '/api/refresh':
            # En web no se lee C:\ del servidor; se recarga data.json.
            load_data_file()
            self._send_json({'ok': True, 'rows': len(CACHE['data']), 'loaded_at': CACHE['loaded_at']})
            return
        return super().do_GET()

    def do_POST(self):
        p = urlparse(self.path).path
        if p != '/api/sync':
            self._send_json({'error': 'Ruta no encontrada'}, 404)
            return
        token = self.headers.get('X-SMAP-TOKEN', '')
        if not SYNC_TOKEN or token != SYNC_TOKEN:
            self._send_json({'error': 'No autorizado'}, 401)
            return
        try:
            n = int(self.headers.get('Content-Length', '0'))
            body = self.rfile.read(n)
            payload = json.loads(body.decode('utf-8'))
            data = payload.get('data') if isinstance(payload, dict) else payload
            if not isinstance(data, list):
                raise ValueError('El payload debe contener una lista data.')
            save_data(data)
            self._send_json({'ok': True, 'rows': len(data), 'updated_at': CACHE['loaded_at']})
        except Exception as e:
            self._send_json({'error': str(e)}, 400)


if __name__ == '__main__':
    os.chdir(ROOT)

    load_data_file()

    print(f'SISTEMA SMAP 2026 WEB')
    print(f'Host: {HOST}')
    print(f'Port: {PORT}')
    print(f'Registros iniciales: {len(CACHE["data"])}', flush=True)

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    server.daemon_threads = True

    print(f'Servidor escuchando en http://{HOST}:{PORT}', flush=True)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
