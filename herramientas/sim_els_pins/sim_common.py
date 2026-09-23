# -*- coding: utf-8 -*-
"""Simulacion del plan de pruebas de Els Pins contra el codigo actual (b80, 12 sep 2026).

Sin ANTHROPIC_API_KEY el clasificador cae a regex y no lee ni el numero de factura,
asi que `fake_ia.py` sustituye la llamada a Claude por lo que devolveria leyendo
bien cada PDF/foto de la carpeta. Todo lo demas (lote, cierre del lote, cruces,
aprobaciones, Oracle, F&B, banco, cierre, fiscal) es el codigo real via test_client.

  ELS_PINS_DIR=/ruta/a/_PRUEBAS_Els_Pins python3.12 herramientas/sim_els_pins/sim_all.py 2 3 4 5 6 7
  ELS_PINS_DIR=... python3.12 herramientas/sim_els_pins/sim_f1.py

Hace copia de las carpetas de datos, deja el checkout como en un Render recien
desplegado, crea el hotel y al final restaura todo (`terminar`). NUNCA a la vez que
otro test que use datos. `edit_plan.py` regenera las 7 `faseN/00_PASOS…` a partir
del `00_PLAN_DE_PRUEBAS.html` (cabecera + seccion + pie).
"""
import json, os, shutil, sys, tempfile, urllib.parse, io
BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, BASE); sys.path.insert(0, os.path.dirname(os.path.abspath(__file__))); os.chdir(BASE)
FIX = os.environ.get('ELS_PINS_DIR') or '/mnt/user-data/uploads/yve01/_PRUEBAS_Els_Pins'
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']
os.environ.pop('ANTHROPIC_API_KEY', None)
import pandas as pd

_copia = None


def empezar():
    """copia de seguridad + datos como en un Render recien desplegado (git limpio)."""
    global _copia
    _copia = tempfile.mkdtemp(prefix='sim_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(_copia, d))
    for d in DIRS:
        shutil.rmtree(d, ignore_errors=True)
    os.system('git checkout -q -- datos-referencia reportes && git clean -qfd datos-referencia facturas-procesadas reportes aprobaciones')
    assert os.path.exists('datos-referencia/inventario.xlsx')
    import dashboard as D
    D.app.config['TESTING'] = True
    cl = D.app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    r = cl.post('/admin/api/hoteles/crear', json={'nombre': 'Hotel Els Pins', 'ciudad': 'Sitges', 'grupo': 'Cadena Llevant', 'habitaciones': 62})
    if r.status_code != 200:
        r = cl.post('/panel-admin/api/hoteles/crear', json={'nombre': 'Hotel Els Pins', 'ciudad': 'Sitges', 'grupo': 'Cadena Llevant', 'habitaciones': 62})
    print('crear hotel:', r.status_code, r.get_json())
    hid = r.get_json()['id']
    global TOK
    TOK = (cl.get('/api/csrf_token').get_json() or {}).get('token')
    print('hotel activo:', post(cl, '/api/hotel_activo', {'hotel': hid}).get_json())
    import fake_ia; fake_ia.instalar()
    # Vision: cliente anthropic falso para la foto
    import anthropic, json as _j, types
    class _Msg:  # noqa
        def __init__(self, txt): self.content = [types.SimpleNamespace(text=txt)]
    class _Fake:
        def __init__(self, *a, **k): self.messages = self
        def create(self, model=None, max_tokens=None, messages=None):
            txt = _j.dumps(fake_ia.IA['foto_factura_bugaderia_sitges.jpg'])
            return _Msg(txt)
    anthropic.Anthropic = _Fake
    return D, cl


TOK = None


def post(cl, url, js):
    return cl.post(url, json=js, headers={'X-CSRF-Token': TOK})


def foto(cl, ruta):
    with open(ruta, 'rb') as fh:
        r = cl.post('/api/scan_documento', data={'image': (fh, os.path.basename(ruta))}, content_type='multipart/form-data', headers={'X-CSRF-Token': TOK})
    print('   foto:', r.status_code, str(r.get_json())[:300])
    return r.get_json()


def terminar():
    for d in DIRS:
        shutil.rmtree(d, ignore_errors=True)
        if os.path.isdir(os.path.join(_copia, d)):
            shutil.copytree(os.path.join(_copia, d), d)
    shutil.rmtree(_copia, ignore_errors=True)
    os.system('git status --short | head')


def lote(cl, rutas, meses=None):
    nombres = []
    for ruta in rutas:
        n = os.path.basename(ruta); nombres.append(n)
        with open(ruta, 'rb') as fh:
            r = cl.post('/api/upload_facturas', data={'files': [(fh, n)]}, content_type='multipart/form-data')
            assert r.status_code == 200, (n, r.status_code, r.get_data(as_text=True)[:200])
    q = '/api/procesar_batch_stream?archivos=' + urllib.parse.quote(json.dumps(nombres))
    if meses:
        q += '&meses=' + urllib.parse.quote(json.dumps(meses))
    r = cl.get(q); txt = r.get_data(as_text=True); r.close()
    lineas = [l[6:] for l in txt.splitlines() if l.startswith('data: ')]
    for l in lineas:
        print('   |', l[:220])
    pend = [l.split(':', 1)[1] for l in lineas if l.startswith('CIERRE_PENDIENTE:')]
    if pend:
        lineas += cierre(cl, pend[0])
    return lineas


def cierre(cl, pasos):
    r = cl.get('/api/cerrar_pipeline_stream?pasos=' + urllib.parse.quote(pasos)); txt = r.get_data(as_text=True); r.close()
    lineas = [l[6:] for l in txt.splitlines() if l.startswith('data: ')]
    for l in lineas:
        print('   c|', l[:220])
    return lineas


def gj(cl, url):
    r = cl.get(url)
    try:
        return r.get_json()
    except Exception:
        return {'_status': r.status_code, '_txt': r.get_data(as_text=True)[:300]}


def pj(x, n=4000):
    s = json.dumps(x, ensure_ascii=False, indent=1, default=str)
    print(s[:n])
