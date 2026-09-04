# -*- coding: utf-8 -*-
"""El panel F&B con ventas en dd/mm/aaaa (dia > 12) no puede caerse.

  python3.12 tests/test_fb_fechas_es.py
  python3.12 tests/test_fb_fechas_es.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DD = os.path.join(BASE, 'datos-referencia')
FICHEROS = ('ventas_fb_diarias.xlsx', 'recetas.xlsx', 'inventario.xlsx', 'mermas.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        import provisiones
        provisiones._fecha = lambda v: pd.to_datetime(str(v), format='%m/%d/%Y').date()   # vuelve al formato americano
    tmp = tempfile.mkdtemp(prefix='fb_')
    for f in FICHEROS:
        if os.path.exists(os.path.join(DD, f)):
            shutil.copy(os.path.join(DD, f), os.path.join(tmp, f))
    try:
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        rec = os.path.join(tmp, 'rec.xlsx')
        pd.DataFrame([{'receta': 'Tortilla', 'ingrediente': 'Huevos', 'cantidad': 3, 'unidad': 'ud', 'coste': 0.75, 'PVP': 7.5, 'categoria': 'Entrantes'}]).to_excel(rec, index=False)
        with open(rec, 'rb') as fh:
            r = cl.post('/fb/api/upload_recetas', data={'file': (fh, 'rec.xlsx')}, content_type='multipart/form-data', headers=H)
        ok(r.status_code == 200 and (r.get_json() or {}).get('ok'), 'recetario subido')
        ven = os.path.join(tmp, 'ventas_pos_test.xlsx')
        pd.DataFrame([{'fecha': '13/08/2026', 'plato': 'Tortilla', 'categoria': 'Entrantes', 'unidades': 10, 'precio_unitario': 7.5, 'total': 75.0},
                      {'fecha': '30/08/2026', 'plato': 'Tortilla', 'categoria': 'Entrantes', 'unidades': 4, 'precio_unitario': 7.5, 'total': 30.0}]).to_excel(ven, index=False)
        with open(ven, 'rb') as fh:
            cl.post('/api/upload_facturas', data={'files': [(fh, 'ventas_pos_test.xlsx')]}, content_type='multipart/form-data')
        txt = cl.get('/api/procesar_batch_stream?archivos=' + urllib.parse.quote(json.dumps(['ventas_pos_test.xlsx']))).get_data(as_text=True)
        ok('✓ F&B' in txt, 'ventas con dia 13 y 30 integradas por la capa 1')
        d = cl.get('/fb/api/resultados').get_json() or {}
        ok(d.get('ok') is True, f"/fb/api/resultados responde ok (error: {str(d.get('error'))[:80]})")
        vd = (d.get('ventas_diarias') or {}).get('fechas') or []
        ok('2026-08-13' in vd and '2026-08-30' in vd, f'las fechas se leen dia primero: {vd}')
    finally:
        for f in FICHEROS:
            dst = os.path.join(DD, f)
            if os.path.exists(os.path.join(tmp, f)):
                shutil.copy(os.path.join(tmp, f), dst)
            elif os.path.exists(dst):
                os.remove(dst)
        shutil.rmtree(tmp, ignore_errors=True)
        for f in ('ventas_pos_test.xlsx',):
            for dd in ('uploads', 'facturas-entrada'):
                pth = os.path.join(BASE, dd, f)
                if os.path.exists(pth):
                    os.remove(pth)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
