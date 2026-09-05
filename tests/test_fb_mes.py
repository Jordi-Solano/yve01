# -*- coding: utf-8 -*-
"""F&B Cost con filtro de mes (Jordi, sep 2026): /fb/api/mermas?mes= y
/fb/api/resultados?mes= devuelven solo el mes; sin mes, todo (como antes).

  python3.12 tests/test_fb_mes.py
  python3.12 tests/test_fb_mes.py --sabotaje
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DATOS = os.path.join(BASE, 'datos-referencia')
MER = os.path.join(DATOS, 'mermas.xlsx')
VEN = os.path.join(DATOS, 'ventas_fb_diarias.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import tab_fb_dashboard as FB
    if SABOTAJE:
        FB._solo_mes = lambda df, mes, col='fecha': df

    tmp = tempfile.mkdtemp(prefix='fbm_'); copias = {}
    for f in (MER, VEN):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        import dashboard as D
        D._guardar_fb_del_hotel(pd.DataFrame([
            {"fecha": "03/08/2026", "ingrediente": "Merluza", "categoria": "Pescados", "cantidad_merma": 4, "unidad": "kg", "causa": "caducado", "coste_merma": 36.0},
            {"fecha": "02/09/2026", "ingrediente": "Pollo", "categoria": "Carnes", "cantidad_merma": 2, "unidad": "kg", "causa": "deterioro", "coste_merma": 9.0},
        ]), 'mermas.xlsx')
        D._guardar_fb_del_hotel(pd.DataFrame([
            {"fecha": "05/08/2026", "plato": "Pollo al ast", "categoria": "Carnes", "unidades": 10, "precio_unitario": 14.0, "total": 140.0},
            {"fecha": "02/09/2026", "plato": "Pollo al ast", "categoria": "Carnes", "unidades": 5, "precio_unitario": 14.0, "total": 70.0},
        ]), 'ventas_fb_diarias.xlsx')
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        todo = cl.get('/fb/api/mermas').get_json()
        ago = cl.get('/fb/api/mermas?mes=2026-08').get_json()
        sep = cl.get('/fb/api/mermas?mes=2026-09').get_json()
        oct_ = cl.get('/fb/api/mermas?mes=2026-10').get_json()
        ok(todo.get('total') == 45.0 and len(todo.get('mermas', [])) == 2, f"sin mes: todas ({todo.get('total')})")
        ok(ago.get('total') == 36.0 and len(ago.get('mermas', [])) == 1, f"agosto: 36,00 ({ago.get('total')})")
        ok(sep.get('total') == 9.0, f"septiembre: 9,00 ({sep.get('total')})")
        ok(oct_.get('ok') and oct_.get('total') == 0 and oct_.get('mermas') == [], "octubre: vacio, sin error")
        r_todo = cl.get('/fb/api/resultados').get_json()
        r_ago = cl.get('/fb/api/resultados?mes=2026-08').get_json()
        r_oct = cl.get('/fb/api/resultados?mes=2026-10').get_json()
        ok(r_todo.get('ok') and r_todo['resumen']['total_ventas'] == 210.0, f"resultados sin mes: ventas {r_todo.get('resumen', {}).get('total_ventas')}")
        ok(r_ago.get('ok') and r_ago['resumen']['total_ventas'] == 140.0 and r_ago['resumen']['coste_mermas'] == 36.0 and r_ago.get('mes') == '2026-08',
           f"resultados de agosto: ventas {r_ago.get('resumen', {}).get('total_ventas')} · mermas {r_ago.get('resumen', {}).get('coste_mermas')}")
        ok(r_oct.get('ok') and r_oct.get('vacio'), "octubre sin ventas: respuesta 'vacio', no un panel con ceros")
        html = cl.get('/').get_data(as_text=True)
        ok('id="fb-mes"' in html and "_fbMesQS()" in html and "fb.sinVentasMes" in html, "selector de mes en F&B Cost y los dos loaders lo usan")
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_fbm.js', 'w', encoding='utf-8').write(b)
            rc = subprocess.run(['node', '--check', '/tmp/_fbm.js'], capture_output=True, text=True)
            if rc.returncode:
                ok(False, f"JS roto: {rc.stderr[:100]}"); break
    finally:
        for f in (MER, VEN):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True)
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
