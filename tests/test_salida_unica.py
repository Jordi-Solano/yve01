# -*- coding: utf-8 -*-
"""Pieza 10 (Jordi, sep 2026): UNA sola salida de descargas.
Todas las descargas (Excel AP, aging, cierre, banco, GL, paquete...) viven en
el menu ⚙️ con un selector de apartado que solo enseña los que el rol puede
ver. Fuera los botones de descarga sueltos de cada panel.

  python3.12 tests/test_salida_unica.py
  python3.12 tests/test_salida_unica.py --sabotaje
"""
import json
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        # vuelve un boton suelto y el selector deja de filtrar por rol
        html = html.replace('<span id="ar-bonos-resumen"', '<a href="/api/exportar/bonos" class="btn-ref">⬇️ Excel</a><span id="ar-bonos-resumen"')
        html = html.replace("return a.tab === 'general' ? _rolVeTodo() : _rolVeApartado(a.tab);", "return true;")

    # 1. ningun enlace de descarga suelto fuera del catalogo
    cat_ini = html.index('var _DESCARGAS = ['); cat_fin = html.index('function _mesDescarga')
    fuera = html[:cat_ini] + html[cat_fin:]
    sueltos = re.findall(r'href="(/api/(?:exportar|reportes|oracle/export|inventarios/hoja)[^"]*)"', fuera)
    ok(not sueltos, f"sin botones de descarga sueltos en los paneles: {sueltos[:5]}")
    ok('id="dl-apartado"' in html and 'id="dl-lista"' in html and 'menu.descargas' in html, "el menu ⚙️ tiene el selector de apartado y la lista")
    for viejo in ('id="prov-descarga"', 'id="cierre-excel"', 'id="paq-excel"', 'id="cbanco-excel"', 'id="inv-excel"', 'id="fis-excel"', 'id="inm-excel"', 'id="inv-hoja"'):
        if viejo in html:
            ok(False, f"queda el boton viejo {viejo}"); break
    else:
        ok(True, "los ocho botones con id del Cierre/AP han desaparecido")

    # 2. el catalogo, ejecutado con node: cada ruta existe y el rol filtra
    js_rol = html[html.index('var _USER_ROL ='):html.index('(function() {', html.index('var _USER_ROL ='))]
    js_cat = html[cat_ini:html.index('function _pintarDescargas')]
    from werkzeug.routing import MapAdapter as _MA  # noqa
    _adapter = app.url_map.bind('localhost')
    def _existe(u):
        try:
            _adapter.match(u.split('?')[0], method='GET'); return True
        except Exception:
            return False
    prog = (js_rol + js_cat + """
var out = {};
['admin','financial_controller','income_auditor','fb_manager','jefe_otras'].forEach(function(r){ _USER_ROL = r; out[r] = _apartadosDescarga().map(function(a){ return a.tab; }); });
var urls = []; _DESCARGAS.forEach(function(a){ a.items.forEach(function(i){ urls.push(i.u); }); });
console.log(JSON.stringify({roles: out, urls: urls}));
""")
    open('/tmp/_su.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_su.js'], capture_output=True, text=True)
    ok(rc.returncode == 0, f"el catalogo se ejecuta ({rc.stderr[:80]})")
    d = json.loads(rc.stdout or '{}') if rc.returncode == 0 else {'roles': {}, 'urls': []}
    faltan = [u for u in d['urls'] if not _existe(u)]
    ok(d['urls'] and not faltan, f"las {len(d['urls'])} rutas del catalogo existen en Flask (faltan: {faltan})")
    esperado = {'ar', 'ap', 'drr', 'banco', 'fb', 'ar_real', 'multi_hotel', 'cierre'}
    ok(set(d['urls']) >= {'/api/exportar/ap', '/api/exportar/aging_ap', '/api/exportar/cierre', '/api/exportar/banco', '/api/oracle/export_excel', '/api/exportar/cierre_paquete', '/api/exportar/asientos'},
       "estan las que pidio Jordi: Excel AP, aging, cierre, banco, GL, paquete, libro diario")
    r = d['roles']
    ok(set(r.get('admin', [])) == esperado | {'general'} and set(r.get('financial_controller', [])) == esperado | {'general'}, f"admin y FC ven todos los apartados: {r.get('admin')}")
    ok(r.get('income_auditor') == ['drr'], f"income_auditor solo DRR: {r.get('income_auditor')}")
    ok(r.get('fb_manager') == ['ap', 'fb'], f"fb_manager AP y F&B: {r.get('fb_manager')}")
    ok(r.get('jefe_otras') == ['ap'], f"jefe_otras solo AP: {r.get('jefe_otras')}")
    ok("_TABS_ROL[rol] || []" in html and "var VISIBLE = {" not in html, "las pestañas usan el MISMO mapa de roles que las descargas")
    ok("_pintarDescargas();" in html.split('function toggleMenu')[1][:400], "el menu se repinta al abrirse (el mes es el de ese momento)")
    ok("mes: 'cierre-mes'" in html and "_mesDescarga(it.mes)" in html, "las descargas del cierre llevan el mes del selector")

    # 3. las rutas responden con la sesion de admin (no un 500)
    malos = []
    for u in ('/api/exportar/ap', '/api/exportar/aging_ap', '/api/exportar/banco', '/api/exportar/cierre?mes=2026-08', '/api/exportar/cierre_paquete?mes=2026-08', '/api/oracle/export_excel'):
        try:
            code = cl.get(u).status_code
        except Exception as e:
            code = f"EXC {str(e)[:40]}"
        if not isinstance(code, int) or code >= 500:
            malos.append((u, code))
    ok(not malos, f"ninguna descarga del catalogo da 500: {malos}")
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_su2.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_su2.js'], capture_output=True, text=True)
        if rc.returncode:
            ok(False, f"JS roto: {rc.stderr[:100]}"); break
    for lang in ('en', 'ca', 'fr', 'de', 'it', 'pt'):
        dd = json.load(open(os.path.join(BASE, 'static', 'i18n', f'{lang}.json'), encoding='utf-8'))
        if 'menu.descargas' not in dd:
            ok(False, f"i18n {lang} sin menu.descargas"); break
    else:
        ok(True, "i18n: 6 idiomas con menu.descargas")
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
