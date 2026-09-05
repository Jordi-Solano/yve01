# -*- coding: utf-8 -*-
"""Hallazgo (g): el panel AP no recargaba al entrar en la pestaña tras generar
el demo (o cambiar de hotel). Ahora AP esta en `_CARGADORES` y el demo invalida
los paneles. Se ejecuta el JS de verdad con node sobre el HTML servido.

  python3.12 tests/test_ap_recarga.py
  python3.12 tests/test_ap_recarga.py --sabotaje
"""
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
    cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace("  ap:          function(){ return loadAP(); },\n", "")
    m1 = re.search(r"var _CARGADORES = \{.*?\n\};", html, re.S)
    m2 = re.search(r"function _cargarPanel\(tab, panel, forzar\) \{.*?\n\}", html, re.S)
    m3 = re.search(r"function _invalidarPaneles\(\) \{.*?\n\}", html, re.S)
    ok(m1 and m2 and m3, "el panel sirve _CARGADORES, _cargarPanel e _invalidarPaneles")
    gen = re.search(r"async function generarDemo\(\) \{.*?\n\}", html, re.S)
    ok(gen and "_invalidarPaneles()" in gen.group(0), "generarDemo invalida los paneles cacheados")
    if m1 and m2 and m3:
        js = '''
var llamadas = [];
function loadAP(){ llamadas.push('ap'); return Promise.resolve(); }
function loadFBTab(){ return Promise.resolve(); } function cargarARRealData(){ return Promise.resolve(); }
function loadDRR(){ return Promise.resolve(); } function loadBanco(){ return Promise.resolve(); }
function loadNotifConfig(){ return Promise.resolve(); } function loadMultiHotel(){ return Promise.resolve(); } function loadCierre(){ return Promise.resolve(); }
function _pintarYa(){} var _fbLoaded = {}; var _mh_loaded = false, _mhClasicaLoaded = false;
var document = { getElementById: function(){ return {id:'panel-ap'}; }, querySelector: function(){ return {id:'panel-ap', classList:{}}; } };
var window = {};
var _panelCargado = {};
''' + m1.group(0) + "\n" + m2.group(0) + "\n" + m3.group(0) + '''
// arranque: AP cargado y marcado
loadAP(); _panelCargado.ap = true;
_cargarPanel('ap', null, false);           // pulsar la pestaña: NO vuelve a pedir
var trasPulsar = llamadas.length;
_invalidarPaneles();                       // demo generado / cambio de hotel (AP a la vista)
var trasDemo = llamadas.length;
_cargarPanel('ap', null, false);           // ya repoblado: no pide otra vez
console.log(JSON.stringify({trasPulsar: trasPulsar, trasDemo: trasDemo, final: llamadas.length}));
'''
        open('/tmp/_apr.js', 'w', encoding='utf-8').write(js)
        out = subprocess.run(['node', '/tmp/_apr.js'], capture_output=True, text=True)
        try:
            import json; r = json.loads(out.stdout.strip().splitlines()[-1])
        except Exception:
            r = {}; ok(False, f"node: {out.stderr[:150]}")
        ok(r.get("trasPulsar") == 1, f"pulsar AP ya cargado no vuelve a pedir ({r.get('trasPulsar')} llamada)")
        ok(r.get("trasDemo") == 2, f"tras el demo, AP a la vista se repuebla solo ({r.get('trasDemo')} llamadas)")
        ok(r.get("final") == 2, "y no se pide dos veces")
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
