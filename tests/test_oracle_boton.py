# -*- coding: utf-8 -*-
"""BOMBA 2 (parte a) · el boton de Oracle existe y dice en que modo trabaja.

ANTES: `procesarOracle()` estaba escrita y traducida a 7 idiomas, pero
`btnOracle` no existia en el HTML: nadie podia pulsarla. Y `/api/oracle/status`
decidia "real" solo con `ORACLE_BASE_URL`, mientras que `oracle_auth` (lo que
ejecuta el pipeline) simula si faltan CLIENT_ID/SECRET o la URL es la de
plantilla: dos criterios distintos para la misma pregunta.

AHORA: el boton esta en la barra de AP, con un chip que dice "simulacion · sin
Oracle real" cuando toca; el estado sale de `oracle_auth.is_simulation()` en
cada peticion; y el resultado del pipeline en simulacion no canta
"contabilizacion completada".

  python3.12 tests/test_oracle_boton.py
  python3.12 tests/test_oracle_boton.py --sabotaje
"""
import json
import os
import re
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv
CLAVES = ('oracle.chipSim', 'oracle.chipReal', 'oracle.titleSim', 'oracle.titleReal',
          'oracle.avisoSim', 'oracle.simOk')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard
    import oracle_auth
    if SABOTAJE:
        # El pipeline pasa a "real" y la pantalla NO debe seguir diciendo
        # simulacion: si el endpoint hubiera cacheado el modo, no se enteraria.
        oracle_auth.is_simulation = lambda: False
    app = dashboard.app
    app.config['TESTING'] = True
    c = app.test_client()
    assert c.post('/api/login', json={'username': 'admin',
                                      'password': 'admin123'}).status_code == 200

    html = c.get('/').get_data(as_text=True)
    js = '\n'.join(re.findall(r'<script[^>]*>(.*?)</script>', html, re.S))
    sin_js = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.S)

    # ── 1 · el boton existe y esta cableado ──────────────────────────
    ok('id="btnOracle"' in sin_js, 'hay un boton btnOracle en el HTML')
    ok(re.search(r'id="btnOracle"[^>]*onclick="procesarOracle\(\)"', sin_js) or
       re.search(r'onclick="procesarOracle\(\)"[^>]*id="btnOracle"', sin_js),
       'el boton llama a procesarOracle()')
    ok('id="oracle-modo-chip"' in sin_js, 'hay un chip para decir el modo')
    ok('function _cargarModoOracle' in js and "_cargarModoOracle()" in js.split('async function loadAP')[1][:400],
       'loadAP consulta el modo de Oracle')

    # ── 2 · el estado dice lo que dice oracle_auth, en cada peticion ──
    st = c.get('/api/oracle/status').get_json() or {}
    ok(st.get('simulacion') is True and st.get('mode') == 'simulation',
       f"/api/oracle/status dice simulacion (sin credenciales): {st}")
    ok(st.get('simulacion') == bool(oracle_auth.is_simulation()),
       'el estado coincide con oracle_auth.is_simulation()')

    # ── 3 · el resultado en simulacion no canta contabilizacion ──────
    m = re.search(r'function procesarOracle\(\)\s*\{(.*?)\n\}', js, re.S)
    cuerpo = m.group(1) if m else ''
    ok("t('oracle.simOk'" in cuerpo and '_oracleSim' in cuerpo,
       'el titulo final distingue simulacion de contabilizacion real')
    ok("t('oracle.avisoSim'" in cuerpo, 'antes de lanzar avisa de que es simulacion')

    # ── 3b · cada getElementById de procesarOracle apunta a un id que EXISTE
    # (la primera version usaba los ids de un modal viejo y reventaba al pulsar)
    ids = set(re.findall(r"getElementById\('([^']+)'\)", cuerpo))
    faltan = sorted(i for i in ids if f'id="{i}"' not in sin_js)
    ok(ids and not faltan, f'todos los ids de procesarOracle existen en el HTML (faltan: {faltan})')

    # ── 4 · traducciones en los 6 idiomas ────────────────────────────
    for lang in ('en', 'ca', 'fr', 'de', 'it', 'pt'):
        with open(os.path.join(BASE, 'static', 'i18n', f'{lang}.json'), encoding='utf-8') as fh:
            d = json.load(fh)
        ok(all(k in d and d[k].strip() for k in CLAVES), f'{lang}.json tiene las {len(CLAVES)} claves')

    # ── 5 · el JS servido parsea ─────────────────────────────────────
    tmpjs = os.path.join(tempfile.mkdtemp(prefix='orb_'), 'servido.js')
    with open(tmpjs, 'w', encoding='utf-8') as fh:
        fh.write(js)
    res = subprocess.run(['node', '--check', tmpjs], capture_output=True)
    ok(res.returncode == 0, 'node --check del JS servido: ' +
       (res.stderr.decode()[:200] if res.returncode else 'parsea'))

    # ── 6 · ningun oracle_* tocado ───────────────────────────────────
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True,
                          text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_')], 'ningun oracle_* tocado')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else
              '*** SABOTAJE SIN EFECTO: la prueba no protege nada ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
