# -*- coding: utf-8 -*-
"""BOMBA 1 · el boton "Aprobar Match OK" del panel de AP.

ANTES: `/api/ap/aprobar_lote` buscaba una columna `aprobacion` en los
informes `matching_*.xlsx` que NINGUN modulo escribe, asi que `aprobadas`
valia siempre 0... y el JS hacia `d.aprobadas || nums.length`, o sea que
enseñaba "✓ 5 facturas aprobadas" habiendo aprobado CERO. Encima el fetch
no mandaba el token de CSRF (en Render: 403).

AHORA: aprueba de verdad, por el MISMO registro que usa la pantalla
"Facturas por aprobar" (`aprobaciones_ap.xlsx`, que es lo que lee Oracle),
solo lo que el servidor comprueba que cuadra y esta pendiente, y devuelve
las cifras reales. El mensaje del navegador no puede inventarse nada.

  python3.12 tests/test_aprobar_match_ok.py
  python3.12 tests/test_aprobar_match_ok.py --sabotaje
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                       # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = datetime.now().strftime('%Y%m%d')

PDIR = os.path.join(BASE, 'facturas-procesadas')
ADIR = os.path.join(BASE, 'aprobaciones')
RDIR = os.path.join(BASE, 'reportes')
CONTAB = os.path.join(PDIR, f'facturas_contabilizadas_{HOY}.xlsx')
APRO = os.path.join(ADIR, 'aprobaciones_ap.xlsx')
AUDIT = os.path.join(BASE, 'datos-referencia', 'audit_log.json')


def fila(n, prov, tot, estado, archivo=None):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'archivo': archivo or f'{n}.pdf', 'estado_matching': estado,
            'cuenta_contable': '600', 'tipo_proveedor': 'OTRAS',
            'fecha_factura': '01/09/2026', 'hotel_id': ''}


# Lo que hay en el panel: dos que cuadran, una por albaran, una con
# discrepancia, una ya RECHAZADA que tambien cuadra, y una SIN numero.
FACTURAS = [
    fila('F-OK-1', 'Makro', 121.0, 'MATCH_3WAY_OK'),
    fila('F-OK-2', 'Lavanderia', 242.0, 'MATCH_3WAY_OK'),
    fila('F-ALB', 'Pescados Rias', 363.0, 'MATCH_ALBARAN_OK'),
    fila('F-DISC', 'Endesa', 484.0, 'DIFERENCIA_IMPORTE'),
    fila('F-RECH', 'Carnes Sur', 605.0, 'MATCH_3WAY_OK'),
    fila('', 'Frutas Norte', 77.0, 'MATCH_3WAY_OK', archivo='foto_frutas.jpg'),
]


class Guardado:
    """Copia y devuelve tal cual todo lo que la prueba toca."""

    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='amok_')
        self.items = {}

    def guarda(self, ruta):
        existia = os.path.exists(ruta)
        copia = os.path.join(self.tmp, str(len(self.items)))
        if existia:
            (shutil.copytree if os.path.isdir(ruta) else shutil.copy)(ruta, copia)
        self.items[ruta] = (existia, copia)

    def restaura(self):
        for ruta, (existia, copia) in self.items.items():
            if os.path.isdir(ruta):
                shutil.rmtree(ruta)
            elif os.path.exists(ruta):
                os.remove(ruta)
            if existia:
                (shutil.copytree if os.path.isdir(copia) else shutil.copy)(copia, ruta)
        shutil.rmtree(self.tmp, ignore_errors=True)


def montar():
    os.makedirs(PDIR, exist_ok=True)
    os.makedirs(ADIR, exist_ok=True)
    for f in os.listdir(PDIR):
        if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
            os.remove(os.path.join(PDIR, f))
    for f in os.listdir(RDIR):
        if f.startswith('matching_') and f.endswith('.xlsx'):
            os.remove(os.path.join(RDIR, f))
    pd.DataFrame(FACTURAS).to_excel(CONTAB, index=False)
    pd.DataFrame([{'fecha_hora': '01/09/2026 09:00:00', 'numero_factura': 'F-RECH',
                   'clave_factura': 'F-RECH', 'accion': 'RECHAZADA',
                   'comentario': 'precio', 'departamento': 'Cocina',
                   'aprobador': 'jefe', 'hotel_id': ''}]
                 ).to_excel(APRO, index=False)


def js_servido(c):
    html = c.get('/').get_data(as_text=True)
    return '\n'.join(re.findall(r'<script[^>]*>(.*?)</script>', html, re.S))


def main():
    g = Guardado()
    for r in (PDIR, ADIR, AUDIT):
        g.guarda(r)
    for f in os.listdir(RDIR):
        if f.startswith('matching_') and f.endswith('.xlsx'):
            g.guarda(os.path.join(RDIR, f))
    fallos = 0
    try:
        montar()
        if SABOTAJE:
            # El escritor "dice que si" sin escribir: la bomba original.
            import app_aprobacion_ap as _P
            _P.registrar_acciones = lambda filas: len(filas or [])
        import dashboard
        app = dashboard.app
        app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin',
                                          'password': 'admin123'}).status_code == 200
        TOK = (c.get('/api/csrf_token').get_json() or {}).get('token')
        assert TOK

        def ok(cond, msg):
            nonlocal fallos
            print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
            if not cond:
                fallos += 1

        # ── 0 · lo que el panel enseña ANTES ──────────────────────────
        facts = c.get('/api/facturas_ap').get_json()
        stats = c.get('/api/stats_ap').get_json()
        ok(len(facts) == 6 and stats['aprobadas'] == 0 and stats['rechazadas'] == 1,
           f"punto de partida: 6 facturas, 0 aprobadas, 1 rechazada ({stats['aprobadas']}/{stats['rechazadas']})")
        claves = {f['numero_factura']: f.get('clave') for f in facts}
        ok(all(f.get('clave') for f in facts),
           'cada factura del panel viaja con su `clave` (la sin numero tambien)')
        ok(claves.get('N/D') == 'foto_frutas.jpg',
           f"la factura sin numero usa el fichero como clave: {claves.get('N/D')!r}")

        # ── 1 · el boton, como lo manda el navegador ──────────────────
        # El navegador manda TODO lo que cree que cuadra; el servidor decide.
        pedidas = [f['clave'] for f in facts if f['estado'] in
                   ('MATCH_3WAY_OK', 'MATCH_ALBARAN_OK', 'MATCH_CORRECTO', 'MATCH_PO_OK')]
        pedidas.append('F-DISC')          # un cliente malicioso o un bug: no debe colar
        r = c.post('/api/ap/aprobar_lote', json={'facturas': pedidas},
                   headers={'X-CSRF-Token': TOK})
        d = r.get_json() or {}
        print('    respuesta:', d)
        ok(r.status_code == 200 and d.get('ok') is True, f'HTTP {r.status_code}')
        # F-OK-1, F-OK-2, F-ALB y la sin numero: 4. F-RECH ya decidida, F-DISC no cuadra.
        ok(d.get('aprobadas') == 4, f"aprueba 4 de verdad (dice {d.get('aprobadas')})")
        ok(d.get('ya_decididas') == 1, f"1 ya tenia decision y se respeta ({d.get('ya_decididas')})")
        ok(d.get('no_cuadran') == 1, f"1 no cuadra y se rechaza aunque la pidan ({d.get('no_cuadran')})")

        # ── 2 · el registro que lee Oracle y la pantalla de aprobar ───
        apro = pd.read_excel(APRO)
        nuevas = apro[apro['accion'].astype(str) == 'APROBADA']
        ok(sorted(nuevas['clave_factura'].astype(str)) ==
           ['F-ALB', 'F-OK-1', 'F-OK-2', 'foto_frutas.jpg'],
           f"aprobaciones_ap.xlsx tiene las 4 claves: {sorted(nuevas['clave_factura'].astype(str))}")
        ok(set(nuevas.columns) >= {'fecha_hora', 'numero_factura', 'clave_factura', 'accion',
                                   'comentario', 'departamento', 'aprobador', 'hotel_id'},
           'misma forma de fila que "Facturas por aprobar" (Oracle lee numero_factura)')
        ok('F-DISC' not in set(apro['clave_factura'].astype(str)), 'F-DISC no se ha escrito')
        rech = apro[apro['clave_factura'].astype(str) == 'F-RECH']
        ok(len(rech) == 1 and rech['accion'].iloc[0] == 'RECHAZADA',
           'F-RECH sigue RECHAZADA, una sola fila')
        ok(all(str(a) == 'admin' for a in nuevas['aprobador']),
           'el aprobador es el usuario logueado, no un texto fijo')
        sin_num = nuevas[nuevas['clave_factura'].astype(str) == 'foto_frutas.jpg']
        ok(len(sin_num) == 1 and str(sin_num['numero_factura'].iloc[0]) == 'foto_frutas.jpg',
           'la sin numero falla en cerrado para Oracle: numero_factura = clave')

        # ── 3 · el panel DESPUES ──────────────────────────────────────
        stats2 = c.get('/api/stats_ap').get_json()
        # Sep 2026 (hallazgo (a) de la bomba 1): el tile cruza por la CLAVE, asi
        # que la aprobada SIN numero tambien cuenta: 4, no 3.
        ok(stats2['aprobadas'] == 4,
           f"el tile Aprobadas pasa de 0 a 4 (3 con numero + la sin numero; dice {stats2['aprobadas']})")
        facts2 = {f['numero_factura']: f['accion'] for f in c.get('/api/facturas_ap').get_json()}
        ok(facts2.get('F-OK-1') == 'APROBADA' and facts2.get('F-ALB') == 'APROBADA'
           and facts2.get('F-DISC') == '' and facts2.get('F-RECH') == 'RECHAZADA',
           f"la columna Aprobacion del panel dice la verdad: {facts2}")

        # ── 4 · segunda pulsacion: no aprueba dos veces ───────────────
        r2 = c.post('/api/ap/aprobar_lote', json={'facturas': pedidas},
                    headers={'X-CSRF-Token': TOK}).get_json()
        ok(r2.get('aprobadas') == 0 and r2.get('ya_decididas') == 5,
           f"repetir no duplica: 0 nuevas, 5 ya decididas ({r2})")
        ok(len(pd.read_excel(APRO)) == 5, 'el fichero sigue con 5 filas')

        # ── 5 · sin CSRF se rechaza (en Render era un 403 silencioso) ─
        r3 = c.post('/api/ap/aprobar_lote', json={'facturas': pedidas})
        ok(r3.status_code == 403, f'sin token de CSRF: {r3.status_code}')

        # ── 6 · el JS que recibe el navegador ─────────────────────────
        js = js_servido(c)
        ok('d.aprobadas || nums.length' not in js,
           'el navegador ya no se inventa el numero de aprobadas')
        m = re.search(r'async function aprobarMatchOK\(\)\s*\{(.*?)\n\}', js, re.S)
        cuerpo = m.group(1) if m else ''
        ok('_postJson(' in cuerpo, 'el boton manda el token de CSRF (_postJson)')
        ok("fetch('/api/ap/aprobar_lote'" not in cuerpo, 'sin fetch a pelo')
        with open(os.devnull, 'w') as dn:
            tmpjs = os.path.join(g.tmp, 'servido.js')
            with open(tmpjs, 'w', encoding='utf-8') as fh:
                fh.write(js)
            res = subprocess.run(['node', '--check', tmpjs], stdout=dn, stderr=subprocess.PIPE)
        ok(res.returncode == 0, 'node --check del JS servido: ' +
           (res.stderr.decode()[:200] if res.returncode else 'parsea'))

        # ── 7 · Oracle intacto ────────────────────────────────────────
        diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True,
                              text=True, cwd=BASE).stdout.split()
        ok(not [f for f in diff if f.startswith('oracle_')], 'ningun oracle_* tocado')
    finally:
        g.restaura()
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else
              '*** SABOTAJE SIN EFECTO: la prueba no protege nada ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
