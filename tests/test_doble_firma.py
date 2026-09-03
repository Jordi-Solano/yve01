# -*- coding: utf-8 -*-
"""OLA A · doble firma por importe (>500 EUR: dos personas).

La primera firma se guarda como FIRMA_1 y Oracle NO la ve (solo lee
APROBADA). La segunda, de OTRA persona, escribe APROBADA. Vale para
"Facturas por aprobar" y para el boton "Aprobar Match OK" del panel.

  python3.12 tests/test_doble_firma.py
  python3.12 tests/test_doble_firma.py --sabotaje
"""
import os
import shutil
import sys
import tempfile
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                    # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = datetime.now().strftime('%Y%m%d')
PDIR = os.path.join(BASE, 'facturas-procesadas')
ADIR = os.path.join(BASE, 'aprobaciones')
RDIR = os.path.join(BASE, 'reportes')
DDIR = os.path.join(BASE, 'datos-referencia')
APRO = os.path.join(ADIR, 'aprobaciones_ap.xlsx')
CFG = os.path.join(DDIR, 'config_aprobaciones.json')
AUDIT = os.path.join(DDIR, 'audit_log.json')


def fila(n, prov, tot, estado='MATCH_3WAY_OK'):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'archivo': f'{n}.pdf', 'estado_matching': estado, 'cuenta_contable': '600',
            'tipo_proveedor': 'OTRAS', 'fecha_factura': '01/09/2026', 'hotel_id': ''}


class Guardado:
    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='df_')
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


def main():
    g = Guardado()
    for r in (PDIR, ADIR, CFG, AUDIT):
        g.guarda(r)
    for f in os.listdir(RDIR):
        if f.startswith('matching_') and f.endswith('.xlsx'):
            g.guarda(os.path.join(RDIR, f))
    fallos = 0
    try:
        os.makedirs(PDIR, exist_ok=True)
        os.makedirs(ADIR, exist_ok=True)
        for f in os.listdir(PDIR):
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PDIR, f))
        for f in os.listdir(RDIR):
            if f.startswith('matching_') and f.endswith('.xlsx'):
                os.remove(os.path.join(RDIR, f))
        if os.path.exists(APRO):
            os.remove(APRO)
        if os.path.exists(CFG):
            os.remove(CFG)
        pd.DataFrame([fila('F-300', 'Makro', 300.0), fila('F-1200', 'Otis', 1200.0),
                      fila('F-1500', 'Endesa', 1500.0), fila('F-DISC', 'Frutas', 900.0, 'DIFERENCIA_IMPORTE')]
                     ).to_excel(os.path.join(PDIR, f'facturas_contabilizadas_{HOY}.xlsx'), index=False)

        import app_aprobacion_ap as P
        if SABOTAJE:
            P.necesita_doble_firma = lambda total: False      # se olvida del umbral
        import dashboard
        import oracle_lector_facturas as ORA
        app = dashboard.app
        app.config['TESTING'] = True

        def ok(cond, msg):
            nonlocal fallos
            print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
            if not cond:
                fallos += 1

        def cliente(user, pwd):
            c = app.test_client()
            assert c.post('/api/login', json={'username': user, 'password': pwd}).status_code == 200
            tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
            return c, tok

        def accion(c, tok, clave, tipo='APROBADA'):
            return c.post('/aprobaciones-ap/api/accion', json={
                'clave': clave, 'numero_factura': clave, 'accion': tipo,
                'comentario': 'ok', 'departamento': 'Administracion'},
                headers={'X-CSRF-Token': tok})

        def aprobadas_oracle():
            aprobadas, _ = ORA.cargar_aprobaciones_ap()   # (set, df)
            return {str(x) for x in aprobadas}

        admin, ta = cliente('admin', 'admin123')
        fc, tf = cliente('fc_user', 'hotel2024')
        ok(P.umbral_doble_firma() == 500.0, f'umbral por defecto 500 ({P.umbral_doble_firma()})')

        # ── 1 · por debajo del umbral: una firma basta ──────────────
        r = accion(admin, ta, 'F-300').get_json()
        ok(r.get('ok') and r.get('accion') == 'APROBADA' and r.get('doble_firma') is False, f'300 EUR: APROBADA a la primera {r}')

        # ── 2 · por encima: primera firma, y Oracle no la ve ────────
        r = accion(admin, ta, 'F-1200').get_json()
        ok(r.get('ok') and r.get('accion') == 'FIRMA_1' and r.get('falta_segunda'), f'1.200 EUR: FIRMA_1 {r}')
        ora = aprobadas_oracle()
        ok('F-300' in ora and 'F-1200' not in ora, f'Oracle ve F-300 y NO F-1200: {sorted(ora)}')
        pend = admin.get('/aprobaciones-ap/api/facturas').get_json()
        f1200 = next((x for x in pend if x['clave'] == 'F-1200'), None)
        ok(f1200 is not None and f1200['firmas'] == 1 and f1200['firma1_por'] == 'admin' and f1200['doble_firma'],
           f'sigue pendiente con 1/2 firmas de admin: {f1200 and (f1200["firmas"], f1200["firma1_por"])}')
        ok(not any(x['clave'] == 'F-300' for x in pend), 'F-300 ya no esta en pendientes')
        st = admin.get('/aprobaciones-ap/api/stats').get_json()
        ok(st.get('segunda_firma') == 1 and st.get('umbral_doble_firma') == 500.0, f"stats: {st.get('segunda_firma')} esperando segunda firma")

        # ── 3 · la misma persona no puede firmar dos veces ──────────
        rr = accion(admin, ta, 'F-1200')
        ok(rr.status_code == 409 and 'otra persona' in (rr.get_json() or {}).get('error', ''), f'admin otra vez: {rr.status_code}')
        ok(len(pd.read_excel(APRO)) == 2, 'no se ha escrito nada mas')

        # ── 4 · la segunda persona aprueba de verdad ────────────────
        r = accion(fc, tf, 'F-1200').get_json()
        ok(r.get('ok') and r.get('accion') == 'APROBADA' and r.get('firma1_por') == 'admin', f'fc_user: APROBADA 2/2 {r}')
        ora = aprobadas_oracle()
        ok('F-1200' in ora, 'ahora Oracle SI la ve')
        apro = pd.read_excel(APRO)
        ok(sorted(apro['aprobador'].astype(str)) == ['admin', 'admin', 'fc_user'], f"el aprobador es la persona logueada: {sorted(apro['aprobador'].astype(str))}")
        hist = admin.get('/aprobaciones-ap/api/historial').get_json()
        ok(any(h['accion'] == 'FIRMA_1' and h['vigente'] is False for h in hist), 'la primera firma queda en el historial, no vigente')

        # ── 5 · el boton del panel sigue la misma regla ─────────────
        r = admin.post('/api/ap/aprobar_lote', json={'facturas': ['F-1500', 'F-DISC']}, headers={'X-CSRF-Token': ta}).get_json()
        ok(r.get('aprobadas') == 0 and r.get('primera_firma') == 1 and r.get('no_cuadran') == 1, f'lote de admin: 1 primera firma {r}')
        r = admin.post('/api/ap/aprobar_lote', json={'facturas': ['F-1500']}, headers={'X-CSRF-Token': ta}).get_json()
        ok(r.get('aprobadas') == 0 and r.get('esperan_segunda') == 1, f'lote de admin otra vez: espera a otra persona {r}')
        ok('F-1500' not in aprobadas_oracle(), 'Oracle sigue sin ver F-1500')
        r = fc.post('/api/ap/aprobar_lote', json={'facturas': ['F-1500']}, headers={'X-CSRF-Token': tf}).get_json()
        ok(r.get('aprobadas') == 1 and r.get('primera_firma') == 0, f'lote de fc_user: la segunda firma aprueba {r}')
        ok('F-1500' in aprobadas_oracle(), 'y Oracle ya la ve')

        # ── 6 · el umbral es configurable ───────────────────────────
        with open(CFG, 'w') as fh:
            fh.write('{"umbral_doble_firma": 5000}')
        ok(P.umbral_doble_firma() == 5000.0 and not P.necesita_doble_firma(1200), 'config_aprobaciones.json manda')
        os.remove(CFG)

        # ── 7 · el gate de Oracle no se ha tocado ───────────────────
        import subprocess
        diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
        ok(not [f for f in diff if f.startswith('oracle_')], 'ningun oracle_* tocado')
        html = admin.get('/aprobaciones-ap/').get_data(as_text=True)
        ok('b-firma' in html and 'PRIMERA FIRMA' in html, 'la pantalla pinta el badge de firmas')
    finally:
        g.restaura()
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
