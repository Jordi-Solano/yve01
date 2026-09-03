# -*- coding: utf-8 -*-
"""OLA A · zona Oracle con OK explicito de Jordi (3 sep 2026):

  (2) la SIMULACION respeta la puerta APROBADA igual que el real;
  (3) el paso 4/4 no falla ni con demo (facturas en facturas_ap_*) ni con cero
      facturas;
  (4) el exportador GL solo exporta lo que el pipeline produjo de verdad.

Todo sobre un arbol temporal: no toca los datos del repo.

  python3.12 tests/test_oracle_simulacion_honesta.py
  python3.12 tests/test_oracle_simulacion_honesta.py --sabotaje
"""
import io
import contextlib
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_TENANT'] = 'default'

import pandas as pd                                    # noqa: E402
import oracle_actualizar_estado as OAE                 # noqa: E402
import oracle_lector_facturas as ORA                   # noqa: E402
import oracle_pipeline as OP                           # noqa: E402
import oracle_export_dryrun as OX                      # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def fila(n, prov, tot):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'base_imponible': round(tot / 1.21, 2), 'cuota_iva': round(tot - tot / 1.21, 2),
            'hotel_id': '', 'fecha_factura': '01/09/2026', 'fecha': '01/09/2026',
            'cuenta_contable': '600', 'cuenta_debe_gasto': '600', 'estado_asignacion': 'ASIGNADA',
            'tipo_proveedor': 'FB'}


def montar(raiz, facturas, aprobadas):
    pdir, apdir, rdir, ddir = [os.path.join(raiz, d) for d in ('facturas-procesadas', 'aprobaciones', 'reportes', 'datos-referencia')]
    for d in (pdir, apdir, rdir, ddir):
        os.makedirs(d, exist_ok=True)
    if facturas:
        # como el demo: SOLO facturas_ap_*, ningun facturas_contabilizadas_*
        pd.DataFrame(facturas).to_excel(os.path.join(pdir, 'facturas_ap_20260901.xlsx'), index=False)
    if aprobadas:
        pd.DataFrame([{'numero_factura': n, 'accion': 'APROBADA', 'aprobador': 'admin',
                       'fecha_hora': '2026-09-01 09:00'} for n in aprobadas]
                     ).to_excel(os.path.join(apdir, 'aprobaciones_ap.xlsx'), index=False)
    ORA.PROCESADAS_DIR = Path(pdir)
    ORA.REPORTES_DIR = Path(rdir)
    ORA.APROBACIONES_DIR = Path(apdir)
    OAE.PROCESADAS_DIR = Path(pdir)
    OAE.REGISTRO_FILE = Path(ddir) / 'oracle_contabilizadas.json'
    OAE.ASIENTOS_FILE = Path(rdir) / 'oracle_asientos_producidos.json'
    import oracle_crear_asientos as OC
    OC.REPORTES_DIR = Path(rdir)
    return pdir, rdir


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    ok(ORA.is_simulation(), 'el sandbox esta en simulacion (sin credenciales Oracle)')
    if SABOTAJE:
        ORA.cargar_aprobaciones_ap = lambda: ({'F-1', 'F-2'}, pd.DataFrame())   # todo aprobado
    raiz = tempfile.mkdtemp(prefix='orasim_')
    try:
        # ── (2) la simulacion respeta la puerta ───────────────────────
        pdir, rdir = montar(raiz, [fila('F-1', 'Makro', 121.0), fila('F-2', 'Otis', 605.0)], aprobadas=['F-1'])
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            batches, bloq, df = ORA.preparar_facturas_para_oracle()
        nums = sorted(b['numero_factura'] for b in batches)
        ok(nums == ['F-1'], f'simulacion: solo la APROBADA entra ({nums})')
        ok(any(b['numero_factura'] == 'F-2' and 'no aprobada' in b['motivo'] for b in bloq), 'la no aprobada queda bloqueada, con motivo')
        ok('ignora aprobación' not in buf.getvalue() and 'se respeta' in buf.getvalue(), 'el log ya no dice "ignora aprobacion"')

        # ── (3) paso 4/4 marca en facturas_ap_* (demo) ────────────────
        with contextlib.redirect_stdout(io.StringIO()):
            stats = OP.run_pipeline()
        ok(stats['errores'] == 0 and stats['contabilizadas'] == 1 and stats['facturas_bloqueadas'] == 1,
           f"pipeline sin errores: {stats['contabilizadas']} contabilizada (sim), {stats['facturas_bloqueadas']} bloqueada, errores {stats['errores']}")
        dfa = pd.read_excel(os.path.join(pdir, 'facturas_ap_20260901.xlsx'))
        est = dict(zip(dfa['numero_factura'].map(str), dfa['oracle_status'].map(lambda v: '' if v != v else str(v))))
        ok(est.get('F-1') == 'CONTABILIZADA_SIM' and est.get('F-2') == '', f'marcador escrito en facturas_ap_*: {est}')
        ok(not OAE.ya_contabilizadas_registro(), 'la simulacion NO entra en el registro de contabilizadas de verdad')

        # ── (4) el exportador exporta lo producido ────────────────────
        prod = OAE.asientos_producidos()
        ok(len(prod) == 1 and prod[0]['numero_factura'] == 'F-1' and len(prod[0]['journal_lines']) == 3, f'asientos producidos guardados: {len(prod)}')
        rows = OX.gl_rows(OX.asientos_exportables())
        dr = sum(float(r['ENTERED_DR'] or 0) for r in rows); cr = sum(float(r['ENTERED_CR'] or 0) for r in rows)
        ok(len(rows) == 3 and abs(dr - cr) < 0.01 and abs(cr - 121.0) < 0.01 and rows[0]['ATTRIBUTE1'] == 'F-1',
           f'GL: 3 lineas cuadradas ({dr:.2f} / {cr:.2f}), sin nada de F-2')
        ok(not any('inventad' in l or 'proveedores.xlsx' in l for l in open(os.path.join(BASE, 'oracle_export_dryrun.py'), encoding='utf-8').read().split('\n') if 'importe_mensual_estimado' in l),
           'ya no queda el generador de asientos inventados')
        # segunda ejecucion: F-1 ya esta CONTABILIZADA_SIM en el Excel -> no se duplica en el export
        with contextlib.redirect_stdout(io.StringIO()):
            stats2 = OP.run_pipeline()
        ok(stats2['errores'] == 0, f'segunda ejecucion sin errores ({stats2["contabilizadas"]} contabilizadas)')
        ok(len(OX.asientos_exportables()) == 1, 'una factura = un asiento exportable aunque el pipeline corra dos veces')

        # ── (3b) cero facturas: no es error ───────────────────────────
        shutil.rmtree(raiz, ignore_errors=True)
        montar(raiz, [], aprobadas=[])
        with contextlib.redirect_stdout(io.StringIO()):
            stats0 = OP.run_pipeline()
        ok(stats0['errores'] == 0 and stats0['facturas_leidas'] == 0, f"cero facturas: errores {stats0['errores']}, nada que contabilizar")
        with contextlib.redirect_stdout(io.StringIO()):
            df0, ruta0, st0 = OAE.actualizar_estados([{'numero_factura': 'X', 'estado': 'CONTABILIZADA_SIM'}])
        ok(ruta0 == '' and st0['sin_excel'] == 1, 'actualizar_estados sin ningun Excel: aviso, no excepcion')
        ok(OX.gl_rows(OX.asientos_exportables()) == [], 'sin pipeline, nada que exportar')

        # ── endpoints ─────────────────────────────────────────────────
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        r = c.get('/api/oracle/export_excel')
        ok(r.status_code == 404 and 'no ha producido' in (r.get_json() or {}).get('error', ''), f'/api/oracle/export_excel sin asientos: {r.status_code} honesto')
        st = c.get('/api/oracle/status').get_json()
        ok(st.get('simulacion') is True and 'mode' in st, "/api/oracle/status sigue siendo el del dashboard")
        html = c.get('/').get_data(as_text=True)
        ok('/api/oracle/export_excel' in html and 'oracle.exportGl' in html, 'enlace "Asientos GL" junto al boton de Oracle')
    finally:
        shutil.rmtree(raiz, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
