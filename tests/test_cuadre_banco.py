# -*- coding: utf-8 -*-
"""OLA B · bloque 2: cuadre de banco por pestañas (AR, AP, tarjetas, caja, varios).

  python3.12 tests/test_cuadre_banco.py
  python3.12 tests/test_cuadre_banco.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                    # noqa: E402
import cuadre_banco as CB              # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
MES = '2026-08'
DDIR = os.path.join(BASE, 'datos-referencia')
MANUAL = os.path.join(DDIR, 'cuadre_banco_manual.json')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        _c = CB.clasificar
        CB.clasificar = lambda mov, p, m, pr: ('AP', 'todo-ap')     # lo mete todo en AP

    bk = pd.DataFrame([
        {'fecha': '2026-08-02', 'concepto': 'Pago Makro Cash - Fra. F-1', 'importe': -121.0, 'saldo': 9879.0, 'estado': 'CONCILIADO', 'origen': 'AP', 'factura_ref': 'F-1'},
        {'fecha': '2026-08-03', 'concepto': 'TRANSF ENDESA ENERGIA', 'importe': -300.0, 'saldo': 9579.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-04', 'concepto': 'BOOKING.COM BV LIQUIDACION', 'importe': 2500.0, 'saldo': 12079.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-05', 'concepto': 'LIQ. TARJETA REDSYS 04/08', 'importe': 640.0, 'saldo': 12719.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-06', 'concepto': 'INGRESO EFECTIVO OFICINA 123', 'importe': 900.0, 'saldo': 13619.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-07', 'concepto': 'NOMINA AGOSTO', 'importe': -4200.0, 'saldo': 9419.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-08', 'concepto': 'COMISION MANTENIMIENTO CTA', 'importe': -12.0, 'saldo': 9407.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-09', 'concepto': 'TRF 4471 XZ', 'importe': -77.0, 'saldo': 9330.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-08-10', 'concepto': 'Cobro ACME - RES-2026-0301', 'importe': 1100.0, 'saldo': 10430.0, 'estado': 'CONCILIADO', 'origen': 'AR', 'factura_ref': 'FAC-1'},
        {'fecha': '2026-08-11', 'concepto': 'DEVOLUCION BOOKING', 'importe': -50.0, 'saldo': 10380.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
        {'fecha': '2026-07-30', 'concepto': 'VIEJO', 'importe': -1.0, 'saldo': 10000.0, 'estado': 'PENDIENTE', 'origen': '', 'factura_ref': ''},
    ])
    vf = pd.DataFrame([{'fecha': '2026-08-04', 'total_venta': 400.0}, {'fecha': '2026-08-05', 'total_venta': 240.0}, {'fecha': '2026-09-01', 'total_venta': 999.0}])
    provs = ['makro', 'endesa']
    r = CB.cuadrar(MES, bk, vf, CB.palabras(), {}, provs)
    P = {m['concepto'][:12]: m['pestana'] for m in r['movimientos']}
    ok(r['n'] == 10 and 'VIEJO' not in [m['concepto'] for m in r['movimientos']], f"{r['n']} movimientos del mes (el de julio fuera)")
    ok(P['Pago Makro C'] == 'AP' and P['Cobro ACME -'] == 'AR', 'lo conciliado va por su origen (AP/AR)')
    ok(P['TRANSF ENDES'] == 'AP', 'proveedor conocido en el concepto -> AP')
    ok(P['BOOKING.COM '] == 'AR' and P['DEVOLUCION B'] == 'AR', 'la OTA es AR con cualquier signo (cobro y devolucion)')
    ok(P['LIQ. TARJETA'] == 'TARJETAS' and P['INGRESO EFEC'] == 'CAJA', 'tarjetas y caja por concepto')
    ok(P['NOMINA AGOST'] == 'VARIOS' and P['COMISION MAN'] == 'VARIOS', 'nomina y comision bancaria -> varios')
    ok(P['TRF 4471 XZ'] == 'SIN_CLASIFICAR', 'lo que no se entiende queda SIN_CLASIFICAR, no se reparte')
    ps = r['pestanas']
    ok(ps['AP']['total'] == -421.0 and ps['AP']['justificado'] == -121.0 and ps['AP']['estado'] == 'PENDIENTE' and ps['AP']['n_sin_factura'] == 1, f"AP: {ps['AP']}")
    ok(ps['AR']['total'] == 3550.0 and ps['AR']['justificado'] == 1100.0 and ps['AR']['estado'] == 'PENDIENTE' and ps['AR']['n'] == 3, f"AR: {ps['AR']}")
    ok(ps['TARJETAS']['justificado'] == 640.0 and ps['TARJETAS']['diferencia'] == 0.0 and ps['TARJETAS']['estado'] == 'INFO', f"tarjetas vs TPV: {ps['TARJETAS']}")
    ok(ps['CAJA']['estado'] == 'SIN_DATO' and ps['CAJA']['total'] == 900.0, 'caja: sin dato (no se inventa arqueo)')
    ok(ps['SIN_CLASIFICAR']['n'] == 1 and r['ok'] is False, 'con algo sin clasificar el cuadre no esta OK')
    ok(r['saldo_final'] == 10380.0 and r['fecha_saldo'] == '2026-08-11', f"saldo final del extracto: {r['saldo_final']} ({r['fecha_saldo']})")
    # asignacion manual manda
    clave = next(m['clave'] for m in r['movimientos'] if m['concepto'].startswith('TRF 4471'))
    r2 = CB.cuadrar(MES, bk, vf, CB.palabras(), {clave: 'VARIOS'}, provs)
    ok(next(m for m in r2['movimientos'] if m['clave'] == clave)['pestana'] == 'VARIOS' and r2['pestanas']['SIN_CLASIFICAR']['n'] == 0, 'la asignacion manual manda y vacia SIN_CLASIFICAR')
    r3 = CB.cuadrar(MES, bk, vf, CB.palabras(), {clave: 'CAJA'}, provs)
    ok(r3['pestanas']['CAJA']['n'] == 2, 'y se puede mandar a cualquier pestaña')
    buf, nombre = CB.exportar_excel(r)
    ok(set(pd.read_excel(buf, sheet_name=None)) == {'Resumen', *CB.PESTANAS} and nombre == 'cuadre_banco_2026-08.xlsx', 'Excel con resumen + una hoja por pestaña')

    # ── endpoints (con copia del fichero de manuales) ──────────────
    tmp = tempfile.mkdtemp(prefix='cb_')
    existia = os.path.exists(MANUAL)
    if existia:
        shutil.copy(MANUAL, os.path.join(tmp, 'm.json'))
    try:
        if os.path.exists(MANUAL):
            os.remove(MANUAL)
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
        rr = c.get('/api/cuadre_banco?mes=2026-08'); d = rr.get_json() or {}
        ok(rr.status_code == 200 and d.get('ok_api') and 'pestanas' in d and set(d['pestanas']) == set(CB.PESTANAS), f'/api/cuadre_banco ({rr.status_code})')
        ra = c.post('/api/cuadre_banco/asignar', json={'clave': '2026-08-09|TRF 4471 XZ|-77.00', 'pestana': 'VARIOS'}, headers={'X-CSRF-Token': tok})
        ok(ra.status_code == 200 and json.load(open(MANUAL, encoding='utf-8')).get('2026-08-09|TRF 4471 XZ|-77.00') == 'VARIOS', 'asignar guarda en cuadre_banco_manual.json')
        rb = c.post('/api/cuadre_banco/asignar', json={'clave': 'x', 'pestana': 'NOEXISTE'}, headers={'X-CSRF-Token': tok})
        ok(rb.status_code == 400, 'pestaña desconocida -> 400')
        rq = c.post('/api/cuadre_banco/asignar', json={'clave': '2026-08-09|TRF 4471 XZ|-77.00', 'pestana': 'SIN_CLASIFICAR'}, headers={'X-CSRF-Token': tok})
        ok(rq.status_code == 200 and '2026-08-09|TRF 4471 XZ|-77.00' not in json.load(open(MANUAL, encoding='utf-8')), 'volver a SIN_CLASIFICAR quita la asignacion')
        rx = c.get('/api/exportar/cuadre_banco?mes=2026-08')
        ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/cuadre_banco da un xlsx')
        html = c.get('/').get_data(as_text=True)
        ok('id="card-cierre-banco"' in html and 'function loadCuadreBanco' in html and 'loadCuadreBanco();' in html, 'tarjeta en la pestaña Cierre')
    finally:
        if os.path.exists(MANUAL):
            os.remove(MANUAL)
        if existia:
            shutil.copy(os.path.join(tmp, 'm.json'), MANUAL)
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
