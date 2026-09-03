# -*- coding: utf-8 -*-
"""OLA A · aging AP: a quien debemos y desde cuando.

Facturas de proveedor + liquidaciones OTA sin pagar (= sin conciliar en el
banco), por tramos de antiguedad. Se prueba la funcion pura con DataFrames
y los dos endpoints contra dashboard.app.

  python3.12 tests/test_aging_ap.py
  python3.12 tests/test_aging_ap.py --sabotaje
"""
import os
import subprocess
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                # noqa: E402
import aging_ap as AG              # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = date(2026, 9, 3)


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        AG._pagadas = lambda df: set()          # se olvida del banco: todo pendiente

    df_ap = pd.DataFrame([
        {'numero_factura': 'F-10', 'nombre_proveedor': 'Makro', 'fecha_factura': '25/08/2026', 'total_factura': 121.0, 'accion': 'APROBADA', 'hotel_id': ''},
        {'numero_factura': 'F-45', 'nombre_proveedor': 'Makro', 'fecha_factura': '20/07/2026', 'total_factura': 242.0, 'accion': '', 'hotel_id': ''},
        {'numero_factura': 'F-70', 'nombre_proveedor': 'Endesa', 'fecha_factura': '2026-06-25', 'total_factura': 484.0, 'accion': '', 'hotel_id': ''},
        {'numero_factura': 'F-120', 'nombre_proveedor': 'Otis', 'fecha_factura': '01/05/2026', 'total_factura': 100.0, 'accion': 'APROBADA', 'hotel_id': ''},
        {'numero_factura': 'F-PAG', 'nombre_proveedor': 'Otis', 'fecha_factura': '01/05/2026', 'total_factura': 999.0, 'accion': 'APROBADA', 'hotel_id': ''},
        {'numero_factura': '', 'nombre_proveedor': 'Frutas', 'fecha_factura': '', 'total_factura': 50.0, 'accion': '', 'hotel_id': ''},
    ])
    df_ar = pd.DataFrame([
        {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'fecha': '2026-08-15', 'importe_comision': 300.0, 'hotel_id': ''},
        {'numero_factura': 'BK-0', 'nombre_ota': 'Booking.com', 'fecha': '2026-08-15', 'importe_comision': 0, 'hotel_id': ''},
    ])
    df_b = pd.DataFrame([
        {'fecha': '2026-08-30', 'concepto': 'OTIS', 'importe': -999.0, 'estado': 'CONCILIADO', 'factura_ref': 'F-PAG'},
        {'fecha': '2026-08-30', 'concepto': 'x', 'importe': -1.0, 'estado': 'PENDIENTE', 'factura_ref': 'F-10'},
    ])
    r = AG.calcular_aging(df_ap, df_ar, df_b, hoy=HOY)
    nums = {f['numero_factura']: f for f in r['filas']}
    ok('F-PAG' not in nums, 'la conciliada en el banco NO cuenta')
    ok('F-10' in nums and nums['F-10']['tramo'] == '0-30', 'F-10 (9 dias) en 0-30 aunque el banco la tenga PENDIENTE')
    ok(nums['F-45']['tramo'] == '31-60', f"F-45 (45 dias): {nums['F-45']['tramo']}")
    ok(nums['F-70']['tramo'] == '61-90', f"F-70 (70 dias, fecha ISO): {nums['F-70']['tramo']}")
    ok(nums['F-120']['tramo'] == '>90', f"F-120 (125 dias): {nums['F-120']['tramo']}")
    ok(nums['N/D']['tramo'] == 'sin fecha' and r['sin_fecha'] == 1, 'sin fecha se cuenta aparte, no se inventa tramo')
    ok('BK-1' in nums and nums['BK-1']['origen'] == 'OTA' and 'BK-0' not in nums, 'la OTA entra por su comision; comision 0 fuera')
    ok(r['total'] == round(121 + 242 + 484 + 100 + 50 + 300, 2), f"total {r['total']}")
    ok(r['tramos']['>90'] == 100.0 and r['tramos']['0-30'] == 421.0, f"tramos {r['tramos']}")
    ok(r['mas_de_60'] == 584.0, f"mas de 60 dias: {r['mas_de_60']}")
    makro = next(p for p in r['por_acreedor'] if p['acreedor'] == 'Makro')
    ok(makro['n'] == 2 and makro['importe'] == 363.0 and makro['sin_aprobar'] == 1 and makro['mas_antigua'] == '2026-07-20',
       f"Makro agrupado: {makro}")
    ok(r['por_acreedor'][0]['acreedor'] == 'Otis', 'ordenado por la mas antigua primero')
    r2 = AG.calcular_aging(df_ap, df_ar, None, hoy=HOY)
    ok(r2['n'] == 7 and 'F-PAG' in {f['numero_factura'] for f in r2['filas']}, 'sin banco todo cuenta como pendiente')

    buf, nombre = AG.exportar_excel(r)
    hojas = pd.read_excel(buf, sheet_name=None)
    ok(set(hojas) == {'Resumen', 'Por acreedor', 'Facturas'} and nombre.startswith('aging_ap_'), 'Excel con 3 hojas')

    import dashboard
    app = dashboard.app
    app.config['TESTING'] = True
    c = app.test_client()
    assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    rr = c.get('/api/aging_ap')
    d = rr.get_json() or {}
    ok(rr.status_code == 200 and d.get('ok') and 'tramos' in d and 'por_acreedor' in d, f'/api/aging_ap ({rr.status_code})')
    re_ = c.get('/api/exportar/aging_ap')
    ok(re_.status_code == 200 and re_.data[:2] == b'PK', '/api/exportar/aging_ap da un xlsx')
    html = c.get('/').get_data(as_text=True)
    ok('id="card-aging-ap"' in html and 'function loadAgingAP' in html, 'tarjeta y JS en el panel AP')

    import ast
    tree = ast.parse(open(os.path.join(BASE, 'aging_ap.py'), encoding='utf-8').read())
    escribe = [n.attr for n in ast.walk(tree) if isinstance(n, ast.Attribute)
               and n.attr in ('to_excel', 'to_csv', 'remove', 'unlink')]
    ok(escribe == ['to_excel'] * 3, f'aging_ap.py solo escribe en el BytesIO: {escribe}')
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
