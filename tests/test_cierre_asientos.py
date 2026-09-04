# -*- coding: utf-8 -*-
"""OLA B · bloque 1: asientos del mes y reconciliacion de cuentas.

Funciones puras con DataFrames controlados (cada fuente con un caso dentro y
otro fuera del mes) + endpoints contra dashboard.app + la pestaña.

  python3.12 tests/test_cierre_asientos.py
  python3.12 tests/test_cierre_asientos.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                    # noqa: E402
import cierre_mes as CM                # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
MES = '2026-08'


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        CM._Diario.nuevo.__defaults__  # existe
        _orig = CM._Diario.nuevo
        def _roto(self, fecha, concepto, documento, origen, lineas, hotel=""):
            lineas = [(c, d, h * 2 if c == "400" else h) for c, d, h in lineas]   # el haber de proveedores se duplica
            return _orig(self, fecha, concepto, documento, origen, lineas, hotel)
        CM._Diario.nuevo = _roto

    fuentes = {
        'ap': pd.DataFrame([
            {'numero_factura': 'F-1', 'nombre_proveedor': 'Makro', 'fecha_factura': '05/08/2026', 'base_imponible': 100.0, 'cuota_iva': 21.0, 'total_factura': 121.0, 'cuenta_debe_gasto': '600', 'tipo_proveedor': 'FB', 'hotel_id': ''},
            {'numero_factura': 'F-2', 'nombre_proveedor': 'Endesa', 'fecha_factura': '2026-08-20', 'base_imponible': 0, 'cuota_iva': 0, 'total_factura': 242.0, 'cuenta_contable': 629.0, 'porcentaje_iva': 21, 'tipo_proveedor': 'OTRAS', 'hotel_id': ''},
            {'numero_factura': 'F-3', 'nombre_proveedor': 'Otis', 'fecha_factura': '01/09/2026', 'base_imponible': 50, 'cuota_iva': 10.5, 'total_factura': 60.5, 'cuenta_contable': '622', 'hotel_id': ''},   # fuera del mes
            {'numero_factura': 'F-4', 'nombre_proveedor': 'Raro', 'fecha_factura': '10/08/2026', 'base_imponible': 81.82, 'cuota_iva': 17.18, 'total_factura': 99.0, 'cuenta_contable': '627', 'hotel_id': ''},   # cuenta fuera del plan
            {'numero_factura': 'F-6', 'nombre_proveedor': 'NoCuadra', 'fecha_factura': '12/08/2026', 'base_imponible': 10, 'cuota_iva': 1, 'total_factura': 60.0, 'cuenta_contable': '629', 'hotel_id': ''},   # no cuadra
            {'numero_factura': 'F-5', 'nombre_proveedor': 'SinTotal', 'fecha_factura': '11/08/2026', 'total_factura': 0, 'hotel_id': ''},
        ]),
        'ar_ota': pd.DataFrame([
            {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'fecha': '2026-08-15', 'importe_comision': 1000.0, 'hotel_id': ''},
            {'numero_factura': 'BK-0', 'nombre_ota': 'Booking.com', 'fecha': '2026-08-15', 'importe_comision': 0, 'hotel_id': ''},
            {'numero_factura': 'EX-9', 'nombre_ota': 'Expedia', 'fecha': '2026-07-15', 'importe_comision': 500.0, 'hotel_id': ''},
        ]),
        'ventas_fb': pd.DataFrame([
            {'fecha': '2026-08-01', 'total_venta': 110.0}, {'fecha': '2026-08-01', 'total_venta': 220.0},
            {'fecha': '2026-08-02', 'total_venta': 55.0}, {'fecha': '2026-09-02', 'total_venta': 999.0},
        ]),
        'reservas': pd.DataFrame([
            {'numero_reserva': 'FAC-1', 'cliente': 'ACME', 'fecha_emision': '2026-08-10', 'importe_habitaciones': 880.0, 'importe_fb': 220.0, 'importe_extras': 0, 'total': 1100.0, 'estado': 'COBRADO', 'fecha_cobro': '2026-08-25', 'hotel_id': ''},
            {'numero': 'GRP-DEMO', 'cliente': 'Demo SA', 'fecha_emision': '2026-08-12', 'importe': 550.0, 'estado': 'PENDIENTE', 'hotel_id': ''},
            {'numero_reserva': 'GRP-77', 'cliente': 'Pendiente', 'fecha_emision': '', 'total': 5000.0, 'estado': 'PENDIENTE_FACTURA', 'hotel_id': ''},
        ]),
        'banco': pd.DataFrame([
            {'fecha': '2026-08-06', 'concepto': 'MAKRO', 'importe': -121.0, 'estado': 'CONCILIADO', 'factura_ref': 'F-1'},
            {'fecha': '2026-08-26', 'concepto': 'ACME', 'importe': 1100.0, 'estado': 'CONCILIADO', 'factura_ref': 'FAC-1'},
            {'fecha': '2026-08-28', 'concepto': 'DESCONOCIDO', 'importe': -300.0, 'estado': 'PENDIENTE', 'factura_ref': ''},
            {'fecha': '2026-07-28', 'concepto': 'VIEJO', 'importe': -10.0, 'estado': 'CONCILIADO', 'factura_ref': 'X'},
        ]),
        'provisiones': [{'asientos': [
            {'fecha': '2026-08-31', 'cuenta': '600', 'concepto': 'Provision 2026-08 albaranes — Frutas', 'debe': 80.0, 'haber': 0},
            {'fecha': '2026-08-31', 'cuenta': '4009', 'concepto': 'Provision 2026-08 albaranes — Frutas', 'debe': 0, 'haber': 80.0}]}],
    }
    plan = CM.plan_cuentas()
    res = CM.generar_asientos(MES, fuentes, plan, CM.config_cierre())
    A = res['asientos']
    def lineas(origen=None, cuenta=None):
        return [a for a in A if (origen is None or a['origen'] == origen) and (cuenta is None or a['cuenta'] == cuenta)]

    ok(res['cuadra'] and res['debe'] == res['haber'], f"el Diario cuadra: {res['debe']} / {res['haber']}")
    ok(res['fuentes'] == {'ap': 3, 'ar_ota': 1, 'ventas_fb': 2, 'ar_facturas': 2, 'ar_cobros': 0, 'banco': 2, 'provisiones': 1}, f"fuentes: {res['fuentes']}")
    ok(res['saltados']['ap_sin_cuadrar'] == 1 and res['saltados']['ap_sin_total'] == 1, f"AP saltadas con motivo: {res['saltados']}")
    f2 = [a for a in A if a['documento'] == 'F-2']
    ok(any(a['cuenta'] == '629' and a['debe'] == 200.0 for a in f2) and any(a['cuenta'] == '472' and a['debe'] == 42.0 for a in f2), 'F-2 sin base: se deriva del total al 21 % y la cuenta 629.0 se lee como 629')
    bk1 = [a for a in A if a['documento'] == 'BK-1']
    ok(sorted((a['cuenta'], a['debe'], a['haber']) for a in bk1) == [('410', 0.0, 1000.0), ('472', 210.0, 0.0), ('477', 0.0, 210.0), ('628', 1000.0, 0.0)], f'OTA con inversion del sujeto pasivo: {sorted((a["cuenta"], a["debe"], a["haber"]) for a in bk1)}')
    res_inc = CM.generar_asientos(MES, {'ar_ota': fuentes['ar_ota']}, plan, {'otas': {'booking': 'es'}})
    ok(any(a['cuenta'] == '628' and a['debe'] == 826.45 for a in res_inc['asientos']), 'OTA marcada como española: IVA incluido, base 826,45')
    ok(CM.regimen_ota('HotelBeds') == 'es' and CM.regimen_ota('Booking.com') == 'ue' and CM.regimen_ota('Expedia') == 'no_ue' and CM.regimen_ota('OTA rara') == 'es', 'regimen por OTA: lo desconocido es "es", no se inventa ISP')
    tpv = lineas('FB', '570')
    ok([a['debe'] for a in tpv] == [330.0, 55.0], f'TPV agrupado por dia: {[a["debe"] for a in tpv]}')
    ok(any(a['cuenta'] == '700' and a['haber'] == 300.0 for a in lineas('FB')) and any(a['cuenta'] == '477' and a['haber'] == 30.0 for a in lineas('FB')), 'ventas F&B al 10 %')
    fac1 = [a for a in A if a['documento'] == 'FAC-1' and a['origen'] == 'AR']
    ok(any(a['cuenta'] == '705' and a['haber'] == 800.0 for a in fac1) and any(a['cuenta'] == '700' and a['haber'] == 200.0 for a in fac1) and any(a['cuenta'] == '430' and a['debe'] == 1100.0 for a in fac1), 'factura AR: 705/700 al 10 % y 430 por el total')
    ok(not any(a['cuenta'] == '572' and a['concepto'].startswith('Cobro') for a in fac1), 'cobro AR por fecha_cobro NO se asienta si el banco ya lo tiene conciliado (evita duplicar 572)')
    demo = [a for a in A if a['documento'] == 'GRP-DEMO']
    ok(any(a['cuenta'] == '430' and a['debe'] == 550.0 for a in demo), 'esquema viejo de reservas (numero/importe/PENDIENTE) tambien se asienta')
    ok(not any(a['documento'] == 'GRP-77' for a in A) and not any(a['documento'] in ('F-3', 'EX-9') for a in A), 'PENDIENTE_FACTURA y lo de otro mes quedan fuera')
    bk = lineas('BANCO')
    ok(len(bk) == 4 and any(a['cuenta'] == '400' and a['debe'] == 121.0 for a in bk) and any(a['cuenta'] == '430' and a['haber'] == 1100.0 for a in bk), 'banco: solo lo CONCILIADO del mes, pago a 400 y cobro a 430')
    ok(len(lineas('PROVISION')) == 2 and any(a['cuenta'] == '4009' for a in A), 'provisiones incorporadas')
    ok(res['cuentas_fuera_plan'] == ['627'], f"cuentas fuera de plan detectadas: {res['cuentas_fuera_plan']}")
    my = {m['cuenta']: m for m in CM.mayor(A, plan)}
    ok(my['400']['haber'] == 462.0 and my['400']['debe'] == 121.0 and my['400']['saldo'] == -341.0, f"mayor 400: {my['400']}")

    rec = CM.reconciliar(MES, res, fuentes, drr={'rooms_revenue_mtd': 1300.0, 'concepto': 'Rooms Revenue'})
    ch = {c['concepto'][:20]: c for c in rec['checks']}
    est = {c['cuenta'] + '|' + c['concepto'][:25]: c['estado'] for c in rec['checks']}
    ok(rec['checks'][0]['estado'] == 'CUADRA', 'check 0: diario cuadra')
    c400 = next(c for c in rec['checks'] if c['cuenta'] == '400')
    ok(c400['estado'] == 'DIFERENCIA' and c400['justificado'] == 522.0 and c400['libro'] == 462.0, f"400: facturado 522 vs asentado 462 (F-6 no cuadra): {c400['estado']} {c400['diferencia']}")
    c572 = [c for c in rec['checks'] if c['cuenta'] == '572']
    ok(c572[0]['estado'] == 'CUADRA' and c572[0]['libro'] == 979.0, f"572 conciliado neto 979: {c572[0]}")
    ok(c572[1]['estado'] == 'PENDIENTE' and '1 movimiento' in c572[1]['nota'], f"572 sin conciliar: {c572[1]['nota']}")
    c705 = next(c for c in rec['checks'] if c['cuenta'] == '705')
    ok(c705['estado'] == 'DIFERENCIA' and c705['libro'] == 1300.0 and c705['diferencia'] == 0.0 or c705['estado'] == 'CUADRA', f"705 vs DRR: libro {c705['libro']} justificado {c705['justificado']} -> {c705['estado']}")
    c700 = next(c for c in rec['checks'] if c['cuenta'] == '570/700')
    ok(c700['estado'] == 'CUADRA' and c700['justificado'] == 385.0, f"TPV: {c700}")
    ok(any(c['estado'] == 'REVISAR' and c['cuenta'] == '627' for c in rec['checks']), 'la 627 sale como REVISAR')
    rec0 = CM.reconciliar(MES, res, fuentes, drr=None)
    ok(next(c for c in rec0['checks'] if c['cuenta'] == '705')['estado'] == 'SIN_DATO', 'sin DRR: 705 = sin dato, no se inventa')

    buf, nombre = CM.exportar_excel(res, rec)
    ok(set(pd.read_excel(buf, sheet_name=None)) == {'Libro Diario', 'Mayor', 'Reconciliacion', 'Resumen'} and nombre == 'cierre_2026-08_asientos.xlsx', 'Excel con 4 hojas')

    # ── endpoints + pestaña ─────────────────────────────────────────
    import dashboard as D
    app = D.app; app.config['TESTING'] = True
    c = app.test_client()
    assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    r = c.get('/api/cierre/asientos?mes=2026-08')
    d = r.get_json() or {}
    ok(r.status_code == 200 and d.get('ok') and 'reconciliacion' in d and 'mayor' in d and 'cuadra' in d, f'/api/cierre/asientos ({r.status_code})')
    rx = c.get('/api/exportar/cierre?mes=2026-08')
    ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/cierre da un xlsx')
    html = c.get('/').get_data(as_text=True)
    ok('id="tab-cierre"' in html and 'id="panel-cierre"' in html and 'function loadCierre' in html and "cierre:      function(){ return loadCierre(); }" in html, 'pestaña Cierre con su cargador')
    import ast
    tree = ast.parse(open(os.path.join(BASE, 'cierre_mes.py'), encoding='utf-8').read())
    escribe = [n.attr for n in ast.walk(tree) if isinstance(n, ast.Attribute) and n.attr in ('to_excel', 'to_csv', 'remove', 'unlink')]
    ok(escribe == ['to_excel'] * 4, f'cierre_mes.py solo escribe en el BytesIO: {escribe}')
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
