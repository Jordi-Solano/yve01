# -*- coding: utf-8 -*-
"""OLA B · bloque 6: fiscal (modelo 303, 349 y libros SII).

  python3.12 tests/test_fiscal.py
  python3.12 tests/test_fiscal.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                    # noqa: E402
import fiscal as FI                    # noqa: E402
import cierre_mes as CM                # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        # el 303 deja de mirar el regimen: toda OTA como española (IVA incluido)
        FI.regimen_ota = lambda nombre, cfg=None: 'es'

    fu = {
        'ap': pd.DataFrame([
            {'numero_factura': 'A-1', 'nombre_proveedor': 'Makro', 'fecha_factura': '05/08/2026', 'total_factura': 1210.0, 'base_imponible': 1000.0, 'cuota_iva': 210.0, 'porcentaje_iva': 21, 'tipo_proveedor': 'FB'},
            {'numero_factura': 'A-2', 'nombre_proveedor': 'Panaderia', 'fecha_factura': '06/08/2026', 'total_factura': 110.0, 'porcentaje_iva': 10},
            {'numero_factura': 'A-JUL', 'nombre_proveedor': 'Otro', 'fecha_factura': '06/07/2026', 'total_factura': 500.0},
        ]),
        'ar_ota': pd.DataFrame([
            {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'fecha': '31/08/2026', 'importe_comision': 300.0},
            {'numero_factura': 'EX-1', 'nombre_ota': 'Expedia', 'fecha': '31/08/2026', 'importe_comision': 200.0},
            {'numero_factura': 'HB-1', 'nombre_ota': 'Hotelbeds', 'fecha': '31/08/2026', 'importe_comision': 121.0},
        ]),
        'ventas_fb': pd.DataFrame([
            {'fecha': '10/08/2026', 'total_venta': 550.0}, {'fecha': '10/08/2026', 'total_venta': 550.0},
            {'fecha': '11/08/2026', 'total_venta': 110.0},
        ]),
        'reservas': pd.DataFrame([
            {'numero_reserva': 'R-1', 'cliente': 'Empresa SA', 'estado': 'FACTURADA', 'total': 1100.0, 'fecha_emision': '20/08/2026', 'importe_habitaciones': 880.0, 'importe_fb': 220.0},
            {'numero_reserva': 'R-P', 'cliente': 'Pend', 'estado': 'PENDIENTE_FACTURA', 'total': 999.0, 'fecha_emision': '20/08/2026'},
        ]),
        'banco': pd.DataFrame(), 'provisiones': [],
    }
    cfg = CM.config_cierre('/nonexistent')
    res = FI.calcular('2026-08', fu, cfg, {'nif': {'booking': 'NL805734958B01'}, 'periodicidad': 'mensual'})
    c = {x['clave']: x for x in res['m303']['casillas']}
    # devengado 10 %: TPV 1210 -> base 1100 / 110 ; R-1 1100 -> base 1000 / 100  => 2100 / 210
    ok(c['dev_10']['base'] == 2100.0 and c['dev_10']['cuota'] == 210.0, f"04/06 (10 %): {c['dev_10']['base']} / {c['dev_10']['cuota']}")
    ok(c['dev_21']['base'] == 0 and c['dev_21']['cuota'] == 0, '07/09 (21 %) vacio: ninguna venta al 21')
    # Booking UE: 300 base, 63 cuota en 10/11 y en 36/37 ; Expedia no_ue: 200 / 42 en 12/13 y en 28/29
    ok(c['dev_aib']['base'] == 300.0 and c['dev_aib']['cuota'] == 63.0 and c['ded_aib']['cuota'] == 63.0, f"Booking (UE) en 10/11 y 36/37: {c['dev_aib']} / {c['ded_aib']}")
    ok(c['dev_isp']['base'] == 200.0 and c['dev_isp']['cuota'] == 42.0, f"Expedia (no UE) en 12/13: {c['dev_isp']}")
    # deducible interior: Makro 210 + Panaderia 10 + Hotelbeds 21 + Expedia ISP 42 = 283 ; base 1000+100+100+200 = 1400
    ok(c['ded_int']['base'] == 1400.0 and c['ded_int']['cuota'] == 283.0, f"28/29: {c['ded_int']}")
    m = res['m303']
    ok(m['c27_devengado'] == 315.0 and m['c45_deducible'] == 346.0 and m['c46_resultado'] == -31.0 and m['signo'] == 'A COMPENSAR',
       f"303: devengado {m['c27_devengado']} deducible {m['c45_deducible']} resultado {m['c46_resultado']} {m['signo']}")
    ok(res['m349']['n'] == 1 and res['m349']['filas'][0]['operador'] == 'Booking.com' and res['m349']['filas'][0]['base'] == 300.0
       and res['m349']['filas'][0]['nif'] == 'NL805734958B01' and res['m349']['filas'][0]['clave'] == 'S', f"349: {res['m349']['filas']}")
    sii = res['sii']
    ok(sii['n_expedidas'] == 3 and sii['n_recibidas'] == 5, f"SII: {sii['n_expedidas']} expedidas (2 TPV + 1 fra), {sii['n_recibidas']} recibidas (2 AP + 3 OTA)")
    isp = {r['nombre']: r['inversion_sujeto_pasivo'] for r in sii['recibidas']}
    ok(isp['Booking.com'] == 'S' and isp['Expedia'] == 'S' and isp['Hotelbeds'] == 'N' and isp['Makro'] == 'N', f"ISP marcado en SII: {isp}")
    ok(all(r['numero'] != 'A-JUL' for r in sii['recibidas']), 'la factura de julio no entra en agosto')
    ok('Expedia' in res['nif_pendientes'] and 'Makro' in res['nif_pendientes'] and 'Booking.com' not in res['nif_pendientes'], f"NIF pendientes: {res['nif_pendientes']}")
    ok(res['estado'] == 'PENDIENTE' and any('67' in a for a in res['avisos']), f"estado {res['estado']}, aviso casilla 67")
    v = FI.calcular('2026-08', {}, cfg, None)
    ok(v['estado'] == 'SIN_DATO' and v['m303']['c46_resultado'] == 0.0 and v['m303']['signo'] == 'SIN ACTIVIDAD', 'sin fuentes: SIN_DATO y 0')

    # el 303 tiene que decir lo mismo que el libro (477 / 472) de los asientos del mes
    asi = CM.generar_asientos('2026-08', fu, None, cfg)
    rep = round(sum(a['haber'] - a['debe'] for a in asi['asientos'] if a['cuenta'] == '477'), 2)
    sop = round(sum(a['debe'] - a['haber'] for a in asi['asientos'] if a['cuenta'] == '472'), 2)
    ok(rep == m['c27_devengado'] and sop == m['c45_deducible'], f"303 = libro: 477 {rep} vs 27 {m['c27_devengado']} · 472 {sop} vs 45 {m['c45_deducible']}")
    pq = FI.resumen_para_paquete.__doc__ and True
    ok(pq and set(FI.hojas(res)) == {'Modelo 303', 'Modelo 349', 'SII expedidas', 'SII recibidas', 'Avisos'}, 'hojas para el paquete')
    bx, nx = FI.exportar_excel(res)
    x = pd.read_excel(bx, sheet_name=None)
    ok(nx == 'fiscal_2026-08.xlsx' and len(x) == 5 and float(x['Modelo 303'].iloc[-1]['cuota']) == -31.0, 'Excel con 5 hojas y resultado en la ultima fila del 303')

    # ── endpoints ─────────────────────────────────────────────────────
    import dashboard as D
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    g = cl.get('/api/fiscal?mes=2026-08').get_json() or {}
    ok(g.get('ok') and 'm303' in g and 'm349' in g and 'sii' in g and g['m303'].get('signo'), f"/api/fiscal responde: estado {g.get('estado')}, {g.get('cifra')}")
    ok(g.get('libro') is None or isinstance(g['libro'].get('cuadra'), bool), '/api/fiscal trae el cuadre con el libro (o None si no hay)')
    rx = cl.get('/api/exportar/fiscal?mes=2026-08')
    ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/fiscal da un xlsx')
    p = cl.get('/api/cierre/paquete?mes=2026-08').get_json() or {}
    it = next((i for i in p.get('checklist', []) if i.get('clave') == 'fiscal'), None)
    ok(it is not None and 'pendiente del bloque fiscal' not in str(it.get('cifra', '')) and '303' in str(it.get('cifra', '')), f"paquete: fiscal = {it and it.get('cifra')}")
    html = cl.get('/').get_data(as_text=True)
    ok('id="card-cierre-fiscal"' in html and 'function loadFiscal' in html and 'loadFiscal();' in html, 'tarjeta en la pestaña Cierre')
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
