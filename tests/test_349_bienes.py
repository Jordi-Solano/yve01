# -*- coding: utf-8 -*-
"""349 con compras de bienes a proveedores de la UE (b78, finanzas 7 sep 2026).
Una factura AP cuyo proveedor lleva NIF-IVA comunitario (prefijo de otro pais
de la UE) es una adquisicion intracomunitaria de bienes:
  - 303: devengado 10-11 y deducible 36-37 sobre la base (no 28-29)
  - 349: fila con clave 'A' (las comisiones OTA siguen con clave 'S')
  - SII recibidas: clave 09
  - un proveedor español (ESB…) o sin NIF no va al 349
  - el aviso de "casillas no verificadas" esta

  python3.12 tests/test_349_bienes.py
  python3.12 tests/test_349_bienes.py --sabotaje
"""
import os
import sys

import pandas as pd

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

    import fiscal as FI
    import cierre_mes as CM
    if SABOTAJE:
        FI.pais_nif_ue = lambda nif: ''
    ok(FI.pais_nif_ue('DE123456789') == 'DE' and FI.pais_nif_ue('FR 12 345678901') == 'FR' and FI.pais_nif_ue('EL123456789') == 'EL' and FI.pais_nif_ue('XI123456789') == 'XI',
       'NIF-IVA comunitario: DE, FR (con espacios), EL (Grecia), XI (Irlanda del Norte)')
    ok(FI.pais_nif_ue('ESB12345678') == '' and FI.pais_nif_ue('B12345678') == '' and FI.pais_nif_ue('GB123456789') == '' and FI.pais_nif_ue('CHE123') == '' and FI.pais_nif_ue('') == '',
       'ES, sin prefijo, GB (fuera de la UE), CHE (Suiza) y vacio: no son comunitarios')
    fu = {
        'ap': pd.DataFrame([
            {'numero_factura': 'A-1', 'nombre_proveedor': 'Makro', 'NIF_proveedor': 'ESA28647451', 'fecha_factura': '05/08/2026', 'total_factura': 1210.0, 'base_imponible': 1000.0, 'cuota_iva': 210.0, 'porcentaje_iva': 21},
            {'numero_factura': 'DE-1', 'nombre_proveedor': 'Weinkellerei Mosel GmbH', 'NIF_proveedor': 'DE811907980', 'fecha_factura': '12/08/2026', 'total_factura': 2000.0, 'base_imponible': 2000.0, 'cuota_iva': 0.0},
            {'numero_factura': 'DE-2', 'nombre_proveedor': 'Weinkellerei Mosel GmbH', 'NIF_proveedor': 'DE811907980', 'fecha_factura': '20/08/2026', 'total_factura': 500.0},
            {'numero_factura': 'FR-1', 'nombre_proveedor': 'Fromagerie SARL', 'NIF_proveedor': 'FR40303265045', 'fecha_factura': '15/08/2026', 'total_factura': 300.0, 'base_imponible': 300.0},
            {'numero_factura': 'SIN', 'nombre_proveedor': 'Sin nif', 'fecha_factura': '15/08/2026', 'total_factura': 121.0},
        ]),
        'ar_ota': pd.DataFrame([
            {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'fecha': '31/08/2026', 'importe_comision': 300.0},
        ]),
        'ventas_fb': pd.DataFrame(), 'reservas': pd.DataFrame(), 'banco': pd.DataFrame(), 'provisiones': [],
    }
    cfg = CM.config_cierre('/nonexistent')
    res = FI.calcular('2026-08', fu, cfg, {'nif': {'booking': 'NL805734958B01'}, 'periodicidad': 'mensual'})
    c = {x['clave']: x for x in res['m303']['casillas']}
    # AIB de bienes: 2000 + 500 + 300 = 2800 base, cuota 21 % = 588; + Booking 300 (servicios UE) -> 10-11 = 3100 / 651
    ok(c['dev_aib']['base'] == 3100.0 and c['dev_aib']['cuota'] == 651.0, f"303 casillas 10-11: bienes UE (2.800) + servicios OTA UE (300) = {c['dev_aib']['base']} / {c['dev_aib']['cuota']}")
    ok(c['ded_aib']['base'] == 3100.0 and c['ded_aib']['cuota'] == 651.0, f"303 casillas 36-37: lo mismo deducible = {c['ded_aib']['base']} / {c['ded_aib']['cuota']}")
    # 28-29 solo Makro (1000/210) y el sin NIF (100/21)
    ok(c['ded_int']['base'] == 1100.0 and c['ded_int']['cuota'] == 231.0, f"303 casillas 28-29: solo los proveedores españoles ({c['ded_int']['base']} / {c['ded_int']['cuota']})")
    f349 = {(x['operador'], x['clave']): x for x in res['m349']['filas']}
    ok(('Weinkellerei Mosel GmbH', 'A') in f349 and f349[('Weinkellerei Mosel GmbH', 'A')]['base'] == 2500.0 and f349[('Weinkellerei Mosel GmbH', 'A')]['nif'] == 'DE811907980',
       f"349: proveedor aleman con clave A y sus dos facturas sumadas ({f349.get(('Weinkellerei Mosel GmbH', 'A'))})")
    ok(('Fromagerie SARL', 'A') in f349 and f349[('Fromagerie SARL', 'A')]['base'] == 300.0, '349: proveedor frances con clave A')
    ok(('Booking.com', 'S') in f349, '349: la comision OTA sigue con clave S (servicios)')
    ok(not any(o in ('Makro', 'Sin nif') for o, _ in f349), '349: ni el proveedor español ni el que no tiene NIF')
    ok(res['m349']['n'] == 3 and res['m349']['total_base'] == 3100.0, f"349: 3 operadores, base total 3.100 ({res['m349']['n']}, {res['m349']['total_base']})")
    rec = {r['numero']: r for r in res['sii']['recibidas']}
    ok(rec['DE-1']['clave_regimen'] == '09' and rec['DE-1']['base'] == 2000.0 and rec['DE-1']['cuota'] == 420.0 and rec['A-1']['clave_regimen'] == '01',
       f"SII recibidas: la compra UE con clave 09 y cuota autoliquidada, la española con 01 ({rec['DE-1']['clave_regimen']}, {rec['DE-1']['cuota']})")
    ok(any('NO estan verificados' in a for a in res['avisos']), 'aviso: enfoque validado, numeros de casilla no verificados')
    ok('NO VERIFICADOS' in FI.__doc__, 'la validacion de finanzas esta apuntada en la cabecera de fiscal.py')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
