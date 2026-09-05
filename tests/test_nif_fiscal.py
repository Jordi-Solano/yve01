# -*- coding: utf-8 -*-
"""El NIF que lee la IA se conserva y alimenta el 349 y el SII (Jordi, sep 2026):
  AP  -> `NIF_proveedor` (mayusculas, como lo escribe el lector)
  OTA -> `nif_ota` (patron nuevo de lector_ota)
  AR  -> ficha de cliente > bono de la misma agencia (`NIF_agencia`) > config

  python3.12 tests/test_nif_fiscal.py
  python3.12 tests/test_nif_fiscal.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

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
    import lector_ota as LO
    if SABOTAJE:
        FI._col = lambda r, *n: ""        # vuelve a perder lo que leyo la IA

    for txt, esp in [("Booking.com B.V. · Amsterdam · NL805734958B01\nNúmero de factura: 1", "NL805734958B01"),
                     ("Hotelbeds Spain SLU  CIF: B-57218372", "B-57218372"),
                     ("Expedia · Seattle WA · VAT ID (none - non-EU)\nInvoice Number: EXP-1", "NO_ENCONTRADO")]:
        got = LO.buscar_campo(txt, "nif_ota")
        ok(got == esp, f"lector_ota nif_ota: {got!r}")

    fu = {
        'ap': pd.DataFrame([{'numero_factura': 'EL-88213', 'nombre_proveedor': 'Energia Llevant SA', 'NIF_proveedor': 'A-58012345', 'fecha': '31/08/2026',
                             'total_factura': 1175.8, 'base_imponible': 971.74, 'cuota_iva': 204.06, 'porcentaje_iva': 21, 'tipo_proveedor': 'OTRAS'},
                            {'numero_factura': 'X-1', 'nombre_proveedor': 'Sin Nif SL', 'NIF_proveedor': 'NO_ENCONTRADO', 'fecha': '31/08/2026', 'total_factura': 121.0, 'porcentaje_iva': 21}]),
        'ar_ota': pd.DataFrame([{'numero_factura': '2410099001', 'nombre_ota': 'Booking.com', 'nif_ota': 'NL805734958B01', 'fecha': '31/08/2026', 'importe_comision': 1200.0},
                                {'numero_factura': 'EXP-1', 'nombre_ota': 'Expedia', 'nif_ota': 'NO_ENCONTRADO', 'fecha': '31/08/2026', 'importe_comision': 720.0},
                                {'numero_factura': 'HB-1', 'nombre_ota': 'Hotelbeds', 'nif_ota': 'B-66432109', 'fecha': '31/08/2026', 'importe_comision': 121.0}]),
        'ventas_fb': pd.DataFrame(),
        'reservas': pd.DataFrame([{'numero_reserva': 'R-1', 'cliente': 'Viatges Mediterrani', 'estado': 'FACTURADA', 'total': 1320.0, 'fecha_emision': '20/08/2026'},
                                  {'numero_reserva': 'R-2', 'cliente': 'Empresa Ficha SA', 'estado': 'FACTURADA', 'total': 110.0, 'fecha_emision': '20/08/2026'},
                                  {'numero_reserva': 'R-3', 'cliente': 'Nadie SL', 'estado': 'FACTURADA', 'total': 110.0, 'fecha_emision': '20/08/2026'}]),
        'clientes': pd.DataFrame([{'nombre_cliente': 'Empresa Ficha SA', 'nif': 'B-11111111'}]),
        'bonos': pd.DataFrame([{'numero_bono': 'VM-7781', 'agencia': 'Viatges Mediterrani', 'NIF_agencia': 'B-62233445', 'importe_total': 1320.0}]),
        'banco': pd.DataFrame(), 'provisiones': [],
    }
    cfg = CM.config_cierre('/nonexistent')
    res = FI.calcular('2026-08', fu, cfg, {'nif': {}, 'nif_propio': 'B-66432109', 'periodicidad': 'mensual'})
    rec = {r['nombre']: r['nif'] for r in res['sii']['recibidas']}
    exp = {r['nombre']: r['nif'] for r in res['sii']['expedidas']}
    ok(rec.get('Energia Llevant SA') == 'A58012345', f"SII recibidas: NIF del proveedor que leyo la IA → {rec.get('Energia Llevant SA')}")
    ok(rec.get('Booking.com') == 'NL805734958B01', f"SII recibidas: NIF-IVA de Booking del PDF → {rec.get('Booking.com')}")
    ok(res['m349']['filas'] and res['m349']['filas'][0]['nif'] == 'NL805734958B01', f"349: Booking con su NIF-IVA sin config → {res['m349']['filas'][0]['nif'] if res['m349']['filas'] else '?'}")
    ok(rec.get('Hotelbeds') == '', "el NIF del propio hotel pillado por error NO se usa como NIF de la OTA")
    ok(exp.get('Viatges Mediterrani') == 'B62233445', f"SII expedidas: NIF de la agencia del bono → {exp.get('Viatges Mediterrani')}")
    ok(exp.get('Empresa Ficha SA') == 'B11111111', f"SII expedidas: NIF de la ficha de cliente → {exp.get('Empresa Ficha SA')}")
    pend = set(res['nif_pendientes'])
    ok(pend == {'Sin Nif SL', 'Expedia', 'Hotelbeds', 'Nadie SL'}, f"pendientes solo los que de verdad no tienen: {sorted(pend)}")
    # recoger_fuentes trae clientes y bonos
    src = open(os.path.join(BASE, 'cierre_mes.py'), encoding='utf-8').read()
    ok('("clientes", "clientes_credito.xlsx")' in src and '("bonos", "bonos_agencia.xlsx")' in src, "recoger_fuentes carga la ficha de clientes y los bonos")
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
