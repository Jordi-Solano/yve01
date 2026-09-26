# -*- coding: utf-8 -*-
"""b107 (finanzas, 26 sep 2026): decisiones VALIDADAS que no deben moverse sin querer.

  1. Salas de un contrato de grupo: cuenta propia (7052), separada de alojamiento (705) y
     de F&B (700), IVA 21 %.
  2. Fecha contable de AP = fecha de registro, y el IVA (303/SII) tambien por esa fecha.
  3. 30 dias de pago por defecto.

  python3.12 tests/test_validado_finanzas.py
  python3.12 tests/test_validado_finanzas.py --sabotaje
"""
import inspect
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

    import cierre_mes as CM
    import fiscal as FI
    import almacen_datos as AD
    if SABOTAJE:
        # salas otra vez al alojamiento, AP por fecha de factura y 60 dias
        CM.CUENTA_SALAS = '705'; CM.IVA_SALAS = 10.0
        CM.AP_FECHA_DEFECTO = 'factura'
        AD.DIAS_PAGO_DEFECTO = 60

    # 1. salas
    cfg = CM.config_cierre('/nonexistent')
    ok(CM.CUENTA_SALAS == '7052' and cfg.get('cuenta_salas') == '7052' and 'salas' in CM.CUENTAS_BASE.get('7052', '').lower(),
       f"salas: cuenta 7052 por defecto y en el plan ({CM.CUENTA_SALAS}, {cfg.get('cuenta_salas')})")
    fila = {'tipo': 'CONTRATO_GRUPO', 'total': 2860.0, 'importe_habitaciones': 1100.0, 'importe_fb': 550.0, 'importe_extras': 1210.0}
    dg = CM.desglose_factura_ar(fila, {**cfg, 'cuenta_salas': CM.CUENTA_SALAS})
    pc = {k: round(v, 2) for k, v in (dg.get('por_cuenta') or {}).items()}
    ok(pc == {'705': 1000.0, '700': 500.0, '7052': 1000.0}, f"factura de grupo: alojamiento 705, F&B 700 y salas 7052 por separado ({pc})")
    sal = [t for t in dg.get('tramos') or [] if t.get('cuenta') == '7052']
    ok(sal and float(sal[0]['pct']) == 21.0 and round(sal[0]['cuota'], 2) == 210.0 and round(dg['cuota'], 2) == 360.0,
       f"salas al 21 % (cuota {sal and sal[0]['cuota']}; total IVA {dg.get('cuota')})")

    # 2. fecha contable (y el IVA por la misma fecha)
    ok(CM.AP_FECHA_DEFECTO == 'contable' and CM.criterio_fecha_ap({}) == 'contable', f"AP entra por fecha contable ({CM.AP_FECHA_DEFECTO})")
    r = {'fecha_factura': '06/07/2026', 'fecha_contable': '2026-09-02'}
    ok(CM.fecha_ap(r, CM.criterio_fecha_ap({})) == '2026-09-02', "una factura de julio registrada en septiembre va a septiembre")
    fu = {'ap': pd.DataFrame([{'numero_factura': 'V-1', 'nombre_proveedor': 'Makro', 'fecha_factura': '06/07/2026', 'fecha_contable': '2026-09-02',
                               'total_factura': 1210.0, 'base_imponible': 1000.0, 'cuota_iva': 210.0, 'porcentaje_iva': 21}]),
          'ar_ota': pd.DataFrame(), 'ventas_fb': pd.DataFrame(), 'reservas': pd.DataFrame(), 'banco': pd.DataFrame(), 'provisiones': []}
    cf = dict(cfg); cf.pop('ap_fecha', None)
    iva = {m: sum(x['cuota'] for x in FI.calcular(m, fu, cf, None)['m303']['casillas'] if x['clave'].startswith('ded')) for m in ('2026-07', '2026-09')}
    ok(iva['2026-07'] == 0 and iva['2026-09'] == 210.0, f"el IVA soportado se deduce en el mes de registro ({iva})")

    # 3. 30 dias por defecto
    dflt = inspect.signature(AD.aplicar_ajustes_ap).parameters['dias_defecto'].default
    ok(AD.DIAS_PAGO_DEFECTO == 30 and dflt == 30, f"30 dias de pago por defecto ({AD.DIAS_PAGO_DEFECTO}, {dflt})")
    df = AD.aplicar_ajustes_ap(pd.DataFrame([{'numero_factura': 'X-1', 'nombre_proveedor': 'Nadie SL', 'fecha_factura': '01/09/2026', 'total_factura': 100.0}]),
                               ajustes={}, dias_prov={}, dias_defecto=AD.DIAS_PAGO_DEFECTO)
    f0 = df.iloc[0].to_dict()
    ok(int(f0.get('dias_pago')) == 30 and str(f0.get('vencimiento'))[:10] == '2026-10-01', f"vence a los 30 dias ({f0.get('dias_pago')}, {f0.get('vencimiento')})")

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
