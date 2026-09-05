# -*- coding: utf-8 -*-
"""FC real deja de ser una copia del FC teorico (Jordi, sep 2026):
  con recuento del mes  -> (inicial + compras F&B - final) / ventas del mes
  sin recuento          -> escandallo + mermas (aproximacion, y se dice)

  python3.12 tests/test_fc_real.py
  python3.12 tests/test_fc_real.py --sabotaje
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

    import tab_fb_dashboard as FB
    if SABOTAJE:
        FB.fc_real_mes = lambda *a, **k: {'metodo': None, 'fc_real_pct': None, 'mermas_mes': 0, 'consumo_real': None, 'mes': '', 'ventas_mes': 0, 'sin_explicar': None, 'compras_fb': None}

    # inventario de agosto: inicial 270+360 = 630, final 189+279 = 468
    inv = pd.DataFrame([
        {"ingrediente": "Pollo", "categoria": "Carnes", "coste_unitario": 4.5, "stock_inicial_kg_l": 60, "stock_actual_kg_l": 42, "unidad": "kg", "mes": "2026-08"},
        {"ingrediente": "Merluza", "categoria": "Pescados", "coste_unitario": 9.0, "stock_inicial_kg_l": 40, "stock_actual_kg_l": 31, "unidad": "kg", "mes": "2026-08"},
    ])
    # compras F&B de agosto: 200 de base
    ap = pd.DataFrame([{"nombre_proveedor": "Garraf", "tipo_proveedor": "FB", "fecha": "12/08/2026", "base_imponible": 200.0, "total_factura": 220.0},
                       {"nombre_proveedor": "Luz", "tipo_proveedor": "OTRAS", "fecha": "12/08/2026", "base_imponible": 999.0, "total_factura": 1200.0}])
    # ventas: 2.000 en agosto, 500 en septiembre (no cuentan para agosto)
    ven = pd.DataFrame([{"fecha": "05/08/2026", "plato": "Pollo al ast", "categoria": "Carnes", "unidades_vendidas": 100, "total_venta": 1400.0},
                        {"fecha": "20/08/2026", "plato": "Merluza", "categoria": "Pescados", "unidades_vendidas": 30, "total_venta": 600.0},
                        {"fecha": "02/09/2026", "plato": "Pollo al ast", "categoria": "Carnes", "unidades_vendidas": 40, "total_venta": 500.0}])
    mer = pd.DataFrame([{"fecha": "03/08/2026", "ingrediente": "Merluza", "coste_merma": 36.0}, {"fecha": "02/09/2026", "ingrediente": "Pollo", "coste_merma": 50.0}])
    rec = pd.DataFrame(columns=["receta", "ingrediente", "cantidad", "unidad", "coste", "PVP", "categoria"])

    # consumo real agosto = 630 + 200 - 468 = 362 ; ventas agosto = 2000 -> 18,10 %
    r = FB.fc_real_mes("2026-08", inv, ven, mer, ap, {}, coste_escandallo_mes=300.0)
    ok(r["metodo"] == "inventario" and r["consumo_real"] == 362.0, f"metodo {r['metodo']} · consumo real {r['consumo_real']} (630 + 200 − 468)")
    ok(r["ventas_mes"] == 2000.0 and r["fc_real_pct"] == 18.1, f"ventas del mes {r['ventas_mes']} · FC real {r['fc_real_pct']} %")
    ok(r["mermas_mes"] == 36.0 and r["sin_explicar"] == 26.0, f"mermas del mes {r['mermas_mes']} · sin explicar {r['sin_explicar']} (362 − 300 − 36)")
    ok(r["compras_fb"] == 200.0, f"solo compras de proveedores FB: {r['compras_fb']}")
    # septiembre: no hay recuento de septiembre -> aproximacion escandallo + mermas
    r9 = FB.fc_real_mes("2026-09", inv, ven, mer, ap, {}, coste_escandallo_mes=120.0)
    ok(r9["metodo"] == "escandallo+mermas" and r9["consumo_real"] == 170.0 and r9["mermas_mes"] == 50.0, f"sin recuento: {r9['metodo']} · {r9['consumo_real']} (120 + 50)")
    # sin stock inicial no hay metodo inventario
    r_si = FB.fc_real_mes("2026-08", inv.drop(columns=["stock_inicial_kg_l"]), ven, mer, ap, {}, coste_escandallo_mes=300.0)
    ok(r_si["metodo"] == "escandallo+mermas", f"sin stock inicial no se inventa el consumo: {r_si['metodo']}")

    # a traves de resumen_fb (lo que consume el panel): real != teorico
    _, _, res = FB.resumen_fb(rec, inv, ven, mer, mes="2026-08", df_ap=ap)
    ok(res["fc_real_pct"] == 18.1 and res["fc_real_pct"] != res["fc_teorico_pct"], f"resumen_fb: FC real {res['fc_real_pct']} vs teorico {res['fc_teorico_pct']}")
    ok(res.get("fc_real_detalle", {}).get("metodo") == "inventario", "el panel recibe el metodo con el que se calculo")
    _, _, res_auto = FB.resumen_fb(rec, inv, ven, mer, df_ap=ap)
    ok(res_auto["fc_real_detalle"].get("mes") == "2026-08", f"sin mes pedido, coge el mes con mas ventas: {res_auto['fc_real_detalle'].get('mes')}")

    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
    ok("fc_real_detalle" in src and "fb.fcRealAprox" in src, "el panel dice de donde sale el FC real")
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
