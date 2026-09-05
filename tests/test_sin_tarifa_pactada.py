# -*- coding: utf-8 -*-
"""Decision de Jordi (ronda de pruebas, fase 7): una OTA que SI se reconoce en
la factura (Booking, Expedia) pero sin contrato subido es "sin tarifa pactada",
no "OTA no reconocida". Solo cuando ni hay nombre de OTA es desconocida. Y el
detector de DI sigue mirando el certificado de esa OTA aunque no haya tarifa.

  python3.12 tests/test_sin_tarifa_pactada.py
  python3.12 tests/test_sin_tarifa_pactada.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
NF = "NO_ENCONTRADO"


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import verificador_comisiones as V
    import detector_doble_imposicion as DI
    if SABOTAJE:
        _orig = V.buscar_tarifa
        V.buscar_tarifa = lambda df, o, h: (None, "OTA_DESCONOCIDA") if _orig(df, o, h)[1] == "SIN_TARIFA_PACTADA" else _orig(df, o, h)

    com = pd.DataFrame([{"OTA": "Expedia", "Porcentaje_Comision": 18, "Mercado": "Extranjero", "Hotel": ""},
                        {"OTA": "Airbnb", "Porcentaje_Comision": 14, "Mercado": "Extranjero", "Hotel": "Hotel Otro"}])
    com["OTA_norm"] = com["OTA"].map(V._norm); com["Hotel_norm"] = com["Hotel"].map(V._norm)
    fila = lambda ota, hotel=NF: pd.Series({"archivo": "x.pdf", "nombre_ota": ota, "nombre_hotel": hotel, "porcentaje_comision": "15", "importe_bruto": "8000"})
    r_b = V.verificar_factura(fila("Booking.com"), com)
    r_nf = V.verificar_factura(fila(NF), com)
    r_e = V.verificar_factura(fila("Expedia"), com)
    r_a = V.verificar_factura(fila("Airbnb", "Els Pins"), com)
    ok(r_b["estado"] == "SIN_TARIFA_PACTADA", f"Booking reconocida sin contrato → {r_b['estado']}")
    ok(r_nf["estado"] == "OTA_DESCONOCIDA", f"factura sin nombre de OTA → {r_nf['estado']}")
    ok(r_e["estado"] == "DISCREPANCIA" or r_e["estado"] == "COBRO_POR_DEBAJO", f"Expedia con contrato se compara → {r_e['estado']}")
    ok(r_a["estado"] == "SIN_TARIFA_HOTEL", f"Airbnb con tarifa solo de otro hotel → {r_a['estado']}")
    # DI: con SIN_TARIFA_PACTADA el detector clasifica por el nombre y busca el certificado
    d = DI.analizar_factura(pd.Series({**r_b, "estado": "SIN_TARIFA_PACTADA", "mercado": NF}))
    ok(d["estado_di"] != "OTA_DESCONOCIDA" and d["tipo_mercado"] == "extranjera", f"DI de Booking sin tarifa → {d['estado_di']} ({d['tipo_mercado']})")
    d2 = DI.analizar_factura(pd.Series({**r_nf, "estado": "OTA_DESCONOCIDA", "mercado": NF}))
    ok(d2["estado_di"] == "OTA_DESCONOCIDA", "DI de una OTA sin nombre sigue siendo desconocida")
    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
    ok("SIN_TARIFA_PACTADA: ['b-unk', '? Sin tarifa pactada']" in src and "(c.SIN_TARIFA_HOTEL||0) + (c.SIN_TARIFA_PACTADA||0)" in src, "panel: badge y linea 'sin tarifa pactada' con el estado nuevo")
    ok("SIN_TARIFA_PACTADA" in open(os.path.join(BASE, 'provisiones.py'), encoding='utf-8').read(), "provisiones: sin tarifa → provisiona lo facturado, no lo pactado")
    ok("SIN_TARIFA_PACTADA" in open(os.path.join(BASE, 'app_aprobacion.py'), encoding='utf-8').read(), "pantalla de aprobar AR con badge")
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
