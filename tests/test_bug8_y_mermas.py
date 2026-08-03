# -*- coding: utf-8 -*-
"""Bug 8 (la comision del contrato de grupo) + M3 (la fecha de las mermas).

BUG 8 — dos mitades:
  1. la comision se escribia SIN `hotel_id`, asi que caia en "sin asignar":
     no salia en ningun hotel y no se podia aprobar;
  2. ese camino no pedia el paso de cierre, asi que la comision se quedaba
     sin cuenta contable y sin asiento.

M3 — la fecha de las mermas se perdia por dos sitios: `_MER_COL_MAP` no tenia
entrada `fecha` (un fichero con 'Fecha' o 'dia' la tiraba) y ninguna puerta
ponia una por defecto.

  python3.12 tests/test_bug8_y_mermas.py
  python3.12 tests/test_bug8_y_mermas.py --sabotaje
"""
import os
import shutil
import sys
import tempfile
from datetime import date, datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                   # noqa: E402
import censo_hoteles                                  # noqa: E402
import lector_contratos_grupo as L                    # noqa: E402
import dashboard                                      # noqa: E402
from tenant_dirs import procesadas_dir                # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = date.today().strftime('%Y-%m-%d')

DATOS = {
    "es_contrato_grupo": True,
    "evento": {"id": "EV-9", "nombre": "Prueba Bug 8"},
    "contrato_numero": "CG-BUG8",
    "hotel": {"nombre": "Hotel Costa Azul"},
    "cliente": {"nombre": "Cliente Prueba S.A."},
    "agencia": {"nombre": "Agencia Prueba S.L.", "cif": "B11111111"},
    "alojamiento": {"fecha_entrada": "2026-09-01", "fecha_salida": "2026-09-03",
                    "noches": 2, "habitaciones": 10, "tarifa_doble": 100,
                    "total_habitaciones": 2000, "iva_pct": 10},
    "fb": {"total": 500, "pax": 20, "dias": 1},
    "salas": {"total": 300},
    "comisiones": {"alojamiento_pct": 10, "salas_pct": 8, "fb_pct": 5},
    "deposito": {"pct": 30},
    "doble_imposicion": False,
}


def con_censo(hoteles, fn):
    """Corre `fn` con un censo de mentira y un hotel activo."""
    reales_h = censo_hoteles.hoteles
    reales_a = censo_hoteles.activo
    censo_hoteles.hoteles = lambda: hoteles
    censo_hoteles.activo = lambda: (hoteles[0]['id'] if hoteles else '')
    try:
        return fn()
    finally:
        censo_hoteles.hoteles = reales_h
        censo_hoteles.activo = reales_a


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: se quita el hotel de la comision y el cierre ***')
    fallos = 0
    ap_file = os.path.join(procesadas_dir(),
                           'facturas_ap_' + datetime.now().strftime('%Y%m%d') + '.xlsx')
    habia = os.path.exists(ap_file)
    copia = ap_file + '.bak-b8'
    if habia:
        shutil.copy(ap_file, copia)
    dd = tempfile.mkdtemp(prefix='b8_')
    try:
        # ── BUG 8 · 1 · el hotel llega a la fila de la comision ──────
        if os.path.exists(ap_file):
            os.remove(ap_file)
        t = L.transformar(DATOS)
        HOTELES = [{'id': 'HTEST01', 'nombre': 'Hotel de Prueba'},
                   {'id': 'HTEST02', 'nombre': 'Otro Hotel'}]

        def _dist():
            fila = dict(DATOS)
            return L.distribuir_contrato(fila, t, dd)

        r = con_censo(HOTELES, _dist)
        df = pd.read_excel(ap_file) if os.path.exists(ap_file) else pd.DataFrame()
        hid = str(df.iloc[0].get('hotel_id', '')) if len(df) else ''
        if SABOTAJE:
            hid = ''            # como estaba antes del arreglo
        ok_hotel = (len(df) == 1 and hid == 'HTEST01')
        print(f"  {'OK ' if ok_hotel else 'FALLA'}  la comision lleva el hotel: {hid!r} "
              f"(esperaba 'HTEST01')")
        if not ok_hotel:
            fallos += 1

        # ── BUG 8 · 1b · nada mas se ha movido ──────────────────────
        f = df.iloc[0].to_dict() if len(df) else {}
        ok_datos = (str(f.get('numero_factura')) == 'COM-CG-BUG8'
                    and round(float(f.get('total_factura', 0)), 2) == round(r['ap'], 2)
                    and str(f.get('tipo')) == 'COMISION_AGENCIA'
                    and str(f.get('nombre_proveedor')) == 'Agencia Prueba S.L.')
        print(f"  {'OK ' if ok_datos else 'FALLA'}  el resto de la fila no se mueve "
              f"({f.get('numero_factura')} · {f.get('total_factura')} € · {f.get('tipo')})")
        if not ok_datos:
            fallos += 1

        # ── BUG 8 · 2 · el contrato PIDE el cierre ──────────────────
        real = L.extraer_contrato_grupo
        L.extraer_contrato_grupo = lambda paths: dict(DATOS, _needs_review=False)
        try:
            if os.path.exists(ap_file):
                os.remove(ap_file)
            res = con_censo(HOTELES, lambda: L.procesar_contrato_grupo(
                [os.path.join(dd, 'x.jpg')], datos_dir=dd))
        finally:
            L.extraer_contrato_grupo = real
        cierre = list(res.get('cierre') or [])
        if SABOTAJE:
            cierre = []
        ok_cierre = ('ap' in cierre and 'ar' in cierre)
        print(f"  {'OK ' if ok_cierre else 'FALLA'}  el contrato pide el cierre: {cierre}")
        if not ok_cierre:
            fallos += 1

        # ── M3 · la fecha de las mermas ─────────────────────────────
        ruta_mer = os.path.join(dashboard._ddir(), 'mermas.xlsx')
        habia_m = os.path.exists(ruta_mer)
        copia_m = ruta_mer + '.bak-m3'
        if habia_m:
            shutil.copy(ruta_mer, copia_m)
        try:
            # a · el fichero TRAE fecha con otro nombre: se conserva
            if os.path.exists(ruta_mer):
                os.remove(ruta_mer)
            df_in = dashboard._normalize_cols(
                pd.DataFrame([{'Producto': 'Merluza', 'dia': '2026-07-01',
                               'cantidad': 2, 'coste': 30}]), dashboard._MER_COL_MAP)
            tiene_col = 'fecha' in df_in.columns
            con_censo(HOTELES, lambda: dashboard._guardar_fb_del_hotel(df_in, 'mermas.xlsx'))
            g = pd.read_excel(ruta_mer)
            suya = str(g.iloc[0]['fecha'])[:10] if 'fecha' in g.columns else ''
            ok_suya = tiene_col and suya == '2026-07-01'
            print(f"  {'OK ' if ok_suya else 'FALLA'}  merma con 'dia' en el fichero: "
                  f"conserva su fecha ({suya!r})")
            if not ok_suya:
                fallos += 1

            # b · el fichero NO trae fecha: se pone la de hoy
            os.remove(ruta_mer)
            df_in2 = dashboard._normalize_cols(
                pd.DataFrame([{'Producto': 'Tomate', 'cantidad': 1, 'coste': 4}]),
                dashboard._MER_COL_MAP)
            con_censo(HOTELES, lambda: dashboard._guardar_fb_del_hotel(df_in2, 'mermas.xlsx'))
            g2 = pd.read_excel(ruta_mer)
            hoy_ok = 'fecha' in g2.columns and str(g2.iloc[0]['fecha'])[:10] == HOY
            print(f"  {'OK ' if hoy_ok else 'FALLA'}  merma SIN fecha: se pone la de hoy "
                  f"({str(g2.iloc[0].get('fecha'))[:10] if 'fecha' in g2.columns else 'sin columna'})")
            if not hoy_ok:
                fallos += 1

            # c · las VENTAS no se tocan (solo era para mermas)
            ruta_v = os.path.join(dashboard._ddir(), 'ventas_fb_diarias.xlsx')
            habia_v = os.path.exists(ruta_v)
            copia_v = ruta_v + '.bak-m3'
            if habia_v:
                shutil.copy(ruta_v, copia_v)
            if os.path.exists(ruta_v):
                os.remove(ruta_v)
            con_censo(HOTELES, lambda: dashboard._guardar_fb_del_hotel(
                pd.DataFrame([{'nombre_plato': 'Paella', 'unidades_vendidas': 2,
                               'total_venta': 40}]), 'ventas_fb_diarias.xlsx'))
            gv = pd.read_excel(ruta_v)
            ok_ventas = 'fecha' not in gv.columns
            print(f"  {'OK ' if ok_ventas else 'FALLA'}  las ventas NO ganan fecha "
                  f"(el arreglo es solo de mermas)")
            if not ok_ventas:
                fallos += 1
            os.remove(ruta_v)
            if habia_v:
                shutil.move(copia_v, ruta_v)
        finally:
            if os.path.exists(ruta_mer):
                os.remove(ruta_mer)
            if habia_m:
                shutil.move(copia_m, ruta_mer)
    finally:
        if os.path.exists(ap_file):
            os.remove(ap_file)
        if habia:
            shutil.move(copia, ap_file)
        shutil.rmtree(dd, ignore_errors=True)

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: {fallos} en rojo.')
            return 0
        print('SABOTAJE MAL.')
        return 1
    if fallos:
        print(f'{fallos} en rojo')
        return 1
    print('Todo OK. La comision lleva hotel y pide cierre; las mermas llevan fecha.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
