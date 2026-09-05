# -*- coding: utf-8 -*-
"""Hallazgo (j) de la Ola B: con el demo, las comisiones OTA daban 0 asientos.
Causa: el demo no escribia `importe_comision_factura` (lo que escribe el
verificador real). Ahora el demo se parece al informe real y el cierre y el
fiscal lo ven.

  python3.12 tests/test_demo_ota_asientos.py
  python3.12 tests/test_demo_ota_asientos.py --sabotaje
"""
import glob
import os
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

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

    import demo_generator as DG
    import cierre_mes as CM
    import fiscal as FI
    tmp = tempfile.mkdtemp(prefix='demo_')
    for d in ('datos', 'reportes', 'proc'):
        os.makedirs(os.path.join(tmp, d), exist_ok=True)
    # el demo escribe en carpetas temporales, no en el checkout
    DG.DATOS = DG._TDir(lambda: os.path.join(tmp, 'datos'))
    DG.REPORTES = DG._TDir(lambda: os.path.join(tmp, 'reportes'))
    DG.PROCESADAS = DG._TDir(lambda: os.path.join(tmp, 'proc'))
    try:
        DG.generar_demo([{"nombre": "Grupo Prueba", "hoteles": ["Hotel Prueba Mar"]}])
    except Exception as e:
        ok(False, f"generar_demo: {e}")
        sys.exit(1)
    fs = sorted(glob.glob(os.path.join(tmp, 'reportes', 'verificacion_*.xlsx')))
    ok(bool(fs), f"el demo escribe el informe de verificacion ({len(fs)})")
    df = pd.read_excel(fs[-1]) if fs else pd.DataFrame()
    if SABOTAJE:
        df = df.drop(columns=[c for c in ("importe_comision_factura", "importe_comision") if c in df.columns])
    ok("importe_comision_factura" in df.columns and "porcentaje_comision" in df.columns, f"columnas del verificador real presentes: {[c for c in df.columns if 'comision' in c]}")
    imp = pd.to_numeric(df.get("importe_comision_factura", pd.Series(dtype=float)), errors="coerce").fillna(0)
    ok(len(df) and (imp > 0).all(), f"todas las facturas OTA del demo traen importe de comision > 0 ({int((imp > 0).sum())}/{len(df)})")
    # el cierre del mes actual las ve (el demo fecha entre hace 2 y 55 dias: al menos alguna cae en el mes)
    from provisiones import _fecha, _mes_a_rango
    # el mes con mas facturas del demo (fecha entre hace 2 y 55 dias)
    meses = pd.Series([(lambda f: f"{f.year:04d}-{f.month:02d}" if f else "")(_fecha(v)) for v in df.get("fecha", [])])
    mes = str(meses[meses != ""].value_counts().idxmax()) if len(meses) and (meses != "").any() else f"{date.today().year:04d}-{date.today().month:02d}"
    ini, fin, _ = _mes_a_rango(mes)
    n_mes = sum(1 for v in df.get("fecha", []) if (lambda f: f is not None and ini <= f <= fin)(_fecha(v)))
    fu = {"ap": pd.DataFrame(), "ar_ota": df, "ventas_fb": pd.DataFrame(), "reservas": pd.DataFrame(), "banco": pd.DataFrame(), "provisiones": []}
    res = CM.generar_asientos(mes, fu)
    ok(n_mes > 0 and res["fuentes"]["ar_ota"] == n_mes, f"cierre {mes}: {res['fuentes']['ar_ota']} asientos de comisiones OTA (en el mes hay {n_mes}), saltados sin importe: {res.get('saltados', {}).get('ar_ota_sin_importe')}")
    fi = FI.calcular(mes, fu, CM.config_cierre('/nonexistent'), {'nif': {}, 'periodicidad': 'mensual'})
    ok(fi["sii"]["n_recibidas"] == n_mes, f"fiscal: {fi['sii']['n_recibidas']} recibidas OTA en el SII")
    import shutil; shutil.rmtree(tmp, ignore_errors=True)
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
