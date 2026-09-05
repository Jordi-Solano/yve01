# -*- coding: utf-8 -*-
"""base + IVA = total al leer una factura AP; si no cuadra se avisa (Jordi, sep 2026).

  python3.12 tests/test_control_importes.py
  python3.12 tests/test_control_importes.py --sabotaje
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
PROC = os.path.join(BASE, 'facturas-procesadas')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import control_importes as CI
    if SABOTAJE:
        CI.comprobar = lambda f: (True, "")

    ok(CI.comprobar({"base_imponible": 1000, "cuota_iva": 210, "total_factura": 1210, "porcentaje_iva": 21}) == (True, ""), "1000 + 210 = 1210 cuadra")
    c, a = CI.comprobar({"base_imponible": 1000, "cuota_iva": 210, "total_factura": 1250, "porcentaje_iva": 21})
    ok(c is False and "1.250,00" in a and "40,00" in a, f"total inflado: {a}")
    c, a = CI.comprobar({"base_imponible": "1.000,00", "cuota_iva": "100,00", "total_factura": "1.100,00", "porcentaje_iva": "21"})
    ok(c is False and "21 %" in a and "210,00" in a, f"IVA mal calculado (texto español): {a}")
    ok(CI.comprobar({"base_imponible": 971.74, "porcentaje_iva": 21, "total_factura": 1175.8}) == (True, ""), "sin cuota: se deduce del porcentaje (971,74 × 21 % → 1.175,80)")
    ok(CI.comprobar({"base_imponible": None, "total_factura": 100})[0] is None, "sin base no se puede comprobar (None, no falso)")
    ok(CI.comprobar({"base_imponible": 100.004, "cuota_iva": 21.0, "total_factura": 121.0})[0] is True, "redondeo de centimos tolerado")
    filas = [{"numero_factura": "A", "base_imponible": 100, "cuota_iva": 21, "total_factura": 121},
             {"numero_factura": "B", "base_imponible": 100, "cuota_iva": 21, "total_factura": 130}]
    mal = CI.anotar(filas)
    ok([f["numero_factura"] for f in mal] == ["B"] and filas[0]["importes_cuadran"] == "SI" and filas[1]["importes_cuadran"] == "NO", "anotar marca SI/NO en cada factura")

    # por el punto unico de guardado, y hasta el panel y la pantalla de aprobar
    tmp = tempfile.mkdtemp(prefix='ci_'); hoy = None
    import dashboard as D
    from datetime import date
    ruta = os.path.join(PROC, f'facturas_ap_{date.today().strftime("%Y%m%d")}.xlsx')
    copia = None
    if os.path.exists(ruta):
        copia = os.path.join(tmp, 'x.xlsx'); shutil.copy(ruta, copia)
    # lo que el guardado APRENDE (proveedores_aprendidos.json) no debe quedar en el checkout
    APR = os.path.join(BASE, 'datos-referencia', 'proveedores_aprendidos.json')
    apr_copia = None
    if os.path.exists(APR):
        apr_copia = os.path.join(tmp, 'apr.json'); shutil.copy(APR, apr_copia)
    try:
        D._guardar_factura_ap([{"archivo": "ci_test.pdf", "numero_factura": "CI-TEST-9", "fecha": "10/08/2026", "nombre_proveedor": "Neteges Test SL",
                                "base_imponible": 500.0, "porcentaje_iva": 21, "cuota_iva": 105.0, "total_factura": 650.0, "tipo_proveedor": "OTRAS",
                                "cuenta_contable": "629", "descripcion_concepto": "Limpieza"}])
        g = pd.read_excel(ruta)
        f = g[g["numero_factura"] == "CI-TEST-9"].iloc[0]
        ok(str(f.get("importes_cuadran")) == "NO" and "605,00" in str(f.get("aviso_importes")), f"guardada con aviso: {f.get('aviso_importes')}")
        ok("⚠" in D._resumen_factura_ap([f.to_dict()]) and "CI-TEST-9" in D._resumen_factura_ap([f.to_dict()]), f"el log lo dice: {D._resumen_factura_ap([f.to_dict()])[:90]}")
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        rows = cl.get('/api/facturas_ap').get_json()
        rows = rows if isinstance(rows, list) else rows.get('facturas', [])
        r = next((x for x in rows if x.get("numero_factura") == "CI-TEST-9"), None)
        ok(r is not None and r.get("importes_cuadran") == "NO", "el panel AP recibe importes_cuadran=NO")
        html = cl.get('/').get_data(as_text=True)
        ok("base + IVA ≠ total" in html, "badge '⚠ base + IVA ≠ total' en la tabla AP")
        lst = cl.get('/aprobaciones-ap/api/facturas?estado=pendientes').get_json() or []
        a = next((x for x in lst if x.get("numero_factura") == "CI-TEST-9"), None)
        ok(a is not None and a.get("importes_cuadran") == "NO", "la pantalla de aprobar lleva el aviso")
        h2 = cl.get('/aprobaciones-ap/').get_data(as_text=True)
        ok("Los importes no cuadran" in h2, "y lo pinta antes de aprobar")
        for pg in ('/', '/aprobaciones-ap/'):
            html = cl.get(pg).get_data(as_text=True)
            for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
                open('/tmp/_ci.js', 'w', encoding='utf-8').write(b)
                rc = subprocess.run(['node', '--check', '/tmp/_ci.js'], capture_output=True, text=True)
                if rc.returncode:
                    ok(False, f"JS roto en {pg}: {rc.stderr[:100]}"); break
    finally:
        if os.path.exists(ruta):
            os.remove(ruta)
        if copia:
            shutil.copy(copia, ruta)
        if os.path.exists(APR):
            os.remove(APR)
        if apr_copia:
            shutil.copy(apr_copia, APR)
        shutil.rmtree(tmp, ignore_errors=True)
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
