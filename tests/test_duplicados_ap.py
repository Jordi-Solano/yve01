# -*- coding: utf-8 -*-
"""Duplicados AP (decision de Jordi, ronda 1): dos documentos con el mismo
numero+proveedor bloquean la aprobacion hasta que alguien elige cual vale
desde /aprobaciones-ap/; la otra queda descartada y fuera de aging y Oracle.

  python3.12 tests/test_duplicados_ap.py
  python3.12 tests/test_duplicados_ap.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402
import almacen_datos as A      # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
PROC = os.path.join(BASE, 'facturas-procesadas')
RAW = os.path.join(PROC, 'facturas_ap_20260101.xlsx')
DUP = os.path.join(BASE, 'datos-referencia', A.DUPLICADOS_FILE)
APR = os.path.join(BASE, 'aprobaciones', 'aprobaciones_ap.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        A.duplicados_resueltos = lambda *a, **k: {}     # la eleccion no se aplica nunca

    tmp = tempfile.mkdtemp(prefix='dup_')
    copias = {}
    for f in (RAW, DUP, APR):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        if os.path.exists(DUP):
            os.remove(DUP)
        os.makedirs(PROC, exist_ok=True)
        pd.DataFrame([
            {"archivo": "nc_original.pdf", "numero_factura": "NC-TEST-118", "fecha": "05/08/2026", "nombre_proveedor": "Neteges Test SL",
             "base_imponible": 240.0, "porcentaje_iva": 21, "cuota_iva": 50.4, "total_factura": 290.4, "tipo_proveedor": "OTRAS",
             "cuenta_contable": "629", "descripcion_concepto": "Limpieza agosto", "hotel_id": ""},
            {"archivo": "nc_copia.pdf", "numero_factura": "NC-TEST-118", "fecha": "05/08/2026", "nombre_proveedor": "Neteges Test SL",
             "base_imponible": 260.0, "porcentaje_iva": 21, "cuota_iva": 54.6, "total_factura": 314.6, "tipo_proveedor": "OTRAS",
             "cuenta_contable": "629", "descripcion_concepto": "Limpieza agosto (reemitida)", "hotel_id": ""},
            {"archivo": "iv.pdf", "numero_factura": "IV-TEST-733", "fecha": "20/08/2026", "nombre_proveedor": "Vila Test SL",
             "base_imponible": 100.0, "porcentaje_iva": 21, "cuota_iva": 21.0, "total_factura": 121.0, "tipo_proveedor": "OTRAS",
             "cuenta_contable": "622", "descripcion_concepto": "Reparacion", "hotel_id": ""},
        ]).to_excel(RAW, index=False)

        df = A.facturas_ap(PROC, os.path.join(BASE, 'reportes'))
        fila = df[df["numero_factura"] == "NC-TEST-118"]
        ok(len(fila) == 1 and int(fila.iloc[0]["duplicados"]) == 2, f"almacen: UNA fila NC-TEST-118 con duplicados={fila.iloc[0]['duplicados'] if len(fila) else '?'}")
        grupos = A.facturas_ap_duplicadas(PROC, os.path.join(BASE, 'reportes'))
        g = next((x for x in grupos if x["numero_factura"] == "NC-TEST-118"), None)
        ok(g is not None and len(g["documentos"]) == 2 and {d["archivo"] for d in g["documentos"]} == {"nc_original.pdf", "nc_copia.pdf"},
           f"grupo de duplicados con los DOS documentos a la vista: {g and [d['archivo'] for d in g['documentos']]}")
        ok(not any(x["numero_factura"] == "IV-TEST-733" for x in grupos), "la factura normal no sale como duplicada")

        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        lst = cl.get('/aprobaciones-ap/api/facturas?estado=pendientes').get_json() or []
        f = next((x for x in lst if x.get("numero_factura") == "NC-TEST-118"), None)
        ok(f is not None and f.get("duplicados") == 2, f"pantalla de aprobar: la fila lleva duplicados={f and f.get('duplicados')}")
        r = cl.post('/aprobaciones-ap/api/accion', json={"clave": f["clave"], "numero_factura": "NC-TEST-118", "accion": "APROBADA", "comentario": "ok", "departamento": "Administracion"}, headers=H)
        ok(r.status_code == 409 and (r.get_json() or {}).get("duplicado"), f"aprobar un duplicado → 409 ({(r.get_json() or {}).get('error', '')[:60]})")
        r = cl.post('/aprobaciones-ap/api/accion', json={"clave": f["clave"], "numero_factura": "NC-TEST-118", "accion": "RECHAZADA", "comentario": "no", "departamento": "Administracion"}, headers=H)
        ok(r.status_code == 409, "rechazar un duplicado → 409 tambien")
        rl = cl.post('/api/ap/aprobar_lote', json={"facturas": [f["clave"]]}, headers=H)
        dl = rl.get_json() or {}
        ok(rl.status_code in (200, 400) and not dl.get("aprobadas"), f"aprobar en lote lo salta: {dl}")
        html = cl.get('/aprobaciones-ap/').get_data(as_text=True)
        ok('id="duplicados"' in html and 'Esta es la buena' in html and 'function cargarDuplicados' in html, "tarjeta 'Duplicados por resolver' con boton 'Esta es la buena'")
        d = cl.get('/aprobaciones-ap/api/duplicados').get_json() or {}
        ok(d.get("ok") and any(x["numero_factura"] == "NC-TEST-118" for x in d.get("grupos", [])), f"/api/duplicados lista el grupo (n={d.get('n')})")
        aging_antes = cl.get('/api/aging_ap').get_json() if cl.get('/api/aging_ap').status_code == 200 else None

        r = cl.post('/aprobaciones-ap/api/duplicados/elegir', json={"clave": g["clave"], "archivo": "nope.pdf"}, headers=H)
        ok(r.status_code == 404, "elegir un fichero que no esta en el grupo → 404")
        r = cl.post('/aprobaciones-ap/api/duplicados/elegir', json={"clave": g["clave"], "archivo": "nc_original.pdf"}, headers=H)
        ok(r.status_code == 200 and (r.get_json() or {}).get("buena") == "nc_original.pdf", "'Esta es la buena' guarda la eleccion")
        ok(os.path.exists(DUP) and json.load(open(DUP, encoding='utf-8')).get(g["clave"], {}).get("descartadas") == ["nc_copia.pdf"], "duplicados_resueltos.json con la descartada")

        df2 = A.facturas_ap(PROC, os.path.join(BASE, 'reportes'))
        fila2 = df2[df2["numero_factura"] == "NC-TEST-118"]
        ok(len(fila2) == 1 and fila2.iloc[0]["archivo"] == "nc_original.pdf" and float(fila2.iloc[0]["total_factura"]) == 290.4 and int(fila2.iloc[0]["duplicados"]) == 0,
           f"tras elegir: queda la buena (290,40) y ya no esta marcada como duplicada → {fila2.iloc[0]['archivo'] if len(fila2) else '?'} / dup={fila2.iloc[0]['duplicados'] if len(fila2) else '?'}")
        ok(not A.facturas_ap_duplicadas(PROC, os.path.join(BASE, 'reportes')), "el grupo desaparece de 'Duplicados por resolver'")
        # la descartada no esta en lo que lee Oracle ni en el aging (todos leen por almacen)
        import oracle_lector_facturas as OL
        try:
            dfo = OL.cargar_facturas_contabilizadas()
            dfo = dfo[0] if isinstance(dfo, tuple) else dfo
            ok(not ((dfo["numero_factura"].astype(str) == "NC-TEST-118") & (dfo["archivo"].astype(str) == "nc_copia.pdf")).any() if not dfo.empty and "archivo" in dfo.columns else True,
               "Oracle no ve la descartada")
        except Exception as e:
            ok(False, f"Oracle lector: {e}")
        ag = cl.get('/api/aging_ap').get_json() if cl.get('/api/aging_ap').status_code == 200 else {}
        txt_ag = json.dumps(ag)
        ok('314.6' not in txt_ag and ('290.4' in txt_ag or not ag), "aging: la descartada (314,60) ya no cuenta")
        r = cl.post('/aprobaciones-ap/api/accion', json={"clave": f["clave"], "numero_factura": "NC-TEST-118", "accion": "APROBADA", "comentario": "ok", "departamento": "Administracion"}, headers=H)
        ok(r.status_code == 200, f"ahora si se puede aprobar la buena ({r.status_code})")
    finally:
        for f in (RAW, DUP, APR):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
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
