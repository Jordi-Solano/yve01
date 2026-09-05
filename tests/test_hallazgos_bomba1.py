# -*- coding: utf-8 -*-
"""Los tres hallazgos apuntados al cerrar la bomba 1 (sep 2026):
  (a) el tile/tabla AP cruza aprobaciones por la CLAVE (numero o fichero): una
      factura sin numero aprobada ya cuenta
  (b) /aprobaciones-ap escribe donde lee el dashboard (rutas por tenant)
  (c) _audit apunta al usuario real, no "sistema"

  python3.12 tests/test_hallazgos_bomba1.py
  python3.12 tests/test_hallazgos_bomba1.py --sabotaje
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

SABOTAJE = '--sabotaje' in sys.argv
PROC = os.path.join(BASE, 'facturas-procesadas')
RAW = os.path.join(PROC, 'facturas_ap_20260103.xlsx')
APR = os.path.join(BASE, 'aprobaciones', 'aprobaciones_ap.xlsx')
AUD = os.path.join(BASE, 'datos-referencia', 'audit_log.json')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    import app_aprobacion_ap as PANEL
    if SABOTAJE:
        D.safe_str = lambda v: "" if v is None or str(v) in ("nan", "None") else str(v)   # inocuo...
        PANEL._apro_file = lambda: os.path.join(tempfile.gettempdir(), 'otro_sitio_aprobaciones.xlsx')   # ...pero el panel escribe en otro sitio

    tmp = tempfile.mkdtemp(prefix='b1_'); copias = {}
    for f in (RAW, APR, AUD):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        for f in (RAW, APR, AUD):
            if os.path.exists(f):
                os.remove(f)
        os.makedirs(PROC, exist_ok=True)
        pd.DataFrame([
            {"archivo": "con_numero.pdf", "numero_factura": "B1-TEST-1", "fecha": "05/08/2026", "nombre_proveedor": "Uno SL", "base_imponible": 100.0, "porcentaje_iva": 21, "cuota_iva": 21.0, "total_factura": 121.0, "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "hotel_id": ""},
            {"archivo": "sin_numero.pdf", "numero_factura": "", "fecha": "06/08/2026", "nombre_proveedor": "Dos SL", "base_imponible": 200.0, "porcentaje_iva": 21, "cuota_iva": 42.0, "total_factura": 242.0, "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "hotel_id": ""},
        ]).to_excel(RAW, index=False)
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        lst = cl.get('/aprobaciones-ap/api/facturas?estado=pendientes').get_json() or []
        sin = next((x for x in lst if x.get("archivo") == "sin_numero.pdf" or x.get("clave") == "sin_numero.pdf"), None)
        ok(sin is not None and sin.get("clave") == "sin_numero.pdf", f"la factura sin numero tiene clave = fichero ({sin and sin.get('clave')})")
        # (b)+(c): aprobar desde /aprobaciones-ap escribe donde lee el dashboard, y el audit lleva al usuario
        for f in (sin, next((x for x in lst if x.get("numero_factura") == "B1-TEST-1"), None)):
            r = cl.post('/aprobaciones-ap/api/accion', json={"clave": f["clave"], "numero_factura": f.get("numero_factura") or f["clave"], "accion": "APROBADA", "comentario": "ok", "departamento": "Administracion"}, headers=H)
            ok(r.status_code == 200, f"aprobar {f['clave']} → {r.status_code}")
        ok(os.path.exists(APR), "(b) el registro de aprobaciones esta en aprobaciones/ del tenant (donde lee el dashboard)")
        stats = cl.get('/api/stats_ap').get_json() or {}
        rows = cl.get('/api/facturas_ap').get_json()
        rows = rows if isinstance(rows, list) else rows.get('facturas', [])
        acc = {x.get("clave"): x.get("accion") for x in rows if x.get("clave") in ("B1-TEST-1", "sin_numero.pdf")}
        ok(acc.get("sin_numero.pdf") == "APROBADA" and acc.get("B1-TEST-1") == "APROBADA", f"(a) el panel AP ve aprobadas las DOS, tambien la sin numero: {acc}")
        n_apro = stats.get("aprobadas", stats.get("n_aprobadas"))
        ok(n_apro is None or int(n_apro) >= 2, f"(a) el tile cuenta {n_apro} aprobadas (>= 2)")
        # (c) audit con usuario real
        D._audit("PRUEBA_B1", "desde fuera de una peticion")
        ent = json.load(open(AUD, encoding='utf-8')) if os.path.exists(AUD) else []
        ok(ent and ent[-1]["accion"] == "PRUEBA_B1" and ent[-1]["usuario"] == "sistema", "(c) fuera de una peticion sigue siendo 'sistema'")
        CFGB = os.path.join(BASE, 'datos-referencia', 'config_banco.json')
        cfg_prev = open(CFGB, encoding='utf-8').read() if os.path.exists(CFGB) else None
        try:
            cl.get('/api/config_banco')
            cur = (cl.get('/api/config_banco').get_json() or {}).get('modo') or 'grupo'
            r = cl.post('/api/config_banco', json={"modo": cur}, headers=H)     # llama a _audit sin usuario explicito
        finally:
            if cfg_prev is None:
                if os.path.exists(CFGB):
                    os.remove(CFGB)
            else:
                open(CFGB, 'w', encoding='utf-8').write(cfg_prev)
        ent = json.load(open(AUD, encoding='utf-8')) if os.path.exists(AUD) else []
        ult = [e for e in ent if e["accion"] == "BANCO_CONFIG"]
        ok(bool(ult) and ult[-1]["usuario"] == "admin", f"(c) dentro de una peticion, _audit apunta al usuario real: {ult and ult[-1]['usuario']}")
        ok(os.path.dirname(os.path.abspath(AUD)) == os.path.abspath(os.path.join(BASE, 'datos-referencia')), "el audit_log vive en datos-referencia del tenant")
    finally:
        for f in (RAW, APR, AUD):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True)
        p2 = os.path.join(tempfile.gettempdir(), 'otro_sitio_aprobaciones.xlsx')
        if os.path.exists(p2):
            os.remove(p2)
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
