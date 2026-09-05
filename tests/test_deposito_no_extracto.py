# -*- coding: utf-8 -*-
"""El deposito previsto de un contrato de grupo se escribia como FILA DEL
EXTRACTO bancario: la pestaña Banco enseñaba un ingreso que el banco no habia
hecho (inventario honesto #14). Ahora va a depositos_previstos.xlsx y la
pestaña Banco lo avisa como "previsto, aun no en el extracto".

  python3.12 tests/test_deposito_no_extracto.py
  python3.12 tests/test_deposito_no_extracto.py --sabotaje
"""
import os, shutil, subprocess, sys, tempfile
from datetime import datetime
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
import pandas as pd                                   # noqa: E402
import censo_hoteles                                  # noqa: E402
import lector_contratos_grupo as L                    # noqa: E402
SABOTAJE = '--sabotaje' in sys.argv
DATOS = {"es_contrato_grupo": True, "evento": {"id": "EV-D", "nombre": "Boda Puig"}, "contrato_numero": "CG-DEP",
         "hotel": {"nombre": "Hotel Els Pins"}, "cliente": {"nombre": "Familia Puig"}, "agencia": {"nombre": "Viatges Test SL", "cif": "B1"},
         "alojamiento": {"fecha_entrada": "2026-09-10", "fecha_salida": "2026-09-12", "noches": 2, "habitaciones": 10, "tarifa_doble": 100, "total_habitaciones": 2000, "iva_pct": 10},
         "fb": {"total": 500, "pax": 20, "dias": 1}, "salas": {"total": 300}, "comisiones": {"alojamiento_pct": 10, "salas_pct": 8, "fb_pct": 5},
         "deposito": {"pct": 30}, "doble_imposicion": False}


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    from tenant_dirs import procesadas_dir, datos_dir
    ap_file = os.path.join(procesadas_dir(), 'facturas_ap_' + datetime.now().strftime('%Y%m%d') + '.xlsx')
    DEP = os.path.join(datos_dir(), 'depositos_previstos.xlsx'); EXT = os.path.join(datos_dir(), 'extracto_banco.xlsx')
    tmp = tempfile.mkdtemp(prefix='dep_'); copias = {}
    for f in (ap_file, DEP, EXT):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    dd = tempfile.mkdtemp(prefix='depdd_')
    ra, rh = censo_hoteles.activo, censo_hoteles.hoteles
    censo_hoteles.hoteles = lambda: [{'id': 'HTEST01', 'nombre': 'Hotel Els Pins'}]; censo_hoteles.activo = lambda: 'HTEST01'
    try:
        for f in (ap_file, DEP, EXT):
            if os.path.exists(f): os.remove(f)
        t = L.transformar(dict(DATOS))
        if SABOTAJE:
            # vuelve a escribir el deposito en el extracto, como antes
            L._append_xlsx(os.path.join(dd, "extracto_banco.xlsx"), {"fecha": "2026-09-05", "concepto": "Depósito previsto 30% · Boda Puig", "importe": 999.0}, dedup_col="concepto")
        res = L.distribuir_contrato(dict(DATOS), t, dd)
        ext = pd.read_excel(os.path.join(dd, 'extracto_banco.xlsx')) if os.path.exists(os.path.join(dd, 'extracto_banco.xlsx')) else pd.DataFrame()
        ok(ext.empty or not ext['concepto'].astype(str).str.contains('Depósito previsto').any(), "el extracto NO lleva el deposito previsto")
        dep = pd.read_excel(os.path.join(dd, 'depositos_previstos.xlsx')) if os.path.exists(os.path.join(dd, 'depositos_previstos.xlsx')) else pd.DataFrame()
        ok(len(dep) == 1 and float(dep.iloc[0]['importe']) == res.get('banco') and str(dep.iloc[0]['estado']) == 'PREVISTO' and str(dep.iloc[0]['hotel_id']) == 'HTEST01',
           f"depositos_previstos.xlsx: 1 fila, {res.get('banco')} €, PREVISTO, con hotel")
        # la pestaña Banco lo avisa (por el tenant real)
        import dashboard as D
        shutil.copy(os.path.join(dd, 'depositos_previstos.xlsx'), DEP)
        cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
        pd.DataFrame([{'fecha': '05/09/2026', 'concepto': 'OTRO', 'importe': 12.0, 'saldo': 12.0}]).to_excel(EXT, index=False)
        d = cl.get('/api/stats_banco').get_json() or {}
        deps = d.get('depositos_previstos') or []
        ok(len(deps) == 1 and deps[0]['estado'] == 'PREVISTO', f"/api/stats_banco lista el deposito como PREVISTO ({deps})")
        # cuando el banco lo ingresa (mismo importe), pasa a EN_EXTRACTO
        pd.DataFrame([{'fecha': '05/09/2026', 'concepto': 'TRANSFERENCIA PUIG', 'importe': float(dep.iloc[0]['importe']), 'saldo': 12.0}]).to_excel(EXT, index=False)
        d = cl.get('/api/stats_banco').get_json() or {}
        ok((d.get('depositos_previstos') or [{}])[0].get('estado') == 'EN_EXTRACTO', "con el ingreso real en el extracto pasa a EN_EXTRACTO")
        html = cl.get('/').get_data(as_text=True)
        ok('bk.depPrevisto' in html and 'depositos_previstos' in html, "la pestaña Banco pinta el aviso")
    finally:
        censo_hoteles.activo, censo_hoteles.hoteles = ra, rh
        for f in (ap_file, DEP, EXT):
            if os.path.exists(f): os.remove(f)
            if f in copias: shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True); shutil.rmtree(dd, ignore_errors=True)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
