# -*- coding: utf-8 -*-
"""Un extracto con cabeceras en mayusculas (Fecha/Concepto/Importe/Saldo) o en
ingles (Date/Description/Amount/Balance) tiene que acabar en las columnas
canonicas: si no, la conciliacion y el cuadre lo ven como filas vacias.

  python3.12 tests/test_banco_cabeceras.py
  python3.12 tests/test_banco_cabeceras.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
BANCO = os.path.join(BASE, 'datos-referencia', 'extracto_banco.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    if SABOTAJE:
        D._normalize_cols = lambda df, m: df       # vuelve a dejar las cabeceras tal cual
    tmp = tempfile.mkdtemp(prefix='bk_')
    existia = os.path.exists(BANCO)
    if existia:
        shutil.copy(BANCO, os.path.join(tmp, 'b.xlsx'))
    try:
        if existia:
            os.remove(BANCO)
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        casos = {
            'extracto_mayusculas.xlsx': pd.DataFrame([{'Fecha': '05/08/2026', 'Concepto': 'RECIBO LUZ FRA EL-1', 'Importe': -100.0, 'Saldo': 900.0}]),
            'bank_statement_en.xlsx': pd.DataFrame([{'Date': '06/08/2026', 'Description': 'CARD SETTLEMENT', 'Amount': 250.0, 'Balance': 1150.0}]),
        }
        for nombre, df in casos.items():
            ruta = os.path.join(tmp, nombre); df.to_excel(ruta, index=False)
            with open(ruta, 'rb') as fh:
                cl.post('/api/upload_facturas', data={'files': [(fh, nombre)]}, content_type='multipart/form-data')
            r = cl.get('/api/procesar_batch_stream?archivos=' + urllib.parse.quote(json.dumps([nombre])))
            ok('✓ Banco' in r.get_data(as_text=True), f'{nombre}: la capa 1 lo integra como banco')
            r.close()      # el generador SSE guarda el log al cerrarse: si no, reescribe despues de la limpieza
        import almacen_datos as ALM
        bk, info = ALM.movimientos_banco(datos_dir=os.path.join(BASE, 'datos-referencia'), reportes_dir=os.path.join(BASE, 'reportes'))
        ok(set(['fecha', 'concepto', 'importe']) <= set(bk.columns) and 'Fecha' not in bk.columns and 'Amount' not in bk.columns,
           f'columnas canonicas y ninguna cruda: {list(bk.columns)[:8]}')
        ok(len(bk) == 2 and bk['importe'].notna().all() and bk['concepto'].notna().all(), f'2 movimientos con importe y concepto: {bk[["fecha", "concepto", "importe"]].to_dict("records")}')
        import cuadre_banco as CB
        c = CB.cuadrar('2026-08', bk)
        ok(c['n'] == 2, f'el cuadre de agosto ve {c["n"]} movimientos (esperados 2)')
    finally:
        if os.path.exists(BANCO):
            os.remove(BANCO)
        if existia:
            shutil.copy(os.path.join(tmp, 'b.xlsx'), BANCO)
        shutil.rmtree(tmp, ignore_errors=True)
        for f in os.listdir(os.path.join(BASE, 'uploads')) if os.path.isdir(os.path.join(BASE, 'uploads')) else []:
            if f in ('extracto_mayusculas.xlsx', 'bank_statement_en.xlsx'):
                os.remove(os.path.join(BASE, 'uploads', f))
        # y dejar archivos_procesados.json como estaba (si no, queda en el checkout).
        # gc.collect() antes: el generador SSE guarda el log al cerrarse.
        try:
            import gc as _gc, json as _j
            _gc.collect()
            _reg = os.path.join(BASE, 'datos-referencia', 'archivos_procesados.json')
            _d = _j.load(open(_reg, encoding='utf-8'))
            for k in ('extracto_mayusculas.xlsx', 'bank_statement_en.xlsx'):
                _d.pop(k, None)
            _j.dump(_d, open(_reg, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        except Exception:
            pass
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
