# -*- coding: utf-8 -*-
"""OLA B · bloque 3: inventarios de cierre (alimentos, bebidas, licores, guest supplies).

  python3.12 tests/test_inventarios_cierre.py
  python3.12 tests/test_inventarios_cierre.py --sabotaje
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                    # noqa: E402
import inventarios as INV              # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
MES = '2026-08'
DDIR = os.path.join(BASE, 'datos-referencia')
INVF = os.path.join(DDIR, 'inventario.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        INV.familia = lambda c, n='', cfg=None: 'ALIMENTOS'      # todo a alimentos

    df_inv = pd.DataFrame([
        {'ingrediente': 'arroz bomba', 'categoria': 'Secos', 'coste_unitario': 3.2, 'stock_inicial_kg_l': 50, 'stock_actual_kg_l': 30, 'unidad': 'kg', 'hotel_id': ''},
        {'ingrediente': 'gambas', 'categoria': 'Mariscos', 'coste_unitario': 22.0, 'stock_inicial_kg_l': 10, 'stock_actual_kg_l': 12, 'unidad': 'kg', 'hotel_id': ''},
        {'ingrediente': 'Gin Tanqueray 70cl', 'categoria': '', 'coste_unitario': 18.0, 'stock_inicial_kg_l': 6, 'stock_actual_kg_l': 4, 'unidad': 'ud', 'hotel_id': ''},
        {'ingrediente': 'Vino Rioja crianza', 'categoria': 'Bebidas', 'coste_unitario': 7.5, 'stock_inicial_kg_l': 40, 'stock_actual_kg_l': 40, 'unidad': 'ud', 'hotel_id': ''},
        {'ingrediente': 'gel de baño 30ml', 'categoria': 'Amenities', 'coste_unitario': 0.4, 'stock_inicial_kg_l': 1000, 'stock_actual_kg_l': 700, 'unidad': 'ud', 'hotel_id': ''},
        {'ingrediente': 'gelatina neutra', 'categoria': 'Secos', 'coste_unitario': 0, 'stock_inicial_kg_l': 2, 'stock_actual_kg_l': 1, 'unidad': 'kg', 'hotel_id': ''},
        {'ingrediente': 'lejia', 'categoria': 'Limpieza', 'coste_unitario': 1.0, 'stock_inicial_kg_l': 5, 'stock_actual_kg_l': None, 'unidad': 'l', 'hotel_id': ''},
    ])
    df_ap = pd.DataFrame([
        {'numero_factura': 'F-1', 'tipo_proveedor': 'FB', 'fecha_factura': '05/08/2026', 'base_imponible': 500.0, 'total_factura': 550.0},
        {'numero_factura': 'F-2', 'tipo_proveedor': 'OTRAS', 'fecha_factura': '05/08/2026', 'base_imponible': 900.0, 'total_factura': 1089.0},
        {'numero_factura': 'F-3', 'tipo_proveedor': 'FB', 'fecha_factura': '05/09/2026', 'base_imponible': 100.0, 'total_factura': 110.0},
        {'numero_factura': 'F-4', 'tipo_proveedor': 'FB', 'fecha_factura': '20/08/2026', 'base_imponible': 0, 'total_factura': 110.0},
    ])
    r = INV.valorar(MES, df_inv, df_ap, coste_teorico_fb=600.0)
    fam = {f['familia']: f for f in r['familias']}
    ok(set(fam) == {'ALIMENTOS', 'BEBIDAS', 'LICORES', 'GUEST_SUPPLIES', 'OTROS'}, f"familias: {sorted(fam)}")
    ok(fam['ALIMENTOS']['n'] == 3 and fam['LICORES']['n'] == 1 and fam['BEBIDAS']['n'] == 1 and fam['GUEST_SUPPLIES']['n'] == 1, 'cada articulo en su familia (gelatina no es gel; gin es licor)')
    ok(fam['ALIMENTOS']['valor_inicial'] == 380.0 and fam['ALIMENTOS']['valor_final'] == 360.0 and fam['ALIMENTOS']['variacion'] == -20.0, f"alimentos: {fam['ALIMENTOS']}")
    ok(r['resumen']['compras_fb'] == 600.0 and r['resumen']['n_facturas_fb'] == 2, f"compras F&B del mes (solo FB, solo agosto, F-4 sin base -> total/1,10): {r['resumen']['compras_fb']}")
    # consumo real = inicial(380+108+300) + 600 - final(360+72+300) = 656
    ok(r['resumen']['consumo_real_fb'] == 656.0 and r['resumen']['desviacion_fb'] == 56.0 and r['resumen']['desviacion_pct'] == 9.3, f"consumo real {r['resumen']['consumo_real_fb']} vs teorico 600 -> desviacion {r['resumen']['desviacion_fb']} ({r['resumen']['desviacion_pct']} %)")
    asi = {(a['familia'], a['cuenta']): a for a in r['asientos']}
    ok(asi.get(('ALIMENTOS', '610'), {}).get('debe') == 20.0 and asi.get(('ALIMENTOS', '300'), {}).get('haber') == 20.0, 'bajada de existencias: 610 debe / 300 haber')
    ok(asi.get(('GUEST_SUPPLIES', '612'), {}).get('debe') == 120.0 and asi.get(('GUEST_SUPPLIES', '328'), {}).get('haber') == 120.0, 'guest supplies por 328/612')
    ok(('BEBIDAS', '610') not in asi, 'sin variacion no hay asiento')
    ok(all(abs(sum(a['debe'] for a in r['asientos']) - sum(a['haber'] for a in r['asientos'])) < 0.01 for _ in [0]), 'los asientos cuadran')
    rev = {a['articulo']: a['motivo'] for a in r['revisar']}
    ok(rev == {'gelatina neutra': 'sin coste unitario', 'lejia': 'sin recuento final'}, f'a revisar con motivo: {rev}')
    r0 = INV.valorar(MES, df_inv, df_ap, coste_teorico_fb=None)
    ok(r0['resumen']['consumo_teorico_fb'] is None and r0['resumen']['desviacion_fb'] is None, 'sin escandallo no hay desviacion (no se inventa)')
    ok(INV.familia('Secos', 'x', {'secos': 'OTROS'}) == 'OTROS', 'config_inventarios.json manda sobre la palabra')

    buf, nombre = INV.hoja_recuento(df_inv, MES)
    hoja = pd.read_excel(buf)
    ok(list(hoja.columns)[:6] == ['ingrediente', 'categoria', 'familia', 'unidad', 'stock_sistema', 'recuento'] and len(hoja) == 7 and hoja['recuento'].isna().all(), 'hoja de recuento con columna vacia')
    hoja['recuento'] = hoja['recuento'].astype(object)     # regla 9: la columna vacia es float64
    hoja.loc[hoja['ingrediente'] == 'arroz bomba', 'recuento'] = 25
    hoja.loc[hoja['ingrediente'] == 'gambas', 'recuento'] = '11,5'
    b2 = io.BytesIO(); hoja.to_excel(b2, index=False); b2.seek(0)
    dfc, n, saltadas = INV.leer_recuento(b2)
    ok(n == 2 and saltadas == 5 and set(dfc['ingrediente']) == {'arroz bomba', 'gambas'} and float(dfc[dfc['ingrediente'] == 'gambas']['stock_actual_kg_l'].iloc[0]) == 11.5, f'leer_recuento: {n} contados, {saltadas} sin contar, coma decimal ok')
    bx, nx = INV.exportar_excel(r)
    ok(set(pd.read_excel(bx, sheet_name=None)) == {'Resumen', 'Familias', 'Articulos', 'Asientos', 'Revisar'}, 'Excel con 5 hojas')

    # ── endpoints, con copia del inventario real ────────────────────
    tmp = tempfile.mkdtemp(prefix='inv_')
    existia = os.path.exists(INVF)
    if existia:
        shutil.copy(INVF, os.path.join(tmp, 'inv.xlsx'))
    try:
        df_inv.to_excel(INVF, index=False)
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
        rr = c.get('/api/inventarios?mes=2026-08'); d = rr.get_json() or {}
        ok(rr.status_code == 200 and d.get('ok') and d['resumen']['n_articulos'] == 7, f"/api/inventarios ({rr.status_code}): {d.get('resumen', {}).get('n_articulos')} articulos")
        rh = c.get('/api/inventarios/hoja_recuento?mes=2026-08')
        ok(rh.status_code == 200 and rh.data[:2] == b'PK', 'hoja de recuento descargable')
        b3 = io.BytesIO(); hoja.to_excel(b3, index=False); b3.seek(0)
        ru = c.post('/api/inventarios/recuento?mes=2026-08', data={'archivo': (b3, 'recuento.xlsx')}, content_type='multipart/form-data', headers={'X-CSRF-Token': tok})
        du = ru.get_json() or {}
        inv2 = pd.read_excel(INVF)
        arroz = inv2[inv2['ingrediente'] == 'arroz bomba'].iloc[0]
        ok(ru.status_code == 200 and du.get('contados') == 2 and float(arroz['stock_actual_kg_l']) == 25.0 and float(arroz['coste_unitario']) == 3.2 and float(arroz['stock_inicial_kg_l']) == 50.0 and len(inv2) == 7,
           f'subir recuento: stock final actualizado, coste e inicial conservados, sin duplicar ({len(inv2)} filas)')
        rx = c.get('/api/exportar/inventarios?mes=2026-08')
        ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/inventarios da un xlsx')
        html = c.get('/').get_data(as_text=True)
        ok('id="card-cierre-inv"' in html and 'function loadInventarios' in html and 'loadInventarios();' in html, 'tarjeta en la pestaña Cierre')
    finally:
        if os.path.exists(INVF):
            os.remove(INVF)
        if existia:
            shutil.copy(os.path.join(tmp, 'inv.xlsx'), INVF)
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
