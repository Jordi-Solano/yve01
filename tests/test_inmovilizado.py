# -*- coding: utf-8 -*-
"""OLA B · bloque 4: inmovilizado y amortizaciones.

  python3.12 tests/test_inmovilizado.py
  python3.12 tests/test_inmovilizado.py --sabotaje
"""
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                    # noqa: E402
import inmovilizado as IM              # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DDIR = os.path.join(BASE, 'datos-referencia')
INMF = os.path.join(DDIR, IM.FICHERO)


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        IM._mi = lambda d: d.year * 12                    # ignora el mes: las cuotas salen mal

    a = IM.normalizar_activo({'descripcion': 'Portatil', 'categoria': 'INFORMATICA', 'fecha_alta': '15/06/2026', 'coste': 1200})
    ok(a['vida_util_anios'] == 4 and a['cuenta_activo'] == '217' and a['cuenta_amortizacion'] == '2817', f"defaults informatica: {a['vida_util_anios']} años, {a['cuenta_activo']}/{a['cuenta_amortizacion']}")
    m = {mes: IM.amortizacion_activo(a, mes) for mes in ('2026-05', '2026-06', '2026-07', '2030-05', '2030-06')}
    ok(m['2026-05']['estado'] == 'NO_ALTA' and m['2026-05']['cuota'] == 0, 'antes del alta no se amortiza')
    ok(m['2026-06']['cuota'] == 25.0 and m['2026-06']['acumulada'] == 25.0 and m['2026-06']['vnc'] == 1175.0, f"mes de alta: cuota {m['2026-06']['cuota']} (1200/48), VNC {m['2026-06']['vnc']}")
    ok(m['2026-07']['acumulada'] == 50.0, 'acumulada avanza')
    ok(m['2030-05']['acumulada'] == 1200.0 and m['2030-05']['estado'] == 'AMORTIZADO' and m['2030-06']['cuota'] == 0.0, 'termina a los 48 meses y no pasa de 1.200')
    b = IM.normalizar_activo({'descripcion': 'Mesa', 'categoria': 'MOBILIARIO', 'fecha_alta': '2025-01-10', 'coste': 1300, 'valor_residual': 100, 'fecha_baja': '2026-08-20'})
    mb = {mes: IM.amortizacion_activo(b, mes) for mes in ('2026-07', '2026-08')}
    ok(mb['2026-07']['cuota'] == 10.0 and mb['2026-07']['acumulada'] == 190.0, 'residual: (1300-100)/120 = 10/mes; 19 meses = 190')
    ok(mb['2026-08']['cuota'] == 0.0 and mb['2026-08']['estado'] == 'BAJA' and mb['2026-08']['acumulada'] == 190.0, 'con baja en agosto no se amortiza agosto')
    c = IM.normalizar_activo({'descripcion': 'Cocina', 'categoria': 'maquinaria', 'fecha_alta': '01/01/2020', 'coste': 24000, 'vida_util_anios': 5})
    ok(IM.amortizacion_activo(c, '2026-08')['estado'] == 'AMORTIZADO' and IM.amortizacion_activo(c, '2026-08')['vnc'] == 0.0, 'vida util propia manda; ya amortizada')
    try:
        IM.normalizar_activo({'descripcion': 'x', 'coste': 0, 'fecha_alta': '01/01/2026'}); ok(False, 'coste 0 rechazado')
    except ValueError:
        ok(True, 'coste 0 rechazado')
    try:
        IM.normalizar_activo({'descripcion': 'x', 'coste': 100, 'fecha_alta': ''}); ok(False, 'sin fecha rechazado')
    except ValueError:
        ok(True, 'sin fecha rechazado')
    ok(IM.vida_defecto('INFORMATICA', {'vidas': {'INFORMATICA': 3}}) == 3, 'config_inmovilizado.json manda en la vida util')

    df = pd.DataFrame([a, b, c, {'id': 'X', 'descripcion': 'roto', 'categoria': 'OTRO', 'fecha_alta': '', 'coste': 10}])
    ap = pd.DataFrame([
        {'numero_factura': 'F-TV', 'nombre_proveedor': 'MediaMarkt', 'fecha_factura': '10/08/2026', 'base_imponible': 4500.0, 'cuenta_contable': '629'},
        {'numero_factura': 'F-OBRA', 'nombre_proveedor': 'Reformas SL', 'fecha_factura': '12/08/2026', 'base_imponible': 800.0, 'cuenta_contable': 212.0},
        {'numero_factura': 'F-PAN', 'nombre_proveedor': 'Panaderia', 'fecha_factura': '12/08/2026', 'base_imponible': 80.0, 'cuenta_contable': '600'},
        {'numero_factura': 'F-JUL', 'nombre_proveedor': 'Otro', 'fecha_factura': '12/07/2026', 'base_imponible': 9000.0, 'cuenta_contable': '629'},
    ])
    r = IM.amortizar_mes('2026-08', df, ap)
    ok(r['resumen']['n_activos'] == 4 and r['resumen']['n_error'] == 1 and r['resumen']['cuota_mes'] == 25.0, f"resumen: {r['resumen']}")
    ok(len(r['asientos']) == 2 and r['asientos'][0]['cuenta'] == '681' and r['asientos'][0]['debe'] == 25.0 and r['asientos'][1]['cuenta'] == '2817' and r['asientos'][1]['haber'] == 25.0, 'asiento 681 / 2817 solo por lo que se amortiza en agosto')
    pend = {p['numero_factura']: p['motivo'] for p in r['altas_pendientes']}
    ok(set(pend) == {'F-TV', 'F-OBRA'} and 'inmovilizado' in pend['F-OBRA'] and 'alto' in pend['F-TV'], f"altas pendientes: {pend}")
    bx, nx = IM.exportar_excel(r)
    ok(set(pd.read_excel(bx, sheet_name=None)) == {'Resumen', 'Activos', 'Asientos', 'Altas pendientes'}, 'Excel con 4 hojas')

    # ── endpoints con copia del fichero ─────────────────────────────
    tmp = tempfile.mkdtemp(prefix='inm_')
    existia = os.path.exists(INMF)
    if existia:
        shutil.copy(INMF, os.path.join(tmp, 'i.xlsx'))
    try:
        if os.path.exists(INMF):
            os.remove(INMF)
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        r1 = cl.post('/api/inmovilizado/alta', json={'descripcion': 'TV habitaciones', 'categoria': 'MOBILIARIO', 'fecha_alta': '2026-08-10', 'coste': 4500, 'documento': 'F-TV'}, headers=H)
        d1 = r1.get_json() or {}
        ok(r1.status_code == 200 and d1.get('ok') and d1['activo']['id'].startswith('INM-'), f"alta por API: {d1.get('activo', {}).get('id')}")
        r2 = cl.post('/api/inmovilizado/alta', json={'descripcion': 'mal', 'categoria': 'OTRO', 'fecha_alta': '2026-08-10', 'coste': 0}, headers=H)
        ok(r2.status_code == 400, 'alta invalida -> 400')
        g = cl.get('/api/inmovilizado?mes=2026-08').get_json() or {}
        ok(g.get('ok') and g['resumen']['n_activos'] == 1 and g['resumen']['cuota_mes'] == 37.5 and 'categorias' in g, f"/api/inmovilizado: cuota {g.get('resumen', {}).get('cuota_mes')} (4500/120)")
        rb = cl.post('/api/inmovilizado/baja', json={'id': d1['activo']['id'], 'fecha_baja': '2026-09-01'}, headers=H)
        g2 = cl.get('/api/inmovilizado?mes=2026-09').get_json() or {}
        ok(rb.status_code == 200 and g2['activos'][0]['estado'] == 'BAJA' and g2['resumen']['cuota_mes'] == 0.0, 'baja por API: en septiembre ya no amortiza')
        rn = cl.post('/api/inmovilizado/baja', json={'id': 'NOPE', 'fecha_baja': '2026-09-01'}, headers=H)
        ok(rn.status_code == 404, 'baja de un id inexistente -> 404')
        rx = cl.get('/api/exportar/inmovilizado?mes=2026-08')
        ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/inmovilizado da un xlsx')
        html = cl.get('/').get_data(as_text=True)
        ok('id="card-cierre-inm"' in html and 'function loadInmovilizado' in html and 'loadInmovilizado();' in html, 'tarjeta en la pestaña Cierre')
    finally:
        if os.path.exists(INMF):
            os.remove(INMF)
        if existia:
            shutil.copy(os.path.join(tmp, 'i.xlsx'), INMF)
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
