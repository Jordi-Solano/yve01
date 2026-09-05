# -*- coding: utf-8 -*-
"""OLA B · bloque 5: el archivo de fin de mes para la central.

  python3.12 tests/test_paquete_cierre.py
  python3.12 tests/test_paquete_cierre.py --sabotaje
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

import pandas as pd                    # noqa: E402
import paquete_cierre as PQ            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DDIR = os.path.join(BASE, 'datos-referencia')
COM = os.path.join(DDIR, PQ.FICHERO_COMENTARIOS)


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        PQ.cuenta_resultados = lambda mayor: {'ingresos': 0.0, 'gastos': 0.0, 'resultado': 0.0, 'por_cuenta': []}

    mayor = [{'cuenta': '705', 'descripcion': 'Alojamiento', 'debe': 0.0, 'haber': 1000.0},
             {'cuenta': '700', 'descripcion': 'F&B', 'debe': 0.0, 'haber': 300.0},
             {'cuenta': '600', 'descripcion': 'Compras', 'debe': 400.0, 'haber': 0.0},
             {'cuenta': '629', 'descripcion': 'Otros', 'debe': 150.0, 'haber': 0.0},
             {'cuenta': '472', 'descripcion': 'IVA', 'debe': 99.0, 'haber': 0.0}]
    cr = PQ.cuenta_resultados(mayor)
    ok(cr['ingresos'] == 1300.0 and cr['gastos'] == 550.0 and cr['resultado'] == 750.0 and len(cr['por_cuenta']) == 4, f"resultados 7xx-6xx: {cr['ingresos']} - {cr['gastos']} = {cr['resultado']} (el IVA no entra)")

    asientos = {'cuadra': True, 'n_asientos': 12, 'debe': 5000.0, 'avisos': ['1 cuenta fuera de plan']}
    recon = {'ok': False, 'resumen': {'CUADRA': 5, 'DIFERENCIA': 1, 'PENDIENTE': 1, 'SIN_DATO': 1}, 'checks': [{'concepto': 'Proveedores', 'estado': 'DIFERENCIA'}], 'mayor': mayor}
    banco = {'n': 10, 'ok': True, 'sin_clasificar': 0, 'sin_conciliar': 2, 'saldo_final': 9000.0, 'pestanas': {'AP': {'n': 3, 'total': -400.0}, 'AR': {'n': 2, 'total': 900.0}}}
    prov = [{'total': 80.0, 'n': 1}, {'total': 200.0, 'n': 2}]
    inv = {'resumen': {'n_articulos': 7, 'n_revisar': 0, 'valor_final': 1200.0, 'consumo_real_fb': 656.0, 'desviacion_pct': 9.3}}
    inm = {'resumen': {'n_activos': 3, 'n_error': 0, 'altas_pendientes': 1, 'cuota_mes': 62.5, 'vnc_total': 8000.0}}
    aging = {'n': 4, 'total': 3000.0, 'mas_de_60': 0.0}
    coms = {'resumen': {'texto': 'Mes tranquilo'}, 'banco': {'texto': 'Dos transferencias sin identificar'}}
    p = PQ.montar('2026-08', asientos, recon, banco, prov, inv, inm, aging, None, coms, drr={'rooms_revenue_mtd': 1300.0})
    est = {c['clave']: c['estado'] for c in p['checklist']}
    ok(len(p['checklist']) == 8 and est == {'asientos': 'OK', 'reconciliacion': 'PENDIENTE', 'banco': 'OK', 'provisiones': 'OK', 'inventarios': 'OK', 'inmovilizado': 'PENDIENTE', 'aging': 'OK', 'fiscal': 'SIN_DATO'}, f"checklist: {est}")
    ok(p['listo'] is False and p['resumen_checklist']['PENDIENTE'] == 2, 'no esta listo con 2 pendientes')
    ok(p['comentario_general'] == 'Mes tranquilo' and next(c for c in p['checklist'] if c['clave'] == 'banco')['comentario'] == 'Dos transferencias sin identificar', 'los comentarios viajan con cada bloque')
    ok(p['resultado']['resultado'] == 750.0 and p['resultado']['drr_rooms_revenue'] == 1300.0, 'resultado del mes y DRR en el resumen')
    p0 = PQ.montar('2026-08')
    # Lote 3 de Jordi (fase 6): un mes vacio NO esta "listo para la central",
    # esta sin datos. Antes este test exigia listo=True; era el bug.
    ok(all(c['estado'] == 'SIN_DATO' for c in p0['checklist']) and p0['listo'] is False and p0.get('sin_datos') is True,
       'sin bloques: todo SIN_DATO y NO listo (sin_datos=True)')
    buf, nombre = PQ.exportar_excel(p, {'asientos': [{'num': 1}]}, recon, banco, prov, {'familias': [{'familia': 'ALIMENTOS'}]}, {'activos': [{'id': 'X'}]}, {'por_acreedor': [{'acreedor': 'Makro'}]})
    hojas = set(pd.read_excel(buf, sheet_name=None))
    ok({'Portada', 'Checklist', 'Resultados', 'Libro Diario', 'Mayor', 'Reconciliacion', 'Banco pestañas', 'Banco movimientos', 'Provisiones', 'Inventarios', 'Inmovilizado', 'Aging AP'} <= hojas and nombre == 'cierre_2026-08_paquete_central.xlsx', f'Excel del paquete: {len(hojas)} hojas')

    # comentarios: fichero real con copia
    tmp = tempfile.mkdtemp(prefix='paq_')
    existia = os.path.exists(COM)
    if existia:
        shutil.copy(COM, os.path.join(tmp, 'c.json'))
    try:
        if os.path.exists(COM):
            os.remove(COM)
        m = PQ.guardar_comentario('2026-08', 'banco', 'hola', 'admin')
        ok(m['banco']['texto'] == 'hola' and m['banco']['usuario'] == 'admin' and PQ.comentarios('2026-08')['banco']['texto'] == 'hola', 'guardar y leer comentario')
        m2 = PQ.guardar_comentario('2026-08', 'banco', '', 'admin')
        ok('banco' not in m2, 'texto vacio borra el comentario')
        try:
            PQ.guardar_comentario('2026-08', 'nope', 'x'); ok(False, 'seccion desconocida rechazada')
        except ValueError:
            ok(True, 'seccion desconocida rechazada')
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (c.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        rc = c.post('/api/cierre/comentario', json={'mes': '2026-08', 'seccion': 'resumen', 'texto': 'Todo bien'}, headers=H)
        ok(rc.status_code == 200 and json.load(open(COM, encoding='utf-8'))['2026-08']['resumen']['usuario'] == 'admin', 'comentario por API con el usuario logueado')
        rb = c.post('/api/cierre/comentario', json={'mes': '20', 'seccion': 'resumen', 'texto': 'x'}, headers=H)
        ok(rb.status_code == 400, 'mes invalido -> 400')
        r = c.get('/api/cierre/paquete?mes=2026-08'); d = r.get_json() or {}
        ok(r.status_code == 200 and d.get('ok') and len(d.get('checklist', [])) == 8 and d.get('comentario_general') == 'Todo bien', f'/api/cierre/paquete ({r.status_code}) con el comentario')
        rx = c.get('/api/exportar/cierre_paquete?mes=2026-08')
        ok(rx.status_code == 200 and rx.data[:2] == b'PK', 'paquete descargable')
        html = c.get('/').get_data(as_text=True)
        ok('id="card-cierre-paquete"' in html and 'function loadPaquete' in html and 'loadPaquete();' in html, 'tarjeta en la pestaña Cierre')
    finally:
        if os.path.exists(COM):
            os.remove(COM)
        if existia:
            shutil.copy(os.path.join(tmp, 'c.json'), COM)
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
