# -*- coding: utf-8 -*-
"""OLA A · las dos provisiones del cierre (provisiones.py).

  1. Albaranes sin factura: lo entregado hasta fin de mes que el cruce marca
     ALBARAN_SIN_FACTURAR. DEBE gasto / HABER 4009.
  2. Comisiones OTA del mes: liquidaciones verificadas con periodo en el mes,
     por lo pactado si se puede calcular, si no por lo facturado.

Se prueba el modulo con un arbol de datos de mentira (sin tocar el repo) y
los dos endpoints contra dashboard.app con login.

  python3.12 tests/test_provisiones.py
  python3.12 tests/test_provisiones.py --sabotaje
"""
import os
import subprocess
import sys
import tempfile
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                      # noqa: E402
import provisiones as PV                 # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def montar(raiz):
    pdir = os.path.join(raiz, 'facturas-procesadas')
    rdir = os.path.join(raiz, 'reportes')
    ddir = os.path.join(raiz, 'datos-referencia')
    for d in (pdir, rdir, ddir):
        os.makedirs(d, exist_ok=True)
    # Albaranes crudos (etapa 1) + informe de cruce (etapa que gana), con estado.
    alb = [
        {'clave': 'A1', 'numero_albaran': 'ALB-1', 'nombre_proveedor': 'Pescados Rias SL',
         'fecha_entrega': '20/08/2026', 'total_albaran': 1000.0, 'archivo': 'a1.pdf', 'hotel_id': ''},
        {'clave': 'A2', 'numero_albaran': 'ALB-2', 'nombre_proveedor': 'Pescados Rias SL',
         'fecha_entrega': '28/08/2026', 'total_albaran': 250.5, 'archivo': 'a2.pdf', 'hotel_id': ''},
        {'clave': 'A3', 'numero_albaran': 'ALB-3', 'nombre_proveedor': 'Limpiezas Sur',
         'fecha_entrega': '15/08/2026', 'total_albaran': 300.0, 'archivo': 'a3.pdf', 'hotel_id': ''},
        {'clave': 'A4', 'numero_albaran': 'ALB-4', 'nombre_proveedor': 'Makro',
         'fecha_entrega': '02/09/2026', 'total_albaran': 999.0, 'archivo': 'a4.pdf', 'hotel_id': ''},
    ]
    with pd.ExcelWriter(os.path.join(pdir, 'albaranes_20260820.xlsx')) as w:
        pd.DataFrame(alb).to_excel(w, index=False, sheet_name='Albaranes')
        pd.DataFrame([{'clave': 'A1', 'n_linea': 1, 'descripcion': 'Merluza', 'cantidad': 10,
                       'precio_unitario': 100, 'importe': 1000}]).to_excel(w, index=False, sheet_name='Lineas')
    estados = {'ALB-1': 'ALBARAN_SIN_FACTURAR', 'ALB-2': 'ALBARAN_SIN_FACTURAR',
               'ALB-3': 'ALBARAN_FACTURADO', 'ALB-4': 'ALBARAN_SIN_FACTURAR'}
    cruce = [dict(a, estado=estados[a['numero_albaran']], numero_factura='',
                  detalle='x') for a in alb]
    with pd.ExcelWriter(os.path.join(rdir, 'matching_albaran_20260903.xlsx')) as w:
        pd.DataFrame([{'numero_factura': 'F-1', 'estado_matching': 'MATCH_ALBARAN_OK'}]
                     ).to_excel(w, index=False, sheet_name='Facturas')
        pd.DataFrame(cruce).to_excel(w, index=False, sheet_name='Albaranes')
    pd.DataFrame([{'nombre': 'Pescados Rias SL', 'tipo': 'FB'},
                  {'nombre': 'Limpiezas Sur', 'tipo': 'OTRAS'}]
                 ).to_excel(os.path.join(ddir, 'proveedores.xlsx'), index=False)
    # Verificacion OTA
    ver = [
        {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'periodo_inicio': '2026-08-01',
         'importe_bruto': 10000.0, 'porcentaje_pactado': 15.0, 'porcentaje_factura': 18.0,
         'importe_comision_factura': 1800.0, 'discrepancia_euros': 300.0, 'estado': 'DISCREPANCIA', 'hotel_id': ''},
        {'numero_factura': 'EX-1', 'nombre_ota': 'Expedia', 'periodo_inicio': '2026-08-01',
         'importe_bruto': 5000.0, 'porcentaje_pactado': 12.0, 'porcentaje_factura': 12.0,
         'importe_comision_factura': 600.0, 'discrepancia_euros': 0.0, 'estado': 'CORRECTO', 'hotel_id': ''},
        {'numero_factura': 'HB-1', 'nombre_ota': 'Hotelbeds', 'periodo_inicio': '2026-08-10',
         'importe_bruto': 2000.0, 'porcentaje_pactado': 'NO_ENCONTRADO', 'porcentaje_factura': 10.0,
         'importe_comision_factura': 200.0, 'discrepancia_euros': 'NO_ENCONTRADO', 'estado': 'OTA_DESCONOCIDA', 'hotel_id': ''},
        {'numero_factura': 'BK-2', 'nombre_ota': 'Booking.com', 'periodo_inicio': '2026-09-01',
         'importe_bruto': 8000.0, 'porcentaje_pactado': 15.0, 'porcentaje_factura': 15.0,
         'importe_comision_factura': 1200.0, 'discrepancia_euros': 0.0, 'estado': 'CORRECTO', 'hotel_id': ''},
    ]
    with pd.ExcelWriter(os.path.join(rdir, 'verificacion_20260903.xlsx')) as w:
        pd.DataFrame(ver).to_excel(w, index=False, sheet_name='Detalle')
    return pdir, rdir, ddir


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        # El filtro de estado se olvida: provisiona TODO lo entregado.
        _orig = PV._txt
        PV._txt = lambda v: 'ALBARAN_SIN_FACTURAR' if str(v) == 'ALBARAN_FACTURADO' else _orig(v)

    raiz = tempfile.mkdtemp(prefix='prov_')
    pdir, rdir, ddir = montar(raiz)

    # ── 1 · albaranes ─────────────────────────────────────────────────
    a = PV.provision_albaranes('2026-08', None, pdir, rdir, ddir, hoy=date(2026, 9, 3))
    nums = sorted(f['numero_albaran'] for f in a['filas'])
    ok(nums == ['ALB-1', 'ALB-2'], f"agosto: solo los sin facturar entregados hasta el 31: {nums}")
    ok(a['total'] == 1250.5, f"total 1.250,50 (dice {a['total']})")
    ok(a['por_proveedor'] and a['por_proveedor'][0]['cuenta_gasto'] == '600',
       'Pescados Rias (FB) va a la 600')
    ok(a['cuenta_provision']['codigo'] == '4009', 'la contrapartida por defecto es la 4009')
    debe = sum(x['debe'] for x in a['asientos']); haber = sum(x['haber'] for x in a['asientos'])
    ok(abs(debe - haber) < 0.005 and abs(debe - 1250.5) < 0.005, f'asiento cuadrado: debe {debe} = haber {haber}')
    ok(a['sin_cruzar'] == 0, 'todos cruzados en este arbol')
    a9 = PV.provision_albaranes('2026-09', None, pdir, rdir, ddir, hoy=date(2026, 9, 30))
    ok(sorted(f['numero_albaran'] for f in a9['filas']) == ['ALB-1', 'ALB-2', 'ALB-4'],
       'septiembre: entra ALB-4 (entregado el 2/9) y siguen los de agosto sin factura')
    # cuenta configurable
    with open(os.path.join(ddir, 'hotel_config.json'), 'w') as fh:
        fh.write('{"cuenta_provision_albaranes": "20630"}')
    ok(PV.provision_albaranes('2026-08', None, pdir, rdir, ddir)['cuenta_provision']['codigo'] == '20630',
       'la cuenta se puede fijar en hotel_config.json')

    # ── 2 · comisiones ────────────────────────────────────────────────
    c = PV.provision_comisiones('2026-08', None, rdir, ddir)
    por = {f['numero_factura']: f for f in c['filas']}
    ok(sorted(por) == ['BK-1', 'EX-1', 'HB-1'], f"agosto: 3 liquidaciones, BK-2 (septiembre) fuera: {sorted(por)}")
    ok(por['BK-1']['importe_provision'] == 1500.0 and por['BK-1']['base_provision'] == 'pactado',
       f"la discrepancia se provisiona por lo PACTADO (1.500, no 1.800): {por['BK-1']['importe_provision']}")
    ok(por['EX-1']['importe_provision'] == 600.0, 'Expedia correcta: 600')
    ok(por['HB-1']['importe_provision'] == 200.0 and por['HB-1']['base_provision'] == 'facturado',
       'sin tarifa conocida se provisiona lo facturado, y se dice')
    ok(c['total'] == 2300.0 and c['total_facturado'] == 2600.0, f"total {c['total']} / facturado {c['total_facturado']}")
    debe = sum(x['debe'] for x in c['asientos']); haber = sum(x['haber'] for x in c['asientos'])
    ok(abs(debe - haber) < 0.005 and abs(debe - 2300) < 0.005, f'asiento comisiones cuadrado ({debe})')
    ok(all(x['cuenta'] == '628' for x in c['asientos'] if x['debe']), 'el gasto va a la 628')

    # ── 3 · el Excel ──────────────────────────────────────────────────
    buf, nombre = PV.exportar_excel('2026-08', None, procesadas_dir=pdir, reportes_dir=rdir, datos_dir=ddir)
    hojas = pd.read_excel(buf, sheet_name=None)
    ok(set(hojas) == {'Resumen', 'Albaranes', 'Comisiones', 'Asientos'} and nombre == 'provisiones_2026-08.xlsx',
       f'Excel con 4 hojas: {sorted(hojas)}')
    ok(len(hojas['Asientos']) == 2 + 6, f"8 lineas de asiento ({len(hojas['Asientos'])})")

    # ── 4 · endpoints contra la app real ──────────────────────────────
    import dashboard
    app = dashboard.app
    app.config['TESTING'] = True
    cli = app.test_client()
    assert cli.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    r = cli.get('/api/provisiones?mes=2026-08')
    d = r.get_json() or {}
    ok(r.status_code == 200 and d.get('ok') and 'albaranes' in d and 'comisiones' in d,
       f'/api/provisiones responde con las dos ({r.status_code})')
    r2 = cli.get('/api/exportar/provisiones?mes=2026-08')
    ok(r2.status_code == 200 and r2.data[:2] == b'PK', f'/api/exportar/provisiones da un xlsx ({r2.status_code})')
    html = cli.get('/').get_data(as_text=True)
    ok('id="card-provisiones"' in html and 'function loadProvisiones' in html, 'la tarjeta y su JS estan en el panel')
    ok(cli.get('/api/provisiones').status_code == 200, 'sin mes usa el actual')

    # ── 5 · solo lee, nada de Oracle ──────────────────────────────────
    import ast
    tree = ast.parse(open(os.path.join(BASE, 'provisiones.py'), encoding='utf-8').read())
    escribe = [n.attr for n in ast.walk(tree) if isinstance(n, ast.Attribute)
               and n.attr in ('to_excel', 'to_csv', 'to_json', 'remove', 'unlink', 'rmtree')]
    ok(escribe == ['to_excel'] * 4, f'provisiones.py solo escribe en el BytesIO del export: {escribe}')
    imps = [n.names[0].name for n in ast.walk(tree) if isinstance(n, ast.Import)]
    ok(not [i for i in imps if i.startswith('oracle')], 'no importa oracle_*')
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni el clasificador tocados')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
