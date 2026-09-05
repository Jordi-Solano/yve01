# -*- coding: utf-8 -*-
"""OLA A · direct bill: factura a credito vs BONO de agencia.

1. NO REGRESION del clasificador: los 15 tipos que habia (lista y esquema JSON
   de cada uno, byte a byte) y las ramas del enrutador siguen EXACTAMENTE
   igual. La linea base esta en tests/baseline_tipos_clasificador.json, tomada
   con `git stash` sobre el codigo anterior al BONO (regla 23).
2. El BONO: prompt + parser + enrutador (guarda en bonos_agencia.xlsx) +
   cotejo puro + endpoint + tarjeta.

  python3.12 tests/test_bono_direct_bill.py
  python3.12 tests/test_bono_direct_bill.py --sabotaje
"""
import ast
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                     # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DDIR = os.path.join(BASE, 'datos-referencia')
BONOS = os.path.join(DDIR, 'bonos_agencia.xlsx')
CLIENTES = os.path.join(DDIR, 'clientes_credito.xlsx')   # desde sep 2026 el bono crea la ficha del cliente
RESERVAS = os.path.join(DDIR, 'reservas_credito.xlsx')
BASELINE = json.load(open(os.path.join(BASE, 'tests', 'baseline_tipos_clasificador.json'), encoding='utf-8'))


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import lector_facturas_ap as L
    P = L.PROMPT_CLASIFICACION
    if SABOTAJE:
        # se "mejora" el esquema de ALBARAN y se pierde la rama de ROOMING
        P = P.replace('"numero_albaran":"X"', '"numero":"X"')

    # ── 1 · no regresion de los tipos que habia ─────────────────────
    tipos = set(re.findall(r'→ ([A-Z_]+)', P))
    faltan = [t for t in BASELINE['tipos'] if t not in tipos]
    ok(not faltan, f"los {len(BASELINE['tipos'])} tipos anteriores siguen en la clasificacion (faltan: {faltan})")
    ok('BONO' in tipos, 'y ahora tambien BONO')
    lineas = P.split('\n')
    rotos = []
    for tipo, esquemas in BASELINE['esquemas'].items():
        for e in esquemas:
            if e not in lineas:
                rotos.append(tipo)
    ok(not rotos, f'el esquema JSON de cada tipo anterior es identico byte a byte (rotos: {sorted(set(rotos))})')
    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
    fn = next(n for n in ast.walk(ast.parse(src)) if isinstance(n, ast.FunctionDef) and n.name == '_enrutar_tipo_doc')
    ramas = set()
    for n in ast.walk(fn):
        if isinstance(n, ast.Compare) and isinstance(n.left, ast.Name) and n.left.id == '_tipo_doc':
            for c in n.comparators:
                if isinstance(c, ast.Constant):
                    ramas.add(c.value)
                elif isinstance(c, ast.Tuple):
                    ramas.update(e.value for e in c.elts if isinstance(e, ast.Constant))
    if SABOTAJE:
        ramas.discard('ROOMING')
    ok(set(BASELINE['ramas']) <= ramas, f"el enrutador conserva las ramas anteriores (faltan: {sorted(set(BASELINE['ramas']) - ramas)})")
    ok('BONO' in ramas, 'y tiene rama BONO')
    ok('"tipo_documento":"BONO"' in P and '"agencia"' in P and '"numero_bono"' in P, 'esquema JSON del BONO en el prompt')

    # ── 2 · parser ──────────────────────────────────────────────────
    f = L.bono_de_respuesta({'tipo_documento': 'BONO', 'numero_bono': 'V-77', 'agencia': 'Viajes Sol SL',
                             'huesped': 'Ana Puig', 'fecha_entrada': '10/09/2026', 'fecha_salida': '13/09/2026',
                             'noches': 3, 'habitaciones': 1, 'precio_noche': 120.0, 'importe_total': None,
                             'referencia_reserva': 'FAC-2026-CORP-0001'}, 'bono.pdf')
    ok(f['importe_total'] == 360.0 and f['clave'] == 'v-77|viajes sol sl', f"total derivado 3 x 120 = {f['importe_total']}; clave {f['clave']}")
    ok(L.bono_tiene_datos(f) and not L.bono_tiene_datos({'agencia': ''}), 'bono_tiene_datos: pagador + importe/fechas')
    ok(L.bono_de_respuesta({'numero_bono': 'X', 'agencia': 'A', 'precio_noche': 100.0}, 'b')['importe_total'] is None,
       'sin noches no se inventa el total')

    # ── 3 · cotejo puro ─────────────────────────────────────────────
    import matching_bonos as MB
    df_b = pd.DataFrame([
        {'clave': 'V-77|viajes sol sl', 'numero_bono': 'V-77', 'agencia': 'Viajes Sol SL', 'huesped': 'Ana Puig', 'fecha_entrada': '10/09/2026', 'fecha_salida': '13/09/2026', 'importe_total': 360.0, 'referencia_reserva': 'FAC-2026-CORP-0001', 'hotel_id': ''},
        {'clave': 'V-78|acme corp', 'numero_bono': 'V-78', 'agencia': 'ACME Corp', 'huesped': 'Bob', 'fecha_entrada': '01/09/2026', 'fecha_salida': '03/09/2026', 'importe_total': 500.0, 'referencia_reserva': '', 'hotel_id': ''},
        {'clave': 'V-79|tour mundo', 'numero_bono': 'V-79', 'agencia': 'Tour Mundo', 'huesped': 'Cai', 'fecha_entrada': '05/09/2026', 'fecha_salida': '07/09/2026', 'importe_total': 800.0, 'referencia_reserva': '', 'hotel_id': ''},
        {'clave': 'V-80|nadie', 'numero_bono': 'V-80', 'agencia': 'Nadie SA', 'huesped': 'Dan', 'fecha_entrada': '20/09/2026', 'fecha_salida': '22/09/2026', 'importe_total': 300.0, 'referencia_reserva': '', 'hotel_id': ''},
    ])
    df_r = pd.DataFrame([
        {'numero_reserva': 'FAC-2026-CORP-0001', 'cliente': 'Viajes Sol', 'fecha_entrada': '2026-09-10', 'fecha_salida': '2026-09-13', 'total': 360.5, 'estado': 'FACTURADO', 'hotel_id': ''},
        {'numero_reserva': 'FAC-2026-CORP-0002', 'cliente': 'ACME Corp', 'fecha_entrada': '2026-09-01', 'fecha_salida': '2026-09-03', 'total': 650.0, 'estado': 'FACTURADO', 'hotel_id': ''},
        {'numero_reserva': 'FAC-2026-CORP-0003', 'cliente': 'Tour Mundo SL', 'fecha_entrada': '2026-09-06', 'fecha_salida': '2026-09-08', 'total': 800.0, 'estado': 'COBRADO', 'hotel_id': ''},
        {'numero_reserva': 'FAC-2026-CORP-0004', 'cliente': 'Empresa Sin Bono', 'fecha_entrada': '2026-09-02', 'fecha_salida': '2026-09-04', 'total': 999.0, 'estado': 'FACTURADO', 'hotel_id': ''},
        {'numero_reserva': 'GRP-55', 'cliente': 'Grupo Pendiente', 'fecha_entrada': '2026-10-02', 'fecha_salida': '2026-10-04', 'total': 5000.0, 'estado': 'PENDIENTE_FACTURA', 'hotel_id': ''},
    ])
    if SABOTAJE:
        MB._cuadra = lambda a, b: True       # todo cuadra
    r = MB.cotejar(df_b, df_r)
    est = {b['numero_bono']: b for b in r['bonos']}
    ok(est['V-77']['estado'] == 'CUADRA' and est['V-77']['via'] == 'referencia', f"V-77 por referencia, 360,50 vs 360 cuadra (margen 1 EUR): {est['V-77']['estado']}")
    ok(est['V-78']['estado'] == 'DIFERENCIA_IMPORTE' and '650' in est['V-78']['detalle'], f"V-78 por pagador+fechas, importe distinto: {est['V-78']['detalle']}")
    ok(est['V-79']['estado'] == 'DIFERENCIA_FECHAS' and est['V-79']['numero_factura'] == 'FAC-2026-CORP-0003', f"V-79 por pagador+importe, fechas distintas: {est['V-79']['estado']}")
    ok(est['V-80']['estado'] == 'SIN_FACTURA', 'V-80 sin factura')
    sb = [x['numero'] for x in r['facturas_sin_bono']]
    ok(sb == ['FAC-2026-CORP-0004'], f'factura a credito sin bono: {sb} (la PENDIENTE_FACTURA no cuenta)')
    ok(r['resumen']['importe_en_disputa'] == 150.0 and r['resumen']['importe_sin_bono'] == 999.0, f"resumen {r['resumen']}")
    buf, nombre = MB.exportar_excel(r)
    ok(set(pd.read_excel(buf, sheet_name=None)) == {'Resumen', 'Bonos', 'Facturas sin bono'}, 'Excel con 3 hojas')

    # ── 4 · enrutador + endpoint (con copia de seguridad de los ficheros) ──
    tmp = tempfile.mkdtemp(prefix='bono_')
    copias = {}
    for ruta in (BONOS, RESERVAS, CLIENTES):
        if os.path.exists(ruta):
            copias[ruta] = os.path.join(tmp, os.path.basename(ruta)); shutil.copy(ruta, copias[ruta])
    try:
        for ruta in (BONOS,):
            if os.path.exists(ruta):
                os.remove(ruta)
        import dashboard as D
        app = D.app
        app.config['TESTING'] = True
        with app.test_request_context('/'):
            msg, marca, flags = D._enrutar_tipo_doc({'tipo_documento': 'BONO', 'numero_bono': 'V-77', 'agencia': 'Viajes Sol SL',
                                                     'huesped': 'Ana Puig', 'fecha_entrada': '10/09/2026', 'fecha_salida': '13/09/2026',
                                                     'noches': 3, 'habitaciones': 1, 'precio_noche': 120.0}, 'bono.pdf')
            ok(marca == 'BONO_OK' and flags.get('bono') and msg.startswith('✓ Bono'), f'enrutador: {marca} · {msg}')
            msg2, marca2, _ = D._enrutar_tipo_doc({'tipo_documento': 'BONO', 'numero_bono': 'V-77', 'agencia': 'Viajes Sol SL',
                                                   'huesped': 'Ana Puig', 'fecha_entrada': '10/09/2026', 'fecha_salida': '13/09/2026',
                                                   'noches': 3, 'habitaciones': 1, 'precio_noche': 125.0}, 'bono.pdf')
            dfb = pd.read_excel(BONOS)
            ok(len(dfb) == 1 and float(dfb.iloc[0]['importe_total']) == 375.0, f'reprocesar el mismo bono lo actualiza, no lo duplica ({len(dfb)} fila, {dfb.iloc[0]["importe_total"]})')
            msg3, marca3, _ = D._enrutar_tipo_doc({'tipo_documento': 'BONO', 'agencia': ''}, 'vacio.pdf')
            ok(marca3 == 'SKIP' and msg3.startswith('⚠'), 'bono sin datos: SKIP honesto')
            # las ramas viejas siguen respondiendo igual
            m_o, k_o, _ = D._enrutar_tipo_doc({'tipo_documento': 'OTRO', 'descripcion': 'agenda'}, 'x.pdf')
            m_x, k_x, _ = D._enrutar_tipo_doc({'tipo_documento': 'INEXISTENTE'}, 'x.pdf')
            ok(k_o == 'SKIP' and m_o == '⚠ x.pdf: agenda' and k_x == 'SKIP' and 'INEXISTENTE' in m_x, 'OTRO y tipo desconocido como siempre')
        pd.DataFrame([{'numero_reserva': 'FAC-2026-CORP-0001', 'cliente': 'Viajes Sol', 'fecha_entrada': '2026-09-10', 'fecha_salida': '2026-09-13',
                       'habitaciones': 1, 'importe_habitaciones': 375.0, 'importe_fb': 0, 'importe_extras': 0, 'total': 375.0,
                       'estado': 'FACTURADO', 'fecha_emision': '2026-09-13', 'hotel_id': ''}]).to_excel(RESERVAS, index=False)
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        d = c.get('/api/ar_real/bonos').get_json()
        ok(d.get('ok') and len(d['bonos']) == 1 and d['bonos'][0]['estado'] == 'CUADRA' and d['bonos'][0]['numero_factura'] == 'FAC-2026-CORP-0001',
           f"/api/ar_real/bonos: {d.get('bonos') and d['bonos'][0]['estado']}")
        rx = c.get('/api/exportar/bonos')
        ok(rx.status_code == 200 and rx.data[:2] == b'PK', '/api/exportar/bonos da un xlsx')
        html = c.get('/').get_data(as_text=True)
        ok('id="ar-bonos-section"' in html and 'function cargarBonosAR' in html and 'cargarBonosAR();' in html, 'tarjeta y JS en AR Real')
    finally:
        for ruta in (BONOS, RESERVAS, CLIENTES):
            if os.path.exists(ruta):
                os.remove(ruta)
            if ruta in copias:
                shutil.copy(copias[ruta], ruta)
        shutil.rmtree(tmp, ignore_errors=True)

    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_')], 'ningun oracle_* tocado')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
