# -*- coding: utf-8 -*-
"""b84 — la ficha de la factura AP y lo que una persona decide sobre ella.

  - el clasificador guarda periodo_inicio / periodo_fin (el prompt los pide; '' si no)
  - almacen_datos: ajustes_ap.json se aplica al leer (cuenta corregida, vencimiento,
    fecha contable = dia de registro, dias de pago, pagada); fecha_registro sale del
    primer facturas_ap_<dia>.xlsx en el que aparece la factura
  - /api/ap/ficha: concepto, periodo, asiento (cuenta/472/400), lineas, historial
  - /api/ap/ajustar: cuenta (y "aplicar al proveedor" -> proveedores_aprendidos), vencimiento,
    fecha contable, dias de pago; validaciones
  - /api/ap/pagar y /despagar: fecha, cuenta bancaria (se recuerda), quien; sale del aging
  - aging: dias desde el vencimiento; pagadas a mano fuera
  - cierre y 303: la factura entra en el mes de la fecha CONTABLE (config ap_fecha=contable,
    decision de Jordi 23 sep); con ap_fecha=factura, en el de la factura
  - descargas: Excel AP real (antes, 4 filas inventadas), asientos de las marcadas, PDF de la ficha
  - la tabla AP: columna Pago, filtro pagadas / pendientes, checkbox; i18n en 6 idiomas

  python3.12 tests/test_ficha_ap.py
  python3.12 tests/test_ficha_ap.py --sabotaje
"""
import io
import json
import os
import shutil
import sys
import tempfile
from datetime import date, timedelta

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']
PROC = os.path.join(BASE, 'facturas-procesadas')
DD = os.path.join(BASE, 'datos-referencia')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    import almacen_datos as ALM
    import lector_facturas_ap as L
    import cierre_mes as CM
    if SABOTAJE:
        ALM.aplicar_ajustes_ap = lambda df, *a, **k: df     # los ajustes no se aplican al leer

    # 1. el clasificador
    ok('"periodo_inicio":"DD/MM/YYYY","periodo_fin":"DD/MM/YYYY"' in L.PROMPT_CLASIFICACION and 'periodo_inicio"/"periodo_fin" SOLO si' in L.PROMPT_CLASIFICACION, 'el prompt pide periodo_inicio/fin solo si la factura lo dice')
    r = L.normalizar_factura_ap({'numero_factura': 'EL-1', 'fecha': '14/08/2026', 'nombre_proveedor': 'Energia', 'base_imponible': 980, 'porcentaje_iva': 21, 'cuota_iva': 195.8, 'total_factura': 1175.8, 'periodo_inicio': '01/07/2026', 'periodo_fin': '31/07/2026'}, 'f.pdf', {})
    r2 = L.normalizar_factura_ap({'numero_factura': 'X-1', 'fecha': '14/08/2026', 'nombre_proveedor': 'Otro', 'total_factura': 121.0, 'periodo_inicio': None, 'periodo_fin': 'julio'}, 'g.pdf', {})
    ok(r['periodo_inicio'] == '01/07/2026' and r['periodo_fin'] == '31/07/2026' and r2['periodo_inicio'] == '' and r2['periodo_fin'] == '', f"normalizar: periodo guardado; sin periodo → '' (no NO_ENCONTRADO): {r2['periodo_inicio']!r}")

    copia = tempfile.mkdtemp(prefix='fap_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in os.listdir(PROC) if os.path.isdir(PROC) else []:
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PROC, f))
        for f in ('ajustes_ap.json', 'proveedores_aprendidos.json', 'cuentas_bancarias.json', 'config_cierre.json'):
            p = os.path.join(DD, f)
            if os.path.exists(p):
                os.remove(p)
        os.makedirs(PROC, exist_ok=True)
        hoy = date.today()
        # dos dias de registro distintos: la factura de julio entro el 3 de septiembre
        pd.DataFrame([
            {"archivo": "luz.pdf", "numero_factura": "EL-84", "fecha": "14/07/2026", "nombre_proveedor": "Energia Test SA", "NIF_proveedor": "A11111111", "descripcion_concepto": "Consumo electrico julio", "periodo_inicio": "01/07/2026", "periodo_fin": "31/07/2026", "base_imponible": 980.0, "porcentaje_iva": 21, "cuota_iva": 205.8, "total_factura": 1185.8, "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "hotel_id": ""},
        ]).to_excel(os.path.join(PROC, 'facturas_ap_20260903.xlsx'), index=False)
        pd.DataFrame([
            {"archivo": "carne.pdf", "numero_factura": "CS-84", "fecha": "10/09/2026", "nombre_proveedor": "Carnes Test SL", "NIF_proveedor": "B22222222", "descripcion_concepto": "Carne", "periodo_inicio": "", "periodo_fin": "", "base_imponible": 200.0, "porcentaje_iva": 10, "cuota_iva": 20.0, "total_factura": 220.0, "tipo_proveedor": "FB", "cuenta_contable": "600", "hotel_id": ""},
            {"archivo": "luz.pdf", "numero_factura": "EL-84", "fecha": "14/07/2026", "nombre_proveedor": "Energia Test SA", "NIF_proveedor": "A11111111", "descripcion_concepto": "Consumo electrico julio", "periodo_inicio": "01/07/2026", "periodo_fin": "31/07/2026", "base_imponible": 980.0, "porcentaje_iva": 21, "cuota_iva": 205.8, "total_factura": 1185.8, "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "hotel_id": ""},
        ]).to_excel(os.path.join(PROC, 'facturas_ap_20260910.xlsx'), index=False)

        # 2. almacen: fecha de registro y ajustes al leer
        df = ALM.facturas_ap()
        el = df[df['numero_factura'] == 'EL-84'].iloc[0]; cs = df[df['numero_factura'] == 'CS-84'].iloc[0]
        ok(str(el['fecha_registro']) == '2026-09-03' and str(el['fecha_contable']) == '2026-09-03' and str(cs['fecha_registro']) == '2026-09-10', f"fecha_registro = primer dia en que entro (reprocesar no la mueve): {el['fecha_registro']} / {cs['fecha_registro']}")
        ok(str(el['vencimiento']) == '2026-08-13' and int(el['dias_pago']) == 30 and not bool(el['pagada']), f"vencimiento = fecha factura + 30 dias por defecto: {el['vencimiento']}")
        ALM.guardar_ajuste_ap('EL-84', {'cuenta_contable': '628', 'dias_pago': 60}, 'tester')
        df = ALM.facturas_ap(); el = df[df['numero_factura'] == 'EL-84'].iloc[0]
        ok(str(el['cuenta_contable']) == '628' and bool(el['cuenta_ajustada']) and str(el['vencimiento']) == '2026-09-12', f"ajuste: cuenta 628 y 60 dias → vence 12/09: {el['cuenta_contable']} {el['vencimiento']}")
        try:
            ALM.guardar_ajuste_ap('EL-84', {'invento': 1}, 'tester'); ok(False, 'campo no ajustable deberia fallar')
        except ValueError:
            ok(True, 'un campo que no existe se rechaza')

        # 3. la app
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        f = cl.get('/api/ap/ficha?clave=EL-84').get_json()
        ok(f.get('ok') and f['concepto'] == 'Consumo electrico julio' and f['periodo_inicio'] == '2026-07-01' and f['periodo_fin'] == '2026-07-31' and f['nif'] == 'A11111111', 'ficha: concepto, periodo y NIF')
        a = f['asiento']
        ok([l['cuenta'] for l in a['lineas']] == ['628', '472', '400'] and a['lineas'][0]['debe'] == 980.0 and a['lineas'][1]['debe'] == 205.8 and a['lineas'][2]['haber'] == 1185.8 and a['cuadra'] and a['fecha'] == '2026-09-03', f"ficha: asiento 628/472/400 con la cuenta corregida y la fecha contable ({[l['cuenta'] for l in a['lineas']]}, {a['fecha']})")
        ok(f['criterio_fecha'] == 'contable' and f['fecha_contable'] == '2026-09-03' and f['vencimiento'] == '2026-09-12' and f['dias_pago'] == 60 and len(f['historial']) == 2 and f['historial'][0]['usuario'] == 'tester', 'ficha: criterio contable, vencimiento, dias e historial')
        ok(cl.get('/api/ap/ficha?clave=NOEXISTE').status_code == 404, 'ficha de una factura que no existe: 404')
        lista = cl.get('/api/facturas_ap').get_json()
        e = [x for x in lista if x['numero_factura'] == 'EL-84'][0]
        ok(e['concepto'] == 'Consumo electrico julio' and e['periodo_inicio'] and e['vencimiento'] == '2026-09-12' and e['pagada'] is False and e['cuenta_contable'] == '628', 'la lista de AP lleva concepto, periodo, vencimiento y pagada')
        # ajustar por API + aplicar al proveedor
        r = cl.post('/api/ap/ajustar', json={'clave': 'CS-84', 'cuenta_contable': '602', 'aplicar_proveedor': True}, headers=H).get_json()
        apr = json.load(open(os.path.join(DD, 'proveedores_aprendidos.json'), encoding='utf-8')) if os.path.exists(os.path.join(DD, 'proveedores_aprendidos.json')) else {}
        ok(r.get('ok') and r['cuenta_contable'] == '602' and r['aprendido_en'] == 'aprendidos' and apr.get('carnes test', {}).get('cuenta') == '602', f"ajustar cuenta + aplicar al proveedor → proveedores_aprendidos ({r.get('aprendido_en')}, {apr.get('carnes test')})")
        r = cl.post('/api/ap/ajustar', json={'clave': 'CS-84', 'cuenta_contable': 'abc'}, headers=H)
        ok(r.status_code == 400, 'cuenta no numerica: 400')
        r = cl.post('/api/ap/ajustar', json={'clave': 'CS-84', 'fecha_contable': '2026-10-01'}, headers=H).get_json()
        ok(r.get('ok') and r['fecha_contable'] == '2026-10-01', 'fecha contable editable')
        r = cl.post('/api/ap/ajustar', json={'clave': 'CS-84', 'fecha_contable': '2026-09-10'}, headers=H).get_json()
        ok(cl.post('/api/ap/ajustar', json={'clave': 'CS-84', 'vencimiento': 'ayer'}, headers=H).status_code == 400, 'vencimiento no valido: 400')
        # pagar
        r = cl.post('/api/ap/pagar', json={'clave': 'EL-84', 'cuenta': ''}, headers=H)
        ok(r.status_code == 400, 'pagar sin cuenta bancaria: 400')
        r = cl.post('/api/ap/pagar', json={'clave': 'EL-84', 'fecha': '2026-09-15', 'cuenta': 'BBVA principal', 'iban': 'ES12 3456', 'nota': 'transferencia'}, headers=H).get_json()
        ok(r.get('ok') and r['pagada'] and r['pagada_fecha'] == '2026-09-15' and r['pagada_cuenta'] == 'BBVA principal' and r['pagada_por'] == 'admin', f"pagar: fecha, cuenta y quien ({r.get('pagada_por')})")
        cb = cl.get('/api/cuentas_bancarias').get_json()['cuentas']
        ok(cb and cb[0]['nombre'] == 'BBVA principal' and cb[0]['iban'] == 'ES123456', 'la cuenta bancaria se recuerda para la proxima')
        ag = cl.get('/api/aging_ap').get_json()
        nums = [x['numero_factura'] for x in ag['filas']]
        ok('EL-84' not in nums and 'CS-84' in nums and ag.get('n_pagadas_mano') == 1, f"aging: la pagada a mano sale, la otra sigue ({nums})")
        csf = [x for x in ag['filas'] if x['numero_factura'] == 'CS-84'][0]
        ok(csf['vencimiento'] == '2026-10-10' and csf['dias'] == (hoy - date(2026, 10, 10)).days and csf['tramo'] == ('0-30' if csf['dias'] <= 30 else csf['tramo']), f"aging: los dias cuentan desde el vencimiento ({csf['vencimiento']}, {csf['dias']} dias)")
        lista = cl.get('/api/facturas_ap').get_json()
        ok([x for x in lista if x['numero_factura'] == 'EL-84'][0]['pagada'] is True, 'la lista marca la pagada')
        r = cl.post('/api/ap/despagar', json={'clave': 'EL-84'}, headers=H).get_json()
        ok(r.get('ok') and not r['pagada'], 'despagar deja de estar pagada')
        cl.post('/api/ap/pagar', json={'clave': 'EL-84', 'fecha': '2026-09-15', 'cuenta': 'BBVA principal'}, headers=H)

        # 4. cierre y 303 por fecha contable
        asi = cl.get('/api/cierre/asientos?mes=2026-09').get_json()
        docs = sorted({a.get('documento') for a in asi.get('asientos', [])})
        ok(asi.get('ok') and 'EL-84' in docs and 'CS-84' in docs and asi['fuentes']['ap'] == 2, f"cierre de septiembre: la factura de julio registrada el 3 de septiembre entra en SEPTIEMBRE ({docs})")
        asi7 = cl.get('/api/cierre/asientos?mes=2026-07').get_json()
        ok(asi7['fuentes']['ap'] == 0, 'y en julio no esta')
        l628 = [a for a in asi['asientos'] if a.get('documento') == 'EL-84' and str(a.get('cuenta')) == '628']
        ok(len(l628) == 1 and l628[0]['debe'] == 980.0, 'el asiento del cierre usa la cuenta corregida (628)')
        fis = cl.get('/api/fiscal?mes=2026-09').get_json()
        rec = {r_['numero'] for r_ in fis['sii']['recibidas']}
        ok('EL-84' in rec and fis['libro']['iva_soportado_472'] == round(205.8 + 20.0, 2), f"303/SII de septiembre: la factura registrada en septiembre deduce ahi ({fis['libro']['iva_soportado_472']})")
        json.dump({'ap_fecha': 'factura'}, open(os.path.join(DD, 'config_cierre.json'), 'w'))
        asi_f = cl.get('/api/cierre/asientos?mes=2026-07').get_json()
        ok(asi_f['fuentes']['ap'] == 1 and cl.get('/api/cierre/asientos?mes=2026-09').get_json()['fuentes']['ap'] == 1, 'con ap_fecha=factura vuelve al criterio antiguo (julio)')
        os.remove(os.path.join(DD, 'config_cierre.json'))
        # 5. descargas
        r = cl.get('/api/exportar/ap')
        x = pd.read_excel(io.BytesIO(r.data))
        ok(r.status_code == 200 and set(x['numero_factura']) == {'EL-84', 'CS-84'} and 'pagada' in x.columns and 'vencimiento' in x.columns and 'concepto' in x.columns and 'Proveedor A - F&B' not in str(x.values), f"Excel AP: las facturas de verdad con pagada/vencimiento/concepto ({list(x['numero_factura'])})")
        ok(str(x[x['numero_factura'] == 'EL-84']['pagada'].iloc[0]) == 'SI' and str(x[x['numero_factura'] == 'EL-84']['cuenta_bancaria'].iloc[0]) == 'BBVA principal', 'Excel AP: pagada SI y la cuenta bancaria')
        r = cl.get('/api/exportar/ap_asientos?claves=EL-84')
        xs = pd.read_excel(io.BytesIO(r.data), sheet_name=None)
        ok(r.status_code == 200 and list(xs['Asientos']['cuenta'].astype(str)) == ['628', '472', '400'] and len(xs['Facturas']) == 1, f"asientos de las marcadas: solo EL-84, 3 lineas ({list(xs['Asientos']['cuenta'])})")
        r = cl.get('/api/exportar/ap_asientos')
        ok(len(pd.read_excel(io.BytesIO(r.data), sheet_name='Asientos')) == 6, 'sin claves: los asientos de todas (6 lineas)')
        r = cl.get('/api/ap/ficha.pdf?clave=EL-84')
        ok(r.status_code == 200 and r.data[:4] == b'%PDF' and 'factura_EL-84.pdf' in r.headers.get('Content-Disposition', ''), 'PDF de la ficha')
        r = cl.get('/api/exportar/ar')
        ok(r.status_code == 200 and 'BKG-2025-06-001' not in str(pd.read_excel(io.BytesIO(r.data)).values), 'Excel AR: ya no son las tres filas inventadas')
        # 6. pantalla e i18n
        html = cl.get('/').get_data(as_text=True)
        ok(all(s in html for s in ('abrirFichaAP(', '/api/ap/ficha?clave=', 'descargarAsientosAP', '@PAGADA', 'ap-row-cb', 'data-i18n="th.pago"')), 'la pantalla: ficha, filtro pagadas, checkbox, columna Pago')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in ('ap.pagada', 'ap.fechaContable', 'ap.asiento', 'th.pago', 'ap.aplicarProv'))]
        ok(not faltan, f"i18n en los 6 idiomas (faltan {faltan})")
        ok(CM.criterio_fecha_ap({}) == 'contable' and CM.criterio_fecha_ap({'ap_fecha': 'factura'}) == 'factura' and CM.criterio_fecha_ap({'ap_fecha': 'x'}) == 'contable', 'criterio por defecto: contable')
    except Exception as e:
        ok(False, f'excepcion en la prueba: {type(e).__name__}: {e}')
    finally:
        for d in DIRS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)):
                shutil.copytree(os.path.join(copia, d), d)
        shutil.rmtree(copia, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
