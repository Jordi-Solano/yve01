# -*- coding: utf-8 -*-
"""b109 (finanzas, 26 sep 2026): pestaña Tarjetas con un formato de liquidacion generico
(fecha, bruto, comision, neto, referencia), plantilla descargable y el cruce con los abonos
del extracto y con lo cobrado con tarjeta del DRR/TPV. Redsys/Adyen: pendientes de fichero real.

  1. leer: la plantilla, un CSV con ';' y numeros en español, cabeceras con acentos, el
     importe que falta se calcula, comision en negativo, columnas que faltan -> error; una
     liquidacion de OTA o un extracto NO se toman por liquidacion de tarjetas
  2. importar: registro con hotel y fichero; lo ya cargado no se duplica; quitar un fichero
  3. cruce con el extracto: por importe (entre el dia y 4 despues), por referencia, por la
     fecha del concepto ("REDSYS LIQ TARJETAS 20/08"); dias sin abono y abonos sin liquidacion
  4. cruce con DRR (cuentas de tarjetas del Trial Balance) y TPV (si trae la forma de pago)
  5. cuadre de banco: TARJETAS contra el neto liquidado (sin liquidaciones, como antes)
  6. el lote: la reconoce por las CABECERAS aunque el fichero se llame "export.xlsx"
  7. la app: /api/tarjetas, plantilla, Excel, quitar (CSRF), pestaña, i18n, JS

  python3.12 tests/test_tarjetas.py
  python3.12 tests/test_tarjetas.py --sabotaje
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'
DIRS = ['datos-referencia', 'reportes', 'facturas-entrada']
DD = os.path.join(BASE, 'datos-referencia')

LIQ = pd.DataFrame([
    {'fecha': '2026-08-02', 'bruto': 3140.00, 'comision': 19.60, 'neto': 3120.40, 'referencia': 'OP-0802-A', 'fecha_abono': ''},
    {'fecha': '2026-08-17', 'bruto': 4910.00, 'comision': 29.85, 'neto': 4880.15, 'referencia': 'OP-0817-A', 'fecha_abono': '2026-08-18'},
    {'fecha': '2026-08-20', 'bruto': 1000.00, 'comision': 10.00, 'neto': 990.00, 'referencia': 'OP-0820-A', 'fecha_abono': ''},
    {'fecha': '2026-08-25', 'bruto': 500.00, 'comision': 5.00, 'neto': 495.00, 'referencia': 'REM20260825', 'fecha_abono': ''},
    {'fecha': '2026-08-28', 'bruto': 300.00, 'comision': 3.00, 'neto': 297.00, 'referencia': 'OP-0828-A', 'fecha_abono': ''},
])
MOVS = [{'fecha': '2026-08-03', 'concepto': 'REDSYS LIQ TARJETAS 02/08', 'importe': 3120.40, 'clave': 'a'},
        {'fecha': '2026-08-18', 'concepto': 'REDSYS LIQ TARJETAS 17/08', 'importe': 4880.15, 'clave': 'b'},
        {'fecha': '2026-08-21', 'concepto': 'REDSYS LIQ TARJETAS 20/08', 'importe': 1000.00, 'clave': 'c'},
        {'fecha': '2026-08-26', 'concepto': 'LIQ ADYEN REM20260825', 'importe': 490.00, 'clave': 'd'},
        {'fecha': '2026-08-29', 'concepto': 'TPV ABONO 999', 'importe': 50.00, 'clave': 'e'}]


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import tarjetas as TJ
    import cuadre_banco as CB
    if SABOTAJE:
        TJ.DIAS_MARGEN = 0                               # el abono tiene que llegar el mismo dia
        TJ.es_liquidacion = lambda cab: False            # el lote no la reconoce

    # 1. leer
    buf, nombre = TJ.plantilla_excel()
    xl = pd.ExcelFile(buf)
    df, av, err = TJ.leer(buf, nombre)
    ok(nombre == 'plantilla_liquidacion_tarjetas.xlsx' and xl.sheet_names == ['Liquidacion', 'Instrucciones'] and not err and len(df) == 2
       and list(xl.parse('Liquidacion').columns)[:5] == ['fecha', 'bruto', 'comision', 'neto', 'referencia'],
       f"plantilla: hoja Liquidacion (fecha, bruto, comision, neto, referencia...) + Instrucciones; se lee ella misma ({err})")
    csv = ("Fecha operación;Importe bruto;Comisión;Importe neto;Referencia;Tarjeta\n"
           "02/08/2026;1.250,00;-7,50;;R-0001;Visa\n02/08/2026;1.878,14;;1.870,40;R-0002;Mastercard\n;;;;;\n")
    df, av, err = TJ.leer(io.BytesIO(csv.encode('utf-8')), 'liquidacion.csv')
    r = df.to_dict('records') if not err else []
    ok(not err and len(r) == 2 and r[0]['bruto'] == 1250.0 and r[0]['comision'] == 7.5 and r[0]['neto'] == 1242.5 and r[1]['comision'] == 7.74
       and r[0]['fecha'] == '2026-08-02' and r[1]['tarjeta'] == 'Mastercard', f"CSV con ';', 1.250,00, acentos; neto y comision calculados ({r})")
    ok(any('negativo' in a for a in av), f"aviso de la comision en negativo ({av})")
    _, _, err2 = TJ.leer(io.BytesIO(b"fecha,bruto,referencia\n02/08/2026,10,R\n"), 'x.csv')
    ok(err2.startswith('faltan columnas') and 'neto' in err2, f"columnas que faltan: error claro ({err2})")
    ok(not TJ.es_liquidacion(['fecha', 'numero_reserva', 'bruto', 'comision', 'neto', 'referencia'])
       and not TJ.es_liquidacion(['Fecha', 'Concepto', 'Importe', 'Saldo']) and TJ.es_liquidacion(['Fecha', 'Bruto', 'Comisión', 'Neto', 'Referencia']),
       "una liquidacion de OTA (con reserva) o un extracto no son liquidacion de tarjetas")

    # 3. cruce con el extracto (puro)
    cb = TJ.cruce_banco(LIQ, MOVS, date(2026, 8, 1), date(2026, 8, 31))
    por = {f['dia']: f for f in cb['filas']}
    ok(por.get('2026-08-02', {}).get('estado') == 'CUADRA' and por['2026-08-02']['via'] == 'importe',
       f"dia 2: neto 3.120,40 = abono del dia 3 (por importe, dentro de los 4 dias) ({por.get('2026-08-02', {}).get('via')})")
    ok(por.get('2026-08-18', {}).get('estado') == 'CUADRA', "dia 17 con fecha_abono 18: su abono del 18")
    ok(por.get('2026-08-20', {}).get('estado') == 'DIFERENCIA' and por['2026-08-20']['diferencia'] == 10.0 and por['2026-08-20']['via'] == 'fecha del concepto',
       f"dia 20: 'REDSYS LIQ TARJETAS 20/08' enlaza por la fecha del concepto y da la diferencia (+10) ({por.get('2026-08-20')})")
    ok(por.get('2026-08-25', {}).get('via') == 'referencia' and por['2026-08-25']['diferencia'] == -5.0, "dia 25: enlaza por la referencia del concepto (REM20260825), −5")
    ok(por.get('2026-08-28', {}).get('estado') == 'SIN_ABONO' and [m['importe'] for m in cb['sin_liquidacion']] == [50.0],
       "dia 28 sin abono; el abono de 50 sin liquidacion")

    # 4. cruce con DRR / TPV
    tmp = tempfile.mkdtemp(prefix='tj_')
    ruta_drr = os.path.join(tmp, 'drr_procesado_prueba.xlsx')
    pd.DataFrame([
        {'Día': 2, 'Fecha': '2026-08-02', 'Sección': 'ASSETS', 'Cuenta': 'Credit Cards Clearing', 'Débitos': 3140.0, 'Créditos': 0, 'Total': 3140.0},
        {'Día': 2, 'Fecha': '2026-08-02', 'Sección': 'ASSETS', 'Cuenta': 'Cash', 'Débitos': 1800.0, 'Créditos': 0, 'Total': 1800.0},
        {'Día': 17, 'Fecha': '2026-08-17', 'Sección': 'ASSETS', 'Cuenta': 'Visa', 'Débitos': 4900.0, 'Créditos': 0, 'Total': 4900.0},
    ]).to_excel(ruta_drr, sheet_name='Trial_Balance_Completo', index=False)
    drr = TJ.drr_por_dia(ruta_drr)
    ok(drr == {'2026-08-02': 3140.0, '2026-08-17': 4900.0}, f"DRR: cuentas de tarjetas del Trial Balance por dia (sin Cash) ({drr})")
    ventas = pd.DataFrame([{'fecha': '01/08/2026', 'nombre_plato': 'Paella', 'total_venta': 288.0}])
    tpv_pago = pd.DataFrame([{'fecha': '02/08/2026', 'total_venta': 40.0, 'forma_pago': 'Tarjeta'}, {'fecha': '02/08/2026', 'total_venta': 25.0, 'forma_pago': 'Efectivo'}])
    ok(TJ.tpv_por_dia(ventas) is None and TJ.tpv_por_dia(tpv_pago) == {'2026-08-02': 40.0}, "TPV: sin forma de pago no hay dato; con forma de pago, solo lo de tarjeta")
    cv = TJ.cruce_ventas(LIQ, drr, None, date(2026, 8, 1), date(2026, 8, 31))
    pv = {f['fecha']: f for f in cv['filas']}
    ok(pv['2026-08-02']['estado'] == 'CUADRA' and pv['2026-08-17']['diferencia'] == 10.0 and pv['2026-08-20']['estado'] == 'SIN_DATO',
       f"bruto contra DRR: dia 2 cuadra, dia 17 +10, dia 20 sin dato ({[(k, v['estado']) for k, v in pv.items()]})")

    # 5. cuadre de banco
    bk = pd.DataFrame([{'fecha': m['fecha'], 'concepto': m['concepto'], 'importe': m['importe'], 'saldo': None} for m in MOVS]
                      + [{'fecha': '2026-08-05', 'concepto': 'RECIBO ENERGIA', 'importe': -100.0, 'saldo': None}])
    pc = TJ.para_cuadre(LIQ.assign(hotel_id=''), bk, '2026-08')
    cu = CB.cuadrar('2026-08', bk, None, tarjetas=pc)['pestanas']['TARJETAS']
    cu0 = CB.cuadrar('2026-08', bk, None)['pestanas']['TARJETAS']
    ok(pc and cu['justificado'] == round(3120.40 + 4880.15 + 990 + 495 + 297, 2) and cu['total'] == 9540.55 and cu['estado'] == 'DIFERENCIA' and 'Tarjetas' in cu['nota'],
       f"cuadre: TARJETAS contra el neto liquidado ({cu.get('justificado')} vs {cu.get('total')}, {cu.get('estado')})")
    ok(cu0['estado'] == 'INFO' and TJ.para_cuadre(pd.DataFrame(columns=TJ.COLS), bk, '2026-08') is None, "sin liquidaciones, como antes (contra el TPV)")

    copia = tempfile.mkdtemp(prefix='tjc_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in (TJ.FICHERO, 'extracto_banco.xlsx', 'cuadre_banco_manual.json'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        for f in os.listdir('reportes') if os.path.isdir('reportes') else []:
            if f.startswith('drr_procesado_') or f.startswith('extracto_'):
                os.remove(os.path.join('reportes', f))
        # 2. importar
        ruta = os.path.join(tmp, 'liq.xlsx'); LIQ.to_excel(ruta, index=False)
        r1 = TJ.importar(ruta, 'H1', DD, 'liq.xlsx')
        r2 = TJ.importar(ruta, 'H1', DD, 'liq.xlsx')
        st = TJ.leer_store(DD)
        ok(r1['ok'] and r1['nuevas'] == 5 and r2['nuevas'] == 0 and r2['repetidas'] == 5 and len(st) == 5 and set(st['hotel_id']) == {'H1'} and set(st['fichero']) == {'liq.xlsx'},
           f"importar: 5 nuevas, la segunda vez 0 (5 ya estaban); con hotel y fichero ({r1.get('nuevas')}, {r2.get('repetidas')})")
        ok(TJ.quitar_fichero('liq.xlsx', 'H1', DD) == 5 and TJ.leer_store(DD).empty, "quitar el fichero deja el registro vacio")

        # 6-7. la app y el lote
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        pd.DataFrame([{'fecha': m['fecha'], 'concepto': m['concepto'], 'importe': m['importe'], 'saldo': 1000.0 + i} for i, m in enumerate(MOVS)]).to_excel(
            os.path.join(DD, 'extracto_banco.xlsx'), index=False)
        os.makedirs('reportes', exist_ok=True); shutil.copy(ruta_drr, os.path.join('reportes', 'drr_procesado_prueba.xlsx'))
        gen = os.path.join(tmp, 'export_0925.xlsx')
        LIQ.rename(columns={'bruto': 'Importe bruto', 'comision': 'Comisión', 'neto': 'Importe neto', 'referencia': 'Referencia', 'fecha': 'Fecha'}).to_excel(gen, index=False)
        ok(D._destino_capa1('export_0925.xlsx', gen) == 'TARJETAS', "capa 1: la reconoce por las cabeceras aunque el nombre no diga nada")

        def lote(n, p):
            with open(p, 'rb') as fh:
                cl.post('/api/upload_facturas', data={'files': [(fh, n)]}, content_type='multipart/form-data')
            q = '/api/procesar_batch_stream?archivos=' + urllib.parse.quote(json.dumps([n]))
            rr = cl.get(q); txt = rr.get_data(as_text=True); rr.close()
            return txt
        t1 = lote('export_0925.xlsx', gen)
        ok('✓ Tarjetas export_0925.xlsx: 5 operación(es) nuevas' in t1 and '· bruto 9.850,00 · comisión 67,45 · neto 9.782,55' in t1 and 'liquidación(es) de tarjetas' in t1,
           f"lote: {[l for l in t1.splitlines() if 'Tarjetas' in l][:1]}")
        t2 = lote('export_0925.xlsx', gen)
        ok(('0 operación(es) nuevas (5 ya estaban)' in t2 and '· bruto 0,00' not in t2) or 'Saltando' in t2, f"otra vez: no duplica ({[l for l in t2.splitlines() if 'Tarjetas' in l or 'Saltando' in l][:1]})")
        ok(len(TJ.leer_store(DD)) == 5, f"registro con 5 operaciones ({len(TJ.leer_store(DD))})")
        d = cl.get('/api/tarjetas?mes=2026-08').get_json()
        T = d.get('totales') or {}
        ok(d.get('ok') and T.get('n') == 5 and T.get('bruto') == 9850.0 and T.get('comision') == 67.45 and T.get('neto') == 9782.55 and T.get('n_abonos') == 4 and T.get('dias_sin_abono') == 1
           and T.get('abonos_sin_liquidacion') == 1, f"/api/tarjetas: totales y cruce ({T})")
        ok(T.get('ventas_tarjeta') == 8040.0 and d['ventas']['hay_drr'], f"/api/tarjetas: lo cobrado con tarjeta segun el DRR del hotel ({T.get('ventas_tarjeta')})")
        rp = cl.get('/api/tarjetas/plantilla')
        ok(rp.status_code == 200 and 'plantilla_liquidacion_tarjetas.xlsx' in rp.headers.get('Content-Disposition', ''), "plantilla descargable")
        rx = cl.get('/api/exportar/tarjetas?mes=2026-08')
        ok(rx.status_code == 200 and set(pd.ExcelFile(io.BytesIO(rx.data)).sheet_names) >= {'Operaciones', 'Cruce banco', 'Abonos sin liquidacion', 'Cruce ventas', 'Resumen'},
           "Excel del cruce con sus hojas")
        cbj = cl.get('/api/cuadre_banco?mes=2026-08').get_json()
        tp = (cbj.get('pestanas') or {}).get('TARJETAS', {})
        ok(tp.get('justificado') == 9782.55 and tp.get('n_dias_liquidados') == 5, f"/api/cuadre_banco: TARJETAS contra la liquidacion ({tp.get('justificado')}, {tp.get('estado')})")
        ok(cl.post('/api/tarjetas/quitar', json={'fichero': 'export_0925.xlsx'}).status_code == 403, "quitar sin CSRF: 403")
        rq = cl.post('/api/tarjetas/quitar', json={'fichero': 'export_0925.xlsx'}, headers=H).get_json()
        ok(rq.get('ok') and rq.get('quitadas') == 5 and TJ.leer_store(DD).empty, f"quitar el fichero desde la pestaña ({rq})")
        html = cl.get('/').get_data(as_text=True)
        ip = html.index('<div id="panel-tarjetas"'); sec = html[ip:html.index('<!-- /panel-tarjetas -->')]
        ok('id="tab-tarjetas"' in html and "switchTab('tarjetas',this)" in html and all('id="' + i + '"' in sec for i in ('tarj-mes', 'tarj-tiles', 'tarj-banco-body', 'tarj-ventas-body', 'tarj-ops-body', 'tarj-ficheros'))
           and 'href="/api/tarjetas/plantilla"' in sec, "pestaña Tarjetas: mes, tiles, los dos cruces, operaciones, plantilla")
        ok("tarjetas:    function(){ return loadTarjetas(); }" in html and "'caja', 'tarjetas'];" in html and "'/api/exportar/tarjetas'" in html and "tarjetas:'panel-tarjetas'" in html,
           "cargador, precarga, descargas del menu ⚙️ y mapa de paneles")
        ok('openUploadModal(' not in sec and 'type="file"' not in sec, "la pestaña no sube nada: entra por ⚡ Procesar archivos")
        claves = set(re.findall(r"t\('(tarj\.[A-Za-z]+)'", html)) | set(re.findall(r'data-i18n(?:-title)?="(tarj\.[A-Za-z]+)"', html)) | {'tab.tarjetas'}
        faltan = {l: sorted(k for k in claves if k not in json.load(open(f'static/i18n/{l}.json', encoding='utf-8'))) for l in ('en', 'ca', 'fr', 'de', 'it', 'pt')}
        ok(len(claves) > 30 and not any(faltan.values()), f"i18n: {len(claves)} claves en los 6 idiomas (faltan {[(l, v[:3]) for l, v in faltan.items() if v]})")
        i = html.index('var _sseFrags = ['); j = html.index('function _tSSE', i)
        open('/tmp/_tjsse.js', 'w', encoding='utf-8').write(html[i:j] + "\nconsole.log(JSON.stringify(Object.keys(_sseTrans).every(function(k){ return _sseTrans[k].length === _sseFrags.length; })));")
        ok(subprocess.run(['node', '/tmp/_tjsse.js'], capture_output=True, text=True).stdout.strip() == 'true', "mensajes del lote traducidos en los 6 idiomas (tablas del mismo largo)")
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_tj.js', 'w', encoding='utf-8').write(b)
            if subprocess.run(['node', '--check', '/tmp/_tj.js'], capture_output=True, text=True).returncode:
                ok(False, 'JS roto'); break
    except Exception as e:
        import traceback; traceback.print_exc()
        ok(False, f'excepcion en la prueba: {type(e).__name__}: {e}')
    finally:
        for d in DIRS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)):
                shutil.copytree(os.path.join(copia, d), d)
        shutil.rmtree(copia, ignore_errors=True)
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
