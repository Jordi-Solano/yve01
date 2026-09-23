# -*- coding: utf-8 -*-
"""b86 — pestaña Caja: arqueo diario a mano, ingresos de efectivo del extracto, cuadre y cierre.

  - caja.apuntar_arqueo guarda/corrige un dia (contado, sistema, diferencia, quien) y valida
  - caja.ingresos_banco saca del extracto los ingresos CAJA (palabras + asignacion a mano)
  - caja.resumen_mes: totales del mes y "en caja" acumulado (contado − ingresado)
  - /api/caja, /api/caja/arqueo (+borrar), /api/exportar/caja (CSRF, login)
  - el cuadre de banco: la pestaña CAJA se justifica con los arqueos (CUADRA / DIFERENCIA / SIN_DATO)
  - el cierre: cada ingreso de efectivo es 572 (D) / 570 (H); el check 570 compara con lo contado;
    esos movimientos ya no cuentan como "sin conciliar"
  - la pestaña en el dashboard: boton, panel, cargador, descarga, i18n en 6 idiomas, JS que parsea

  python3.12 tests/test_caja.py
  python3.12 tests/test_caja.py --sabotaje
"""
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
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'
DIRS = ['datos-referencia', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import caja as CJ
    import cuadre_banco as CB
    import cierre_mes as CM
    if SABOTAJE:
        CJ.contado_mes = lambda *a, **k: (None, 0)      # los arqueos no justifican nada

    copia = tempfile.mkdtemp(prefix='caja_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in ('caja.xlsx', 'cuadre_banco_manual.json'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        for f in os.listdir('reportes') if os.path.isdir('reportes') else []:
            if f.startswith('conciliacion_'):
                os.remove(os.path.join('reportes', f))
        # extracto con dos ingresos de efectivo (uno por palabra, otro a mano), un cobro y un pago
        pd.DataFrame([
            {'fecha': '2026-09-03', 'concepto': 'INGRESO EFECTIVO OFICINA 0123', 'importe': 400.0, 'saldo': 10400.0, 'referencia': '', 'conciliado': ''},
            {'fecha': '2026-09-10', 'concepto': 'REMESA PROSEGUR', 'importe': 250.0, 'saldo': 10650.0, 'referencia': '', 'conciliado': ''},
            {'fecha': '2026-09-12', 'concepto': 'ABONO VARIOS 77', 'importe': 100.0, 'saldo': 10750.0, 'referencia': '', 'conciliado': ''},
            {'fecha': '2026-09-15', 'concepto': 'PAGO FRA 2026-77 CARNES', 'importe': -300.0, 'saldo': 10450.0, 'referencia': '', 'conciliado': ''},
            {'fecha': '2026-08-28', 'concepto': 'INGRESO EFECTIVO AGOSTO', 'importe': 120.0, 'saldo': 10000.0, 'referencia': '', 'conciliado': ''},
        ]).to_excel(os.path.join(DD, 'extracto_banco.xlsx'), index=False)

        # 1. el modulo puro
        f1 = CJ.apuntar_arqueo('2026-09-02', '310.50', '300', 'sobran 10,50', 'recepcion', '', DD)
        f2 = CJ.apuntar_arqueo('03/09/2026', 280, None, '', 'recepcion', '', DD)
        f3 = CJ.apuntar_arqueo('2026-09-10', 200, 215.25, 'faltan', 'admin', '', DD)
        ok(f1['diferencia'] == 10.5 and f2['diferencia'] is None and f3['diferencia'] == -15.25 and f2['fecha'] == '2026-09-03' and f1['usuario'] == 'recepcion',
           f"apuntar_arqueo: diferencia contado−sistema, fecha dd/mm aceptada, quien ({f1['diferencia']}, {f3['diferencia']})")
        CJ.apuntar_arqueo('2026-09-02', 320, 300, 'corregido', 'admin', '', DD)
        df = CJ.leer(DD)
        ok(len(df) == 3 and float(df[df['fecha'] == '2026-09-02'].iloc[0]['efectivo_contado']) == 320.0, 'corregir el mismo dia sustituye la fila (no duplica)')
        for mal in (('2026-09-02', -5), ('ayer', 10), ('2026-09-02', 'abc'), ('2026-09-02', 10, -1)):
            try:
                CJ.apuntar_arqueo(*mal, datos_dir=DD); ok(False, f'deberia rechazar {mal}'); break
            except ValueError:
                pass
        else:
            ok(True, 'rechaza contado negativo/no numerico, fecha mala y sistema negativo')
        bk, _ = __import__('almacen_datos').movimientos_banco(datos_dir=DD, reportes_dir='reportes')
        ing = CJ.ingresos_banco(bk, '2026-09-01', '2026-09-30', CB.palabras(DD))
        ok([i['importe'] for i in ing] == [400.0, 250.0] and ing[0]['fecha'] == '2026-09-03', f"ingresos_banco: los dos ingresos CAJA del mes por palabra ({[i['importe'] for i in ing]})")
        clave_abono = [m for m in ing] and __import__('almacen_datos').clave_movimiento(bk[bk['concepto'] == 'ABONO VARIOS 77'].iloc[0].to_dict())
        CB.guardar_manual(clave_abono, 'CAJA', DD)
        ing2 = CJ.ingresos_banco_dir(bk, '2026-09-01', '2026-09-30', DD)
        ok([i['importe'] for i in ing2] == [400.0, 250.0, 100.0], f"con la asignacion a mano entra el tercero ({[i['importe'] for i in ing2]})")
        res = CJ.resumen_mes('2026-09', CJ.leer(DD), bk, '', CB.palabras(DD), CB.manuales(DD))
        T = res['totales']
        ok(T['n_dias'] == 3 and T['contado'] == 800.0 and T['sistema'] == 515.25 and T['diferencia'] == 4.75 and T['dias_con_descuadre'] == 2 and T['ingresado'] == 750.0 and T['n_ingresos'] == 3,
           f"resumen_mes: contado 800, sistema 515.25, descuadre +4.75 en 2 dias, ingresado 750 ({T})")
        ok(res['en_caja'] == round(800.0 - 750.0 - 120.0, 2), f"en caja acumulado: contado hasta fin de mes − ingresado (incluye el de agosto): {res['en_caja']}")
        ok(CJ.contado_mes(CJ.leer(DD), '2026-09-01', '2026-09-30') == (800.0, 3) and CJ.contado_mes(CJ.leer(DD), '2026-10-01', '2026-10-31') == (None, 0), 'contado_mes: (800, 3) en sept; None sin arqueos')
        buf, nombre = CJ.exportar_excel(res)
        xl = pd.ExcelFile(buf)
        ok(nombre == 'caja_2026-09.xlsx' and set(xl.sheet_names) == {'Arqueos', 'Ingresos banco', 'Resumen'} and len(xl.parse('Arqueos')) == 3, 'Excel: tres hojas')

        # 2. el cuadre de banco
        cu = CB.cuadrar('2026-09', bk, None, CB.palabras(DD), CB.manuales(DD), [], caja=CJ.contado_mes(CJ.leer(DD), '2026-09-01', '2026-09-30'))
        c = cu['pestanas']['CAJA']
        ok(c['n'] == 3 and c['total'] == 750.0 and c['justificado'] == 800.0 and c['diferencia'] == -50.0 and c['estado'] == 'CUADRA' and c['n_arqueos'] == 3, f"cuadre: CAJA justificada por los arqueos (CUADRA, {c.get('justificado')}, {c.get('diferencia')})")
        c2 = CB.cuadrar('2026-09', bk, None, CB.palabras(DD), CB.manuales(DD), [], caja=(700.0, 3))['pestanas']['CAJA']
        c3 = CB.cuadrar('2026-09', bk, None, CB.palabras(DD), CB.manuales(DD), [], caja=None)['pestanas']['CAJA']
        ok(c2['estado'] == 'DIFERENCIA' and c2['diferencia'] == 50.0 and c3['estado'] == 'SIN_DATO' and 'pestaña Caja' in c3['nota'], 'cuadre: ingresado > contado → DIFERENCIA; sin arqueos → SIN_DATO con aviso')

        # 3. el cierre
        fu = CM.recoger_fuentes('2026-09', None, 'facturas-procesadas', 'reportes', DD)
        ok(len(fu.get('caja_ingresos') or []) == 3 and fu.get('caja') is not None and len(fu['caja']) == 3, f"recoger_fuentes trae los arqueos y los 3 ingresos ({len(fu.get('caja_ingresos') or [])})")
        res_c = CM.generar_asientos('2026-09', fu, CM.plan_cuentas(DD), CM.config_cierre(DD))
        a_caja = [a for a in res_c['asientos'] if a['origen'] == 'CAJA']
        ok(res_c['fuentes'].get('caja') == 3 and len(a_caja) == 6 and sum(a['debe'] for a in a_caja if a['cuenta'] == '572') == 750.0 and sum(a['haber'] for a in a_caja if a['cuenta'] == '570') == 750.0 and all(a['fecha'].startswith('2026-09') for a in a_caja),
           f"cierre: 3 asientos CAJA 572 (D) / 570 (H) por 750 ({res_c['fuentes'].get('caja')})")
        rec = CM.reconciliar('2026-09', res_c, fu, None, CM.config_cierre(DD))
        ch = {c['concepto']: c for c in rec['checks']}
        c570 = [c for c in rec['checks'] if c['cuenta'] == '570']
        ok(c570 and c570[0]['libro'] == 750.0 and c570[0]['justificado'] == 800.0 and c570[0]['diferencia'] == -50.0 and c570[0]['estado'] == 'CUADRA', f"reconciliacion 570: ingresado 750 contra contado 800 → CUADRA ({c570[0] if c570 else None})")
        pend = [c for c in rec['checks'] if c['concepto'].startswith('Movimientos del extracto sin conciliar')][0]
        ok(pend['justificado'] == -300.0 and '1 movimiento' in pend['nota'], f"los ingresos de caja ya no cuentan como sin conciliar: queda solo el pago ({pend['justificado']})")
        c572 = [c for c in rec['checks'] if c['cuenta'] == '572' and 'efectivo' in c['concepto']][0]
        ok(c572['libro'] == 750.0 and c572['justificado'] == 750.0 and c572['estado'] == 'CUADRA', 'reconciliacion 572: incluye los ingresos de efectivo')

        # 4. la app
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        d = cl.get('/api/caja?mes=2026-09').get_json()
        ok(d.get('ok') and d['totales']['n_dias'] == 3 and d['totales']['ingresado'] == 750.0 and d['en_caja'] == -70.0 and len(d['ingresos']) == 3, f"/api/caja: resumen del mes ({d.get('totales')})")
        ok(cl.post('/api/caja/arqueo', json={'fecha': '2026-09-20', 'contado': 50}).status_code == 403, 'POST sin CSRF: 403')
        r = cl.post('/api/caja/arqueo', json={'fecha': '2026-09-20', 'contado': '50', 'sistema': '', 'nota': 'domingo'}, headers=H).get_json()
        ok(r.get('ok') and r['arqueo']['efectivo_contado'] == 50.0 and r['arqueo']['efectivo_sistema'] is None and r['arqueo']['usuario'] == 'admin', f"POST arqueo: guarda con quien ({r.get('arqueo')})")
        ok(cl.post('/api/caja/arqueo', json={'fecha': '2026-09-21', 'contado': '-1'}, headers=H).status_code == 400, 'contado negativo: 400')
        ok(cl.get('/api/caja?mes=2026-09').get_json()['totales']['n_dias'] == 4, 'el nuevo arqueo sale en el resumen')
        rb = cl.post('/api/caja/arqueo/borrar', json={'fecha': '2026-09-20'}, headers=H).get_json()
        ok(rb.get('ok') and cl.get('/api/caja?mes=2026-09').get_json()['totales']['n_dias'] == 3 and cl.post('/api/caja/arqueo/borrar', json={'fecha': '2026-09-20'}, headers=H).status_code == 404, 'borrar: quita el dia; repetir → 404')
        rx = cl.get('/api/exportar/caja?mes=2026-09')
        ok(rx.status_code == 200 and 'caja_2026-09.xlsx' in rx.headers.get('Content-Disposition', ''), '/api/exportar/caja descarga el Excel')
        cb = cl.get('/api/cuadre_banco?mes=2026-09').get_json()
        ok(cb.get('ok_api') and cb['pestanas']['CAJA']['estado'] == 'CUADRA' and cb['pestanas']['CAJA']['justificado'] == 800.0, f"/api/cuadre_banco: CAJA justificada por los arqueos ({cb.get('pestanas', {}).get('CAJA', {}).get('estado')})")
        ci = cl.get('/api/cierre/asientos?mes=2026-09').get_json()
        ok(ci.get('ok') and ci['fuentes'].get('caja') == 3 and any(c['cuenta'] == '570' for c in ci['reconciliacion']['checks']), '/api/cierre/asientos: asientos de caja y check 570')
        an = cl.get('/api/caja?mes=2026-09', environ_base={}).status_code
        anon = app.test_client()
        ok(anon.get('/api/caja?mes=2026-09').status_code in (401, 302), 'sin login no hay caja')

        # 5. la pestaña
        html = cl.get('/').get_data(as_text=True)
        ip = html.index('id="panel-caja"'); fp = html.index('<!-- /panel-caja -->'); sec = html[ip:fp]
        ok('id="tab-caja"' in html and "switchTab('caja',this)" in html, 'hay pestaña Caja')
        ok(all('id="' + i + '"' in sec for i in ('caja-mes', 'caja-tiles', 'caja-k-contado', 'caja-k-encaja', 'caja-form', 'caja-f-fecha', 'caja-f-contado', 'caja-f-sistema', 'caja-f-nota', 'caja-body', 'caja-ingresos')) and sec.count('g-primary') == 1, 'panel: mes, tiles, formulario, tablas; UN primario')
        ok("caja:        function(){ return loadCaja(); }" in html and "caja:'panel-caja'" in html and "'/api/exportar/caja'" in html and "mes: 'caja-mes'" in html, 'cargador, mapa de paneles y descarga en el menu ⚙️')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in ('tab.caja', 'caja.apuntar', 'caja.guardar', 'caja.vacio', 'caja.descargar', 'caja.enCaja'))]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_caja.js', 'w', encoding='utf-8').write(b)
            if subprocess.run(['node', '--check', '/tmp/_caja.js'], capture_output=True, text=True).returncode:
                malos += 1
        ok(malos == 0, 'el JS servido parsea')
    except Exception as e:
        import traceback; traceback.print_exc()
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
