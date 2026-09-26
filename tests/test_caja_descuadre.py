# -*- coding: utf-8 -*-
"""b108 (finanzas, 26 sep 2026): "el descuadre (sobra/falta) si se contabiliza, en una cuenta
propia de 'overs & shorts'. Añadela al plan de cuentas (configurable en config_cierre.json),
asiento automatico al guardar el arqueo, y que salga en el cierre."

  - caja.asiento_descuadre: falta = overs&shorts (D) / 570 (H); sobra = 570 (D) / overs&shorts (H);
    sin efectivo segun sistema, o sin descuadre, no hay asiento
  - la cuenta: 6591 por defecto, en el plan; config_cierre.json -> cuenta_descuadre_caja la
    cambia y la cuenta nueva entra sola en el plan
  - el cierre: un asiento CAJA_DESCUADRE por arqueo con descuadre (solo del hotel), el diario
    cuadra, sale en el mayor y en la reconciliacion (check propio); el check 570 sigue
    contando solo lo ingresado en el banco
  - la app: POST /api/caja/arqueo devuelve el asiento; /api/caja lo trae en cada dia;
    /api/cierre/asientos lo asienta; la pestaña enseña la columna Asiento

  python3.12 tests/test_caja_descuadre.py
  python3.12 tests/test_caja_descuadre.py --sabotaje
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
    import cierre_mes as CM
    if SABOTAJE:
        CJ.asiento_descuadre = lambda fila, cuenta: None      # vuelve el "no se asienta solo"

    # 1. el asiento (puro)
    falta = CJ.asiento_descuadre({'fecha': '2026-09-10', 'efectivo_contado': 200, 'efectivo_sistema': 215.25, 'diferencia': -15.25}, '6591')
    sobra = CJ.asiento_descuadre({'fecha': '2026-09-02', 'efectivo_contado': 320, 'efectivo_sistema': 300}, '6591')
    ok(falta and [tuple(l) for l in falta['lineas']] == [('6591', 15.25, 0.0), ('570', 0.0, 15.25)] and falta['tipo'] == 'falta',
       f"falta: 6591 (D) / 570 (H) por 15,25 ({falta and falta['lineas']})")
    ok(sobra and [tuple(l) for l in sobra['lineas']] == [('570', 20.0, 0.0), ('6591', 0.0, 20.0)] and sobra['tipo'] == 'sobra',
       f"sobra: 570 (D) / 6591 (H) por 20 ({sobra and sobra['lineas']})")
    nada = [CJ.asiento_descuadre({'fecha': '2026-09-03', 'efectivo_contado': 280, 'efectivo_sistema': None}, '6591'),
            CJ.asiento_descuadre({'fecha': '2026-09-04', 'efectivo_contado': 100, 'efectivo_sistema': 100}, '6591')]
    ok(nada == [None, None], "sin efectivo segun sistema, o cuadrado, no hay asiento")
    ok(CJ.asiento_texto(falta) == '6591 D 15,25 / 570 H 15,25' and 'faltan 15,25 EUR' in (falta or {}).get('concepto', ''),
       f"texto del asiento y concepto en formato español ({CJ.asiento_texto(falta)})")

    copia = tempfile.mkdtemp(prefix='cajad_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in ('caja.xlsx', 'cuadre_banco_manual.json', 'config_cierre.json', 'extracto_banco.xlsx'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        # 2. la cuenta, en el plan
        ok(CM.config_cierre(DD)['cuenta_descuadre_caja'] == '6591' and 'overs' in CM.plan_cuentas(DD).get('6591', '').lower(),
           "6591 por defecto y en el plan (Diferencias de caja, overs & shorts)")
        json.dump({'cuenta_descuadre_caja': '6781'}, open(os.path.join(DD, 'config_cierre.json'), 'w'))
        ok(CM.config_cierre(DD)['cuenta_descuadre_caja'] == '6781' and 'overs' in CM.plan_cuentas(DD).get('6781', '').lower(),
           "configurable: con 6781 en config_cierre.json, la 6781 entra sola en el plan")

        # 3. el cierre
        arq = pd.DataFrame([
            {'fecha': '2026-09-02', 'hotel_id': '', 'efectivo_sistema': 300.0, 'efectivo_contado': 320.0, 'diferencia': 20.0},
            {'fecha': '2026-09-03', 'hotel_id': '', 'efectivo_sistema': None, 'efectivo_contado': 280.0, 'diferencia': None},
            {'fecha': '2026-09-10', 'hotel_id': '', 'efectivo_sistema': 215.25, 'efectivo_contado': 200.0, 'diferencia': -15.25},
            {'fecha': '2026-09-11', 'hotel_id': 'OTRO', 'efectivo_sistema': 100.0, 'efectivo_contado': 90.0, 'diferencia': -10.0},
            {'fecha': '2026-08-30', 'hotel_id': '', 'efectivo_sistema': 50.0, 'efectivo_contado': 45.0, 'diferencia': -5.0},
        ])
        fu = {'caja': arq, 'hotel': '', 'caja_ingresos': [{'fecha': '2026-09-05', 'concepto': 'INGRESO EFECTIVO', 'importe': 400.0, 'clave': 'k1', 'estado': 'PENDIENTE', 'hotel_id': ''}]}
        cfg = CM.config_cierre('/nonexistent')
        res = CM.generar_asientos('2026-09', fu, CM.plan_cuentas('/nonexistent'), cfg)
        ad = [a for a in res['asientos'] if a['origen'] == 'CAJA_DESCUADRE']
        ok(res['fuentes'].get('caja_descuadres') == 3 and len(ad) == 6 and res['cuadra'],
           f"grupo: 3 asientos de descuadre (sept, los tres con sistema y descuadre, tambien el del otro hotel), diario cuadrado ({res['fuentes'].get('caja_descuadres')})")
        fu_h = {**fu, 'hotel': 'OTRO'}
        res_h = CM.generar_asientos('2026-09', fu_h, None, cfg)
        ok(res_h['fuentes'].get('caja_descuadres') == 1, f"con hotel activo solo sus arqueos ({res_h['fuentes'].get('caja_descuadres')})")
        fu1 = {**fu, 'caja': arq[arq['hotel_id'] == '']}
        res1 = CM.generar_asientos('2026-09', fu1, CM.plan_cuentas('/nonexistent'), cfg)
        my = {m['cuenta']: m for m in CM.mayor(res1['asientos'])}
        ok(my.get('6591', {}).get('debe') == 15.25 and my.get('6591', {}).get('haber') == 20.0 and 'overs' in my['6591']['descripcion'].lower(),
           f"mayor 6591: 15,25 al debe (falta) y 20 al haber (sobra) ({my.get('6591')})")
        ok(not res1['cuentas_fuera_plan'], f"ninguna cuenta fuera del plan ({res1['cuentas_fuera_plan']})")
        rec = CM.reconciliar('2026-09', res1, fu1, None, cfg)
        cod = [c for c in rec['checks'] if c['cuenta'] == '6591']
        ok(cod and cod[0]['libro'] == 4.75 and cod[0]['justificado'] == 4.75 and cod[0]['estado'] == 'CUADRA',
           f"reconciliacion: overs & shorts asentado (+4,75 neto) contra los descuadres de los arqueos ({cod and cod[0]})")
        c570 = [c for c in rec['checks'] if c['cuenta'] == '570']
        ok(c570 and c570[0]['libro'] == 400.0, f"el check 570 sigue contando solo lo ingresado en el banco ({c570 and c570[0]['libro']})")
        res_c = CM.generar_asientos('2026-09', fu1, None, {**cfg, 'cuenta_descuadre_caja': '6781'})
        ok(any(a['cuenta'] == '6781' for a in res_c['asientos']) and not any(a['cuenta'] == '6591' for a in res_c['asientos']),
           "con otra cuenta configurada, el asiento va a esa")

        # 4. la app
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        os.remove(os.path.join(DD, 'config_cierre.json'))
        r = cl.post('/api/caja/arqueo', json={'fecha': '2026-09-10', 'contado': '200', 'sistema': '215.25', 'nota': 'faltan'}, headers=H).get_json()
        ok(r.get('ok') and r.get('cuenta_descuadre') == '6591' and r.get('asiento') and [list(l) for l in r['asiento']['lineas']] == [['6591', 15.25, 0.0], ['570', 0.0, 15.25]],
           f"al guardar el arqueo sale su asiento ({r.get('asiento')})")
        r2 = cl.post('/api/caja/arqueo', json={'fecha': '2026-09-12', 'contado': '90'}, headers=H).get_json()
        ok(r2.get('ok') and r2.get('asiento') is None, "sin efectivo segun sistema: guardado, sin asiento")
        d = cl.get('/api/caja?mes=2026-09').get_json()
        dias = {x['fecha']: x for x in d.get('dias') or []}
        ok(d.get('cuenta_descuadre') == '6591' and dias.get('2026-09-10', {}).get('asiento') and dias.get('2026-09-12', {}).get('asiento') is None,
           "/api/caja: cada dia con su asiento (o nada)")
        ci = cl.get('/api/cierre/asientos?mes=2026-09').get_json()
        a_ci = [a for a in (ci.get('asientos') or []) if a.get('origen') == 'CAJA_DESCUADRE']
        ok(ci.get('ok') and ci['fuentes'].get('caja_descuadres') == 1 and {a['cuenta'] for a in a_ci} == {'6591', '570'}
           and any(c['cuenta'] == '6591' for c in ci['reconciliacion']['checks']), f"/api/cierre/asientos: el descuadre esta en el diario y en la reconciliacion ({ci.get('fuentes', {}).get('caja_descuadres')})")
        rx = cl.get('/api/exportar/caja?mes=2026-09')
        hoja = pd.read_excel(__import__('io').BytesIO(rx.data), sheet_name='Arqueos')
        ok(rx.status_code == 200 and 'asiento' in hoja.columns and '6591 D 15,25 / 570 H 15,25' in hoja['asiento'].astype(str).tolist(),
           "el Excel de caja lleva el asiento de cada arqueo")
        html = cl.get('/').get_data(as_text=True)
        ok("function _cajaAsientoTxt(a)" in html and "t('caja.asiento', 'Asiento')" in html and "t('caja.guardadoAsiento'" in html and "f.caja_descuadres" in html,
           "la pestaña enseña la columna Asiento, el aviso al guardar lo dice y el cierre cuenta los descuadres")
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8'))
                                                                            for k in ('caja.asiento', 'caja.asientoTit', 'caja.asentadoEn', 'caja.guardadoAsiento', 'cierre.fuentesDescCaja'))]
        ok(not faltan, f"i18n en los 6 idiomas (faltan {faltan})")
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_cajad.js', 'w', encoding='utf-8').write(b)
            if subprocess.run(['node', '--check', '/tmp/_cajad.js'], capture_output=True, text=True).returncode:
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
