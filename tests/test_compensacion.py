# -*- coding: utf-8 -*-
"""b89 — compensar la comision de la agencia contra la factura del grupo (410/430).

Reglas de finanzas (24 sep 2026) que se comprueban:
  - SOLO se compensa si quien paga la factura del grupo es la AGENCIA. Si paga el cliente
    final NO hay compensacion: primero se cobra la factura del grupo y despues se paga la
    de comision (la ficha AP no deja marcarla pagada antes: 409).
  - lo compensado se asienta 410 (D) / 430 (H); el cobro de la factura del grupo es por lo
    que queda (572/430 por total - compensado) y la 430 del grupo queda a cero.
  - maximo = el menor de los dos saldos; total o parcial; se puede anular mientras la
    factura del grupo no este cobrada ni la de comision pagada.
  - lo compensado baja el saldo en AR (aging, cliente) y en AP (aging; a cero, sale).
  - sin saber quien paga, la factura del grupo no se emite (409) ni se compensa.
  - pantalla: bloque de compensacion en AR › Contratos, emitir la pendiente, CSRF, i18n.

  python3.12 tests/test_compensacion.py
  python3.12 tests/test_compensacion.py --sabotaje
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')
PROC = os.path.join(BASE, 'facturas-procesadas')
HOY = date.today()
MES = HOY.strftime('%Y-%m')


def contrato(numero, evento, agencia, nif, pct, pagador):
    return {"es_contrato_grupo": True, "evento": {"id": "", "nombre": evento}, "contrato_numero": numero,
            "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000001"},
            "agencia": {"nombre": agencia, "cif": nif},
            "alojamiento": {"fecha_entrada": "2026-08-07", "fecha_salida": "2026-08-10", "noches": 3, "habitaciones": 20,
                            "total_habitaciones": 11000, "iva_pct": 10},
            "fb": {"total": 0}, "salas": {"total": 0},
            "comisiones": {"modo": None, "alojamiento_pct": pct}, "facturacion": {"pagador": pagador}}


def fila_ap(num, prov, nif, concepto, base, fecha):
    return {"archivo": f"{num}.pdf", "numero_factura": num, "nombre_proveedor": prov, "NIF_proveedor": nif,
            "fecha_factura": fecha, "descripcion_concepto": concepto, "base_imponible": base, "porcentaje_iva": 21,
            "cuota_iva": round(base * .21, 2), "total_factura": round(base * 1.21, 2), "tipo_proveedor": "OTRAS",
            "cuenta_contable": "629", "hotel_id": ""}


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import compensaciones as CMP
    import contratos_grupo as CG
    import lector_contratos_grupo as L
    import cierre_mes as CM
    import almacen_datos as ALM
    import aging_ap as AG
    import tab_contratos as TC
    if SABOTAJE:
        # la regla de finanzas desaparece: se compensa aunque pague el cliente final
        _orig = CMP.puede_compensar
        CMP.puede_compensar = lambda c, *a, **k: _orig(dict(c, pagador={"quien": "agencia"}), *a, **k)

    copia = tempfile.mkdtemp(prefix='cmp_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in os.listdir(PROC):
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PROC, f))
        for f in ('contratos_grupo.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx', 'ajustes_ap.json',
                  'extracto_banco.xlsx', 'compensaciones_ar.json'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        # A: paga la agencia · B: paga el cliente final · C: el contrato no dice quien paga
        for dt in (contrato("CG-2026-0901", "Congreso Cardio", "Viajes Meridiano S.L.", "B28004556", 10, "agencia"),
                   contrato("CG-2026-0902", "Incentivo Pharma", "Eventos Costa DMC", "B66123456", 10, "cliente"),
                   contrato("CG-2026-0903", "Foro Seguros", "Grupos Levante S.L.", "B46999888", 8, None)):
            t = L.transformar(dt, hotel_id="")
            L.guardar(t, DD)
            CG.registrar(dt, t, hotel_id="", datos_dir=DD)
        for c in CG.leer(DD):
            CG.sincronizar_factura(c, DD)
        hoy_es = HOY.strftime('%d/%m/%Y')
        pd.DataFrame([
            fila_ap("VM-2026-501", "VIAJES MERIDIANO SL", "B28004556", "Comisión grupo contrato CG-2026-0901", 1000.0, hoy_es),
            fila_ap("EC-2026-88", "Eventos Costa DMC", "B66123456", "Comisión grupo contrato CG-2026-0902", 1000.0, hoy_es),
            fila_ap("GL-7", "Grupos Levante SL", "B46999888", "Comisión contrato CG-2026-0903", 800.0, hoy_es),
        ]).to_excel(os.path.join(PROC, f"facturas_ap_{HOY.strftime('%Y%m%d')}.xlsx"), index=False)

        por_ctr = {c['contrato']: c for c in TC.vista_contratos('', DD)}
        a, b, c = por_ctr['CG-2026-0901'], por_ctr['CG-2026-0902'], por_ctr['CG-2026-0903']
        ok(a['factura_comision_vista'].get('numero') == 'VM-2026-501' and b['factura_comision_vista'].get('numero') == 'EC-2026-88',
           'cada factura de comision esta unida a su contrato')
        # 1. las reglas, antes de emitir
        ok(not a['compensacion']['puede'] and 'emitir' in a['compensacion']['motivo'],
           f"paga la agencia pero la factura del grupo no esta emitida: aun no se compensa ({a['compensacion']['motivo']})")
        ok(b['compensacion']['regla'] == 'cliente' and not b['compensacion']['puede'] and 'cliente final' in b['compensacion']['motivo'],
           f"paga el cliente final: NO se compensa ({b['compensacion']['motivo']})")
        ok(not c['compensacion']['puede'] and 'quién paga' in c['compensacion']['motivo'],
           f"el contrato no dice quien paga: no se supone ({c['compensacion']['motivo']})")

        # 2. la app
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        r = cl.post('/api/ar_real/emitir_pendiente', json={'numero': 'GRP-CG-2026-0903'}, headers=H)
        ok(r.status_code == 409 and 'quién paga' in (r.get_json() or {}).get('error', ''), 'sin saber quien paga, la factura del grupo NO se emite (409)')
        for n in ('GRP-CG-2026-0901', 'GRP-CG-2026-0902'):
            r = cl.post('/api/ar_real/emitir_pendiente', json={'numero': n}, headers=H)
            ok(r.status_code == 200 and (r.get_json() or {}).get('ok'), f'emitir la factura del grupo {n}')
        r = cl.post('/api/ar_real/emitir_pendiente', json={'numero': 'GRP-CG-2026-0901'}, headers=H)
        ok(r.status_code == 409, 'emitir dos veces: 409')
        bj = cl.get('/api/ar_real/bonos').get_json()
        ok(not [x for x in (bj.get('facturas_sin_bono') or []) if str(x.get('numero', '')).startswith('GRP-')],
           'la factura emitida de un contrato de grupo no sale como "sin bono" (la respalda el contrato)')
        ida = a['id']
        ok(cl.post('/api/ar/compensar', json={'id': ida, 'importe': 100}).status_code == 403, 'compensar sin CSRF: 403')
        r = cl.post('/api/ar/compensar', json={'id': b['id'], 'importe': 100, 'fecha': HOY.isoformat()}, headers=H)
        ok(r.status_code == 409 and 'cliente final' in (r.get_json() or {}).get('error', ''),
           f"paga el cliente final: compensar da 409 con la regla ({r.status_code} {(r.get_json() or {}).get('error', '')[:60]})")
        r = cl.post('/api/ar/compensar', json={'id': ida, 'importe': 5000, 'fecha': HOY.isoformat()}, headers=H)
        ok(r.status_code == 400 and '1210' in (r.get_json() or {}).get('error', '').replace('.', '').replace(',', ''),
           f"mas que el menor de los saldos: 400 ({(r.get_json() or {}).get('error')})")
        ok(cl.post('/api/ar/compensar', json={'id': ida, 'importe': 0}, headers=H).status_code == 400, 'importe 0: 400')
        r = cl.post('/api/ar/compensar', json={'id': ida, 'importe': 500, 'fecha': HOY.isoformat(), 'nota': 'parcial'}, headers=H).get_json()
        ok(r.get('ok') and r['contrato']['compensacion']['total'] == 500.0 and r['contrato']['compensacion']['maximo'] == 710.0,
           f"compensacion parcial de 500: quedan 710 por compensar ({(r.get('contrato') or {}).get('compensacion', {}).get('maximo')})")
        r = cl.post('/api/ar/compensar', json={'id': ida, 'importe': 710, 'fecha': HOY.isoformat()}, headers=H).get_json()
        cp = (r.get('contrato') or {}).get('compensacion') or {}
        ok(r.get('ok') and cp.get('total') == 1210.0 and not cp.get('puede') and len(cp.get('lista') or []) == 2,
           f"el resto (710): la comision queda compensada entera y ya no se puede compensar mas ({cp.get('motivo')})")
        # reflejo en AR y AP
        rv = pd.read_excel(os.path.join(DD, 'reservas_credito.xlsx'))
        fa = rv[rv['numero_reserva'] == 'GRP-CG-2026-0901'].iloc[0]
        ok(round(float(fa['compensado']), 2) == 1210.0, f"factura del grupo: compensado 1.210 ({fa.get('compensado')})")
        aj = ALM.ajustes_ap(DD).get('VM-2026-501') or {}
        ok(aj.get('compensado') == 1210.0, f"factura de comision: compensado 1.210 en sus ajustes ({aj.get('compensado')})")
        fap = next(x for x in ALM.facturas_ap().to_dict('records') if x['numero_factura'] == 'VM-2026-501')
        ok(fap.get('compensado') == 1210.0 and bool(fap.get('compensada')), 'AP: la factura de comision sale compensada')
        ag = AG.calcular_aging(ALM.facturas_ap(), hoy=HOY)
        ok(not [x for x in ag['filas'] if x['numero_factura'] == 'VM-2026-501'] and ag['n_compensadas'] == 1,
           'aging AP: la comision compensada entera ya no se debe')
        eb = next((x for x in ag['filas'] if x['numero_factura'] == 'EC-2026-88'), {})
        ok(eb.get('tras_cobro') == 'GRP-CG-2026-0902', f"aging AP: la del cliente final dice que se paga tras cobrar su grupo ({eb.get('tras_cobro')})")
        fr = cl.get('/api/ar_real/facturas').get_json()
        ga = next(x for x in fr['facturas'] if x['numero'] == 'GRP-CG-2026-0901')
        ok(ga['saldo'] == 9790.0 and ga['compensado'] == 1210.0, f"AR: la factura del grupo queda en 9.790 ({ga['saldo']})")
        ok(round(fr['stats']['pendiente'], 2) == 9790.0 + 11000.0, f"AR: lo pendiente es el saldo, no el total ({fr['stats']['pendiente']})")
        cli = next((x for x in cl.get('/api/ar_real/clientes').get_json()['clientes'] if x['nombre'] == 'Viajes Meridiano S.L.'), {})
        ok(cli.get('saldo_pendiente') == 9790.0, f"ficha del cliente: su saldo es 9.790 ({cli.get('saldo_pendiente')})")
        # no se cambia a "paga el cliente" con compensaciones hechas
        r = cl.post('/api/ar/contratos/decidir', json={'id': ida, 'pagador': 'cliente'}, headers=H)
        ok(r.status_code == 409, 'con compensaciones hechas no se cambia a "paga el cliente" (409)')
        # 3. paga el cliente final: primero se cobra el grupo, despues se paga la comision
        fi = cl.get('/api/ap/ficha?clave=EC-2026-88').get_json()
        ok((fi.get('comision') or {}).get('pagador') == 'cliente' and not fi['comision']['se_puede_pagar'],
           'ficha: la comision del cliente final aun no se puede pagar')
        r = cl.post('/api/ap/pagar', json={'clave': 'EC-2026-88', 'fecha': HOY.isoformat(), 'cuenta': 'BBVA principal'}, headers=H)
        ok(r.status_code == 409 and 'Primero se cobra' in (r.get_json() or {}).get('error', ''),
           f"marcarla pagada antes de cobrar el grupo: 409 ({(r.get_json() or {}).get('error', '')[:70]})")
        ok(cl.post('/api/ar_real/cobrar', json={'numero': 'GRP-CG-2026-0902'}, headers=H).get_json().get('ok'), 'se cobra la factura del grupo B')
        r = cl.post('/api/ap/pagar', json={'clave': 'EC-2026-88', 'fecha': HOY.isoformat(), 'cuenta': 'BBVA principal'}, headers=H)
        ok(r.status_code == 200 and (r.get_json() or {}).get('pagada'), 'y ahora SI se paga la comision')
        # 4. cobro de A por lo que queda; ya no se anula
        rc = cl.post('/api/ar_real/cobrar', json={'numero': 'GRP-CG-2026-0901'}, headers=H).get_json()
        ok(rc.get('ok') and rc.get('cobrado') == 9790.0 and rc.get('compensado') == 1210.0, f"cobro de A: 9.790 ({rc.get('cobrado')})")
        comp_id = cp['lista'][0]['id']
        r = cl.post('/api/ar/compensar/anular', json={'id': ida, 'comp_id': comp_id}, headers=H)
        ok(r.status_code == 409, 'con la factura del grupo cobrada la compensacion ya no se anula (409)')

        # 5. el cierre del mes
        fu = CM.recoger_fuentes(MES, None, PROC, 'reportes', DD)
        res = CM.generar_asientos(MES, fu, CM.plan_cuentas(DD), CM.config_cierre(DD))
        cmp_as = [x for x in res['asientos'] if x['origen'] == 'COMPENSACION']
        ok(sorted((x['cuenta'], x['debe'], x['haber']) for x in cmp_as) == [('410', 500.0, 0.0), ('410', 710.0, 0.0), ('430', 0.0, 500.0), ('430', 0.0, 710.0)],
           f"asientos de compensacion 410 (D) / 430 (H) ({[(x['cuenta'], x['debe'], x['haber']) for x in cmp_as]})")
        cob = [x for x in res['asientos'] if x['origen'] == 'AR' and x['documento'] == 'GRP-CG-2026-0901' and x['concepto'].startswith('Cobro')]
        ok(sorted((x['cuenta'], x['debe'], x['haber']) for x in cob) == [('430', 0.0, 9790.0), ('572', 9790.0, 0.0)],
           f"el cobro de A es por lo que quedaba: 572/430 9.790 ({[(x['cuenta'], x['debe'], x['haber']) for x in cob]})")
        s430 = round(sum(x['debe'] - x['haber'] for x in res['asientos'] if x['cuenta'] == '430' and x['documento'] == 'GRP-CG-2026-0901'), 2)
        ok(s430 == 0.0, f"la 430 de la factura del grupo A queda a cero ({s430})")
        ok(res['fuentes'].get('compensaciones') == 2 and res['cuadra'], 'el diario cuadra y cuenta las 2 compensaciones')
        rec = CM.reconciliar(MES, res, fu, None, CM.config_cierre(DD))
        ch = [x for x in rec['checks'] if x['cuenta'] == '410/430']
        ok(ch and ch[0]['estado'] == 'CUADRA' and ch[0]['libro'] == 1210.0, f"reconciliacion 410/430 cuadra ({ch[0] if ch else None})")
        ok(not [x for x in rec['checks'] if x['cuenta'] in ('400', '410') and x['estado'] == 'DIFERENCIA'], 'la 400 y la 410 siguen cuadrando')
        # funciones puras: anular con el grupo sin cobrar si se puede
        CMP._escribir([], DD)
        ctr_a = next(x for x in CG.leer(DD) if x['contrato'] == 'CG-2026-0901')
        fg = {'numero': 'GRP-X', 'numero_reserva': 'GRP-X', 'hotel_id': '', 'estado': 'FACTURADO', 'total': 2000.0}
        fc = {'numero_factura': 'VM-X', 'total_factura': 300.0}
        ok(CMP.puede_compensar(ctr_a, fg, fc, [], []) == (True, 300.0), 'maximo = el menor de los dos saldos (300)')
        ok(CMP.puede_compensar(ctr_a, dict(fg, estado='COBRADO'), fc, [], [])[0] is False, 'factura del grupo cobrada: no se compensa')
        ok(CMP.puede_compensar(ctr_a, fg, dict(fc, pagada=True), [], [])[0] is False, 'factura de comision pagada: no se compensa')
        ok(CMP.puede_compensar(ctr_a, fg, None, [], [])[0] is False, 'sin factura de comision: no se compensa')

        # 6. la pantalla
        html = cl.get('/').get_data(as_text=True)
        ok(all(x in html for x in ('function _ctrCompensacion(', 'function compensarAR(', 'function anularCompensacionAR(',
                                   "'/api/ar/compensar'", "'/api/ar/compensar/anular'", 'function emitirPendienteAR(',
                                   "'/api/ar_real/emitir_pendiente'", 'ctr-cita')),
           'JS: compensar/anular en contratos, emitir la pendiente y la cita del contrato fuera de la etiqueta')
        claves = ('cmp.compensado', 'cmp.compensada', 'cmp.anular', 'cmp.noCliente', 'cmp.yaCobrada', 'cmp.pagarTras', 'cmp.explica',
                  'cmp.importe', 'cmp.fecha', 'cmp.compensar', 'cmp.maximo', 'cmp.confirmar', 'cmp.queda', 'cmp.trasCobro',
                  'cmp.sinGrupo', 'arreal.emitirPend', 'arreal.emitidaOk', 'cierre.fuentesCmp')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in claves)]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_cmp.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_cmp.js'], capture_output=True, text=True).returncode:
                malos += 1
        ok(malos == 0, 'el JS servido parsea')
        diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
        ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    except Exception as e:
        import traceback; traceback.print_exc()
        ok(False, f'excepcion en la prueba: {type(e).__name__}: {e}')
    finally:
        for d in DIRS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)):
                shutil.copytree(os.path.join(copia, d), d)
            else:
                os.makedirs(d, exist_ok=True)
        shutil.rmtree(copia, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
