# -*- coding: utf-8 -*-
"""b88 — la factura de comision de la agencia: unida a su contrato, esperado vs facturado,
mes = fecha de la factura, 628/472/410, y lo devengado sin factura a la provision 4109.

Reglas de finanzas (24 sep 2026) que se comprueban:
  - Yve no crea la factura de comision: llega la de la agencia (por AP) y se une a SU
    contrato — por la referencia que cite o por eliminacion; lo dudoso lo elige una persona.
  - se compara la base facturada con la comision ESPERADA y se avisa si no coincide.
  - el mes de imputacion de la factura de comision es su FECHA DE FACTURA (aunque el resto de
    AP vaya por fecha de registro desde b84): asientos, reconciliacion y 303.
  - su asiento es 628 (D) / 472 (D) / 410 (H); la 400 ya no la cuenta, y hay un check 410.
  - lo devengado (evento terminado) sin factura a fin de mes va a la provision 628/4109.
  - pantalla: badge en AP, tarjeta en la ficha (la fecha contable no se toca), aging con el
    mes de imputacion, contratos con su factura y unir/separar (CSRF).

  python3.12 tests/test_comision_agencia.py
  python3.12 tests/test_comision_agencia.py --sabotaje
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
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')
PROC = os.path.join(BASE, 'facturas-procesadas')


def contrato(numero, evento, agencia, nif, pct=None, modo=None, salida="2026-08-10", pagador="agencia"):
    return {"es_contrato_grupo": True, "evento": {"id": "", "nombre": evento}, "contrato_numero": numero,
            "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000001"},
            "agencia": {"nombre": agencia, "cif": nif},
            "alojamiento": {"fecha_entrada": "2026-08-07", "fecha_salida": salida, "noches": 3, "habitaciones": 20,
                            "total_habitaciones": 11000, "iva_pct": 10},
            "fb": {"total": 0}, "salas": {"total": 0},
            "comisiones": {"modo": modo, "alojamiento_pct": pct or 0}, "facturacion": {"pagador": pagador}}


def fila_ap(num, prov, nif, concepto, base, fecha, **kw):
    d = {"archivo": f"{num}.pdf", "numero_factura": num, "nombre_proveedor": prov, "NIF_proveedor": nif,
         "fecha_factura": fecha, "descripcion_concepto": concepto, "base_imponible": base, "porcentaje_iva": 21,
         "cuota_iva": round(base * .21, 2), "total_factura": round(base * 1.21, 2), "tipo_proveedor": "OTRAS",
         "cuenta_contable": "629", "hotel_id": ""}
    d.update(kw)
    return d


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import contratos_grupo as CG
    import lector_contratos_grupo as L
    import cierre_mes as CM
    import provisiones as PV
    import almacen_datos as ALM
    if SABOTAJE:
        CG.marcar_comisiones = lambda df, *a, **k: df          # la factura de la agencia no se reconoce

    copia = tempfile.mkdtemp(prefix='com_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in os.listdir(PROC):
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PROC, f))
        for f in ('contratos_grupo.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx', 'ajustes_ap.json', 'extracto_banco.xlsx'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        # contratos: A con 10 % (esperada 1.000), B tarifa neta, C con 10 % de otra agencia sin factura
        for dt in (contrato("CG-2026-0801", "Congreso Cardio", "Viajes Meridiano S.L.", "B28004556", pct=10),
                   contrato("CG-2026-0802", "Incentivo Pharma", "Eventos Costa DMC", "B66123456", modo="neta"),
                   contrato("CG-2026-0803", "Foro Seguros", "Grupos Levante S.L.", "B46999888", pct=8)):
            CG.registrar(dt, L.transformar(dt, hotel_id=""), hotel_id="", datos_dir=DD)
        # AP registrada el 15/09: la de la agencia es del 28/08 y cita el contrato; otra de la agencia sin referencia;
        # una de la DMC del contrato neto; una normal
        pd.DataFrame([
            fila_ap("VM-2026-310", "VIAJES MERIDIANO SL", "B28004556", "Comisión grupo contrato CG-2026-0801", 1000.0, "28/08/2026"),
            fila_ap("VM-2026-311", "Viajes Meridiano, S.L.", "B28004556", "Comisión reserva individual 55/26", 90.0, "30/08/2026"),
            fila_ap("EC-77", "Eventos Costa DMC", "B66123456", "Comision grupo Incentivo Pharma", 300.0, "29/08/2026"),
            fila_ap("END-9", "Endesa Energia SA", "A81948077", "Suministro electrico agosto", 500.0, "31/08/2026"),
        ]).to_excel(os.path.join(PROC, 'facturas_ap_20260915.xlsx'), index=False)

        # 1. unirla a su contrato
        df = ALM.facturas_ap()
        por = {r["numero_factura"]: r for r in df.to_dict("records")}
        a, b, c, e = por["VM-2026-310"], por["VM-2026-311"], por["EC-77"], por["END-9"]
        ok(bool(a.get("es_comision_agencia")) and a.get("comision_estado") == "CUADRA" and a.get("comision_vinculo") == "referencia",
           f"la factura que cita el contrato se une a el y cuadra con lo esperado ({a.get('comision_estado')}, {a.get('comision_vinculo')})")
        ok(not b.get("es_comision_agencia"), "otra factura de la misma agencia sin referencia NO se une sola (hay una que cita el contrato)")
        ok(bool(c.get("es_comision_agencia")) and c.get("comision_estado") == "NO_DEBERIA",
           f"factura de comision en un contrato de tarifa neta: se avisa ({c.get('comision_estado')})")
        ok(not e.get("es_comision_agencia") and str(e.get("fecha_contable")) == "2026-09-15", "una factura normal no se toca (fecha contable = registro)")
        # 2. el mes: fecha de factura, y la cuenta 628
        ok(str(a.get("fecha_contable")) == "2026-08-28" and str(a.get("cuenta_contable")) == "628",
           f"comision: fecha contable = fecha de la factura (28/08, no el registro 15/09) y gasto 628 ({a.get('fecha_contable')}, {a.get('cuenta_contable')})")
        # DIFERENCIA
        d2 = CG.estado_comision(CG.leer(DD)[0], dict(a, base_imponible=1100.0))
        ok(d2["estado"] == "DIFERENCIA" and d2["diferencia"] == 100.0, f"si factura 1.100 y se esperaban 1.000: DIFERENCIA de 100 ({d2})")

        # 3. el cierre: agosto lleva la comision (628/472/410) y la provision de C; septiembre la factura normal
        res8 = CM.generar_asientos('2026-08', CM.recoger_fuentes('2026-08', None, PROC, 'reportes', DD), CM.plan_cuentas(DD), CM.config_cierre(DD))
        asi = [x for x in res8['asientos'] if x['documento'] == 'VM-2026-310']
        cuentas = sorted((x['cuenta'], x['debe'], x['haber']) for x in asi)
        ok(cuentas == [('410', 0.0, 1210.0), ('472', 210.0, 0.0), ('628', 1000.0, 0.0)],
           f"asiento de la comision en AGOSTO: 628 1.000 / 472 210 / 410 1.210 ({cuentas})")
        res9 = CM.generar_asientos('2026-09', CM.recoger_fuentes('2026-09', None, PROC, 'reportes', DD), CM.plan_cuentas(DD), CM.config_cierre(DD))
        ok(not [x for x in res9['asientos'] if x['documento'] == 'VM-2026-310'] and [x for x in res9['asientos'] if x['documento'] == 'END-9'],
           "en septiembre (mes de registro) NO esta la comision; la factura normal SI")
        rec8 = CM.reconciliar('2026-08', res8, CM.recoger_fuentes('2026-08', None, PROC, 'reportes', DD), None, CM.config_cierre(DD))
        ch410 = [x for x in rec8['checks'] if x['cuenta'] == '410']
        ch400 = [x for x in rec8['checks'] if x['cuenta'] == '400'][0]
        ok(ch410 and ch410[0]['estado'] == 'CUADRA' and ch400['estado'] == 'CUADRA',
           f"reconciliacion: la 400 ya no cuenta la comision y el check 410 cuadra ({ch410[0]['libro'] if ch410 else None})")
        prov8 = [x for x in res8['asientos'] if x['origen'] == 'PROVISION' and 'Grupos Levante' in x['concepto']]
        ok(sorted((x['cuenta'], x['debe'], x['haber']) for x in prov8) == [('4109', 0.0, 800.0), ('628', 800.0, 0.0)],
           f"lo devengado sin factura (contrato C, 8 % de 10.000) va a la provision 628/4109 ({[(x['cuenta'], x['debe'], x['haber']) for x in prov8]})")
        pg = PV.provision_comisiones_agencia('2026-08', None, PROC, 'reportes', DD)
        ok(pg['n'] == 1 and pg['total'] == 800.0 and pg['filas'][0]['contrato'] == 'CG-2026-0803',
           "provision: solo el contrato con % sin factura (A ya tiene la suya en agosto; B es tarifa neta)")
        pg7 = PV.provision_comisiones_agencia('2026-07', None, PROC, 'reportes', DD)
        ok(pg7['n'] == 0, "en julio el evento aun no habia terminado: no se provisiona")
        import fiscal as FI
        try:
            f8 = FI.calcular('2026-08', CM.recoger_fuentes('2026-08', None, PROC, 'reportes', DD), CM.config_cierre(DD))
            txt = json.dumps(f8, ensure_ascii=False, default=str)
            ok('VM-2026-310' in txt, "303/SII de agosto: la factura de comision entra por su fecha de factura")
        except Exception as ex:
            ok(False, f"fiscal: {ex}")

        # 4. la app
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        fi = cl.get('/api/ap/ficha?clave=VM-2026-310').get_json()
        ok(fi.get('ok') and fi.get('criterio_fecha') == 'comision' and (fi.get('comision') or {}).get('estado') == 'CUADRA'
           and [l['cuenta'] for l in fi['asiento']['lineas']] == ['628', '472', '410'],
           f"ficha: tarjeta de comision, criterio 'comision' y asiento 628/472/410 ({[l['cuenta'] for l in (fi.get('asiento') or {}).get('lineas', [])]})")
        r = cl.post('/api/ap/ajustar', json={'clave': 'VM-2026-310', 'fecha_contable': '2026-09-15'}, headers=H)
        ok(r.status_code == 400 and 'fecha de factura' in (r.get_json() or {}).get('error', ''), "la fecha contable de una comision de agencia no se puede mover")
        fa = cl.get('/api/facturas_ap').get_json()
        fa = fa if isinstance(fa, list) else (fa.get('facturas') or [])
        fr = next((x for x in fa if x.get('numero_factura') == 'VM-2026-310'), {})
        ok(fr.get('comision_evento') and fr.get('comision_estado') == 'CUADRA', f"lista AP: la fila dice 'comision grupo' ({fr.get('comision_evento')})")
        ag = cl.get('/api/aging_ap').get_json()
        fila_ag = next((p for p in ag.get('por_acreedor', []) if p.get('origen') == 'Comisión grupo'), {})
        ok(fila_ag and '2026-08' in (fila_ag.get('imputacion') or []), f"aging AP: la comision va aparte y dice en que mes se imputa ({fila_ag.get('imputacion')})")
        pv = cl.get('/api/provisiones?mes=2026-08').get_json()
        ok(pv.get('ok') and pv['comisiones_agencia']['total'] == 800.0, "/api/provisiones trae el bloque de comisiones de agencia")
        ctr = cl.get('/api/ar/contratos').get_json()['contratos']
        ca = next(x for x in ctr if x['contrato'] == 'CG-2026-0801')
        cc = next(x for x in ctr if x['contrato'] == 'CG-2026-0803')
        ok(ca['factura_comision_vista'].get('numero') == 'VM-2026-310' and ca['factura_comision_vista'].get('estado') == 'CUADRA',
           "AR › Contratos: el contrato ensena su factura de comision y si cuadra")
        # C no tiene factura: una persona une la VM-2026-311 (no es de su agencia, pero decide ella)
        ok(cl.post('/api/ar/contratos/vincular', json={'id': cc['id'], 'clave': 'VM-2026-311'}).status_code == 403, 'unir sin CSRF: 403')
        ok(cl.post('/api/ar/contratos/vincular', json={'id': cc['id'], 'clave': 'NO-EXISTE'}, headers=H).status_code == 404, 'unir una factura que no esta en AP: 404')
        rv = cl.post('/api/ar/contratos/vincular', json={'id': cc['id'], 'clave': 'VM-2026-311'}, headers=H).get_json()
        ok(rv.get('ok') and rv['contrato']['factura_comision_vista'].get('numero') == 'VM-2026-311' and rv['contrato']['factura_comision_vista'].get('origen') == 'manual',
           'unir a mano: queda unida (origen manual)')
        b2 = next(r for r in ALM.facturas_ap().to_dict('records') if r['numero_factura'] == 'VM-2026-311')
        ok(bool(b2.get('es_comision_agencia')) and str(b2.get('fecha_contable')) == '2026-08-30', 'unida a mano ya es comision de agencia (mes = su fecha de factura)')
        rs = cl.post('/api/ar/contratos/desvincular', json={'id': cc['id'], 'clave': 'VM-2026-311'}, headers=H).get_json()
        ok(rs.get('ok') and not rs['contrato']['factura_comision_vista'] and 'VM-2026-311' not in [x['clave'] for x in rs['contrato']['candidatas']],
           'separar: se suelta y no vuelve a unirse sola')
        # por eliminacion: una agencia con UNA factura libre y UN contrato con % abierto
        CG.registrar(contrato("CG-2026-0804", "Jornada Tech", "Viajes Sur S.L.", "B41000111", pct=10), None, hotel_id="", datos_dir=DD)
        cs = [c for c in CG.leer(DD) if c['contrato'] in ('CG-2026-0804',)]
        enl = CG.enlazar(cs, [fila_ap("VS-1", "Viajes Sur SL", "B41000111", "Comisiones septiembre", 1000.0, "05/09/2026")])
        ok(enl[cs[0]['id']]['clave'] == 'VS-1' and enl[cs[0]['id']]['origen'] == 'unica', 'por eliminacion: una factura libre y un contrato abierto de esa agencia')
        CG.registrar(contrato("CG-2026-0805", "Jornada Tech II", "Viajes Sur S.L.", "B41000111", pct=10), None, hotel_id="", datos_dir=DD)
        cs = [c for c in CG.leer(DD) if c['contrato'] in ('CG-2026-0804', 'CG-2026-0805')]
        enl = CG.enlazar(cs, [fila_ap("VS-1", "Viajes Sur SL", "B41000111", "Comisiones septiembre", 1000.0, "05/09/2026")])
        ok(all(not v['clave'] and v['candidatas'] == ['VS-1'] for v in enl.values()), 'con dos contratos abiertos NO se adivina: queda como candidata en los dos')
        cs1 = [c for c in CG.leer(DD) if c['contrato'] == 'CG-2026-0804']
        enl = CG.enlazar(cs1, [fila_ap("VS-2", "Viajes Sur SL", "B41000111", "Comisiones septiembre", 1250.0, "05/09/2026")])
        ok(not enl[cs1[0]['id']]['clave'] and enl[cs1[0]['id']]['candidatas'] == ['VS-2'], 'sin referencia y con importe distinto NO se une sola: la elige una persona')

        # 5. la pantalla
        html = cl.get('/').get_data(as_text=True)
        ok(all(x in html for x in ('function _ctrFacturaComision(', 'function unirComisionAR(', 'function separarComisionAR(',
                                   "'/api/ar/contratos/vincular'", 'function _fichaComision(', 'comisiones_agencia')),
           'JS: factura de comision en contratos (unir/separar), tarjeta en la ficha y bloque de provision')
        claves = ('ap.comGrupo', 'ap.mesComision', 'com.cuadra', 'com.noCoincide', 'com.noDeberia', 'com.titulo', 'com.factura',
                  'com.imputa', 'com.unir', 'com.separar', 'com.esperando', 'prov.comAgencia', 'prov.comAgVacio')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in claves)]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_com.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_com.js'], capture_output=True, text=True).returncode:
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
