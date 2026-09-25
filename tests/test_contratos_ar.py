# -*- coding: utf-8 -*-
"""b87 — AR > Contratos: la comision segun el contrato, quien paga, y el PDF por el mismo lector.

Respuestas de finanzas (24 sep 2026) que se comprueban:
  1. el campo de comisiones del contrato decide: "tarifa neta" = sin comision (el hotel
     factura el neto); un % = la agencia factura su comision aparte. Si el contrato no lo
     dice, queda PENDIENTE y se pregunta (no se supone).
  2. quien paga la factura del grupo (agencia / cliente) se lee del contrato o se pregunta;
     sin agencia paga el cliente.
  3. Yve YA NO crea la factura de comision en AP (la manda la agencia): guarda la esperada.
  4. un contrato en PDF (o una foto suelta) pasa por el mismo lector y acaba en AR.
  5. la factura del grupo va a nombre de quien paga y lleva el hotel; la ficha AR nace sin
     credito; reprocesar no pisa lo que decidio una persona ni "des-emite" la factura.
  6. la pantalla: tres subpestañas (Contratos · Peticion de credito · Aging AR), decidir con
     CSRF, i18n en 6 idiomas, JS que parsea.

  python3.12 tests/test_contratos_ar.py
  python3.12 tests/test_contratos_ar.py --sabotaje
"""
import base64
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
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'facturas-entrada']
DD = os.path.join(BASE, 'datos-referencia')

BASE_DATOS = {
    "es_contrato_grupo": True,
    "evento": {"id": "EV-77", "nombre": "Congreso Cardio"},
    "contrato_numero": "CG-2026-0917",
    "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000001", "pais": "España", "email": "eventos@norte.es"},
    "agencia": {"nombre": "Viajes Meridiano S.L.", "cif": "B28004556", "email": "grupos@meridiano.es"},
    "alojamiento": {"fecha_entrada": "2026-10-05", "fecha_salida": "2026-10-08", "noches": 3, "habitaciones": 20,
                    "total_habitaciones": 11000, "iva_pct": 10},
    "fb": {"total": 2200, "pax": 40, "dias": 2, "iva_pct": 10},
    "salas": {"total": 1210},
    "comisiones": {"modo": None, "texto": "", "alojamiento_pct": 0, "fb_pct": 0, "salas_pct": 0},
    "facturacion": {"pagador": None, "texto": ""},
    "deposito": {"pct": 0},
    "doble_imposicion": False,
}


def datos(**kw):
    d = json.loads(json.dumps(BASE_DATOS))
    for k, v in kw.items():
        if isinstance(v, dict) and isinstance(d.get(k), dict):
            d[k].update(v)
        else:
            d[k] = v
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
    if SABOTAJE:
        CG.modo_comision = lambda d: {"modo": "porcentaje", "texto": "", "origen": "contrato"}   # "supone" siempre %

    copia = tempfile.mkdtemp(prefix='ctr_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    tmpd = tempfile.mkdtemp(prefix='ctr_dd_')
    try:
        for f in ('contratos_grupo.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx', 'beos_generados.json',
                  'eventos_referencia.json', 'depositos_previstos.xlsx'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))

        # 1. lo que dice el contrato
        m = lambda **kw: CG.modo_comision(datos(**kw))["modo"]      # noqa: E731
        ok(m(comisiones={"modo": "neta"}) == "neta", "modo 'neta' leido del contrato → tarifa neta")
        ok(m(comisiones={"modo": "porcentaje", "alojamiento_pct": 10}) == "porcentaje", "modo 'porcentaje' → comision")
        ok(m(comisiones={"alojamiento_pct": 8}) == "porcentaje", "un % en el campo de comisiones ES el contrato diciendo que hay comision")
        ok(m(comisiones={"texto": "Tarifas netas, no comisionables"}) == "neta", "el texto 'tarifas netas' → neta")
        ok(m() == "", "el contrato no dice nada de comisiones → PENDIENTE (se pregunta, no se supone)")
        ok(m(agencia={"nombre": ""}) == "sin_agencia", "sin agencia no hay comision que decidir")
        p = lambda **kw: CG.pagador_de(datos(**kw))["quien"]         # noqa: E731
        ok(p(facturacion={"pagador": "agencia"}) == "agencia" and p(facturacion={"pagador": "cliente final"}) == "cliente",
           "quien paga: agencia / cliente leidos del contrato")
        ok(p() == "" and p(agencia={"nombre": ""}) == "cliente", "quien paga: si no lo dice, pendiente; sin agencia, el cliente")

        # 2. el lector: comision segun el modo, sin factura en AP, factura a nombre de quien paga
        t_neta = L.transformar(datos(comisiones={"modo": "neta", "alojamiento_pct": 10}, facturacion={"pagador": "agencia"}), hotel_id="HT1")
        t_pct = L.transformar(datos(comisiones={"modo": "porcentaje", "alojamiento_pct": 10, "fb_pct": 5}, facturacion={"pagador": "cliente"}), hotel_id="HT1")
        t_pend = L.transformar(datos(), hotel_id="HT1")
        ok(t_neta["comisiones"]["total"] == 0 and t_pct["comisiones"]["total"] == round(10000 * .10 + 2000 * .05, 2),
           f"tarifa neta: comision 0; %: base sin IVA × % ({t_pct['comisiones']['total']})")
        rv = t_neta["reserva"]
        ok(rv["hotel_id"] == "HT1" and rv["cliente"] == "Viajes Meridiano S.L." and rv["pagador"] == "agencia" and rv["modo_comision"] == "neta",
           "la factura del grupo lleva el hotel y va a nombre de quien paga (la agencia)")
        ok(t_pct["reserva"]["cliente"] == "Laboratorios Norte S.A." and t_pend["cliente"] is None and t_pend["reserva"]["pagador"] == "",
           "paga el cliente → a su nombre; sin saber quien paga no hay ficha AR")
        ok(t_pct["cliente"]["credito_limite"] == 0 and t_pct["cliente"]["estado_ficha"] == "PENDIENTE",
           "la ficha AR nace SIN credito (antes 100.000 € regalados)")

        # 3. procesar: nada en AP, el contrato registrado con lo que falta decidir
        real_ext = L.extraer_contrato_grupo
        L.extraer_contrato_grupo = lambda paths: dict(datos(comisiones={"alojamiento_pct": 10}), _needs_review=False)
        antes_ap = sorted(os.listdir('facturas-procesadas')) if os.path.isdir('facturas-procesadas') else []
        r = L.procesar_contrato_grupo([os.path.join(tmpd, 'p1.jpg')], datos_dir=tmpd)
        despues_ap = sorted(os.listdir('facturas-procesadas')) if os.path.isdir('facturas-procesadas') else []
        ok(r.get("ok") and r["distribucion"].get("ap") is None and antes_ap == despues_ap and r.get("cierre") == ["ar"],
           f"Yve ya no crea factura de comision en AP (la manda la agencia): ap={r['distribucion'].get('ap')}, cierre={r.get('cierre')}")
        regs = CG.leer(tmpd)
        ok(len(regs) == 1 and regs[0]["comision"]["modo"] == "porcentaje" and r.get("pendientes") == ["pagador"]
           and r.get("comision_esperada") == 1000.0,
           f"contrato registrado: %, esperada 1.000 € sin IVA, falta quien paga ({r.get('pendientes')}, {r.get('comision_esperada')})")
        cid = regs[0]["id"]
        CG.decidir(cid, pagador="agencia", usuario="jordi", datos_dir=tmpd)
        CG.sincronizar_factura(CG.buscar(cid, tmpd), tmpd)
        dfr = pd.read_excel(os.path.join(tmpd, 'reservas_credito.xlsx'))
        dfc = pd.read_excel(os.path.join(tmpd, 'clientes_credito.xlsx'))
        ok(str(dfr.iloc[0]["cliente"]) == "Viajes Meridiano S.L." and str(dfr.iloc[0]["pagador"]) == "agencia"
           and list(dfc["nombre_cliente"]) == ["Viajes Meridiano S.L."] and float(dfc.iloc[0]["credito_limite"]) == 0,
           "al decidir 'paga la agencia': la factura pasa a su nombre y nace su ficha AR sin credito")
        # la factura se emite y se reprocesa el contrato: no se "des-emite" ni se pierde la decision
        dfr["estado"] = dfr["estado"].astype(object); dfr["fecha_emision"] = dfr["fecha_emision"].astype(object)
        dfr.loc[0, "estado"] = "FACTURADO"; dfr.loc[0, "fecha_emision"] = "2026-10-08"
        dfr.to_excel(os.path.join(tmpd, 'reservas_credito.xlsx'), index=False)
        r2 = L.procesar_contrato_grupo([os.path.join(tmpd, 'p1.jpg')], datos_dir=tmpd)
        dfr2 = pd.read_excel(os.path.join(tmpd, 'reservas_credito.xlsx'))
        reg2 = CG.leer(tmpd)
        ok(len(reg2) == 1 and reg2[0]["pagador"]["origen"] == "usuario" and r2.get("pendientes") == []
           and str(dfr2.iloc[0]["estado"]) == "FACTURADO" and str(dfr2.iloc[0]["cliente"]) == "Viajes Meridiano S.L.",
           "reprocesar el contrato conserva lo decidido y el estado EMITIDA de la factura")
        L.extraer_contrato_grupo = real_ext

        # 4. el lector acepta PDF (bloque 'document')
        pdf = os.path.join(tmpd, 'contrato.pdf')
        open(pdf, 'wb').write(b'%PDF-1.4\n%fake\n')
        b = L._bloque(pdf)
        ok(b["type"] == "document" and b["source"]["media_type"] == "application/pdf" and base64.b64decode(b["source"]["data"]).startswith(b'%PDF'),
           "un PDF va a la IA como documento entero")
        vistos = {}

        class _Resp:
            content = [type('x', (), {'text': json.dumps(datos(comisiones={"modo": "neta"}))})()]

        class _Cli:
            def __init__(self, **kw):
                self.messages = self

            def create(self, **kw):
                vistos['tipos'] = [c.get('type') for c in kw['messages'][0]['content']]
                vistos['prompt'] = kw['messages'][0]['content'][0]['text']
                return _Resp()
        import anthropic
        real_cli, real_key = anthropic.Anthropic, L._api_key
        anthropic.Anthropic = _Cli; L._api_key = lambda: 'x'
        try:
            dx = L.extraer_contrato_grupo([pdf])
        finally:
            anthropic.Anthropic = real_cli; L._api_key = real_key
        ok(vistos.get('tipos') == ['text', 'document'] and dx.get('comisiones', {}).get('modo') == 'neta',
           f"extraer_contrato_grupo manda el PDF ({vistos.get('tipos')})")
        ok('"modo":null' in vistos.get('prompt', '') and '"facturacion"' in vistos.get('prompt', '') and 'tarifa neta' in vistos.get('prompt', ''),
           "el prompt pide el campo de comisiones (neta/%) y quien paga")

        # 5. la app
        import dashboard as D
        import censo_hoteles
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        # un contrato en PDF por el enrutador del lote
        L.extraer_contrato_grupo = lambda paths: dict(datos(comisiones={"modo": "porcentaje", "alojamiento_pct": 10}), _needs_review=False)
        try:
            msg, marca, flags = D._enrutar_tipo_doc({'tipo_documento': 'CONTRATO', 'evento': 'Congreso'}, 'contrato.pdf', pdf)
        finally:
            L.extraer_contrato_grupo = real_ext
        ok(marca == 'AR_REAL_OK' and flags.get('ar_real') and 'AR › Contratos' in msg and '\n⚠' in msg and 'quién paga' in msg,
           f"un contrato en PDF acaba en AR › Contratos, con lo que falta decidir en su linea ({marca})")
        L.extraer_contrato_grupo = lambda paths: {"es_contrato_grupo": False, "_needs_review": False}
        try:
            msg2, marca2, flags2 = D._enrutar_tipo_doc({'tipo_documento': 'CONTRATO', 'evento': 'Otro evento'}, 'otro.pdf', pdf)
        finally:
            L.extraer_contrato_grupo = real_ext
        ok(marca2 == 'CONTRATO_OK' and not flags2.get('ar_real'), f"si no es un contrato de grupo, sigue como siempre (eventos_referencia): {marca2}")
        # una foto suelta de un contrato (/api/scan_documento): mismo lector, y lo que falta en su linea
        class _RespC:
            content = [type('x', (), {'text': json.dumps({"tipo_documento": "CONTRATO", "evento": "Congreso Cardio", "cliente": "Laboratorios Norte"})})()]

        class _CliC:
            def __init__(self, **kw):
                self.messages = self

            def create(self, **kw):
                return _RespC()
        import io
        anthropic.Anthropic = _CliC
        L.extraer_contrato_grupo = lambda paths: dict(datos(contrato_numero="CG-FOTO-1", comisiones={"modo": "neta"}), _needs_review=False)
        try:
            rs = cl.post('/api/scan_documento', data={'image': (io.BytesIO(b'\xff\xd8\xff fake jpg'), 'contrato_p1.jpg')},
                         headers=H, content_type='multipart/form-data').get_json()
        finally:
            anthropic.Anthropic = real_cli; L.extraer_contrato_grupo = real_ext
        ok(rs.get('ok') and rs.get('guardado') is True and 'AR › Contratos' in (rs.get('mensaje') or '') and 'tarifa neta' in (rs.get('mensaje') or '')
           and 'quién paga' in (rs.get('aviso') or ''),
           f"foto suelta de un contrato → AR › Contratos, con aviso aparte ({(rs.get('mensaje') or '')[:60]} | {rs.get('aviso')})")
        ids_ctr = [c['id'] for c in CG.leer(DD)]
        for c in CG.leer(DD):
            if c.get('contrato') == 'CG-FOTO-1':
                CG._escribir([x for x in CG.leer(DD) if x.get('contrato') != 'CG-FOTO-1'], DD)
        ok(len(ids_ctr) == 2, f"la foto registro su contrato ({len(ids_ctr)} en el registro)")
        d = cl.get('/api/ar/contratos').get_json()
        ok(d.get('ok') and len(d['contratos']) == 1 and d['contratos'][0]['pendientes'] == ['pagador'] and d['n_pendientes'] == 1,
           f"/api/ar/contratos: el contrato con lo que falta ({d.get('n_pendientes')})")
        c0 = d['contratos'][0]
        ok(cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'pagador': 'agencia'}).status_code == 403, 'decidir sin CSRF: 403')
        ok(cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'modo': 'porcentaje', 'pct': {}}, headers=H).status_code == 400
           and cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'modo': 'raro'}, headers=H).status_code == 400
           and cl.post('/api/ar/contratos/decidir', json={'id': 'CG|nada|', 'pagador': 'agencia'}, headers=H).status_code == 404,
           'decidir: % sin porcentaje → 400, modo raro → 400, contrato que no existe → 404')
        rr = cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'pagador': 'cliente'}, headers=H).get_json()
        ok(rr.get('ok') and rr['contrato']['pendientes'] == [] and rr['contrato']['deudor'] == 'Laboratorios Norte S.A.'
           and rr['contrato']['pagador']['origen'] == 'usuario',
           'decidir "paga el cliente": sin pendientes, deudor el cliente, queda quien lo decidio')
        rr = cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'modo': 'neta'}, headers=H).get_json()
        ok(rr.get('ok') and rr['contrato']['esperada'] == 0.0 and rr['contrato']['comision']['modo'] == 'neta',
           'corregir a tarifa neta: comision esperada 0')
        dfc = pd.read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        ok('Laboratorios Norte S.A.' in list(dfc['nombre_cliente']) and float(dfc[dfc['nombre_cliente'] == 'Laboratorios Norte S.A.'].iloc[0]['credito_limite']) == 0,
           'la ficha del que paga existe y sin credito')
        # con un hotel elegido, solo lo suyo
        real_act = censo_hoteles.activo
        censo_hoteles.activo = lambda: 'OTRO-HOTEL'
        try:
            d2 = cl.get('/api/ar/contratos').get_json()
            r404 = cl.post('/api/ar/contratos/decidir', json={'id': c0['id'], 'pagador': 'agencia'}, headers=H).status_code
        finally:
            censo_hoteles.activo = real_act
        ok(d2['contratos'] == [] and r404 == 404, 'con otro hotel elegido no se ve ni se decide el contrato')

        # 6. la pantalla
        html = cl.get('/').get_data(as_text=True)
        ip = html.index('<div id="panel-ar_real"'); fp = html.index('<!-- /panel-ar_real -->'); sec = html[ip:fp]
        ok(all('id="' + i + '"' in sec for i in ('ar-subtabs', 'ar-sub-contratos', 'ar-sub-credito', 'ar-sub-aging', 'ar-contratos-list',
                                                  'ar-contratos-count', 'ar-contratos-resumen', 'ar-clientes-list', 'ar-facturas-tbody', 'ar-bonos-list')),
           'panel: tres subpestañas con los ids de siempre dentro')
        ok(sec.index('id="ar-sub-contratos"') < sec.index('id="ar-contratos-list"') < sec.index('id="ar-sub-credito"') < sec.index('id="ar-clientes-list"')
           < sec.index('id="ar-sub-aging"') < sec.index('id="ar-facturas-tbody"') < sec.index('id="ar-bonos-list"'),
           'cada cosa en su subpestaña: contratos, clientes (credito), facturas y bonos (aging)')
        ok('function arSub(' in html and 'function cargarContratosAR(' in html and 'function decidirContratoAR(' in html
           and "'/api/ar/contratos/decidir'" in html and "querySelectorAll('#fb-subtabs .fb-sub')" in html,
           'JS: arSub, cargar y decidir (por _postJson), y F&B ya no apaga las subpestañas de otros')
        claves = ('arn.contratos', 'arn.credito', 'arn.aging', 'ctr.titulo', 'ctr.sub', 'ctr.vacio', 'ctr.neta', 'ctr.comision',
                  'ctr.faltaModo', 'ctr.faltaPct', 'ctr.pagaAgencia', 'ctr.pagaCliente', 'ctr.faltaPagador', 'ctr.decidirComision',
                  'ctr.decidirPagador', 'ctr.btnPct', 'ctr.esperada', 'ctr.pendientes')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in claves)]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_ctr.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_ctr.js'], capture_output=True, text=True).returncode:
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
            else:
                os.makedirs(d, exist_ok=True)
        shutil.rmtree(copia, ignore_errors=True)
        shutil.rmtree(tmpd, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
