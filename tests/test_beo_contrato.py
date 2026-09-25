# -*- coding: utf-8 -*-
"""b91 — la BEO desde el contrato, con el formato de la BEO de ejemplo.

Lo que se comprueba:
  - una BEO por DIA de evento cuando el contrato trae el programa (funciones con fecha,
    hora y sala); ordenadas por hora; numeracion estable (regenerar no cambia el numero);
  - sin programa en el contrato NO se inventan horas ni salas: una BEO "por confirmar"
    con lo contratado (F&B y salas);
  - el PDF lleva las partes de la BEO de ejemplo (cabecera, datos, funciones, Menu,
    Montaje, Audio Visuales, Miscelaneos, Instrucciones de Facturacion, firmas) y su
    paginacion por BEO (Page x of y);
  - cotejo BEO vs F&B + salas del contrato;
  - rutas (JSON y PDF, hotel), el lector pide el programa sin inventar, pantalla, i18n.

  python3.12 tests/test_beo_contrato.py
  python3.12 tests/test_beo_contrato.py --sabotaje
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

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')


def contrato(numero, programa=True):
    d = {"es_contrato_grupo": True, "evento": {"id": "EV-77", "nombre": "Congreso Cardio 2026"}, "contrato_numero": numero,
         "hotel": {"nombre": "Hotel Els Pins", "direccion": "Passeig Marítim 12, 08870 Sitges", "telefono": "93 811 22 33"},
         "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000002", "contacto": "Marta Soler", "email": "eventos@labnorte.es"},
         "agencia": {"nombre": "Viajes Meridiano S.L.", "cif": "B28004554", "contacto": "Pau Riera", "email": "grupos@meridiano.es",
                     "direccion": "C/ Balmes 200\n08006 Barcelona"},
         "alojamiento": {"fecha_entrada": "2026-10-05", "fecha_salida": "2026-10-08", "noches": 3, "habitaciones": 20,
                         "total_habitaciones": 11000, "iva_pct": 10},
         "fb": {"total": 2200, "pax": 40, "dias": 2, "por_persona_dia": 27.5}, "salas": {"total": 1210, "nombre": "Mediterrània", "montaje": "escuela", "dias": 2},
         "comisiones": {"modo": "porcentaje", "alojamiento_pct": 10},
         "facturacion": {"pagador": "agencia", "texto": "Se factura a la agencia. Depósito del 30 % a la firma."},
         "deposito": {"pct": 30, "cuando": "a la firma"}}
    if programa:
        d["beo"] = {"coordinador": "Gemma Ràfols", "anuncio": "CONGRESO CARDIO 2026", "alergias": ["1 x alergia al marisco", "4 x vegetariano"],
                    "funciones": [
                        {"fecha": "2026-10-05", "hora_inicio": "14:00", "hora_fin": "15:30", "funcion": "Almuerzo buffet", "sala": "Restaurante", "montaje": "Buffet", "pax": 40, "precio_pp": 27.5, "menu": ["Ensalada", "Brochetas de pollo", "Postres"]},
                        {"fecha": "2026-10-05", "hora_inicio": "09:00", "hora_fin": "14:00", "funcion": "Reunión", "sala": "Mediterrània", "montaje": "Escuela", "pax": 40, "alquiler": 605, "av": [{"concepto": "Proyector", "importe": 0}]},
                        {"fecha": "2026-10-06", "hora_inicio": "09:00", "hora_fin": "14:00", "funcion": "Reunión", "sala": "Mediterrània", "montaje": "Escuela", "pax": 40, "alquiler": 605},
                        {"fecha": "2026-10-06", "hora_inicio": "14:00", "hora_fin": "15:30", "funcion": "Almuerzo buffet", "sala": "Restaurante", "montaje": "Buffet", "pax": 40, "precio_pp": 27.5,
                         "menu": [f"Plato {k} de un menú muy largo para que la BEO ocupe dos páginas" for k in range(70)]},
                    ]}
    return d


def texto_pdf(buf):
    ruta = tempfile.mktemp(suffix='.pdf')
    open(ruta, 'wb').write(buf.read())
    try:
        return subprocess.run(['pdftotext', '-layout', ruta, '-'], capture_output=True, text=True).stdout
    finally:
        os.remove(ruta)


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import beo_contrato as B
    import contratos_grupo as CG
    import lector_contratos_grupo as L
    if SABOTAJE:
        B._funciones = lambda datos: []            # el programa del contrato se pierde

    ok(B.fecha_larga('2025-07-03') == 'jueves, 3 julio, 2025', f"fecha como en la BEO de ejemplo ({B.fecha_larga('2025-07-03')})")
    ok(B.eur(2500) == '€2.500,00' and B.eur(0) == '€0,00', 'importes como en el ejemplo (€2.500,00)')

    copia = tempfile.mkdtemp(prefix='beo_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in ('contratos_grupo.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx', B.NUMERACION, 'beos_generados.json'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        for dt in (contrato("CG-2026-0930"), contrato("CG-2026-0931", programa=False)):
            t = L.transformar(dt, hotel_id="")
            L.guardar(t, DD)
            CG.registrar(dt, t, hotel_id="", datos_dir=DD)
        cs = {c['contrato']: c for c in CG.leer(DD)}
        c1, c2 = cs['CG-2026-0930'], cs['CG-2026-0931']

        # 1. una BEO por dia, por orden de hora, numeracion estable
        l1 = B.beos(c1, DD)
        ok([b['fecha'] for b in l1] == ['2026-10-05', '2026-10-06'], f"una BEO por dia de evento ({[b['fecha'] for b in l1]})")
        ok(l1 and [f['funcion'] for f in l1[0]['funciones']] == ['Reunión', 'Almuerzo buffet'], 'funciones del dia por orden de hora')
        ok(l1 and l1[0]['numero'] == B.BEO_INICIAL and l1[1]['numero'] == B.BEO_INICIAL + 1, f"numeradas ({[b['numero'] for b in l1]})")
        ok([b['numero'] for b in B.beos(c1, DD)] == [b['numero'] for b in l1], 'regenerarla no cambia los numeros')
        ok(l1 and l1[0]['fecha_texto'] == 'lunes, 5 octubre, 2026' and l1[0]['cuenta'] == 'Viajes Meridiano S.L.' and l1[0]['master'] == 'CG-2026-0930'
           and l1[0]['coordinador'] == 'Gemma Ràfols', 'datos del evento: dia, cuenta (quien paga), master = contrato, coordinador')
        cot = B.cotejo(c1, l1)
        ok(cot['cuadra'] and cot['total_beo'] == 3410.0 and cot['contrato_fb_salas'] == 3410.0,
           f"cotejo: la BEO suma lo mismo que F&B + salas del contrato ({cot['total_beo']} / {cot['contrato_fb_salas']})")
        # 2. sin programa: no se inventa
        l2 = B.beos(c2, DD)
        ok(len(l2) == 1 and l2[0]['por_confirmar'] and all(not f['hora_inicio'] and not f['fecha'] for f in l2[0]['funciones'])
           and l2[0]['numero'] == B.BEO_INICIAL + 2,
           'sin programa: UNA BEO "por confirmar", sin horas ni dias inventados')
        ok(B.cotejo(c2, l2)['cuadra'], 'y lleva lo contratado (F&B y salas)')
        fac2 = l2[0]['facturacion'] if l2 else []
        ok(any('Depósito del 30 % a la firma' in x for x in fac2) and not any(x.strip().lower() == 'a la firma' for x in fac2),
           f"facturación: el depósito una vez, sin repetir 'a la firma' suelto (b97: {fac2})")
        # 3. el PDF con el formato del ejemplo
        txt = texto_pdf(B.pdf(l1))
        etiquetas = ('Orden del Servicio (BEO)', 'Postear como:', 'Fecha del evento:', 'Cuenta:', 'Contacto:', 'Dirección:', 'Master #:',
                     'Coord. Evento:', 'Hora del evento', 'Función', 'Sala', 'Montaje', 'Agr', 'Gtd', 'Alquiler', 'Cantidad', 'Package',
                     'Descripcion', 'Precio Por', 'Menú', 'Audio Visuales', 'Misceláneos', 'Instrucciones de Facturación',
                     'Firma Autorizada de la Organización', 'Hotel Els Pins Approval', 'Group Catering', 'ALERGIAS ALIMENTARIAS')
        faltan = [e for e in etiquetas if e not in txt]
        ok(not faltan, f'el PDF tiene las partes de la BEO de ejemplo (faltan {faltan})')
        ok(f'BEO #: {B.BEO_INICIAL}' in txt and 'Page 1 of 1' in txt and 'Page 2 of 2' in txt and '€605,00' in txt and '@ €27,50 por persona' in txt,
           'paginacion por BEO (la del segundo dia ocupa 2 paginas) e importes')
        txt2 = texto_pdf(B.pdf(l2))
        ok('Programa por confirmar' in txt2 and 'por confirmar' in txt2, 'la BEO sin programa lo dice')

        # 4. rutas
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        ctr = {c['contrato']: c for c in cl.get('/api/ar/contratos').get_json()['contratos']}
        ok(ctr['CG-2026-0930']['beo']['n'] == 2 and ctr['CG-2026-0931']['beo']['por_confirmar'], 'AR › Contratos: cada contrato dice su BEO')
        r = cl.get(f"/api/ar/contratos/{c1['id']}/beo.pdf")
        ok(r.status_code == 200 and r.mimetype == 'application/pdf' and r.data[:4] == b'%PDF', 'la BEO en PDF')
        j = cl.get(f"/api/ar/contratos/{c1['id']}/beos").get_json()
        ok(j.get('ok') and len(j['beos']) == 2 and j['cotejo']['cuadra'], 'y en JSON con su cotejo')
        ok(cl.get('/api/ar/contratos/NO-EXISTE/beo.pdf').status_code == 404, 'un contrato que no existe: 404')
        import censo_hoteles
        _act = censo_hoteles.activo
        censo_hoteles.activo = lambda: 'OTRO'
        try:
            ok(cl.get(f"/api/ar/contratos/{c1['id']}/beo.pdf").status_code == 404, 'la BEO de otro hotel no se ve (404)')
        finally:
            censo_hoteles.activo = _act
        # 5. el lector pide el programa sin inventar; pantalla; i18n; JS
        ok('"funciones"' in L._PROMPT and 'no te inventes' in L._PROMPT and '"alergias"' in L._PROMPT, 'el lector pide el programa dia a dia (y no inventarlo)')
        html = cl.get('/').get_data(as_text=True)
        ok('function _ctrBeo(' in html and '/beo.pdf' in html, 'AR › Contratos: boton de la BEO')
        # b94: fuera la seccion vieja "BEOs desde contratos" (un resumen con el alojamiento, no una BEO)
        ok('ar-beos-list' not in html and 'cargarBeosAR' not in html and 'BEOs desde contratos' not in html,
           'la seccion vieja "BEOs desde contratos" ya no esta')
        ok(cl.get('/api/ar_real/beos').status_code == 404, 'ni su ruta (/api/ar_real/beos: 404)')
        L.guardar_beo(L.generar_beo(contrato("CG-2026-0932"), None), DD)
        ok(not os.path.exists(os.path.join(DD, 'beos_generados.json')) and os.path.exists(os.path.join(DD, 'eventos_referencia.json')),
           'ni su fichero (beos_generados.json); el cruce de documentos del evento sigue')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8'))
                                                                     for k in ('beo.ver', 'beo.dias', 'beo.porConfirmar', 'beo.cuadra', 'beo.noCuadra'))]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_beo.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_beo.js'], capture_output=True, text=True).returncode:
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
