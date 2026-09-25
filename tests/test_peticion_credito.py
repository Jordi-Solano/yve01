# -*- coding: utf-8 -*-
"""b90 — la peticion de credito sigue el proceso real, y el limite solo sale de ella.

Reglas (finanzas y Jordi, 24 sep 2026) que se comprueban:
  (a) historial de pago de AR de TODO el grupo (facturas, cobro medio, retrasos, vencido);
  (b) referencias obligatorias si ha trabajado con otro hotel del grupo;
  (c) datos fiscales completos y NIF/CIF valido (con su control), informe de Informa
      adjunto (PDF/imagen, con CSRF) o su resultado anotado;
  (d) potencial comercial, limite propuesto y fecha de revision futura;
  (e) SIEMPRE dos firmas: quien pide y Direccion. Solo el rol "direccion" firma como
      directora (ni admin), y nunca quien pidio el credito; rechazar exige nota; una
      rechazada se reabre y hay que firmarla otra vez.
  - el limite de la ficha solo sale de una peticion aprobada: /api/ar_real/cliente ya no
    lo escribe; los de contrato/bono nacen sin credito; los escritos a mano antes quedan
    como "limite sin peticion" (no se borran).
  - al emitir a credito, aviso (no bloquea) si no hay credito o se pasa del limite.
  - rol nuevo "direccion" en Administracion; pantalla; i18n; JS.

  python3.12 tests/test_peticion_credito.py
  python3.12 tests/test_peticion_credito.py --sabotaje
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, timedelta

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')
HOY = date.today()


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import peticiones_credito as PC
    if SABOTAJE:
        # la segunda firma deja de ser de Direccion: cualquiera firma, tambien quien pidio
        PC.puede_firmar_direccion = lambda p, usuario, rol: ((p or {}).get('estado') == 'PENDIENTE_DIRECCION', '')

    # 1. NIF / CIF / NIE (puro)
    ok(PC.validar_nif('12345678Z')[0] and not PC.validar_nif('12345678A')[0], 'DNI: la letra se comprueba')
    ok(PC.validar_nif('X1234567L')[0], 'NIE valido')
    ok(PC.validar_nif('B28004554')[0] and not PC.validar_nif('B28004556')[0], 'CIF: el digito de control se comprueba (B28004554 si, B28004556 no)')
    ok(PC.validar_nif('ES-B28004554')[1] == 'B28004554', 'CIF con prefijo ES y guiones se normaliza')
    ok(PC.validar_nif('FR40303265045', 'Francia')[0], 'fuera de España: identificacion fiscal sin control espanol')

    copia = tempfile.mkdtemp(prefix='cred_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in ('peticiones_credito.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        shutil.rmtree(os.path.join(DD, PC.CARPETA_INFORMA), ignore_errors=True)
        pd.DataFrame([
            {"nombre_cliente": "Viajes Meridiano S.L.", "nif": "B28004556", "credito_limite": 0.0, "dias_pago": 30, "estado_ficha": "PENDIENTE", "origen": "contrato CG-2026-0917", "hotel_id": ""},
            {"nombre_cliente": "Viajes Legado S.L.", "nif": "B12345674", "credito_limite": 9000.0, "dias_pago": 30, "estado_ficha": "COMPLETA", "hotel_id": ""},
        ]).to_excel(os.path.join(DD, 'clientes_credito.xlsx'), index=False)
        hace60 = (HOY - timedelta(days=60)).isoformat()
        pd.DataFrame([
            {"numero_reserva": "FAC-1", "cliente": "VIAJES MERIDIANO, S.L.", "total": 5000.0, "estado": "COBRADO", "fecha_emision": "2026-05-01", "fecha_cobro": "2026-06-20", "hotel_id": "H2"},
            {"numero_reserva": "FAC-2", "cliente": "Viajes Meridiano S.L.", "total": 3000.0, "estado": "COBRADO", "fecha_emision": "2026-07-01", "fecha_cobro": "2026-07-25", "hotel_id": ""},
            {"numero_reserva": "FAC-3", "cliente": "Viajes Meridiano S.L.", "total": 2000.0, "estado": "FACTURADO", "fecha_emision": hace60, "hotel_id": "H2"},
            {"numero_reserva": "GRP-X", "cliente": "Viajes Meridiano S.L.", "total": 9999.0, "estado": "PENDIENTE_FACTURA", "hotel_id": ""},
            {"numero_reserva": "FAC-9", "cliente": "Otra Empresa SA", "total": 100.0, "estado": "COBRADO", "fecha_emision": "2026-07-01", "fecha_cobro": "2026-07-02", "hotel_id": ""},
        ]).to_excel(os.path.join(DD, 'reservas_credito.xlsx'), index=False)

        # 2. (a) historial de AR de todo el grupo
        h = PC.historial_ar("Viajes Meridiano S.L.", "B28004556", "H1", DD, HOY)
        ok(h['facturas'] == 3 and h['cobradas'] == 2 and h['dias_medio_cobro'] == 37.0 and h['cobradas_con_retraso'] == 1,
           f"historial: 3 facturas del grupo (no la pendiente de emitir), cobro medio 37 dias, 1 con retraso ({h['facturas']}, {h['dias_medio_cobro']}, {h['cobradas_con_retraso']})")
        ok(h['vencidas'] == 1 and h['vencido'] == 2000.0 and [x['hotel_id'] for x in h['otros_hoteles']] == ['H2'],
           f"historial: 1 vencida (2.000) y ha trabajado con otro hotel del grupo (H2) ({h['vencido']}, {h['otros_hoteles']})")

        import auth
        ok('direccion' in auth.ROLES_VALIDOS, 'rol nuevo "direccion"')
        for u, rol in (('comercial1', 'financial_controller'), ('directora', 'direccion'), ('director2', 'direccion')):
            r = auth.crear_usuario(u, 'clave-prueba-1', u.capitalize(), f'{u}@hotel.com', rol)
            ok(r is True or r == 'El usuario ya existe', f'alta de {u} ({rol}): {r}')
        import dashboard as D
        app = D.app; app.config['TESTING'] = True

        def cliente(usuario, clave):
            c = app.test_client()
            assert c.post('/api/login', json={'username': usuario, 'password': clave}).status_code == 200, usuario
            tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
            return c, {'X-CSRF-Token': tok}

        com, HC = cliente('comercial1', 'clave-prueba-1')
        dire, HD = cliente('directora', 'clave-prueba-1')
        dir2, HD2 = cliente('director2', 'clave-prueba-1')
        adm, HA = cliente('admin', 'admin123')

        # 3. crear y rellenar
        ok(com.post('/api/credito/peticion', json={'cliente': 'Viajes Meridiano S.L.'}).status_code == 403, 'crear sin CSRF: 403')
        r = com.post('/api/credito/peticion', json={'cliente': 'Viajes Meridiano S.L.'}, headers=HC).get_json()
        p = r['peticion']; pid = p['id']
        ok(r['creada'] and p['estado'] == 'BORRADOR' and p['fiscal']['nif'] == 'B28004556', f"peticion en borrador con los datos de la ficha ({pid})")
        r2 = com.post('/api/credito/peticion', json={'cliente': 'viajes meridiano sl'}, headers=HC).get_json()
        ok(not r2['creada'] and r2['peticion']['id'] == pid, 'pedir otra vez para el mismo cliente devuelve la abierta (no duplica)')
        falt = ' '.join(p['faltan'])
        ok(all(x in falt for x in ('dirección', 'código postal', 'control del CIF', 'Informa', 'potencial', 'límite', 'revisión', 'Referencias')),
           f"faltan: direccion, CP, CIF mal, Informa, potencial, limite, revision y referencias (otro hotel) ({len(p['faltan'])})")
        r = com.post(f'/api/credito/peticion/{pid}/firmar', json={}, headers=HC)
        ok(r.status_code == 409 and 'Falta' in (r.get_json() or {}).get('error', ''), 'firmar incompleta: 409')
        datos = {"referencias": {"otro_hotel": True, "texto": "Hotel Ribera: pagan a 45 días, sin incidencias (Marta, jefa de admin.)"},
                 "fiscal": {"razon_social": "Viajes Meridiano S.L.", "nif": "B28004554", "direccion": "C/ Balmes 200", "cp": "08006",
                            "poblacion": "Barcelona", "pais": "España", "email": "admin@meridiano.es"},
                 "informa": {"resultado": "Rating 8/10, riesgo bajo, límite recomendado 25.000 €", "fecha": HOY.isoformat()},
                 "comercial": {"potencial": "120000", "limite": "20000", "revision": (HOY + timedelta(days=180)).isoformat(), "comentario": "Congresos médicos"}}
        r = com.post(f'/api/credito/peticion/{pid}/guardar', json=dict(datos, comercial=dict(datos['comercial'], revision=HOY.isoformat())), headers=HC).get_json()
        ok(any('futura' in x for x in r['peticion']['faltan']), 'la fecha de revision tiene que ser futura')
        r = com.post(f'/api/credito/peticion/{pid}/guardar', json=datos, headers=HC).get_json()
        ok(r['ok'] and not r['peticion']['faltan'] and r['peticion']['nif_ok'], f"completa: no falta nada ({r['peticion']['faltan']})")
        # informe de Informa (multipart con su CSRF)
        pdf = b'%PDF-1.4 informe de prueba'
        ok(com.post(f'/api/credito/peticion/{pid}/informa', data={'fichero': (io.BytesIO(pdf), 'informa.pdf')}, content_type='multipart/form-data').status_code == 403,
           'adjuntar el informe sin CSRF: 403 (el guardia de /api no mira los multipart; esta ruta si)')
        r = com.post(f'/api/credito/peticion/{pid}/informa', data={'fichero': (io.BytesIO(b'MZ'), 'virus.exe')}, content_type='multipart/form-data', headers=HC)
        ok(r.status_code == 400, 'un .exe no se adjunta: 400')
        r = com.post(f'/api/credito/peticion/{pid}/informa', data={'fichero': (io.BytesIO(pdf), 'Informe Informa.pdf')}, content_type='multipart/form-data', headers=HC).get_json()
        ok(r.get('ok') and r['peticion']['hay_informa_fichero'], 'informe PDF adjunto')
        g = com.get(f'/api/credito/peticion/{pid}/informa')
        ok(g.status_code == 200 and g.data == pdf, 'y se descarga tal cual')

        # 4. (e) firma 1: quien pide
        r = com.post(f'/api/credito/peticion/{pid}/firmar', json={}, headers=HC).get_json()
        p = r['peticion']
        ok(r['ok'] and p['estado'] == 'PENDIENTE_DIRECCION' and p['firmas']['solicitante']['usuario'] == 'comercial1' and p['historial_ar']['facturas'] == 3,
           'firma de quien pide: pendiente de Direccion, con el historial congelado')
        r = com.post(f'/api/credito/peticion/{pid}/guardar', json=datos, headers=HC)
        ok(r.status_code == 409, 'firmada ya no se cambia (409)')
        # firma 2: solo Direccion
        r = com.post(f'/api/credito/peticion/{pid}/direccion', json={'decision': 'aprobar'}, headers=HC)
        ok(r.status_code == 403, f'quien pidio (y sin rol Direccion) no firma como directora: {r.status_code}')
        r = adm.post(f'/api/credito/peticion/{pid}/direccion', json={'decision': 'aprobar'}, headers=HA)
        ok(r.status_code == 403 and 'Dirección' in (r.get_json() or {}).get('error', ''), f'ni el admin: solo el rol Direccion ({r.status_code})')
        lst = dire.get('/api/credito/peticiones').get_json()
        ok(lst['n_para_mi'] == 1 and lst['es_direccion'], 'la directora ve que tiene 1 para firmar')
        r = dire.post(f'/api/credito/peticion/{pid}/direccion', json={'decision': 'rechazar'}, headers=HD)
        ok(r.status_code == 400, 'rechazar sin nota: 400')
        r = dire.post(f'/api/credito/peticion/{pid}/direccion', json={'decision': 'aprobar'}, headers=HD).get_json()
        ok(r.get('ok') and r['peticion']['estado'] == 'APROBADA' and r['peticion']['limite_aprobado'] == 20000.0, 'la directora aprueba: 20.000')
        # la ficha
        fi = pd.read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        m = fi[fi['nombre_cliente'] == 'Viajes Meridiano S.L.'].iloc[0]
        ok(float(m['credito_limite']) == 20000.0 and m['credito_peticion'] == pid and m['nif'] == 'B28004554' and m['estado_ficha'] == 'COMPLETA',
           'el limite aprobado (y el CIF firmado) pasa a la ficha del cliente')
        cls = {c['nombre']: c for c in adm.get('/api/ar_real/clientes').get_json()['clientes']}
        cm, cl2 = cls['Viajes Meridiano S.L.'], cls['Viajes Legado S.L.']
        ok(cm['credito']['origen'] == 'peticion' and cm['credito']['peticion'] == pid and cm['limite_credito'] == 20000.0,
           f"ficha: el limite sale de la peticion {pid}")
        ok(cl2['credito']['origen'] == 'sin_peticion' and cl2['limite_credito'] == 9000.0, 'el limite escrito a mano no se borra: "limite sin peticion"')
        # /api/ar_real/cliente ya no escribe limites
        r = adm.post('/api/ar_real/cliente', json={'nombre': 'Viajes Meridiano S.L.', 'nif': 'B28004554', 'limite': 50000, 'dias_pago': 45}, headers=HA).get_json()
        fi = pd.read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        m = fi[fi['nombre_cliente'] == 'Viajes Meridiano S.L.'].iloc[0]
        ok(r.get('ok') and r.get('aviso') and float(m['credito_limite']) == 20000.0 and m['credito_peticion'] == pid and int(m['dias_pago']) == 45,
           'editar la ficha no toca el limite (aviso) y conserva su peticion')
        r = adm.post('/api/ar_real/cliente', json={'nombre': 'Cliente Nuevo SL', 'limite': 5000}, headers=HA).get_json()
        fi = pd.read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        ok(r.get('ok') and float(fi[fi['nombre_cliente'] == 'Cliente Nuevo SL'].iloc[0]['credito_limite']) == 0.0, 'un cliente nuevo nace sin credito')
        # aviso al emitir a credito (no bloquea)
        r = adm.post('/api/ar_real/emitir_factura', json={'cliente': 'Cliente Nuevo SL', 'fecha_entrada': '2026-10-01', 'fecha_salida': '2026-10-03', 'total': 800}, headers=HA).get_json()
        ok(r.get('ok') and 'no tiene crédito aprobado' in (r.get('aviso_credito') or ''), f"emitir sin credito: se emite y avisa ({r.get('aviso_credito')})")
        r = adm.post('/api/ar_real/emitir_factura', json={'cliente': 'Viajes Meridiano S.L.', 'fecha_entrada': '2026-10-01', 'fecha_salida': '2026-10-03', 'total': 1000}, headers=HA).get_json()
        ok(r.get('ok') and not r.get('aviso_credito'), f"dentro del limite: sin aviso ({r.get('aviso_credito')})")
        r = adm.post('/api/ar_real/emitir_factura', json={'cliente': 'Viajes Meridiano S.L.', 'fecha_entrada': '2026-10-05', 'fecha_salida': '2026-10-08', 'total': 25000}, headers=HA).get_json()
        ok(r.get('ok') and 'pasa de su límite' in (r.get('aviso_credito') or ''), f"se pasa del limite: avisa ({r.get('aviso_credito')})")

        # 5. la directora no firma lo que ella pidio; rechazar y reabrir
        r = dire.post('/api/credito/peticion', json={'cliente': 'Viajes Legado S.L.'}, headers=HD).get_json()
        pid2 = r['peticion']['id']
        d2 = dict(datos, fiscal=dict(datos['fiscal'], razon_social='Viajes Legado S.L.', nif='B12345674'), referencias={"otro_hotel": False, "texto": ""})
        dire.post(f'/api/credito/peticion/{pid2}/guardar', json=d2, headers=HD)
        r = dire.post(f'/api/credito/peticion/{pid2}/firmar', json={}, headers=HD).get_json()
        ok(r.get('ok') and r['peticion']['estado'] == 'PENDIENTE_DIRECCION', 'la directora puede PEDIR credito')
        r = dire.post(f'/api/credito/peticion/{pid2}/direccion', json={'decision': 'aprobar'}, headers=HD)
        ok(r.status_code == 403 and 'no puede firmarlo' in (r.get_json() or {}).get('error', ''), 'pero no firmar como Direccion lo que ella pidio (403)')
        r = dir2.post(f'/api/credito/peticion/{pid2}/direccion', json={'decision': 'rechazar', 'nota': 'Falta el informe completo de Informa'}, headers=HD2).get_json()
        ok(r.get('ok') and r['peticion']['estado'] == 'RECHAZADA', 'otro director la rechaza con nota')
        cls = {c['nombre']: c for c in adm.get('/api/ar_real/clientes').get_json()['clientes']}
        ok(cls['Viajes Legado S.L.']['limite_credito'] == 9000.0, 'rechazada: la ficha no cambia')
        r = dire.post(f'/api/credito/peticion/{pid2}/reabrir', json={}, headers=HD).get_json()
        ok(r.get('ok') and r['peticion']['estado'] == 'BORRADOR' and not r['peticion']['firmas']['solicitante'], 'reabierta: vuelve a borrador y hay que firmarla otra vez')
        # 6. hotel: igualdad estricta
        import censo_hoteles
        _act = censo_hoteles.activo
        censo_hoteles.activo = lambda: 'H2'
        try:
            ok(com.get(f'/api/credito/peticion/{pid}').status_code == 404, 'en otro hotel la peticion no se ve (404)')
        finally:
            censo_hoteles.activo = _act

        # 7. pantalla y administracion
        html = adm.get('/').get_data(as_text=True)
        ok(all(x in html for x in ('function pedirCreditoAR(', 'function abrirPeticionCredito(', 'function firmarPeticionCredito(',
                                   'function direccionPeticionCredito(', 'function subirInformaCredito(', "'/api/credito/peticion", 'id="ar-credito-list"',
                                   'function _credFichaLimite(')) and 'ncl-limite' not in html,
           'JS: peticion de credito (pedir, firmar, Direccion, Informa) y la ficha sin campo de limite')
        adm_html = adm.get('/admin/').get_data(as_text=True)
        ok('value="direccion"' in adm_html, 'Administracion: se puede crear un usuario con rol Direccion')
        claves = ('cred.titulo', 'cred.pedir', 'cred.firmar', 'cred.aprobar', 'cred.rechazar', 'cred.sinPeticion', 'cred.sinCredito',
                  'cred.secHist', 'cred.secRef', 'cred.secFiscal', 'cred.secCom', 'cred.secFirmas', 'cred.informa', 'rol.direccion')
        faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in claves)]
        ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_cred.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_cred.js'], capture_output=True, text=True).returncode:
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
