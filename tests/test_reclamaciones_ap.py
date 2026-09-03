# -*- coding: utf-8 -*-
"""OLA A · reclamar al proveedor una rectificativa o un abono.

Candidatas = facturas con incidencia de matching sin aprobar + facturas
RECHAZADAS. Borrador con las cifras, gate humano, envio (aqui simulado),
idempotencia y registro. No toca aprobaciones_ap.xlsx ni Oracle.

  python3.12 tests/test_reclamaciones_ap.py
  python3.12 tests/test_reclamaciones_ap.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                    # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = datetime.now().strftime('%Y%m%d')
PDIR = os.path.join(BASE, 'facturas-procesadas')
ADIR = os.path.join(BASE, 'aprobaciones')
RDIR = os.path.join(BASE, 'reportes')
DDIR = os.path.join(BASE, 'datos-referencia')
APRO = os.path.join(ADIR, 'aprobaciones_ap.xlsx')
ESTADO = os.path.join(DDIR, 'reclamaciones_ap.json')
CONTACTOS = os.path.join(DDIR, 'reclamaciones_ap_contactos.json')
AUDIT = os.path.join(DDIR, 'audit_log.json')


def fila(n, prov, tot, estado, detalle=''):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'archivo': f'{n}.pdf', 'estado_matching': estado, 'detalle_matching': detalle,
            'cuenta_contable': '600', 'tipo_proveedor': 'FB', 'fecha_factura': '01/09/2026', 'hotel_id': ''}


class Guardado:
    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='rap_')
        self.items = {}

    def guarda(self, ruta):
        existia = os.path.exists(ruta)
        copia = os.path.join(self.tmp, str(len(self.items)))
        if existia:
            (shutil.copytree if os.path.isdir(ruta) else shutil.copy)(ruta, copia)
        self.items[ruta] = (existia, copia)

    def restaura(self):
        for ruta, (existia, copia) in self.items.items():
            if os.path.isdir(ruta):
                shutil.rmtree(ruta)
            elif os.path.exists(ruta):
                os.remove(ruta)
            if existia:
                (shutil.copytree if os.path.isdir(copia) else shutil.copy)(copia, ruta)
        shutil.rmtree(self.tmp, ignore_errors=True)


def main():
    g = Guardado()
    for r in (PDIR, ADIR, ESTADO, CONTACTOS, AUDIT):
        g.guarda(r)
    for f in os.listdir(RDIR):
        if f.startswith('matching_') and f.endswith('.xlsx'):
            g.guarda(os.path.join(RDIR, f))
    fallos = 0
    try:
        os.makedirs(PDIR, exist_ok=True)
        os.makedirs(ADIR, exist_ok=True)
        for f in os.listdir(PDIR):
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PDIR, f))
        for f in os.listdir(RDIR):
            if f.startswith('matching_') and f.endswith('.xlsx'):
                os.remove(os.path.join(RDIR, f))
        for f in (APRO, ESTADO, CONTACTOS):
            if os.path.exists(f):
                os.remove(f)
        pd.DataFrame([
            fila('F-OK', 'Makro Cash & Carry SL', 300.0, 'MATCH_3WAY_OK'),
            fila('F-DIF', 'Makro Cash & Carry SL', 1234.56, 'DIFERENCIA_IMPORTE', 'albaran 1.100,00 vs factura 1.234,56'),
            fila('F-REJ', 'Frutas del Camp SL', 210.0, 'MATCH_3WAY_OK'),
            fila('F-APR', 'Otis SL', 400.0, 'FACTURA_SIN_ALBARAN'),
        ]).to_excel(os.path.join(PDIR, f'facturas_contabilizadas_{HOY}.xlsx'), index=False)

        import reclamaciones_ap as R
        import notificaciones as N
        if SABOTAJE:
            R._cifras_que_faltan = lambda cuerpo, c: []      # deja salir emails sin cifras
        enviados = []
        N.enviar_email = lambda dest, asunto, cuerpo, tipo='general': (enviados.append((dest, asunto, cuerpo, tipo)) or True)

        import dashboard
        app = dashboard.app
        app.config['TESTING'] = True

        def ok(cond, msg):
            nonlocal fallos
            print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
            if not cond:
                fallos += 1

        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
        H = {'X-CSRF-Token': tok}

        # F-REJ la rechazamos y F-APR la aprobamos en "Facturas por aprobar"
        for clave, acc in (('F-REJ', 'RECHAZADA'), ('F-APR', 'APROBADA')):
            r = c.post('/aprobaciones-ap/api/accion', json={'clave': clave, 'numero_factura': clave, 'accion': acc,
                                                           'comentario': 'genero caducado' if acc == 'RECHAZADA' else 'ok',
                                                           'departamento': 'F&B'}, headers=H)
            assert r.status_code == 200, r.get_json()
        apro_antes = pd.read_excel(APRO)

        # ── 1 · candidatas ──────────────────────────────────────────
        d = c.get('/api/reclamaciones_ap/list').get_json()
        items = {i['id']: i for i in d.get('items', [])}
        ok(set(items) == {'F-DIF', 'F-REJ'}, f"reclamables: {sorted(items)} (ni la que cuadra ni la aprobada)")
        ok(items['F-DIF']['tipo'] == 'CORRECCION' and items['F-REJ']['tipo'] == 'ABONO', 'tipo: rectificativa vs abono')
        ok(items['F-DIF']['destinatario'] == 'facturas@makro.es', f"email del proveedor desde proveedores.xlsx: {items['F-DIF']['destinatario']}")
        ok(items['F-REJ']['comentario'] == 'genero caducado', 'la rechazada trae el motivo del rechazo')
        ok(d.get('n_pendientes') == 2 and d.get('total_en_disputa') == 1444.56, f"resumen {d.get('n_pendientes')} · {d.get('total_en_disputa')}")

        # ── 2 · borrador con las cifras ─────────────────────────────
        r = c.post('/api/reclamaciones_ap/generar', json={'id': 'F-DIF'}, headers=H).get_json()
        ok(r.get('ok') and 'F-DIF' in r['cuerpo'] and '1.234,56 €' in r['cuerpo'] and 'rectificativa' in r['cuerpo'].lower(),
           'borrador de rectificativa con numero e importe')
        ok('1.100,00' in r['cuerpo'], 'y con el detalle del matching')
        r2 = c.post('/api/reclamaciones_ap/generar', json={'id': 'F-REJ', 'idioma': 'en'}, headers=H).get_json()
        ok(r2.get('ok') and 'credit note' in r2['cuerpo'].lower() and 'genero caducado' in r2['cuerpo'], 'abono en ingles con el motivo del rechazo')
        r3 = c.post('/api/reclamaciones_ap/generar', json={'id': 'F-OK'}, headers=H)
        ok(r3.status_code == 404, 'no se redacta para una factura que cuadra')

        # ── 3 · sin cifras no sale ──────────────────────────────────
        r = c.post('/api/reclamaciones_ap/aprobar_enviar', json={'id': 'F-DIF', 'destinatario': 'x@makro.es',
                                                                'asunto': 'hola', 'cuerpo': 'Por favor corrijan la factura.'}, headers=H)
        ok(r.status_code == 422 and (r.get_json() or {}).get('sin_cifras') and not enviados, f'sin numero ni importe: {r.status_code}, nada enviado')

        # ── 4 · envio con gate humano ───────────────────────────────
        it = {i['id']: i for i in c.get('/api/reclamaciones_ap/list').get_json()['items']}['F-DIF']
        r = c.post('/api/reclamaciones_ap/aprobar_enviar', json={'id': 'F-DIF', 'destinatario': 'x@makro.es',
                                                                'asunto': it['asunto'], 'cuerpo': it['cuerpo']}, headers=H)
        ok(r.status_code == 200 and (r.get_json() or {}).get('ok') and len(enviados) == 1 and enviados[0][0] == 'x@makro.es',
           f'enviada 1 vez a x@makro.es ({r.status_code})')
        ok(enviados and enviados[0][3] == 'reclamacion_ap' and '1.234,56' in enviados[0][2], 'el email lleva las cifras')
        r = c.post('/api/reclamaciones_ap/aprobar_enviar', json={'id': 'F-DIF', 'destinatario': 'x@makro.es',
                                                                'asunto': it['asunto'], 'cuerpo': it['cuerpo']}, headers=H)
        ok(r.status_code == 409 and (r.get_json() or {}).get('ya_enviada') and len(enviados) == 1, 'segunda vez: 409 y no se reenvia')
        st = json.load(open(ESTADO, encoding='utf-8'))
        ok(st['F-DIF']['estado'] == 'ENVIADA' and st['F-DIF']['enviada_por'] == 'admin', 'estado ENVIADA por admin')
        ok(json.load(open(CONTACTOS, encoding='utf-8')).get('makro cash carry sl') == 'x@makro.es', 'recuerda el email del proveedor')
        it2 = {i['id']: i for i in c.get('/api/reclamaciones_ap/list').get_json()['items']}
        ok(it2['F-DIF']['estado'] == 'ENVIADA' and it2['F-REJ']['destinatario'] == '', 'la lista refleja el envio; sin email conocido queda vacio')
        aud = json.load(open(AUDIT, encoding='utf-8')) if os.path.exists(AUDIT) else []
        ult = [a for a in aud if a.get('accion') == 'RECLAMACION_AP_ENVIADA']
        ok(ult and ult[-1].get('usuario') == 'admin' and 'F-DIF' in ult[-1].get('detalle', ''), f"audit log: {ult[-1] if ult else None}")

        # ── 5 · descartar ───────────────────────────────────────────
        r = c.post('/api/reclamaciones_ap/descartar', json={'id': 'F-REJ'}, headers=H)
        d = c.get('/api/reclamaciones_ap/list').get_json()
        ok(r.status_code == 200 and d['n_pendientes'] == 0 and d['total_en_disputa'] == 1234.56, 'descartada: 0 pendientes y fuera del importe en disputa')

        # ── 6 · no toca la aprobacion ni Oracle ─────────────────────
        ok(pd.read_excel(APRO).equals(apro_antes), 'aprobaciones_ap.xlsx intacto')
        diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
        ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
        html = c.get('/').get_data(as_text=True)
        ok('id="card-recl-ap"' in html and 'function cargarReclamacionesAP' in html and 'cargarReclamacionesAP();' in html,
           'tarjeta y JS en el panel AP, y se carga con el panel')
    finally:
        g.restaura()
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
