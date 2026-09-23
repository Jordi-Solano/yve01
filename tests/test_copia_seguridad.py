# -*- coding: utf-8 -*-
"""b81 — copia de seguridad diaria fuera de Render, con restauracion PROBADA.

  - crear_zip mete las carpetas de datos con un MANIFIESTO
  - hacer_copia sube a los destinos (carpeta local y S3 — el S3 va con un cliente
    falso en memoria que se comporta como boto3) y aplica la retencion
  - restaurar: baja, comprueba el zip, guarda "antes_de_restaurar_…", vacia y
    extrae → los datos quedan IDENTICOS a los de la copia (hash a hash)
  - si el destino falla: estado en rojo + email de aviso (Brevo simulado)
  - endpoints /admin/api/copias (admin), restaurar exige escribir RESTAURAR,
    fc_user recibe 403; /api/health lleva el estado de las copias
  - la hora de la siguiente copia se calcula bien

  python3.12 tests/test_copia_seguridad.py
  python3.12 tests/test_copia_seguridad.py --sabotaje
"""
import hashlib
import io
import json
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta, timezone

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'      # que dashboard no arranque el hilo
for k in ('YVE_BACKUP_S3_BUCKET', 'YVE_BACKUP_DIR', 'YVE_DATA_DIR'):
    os.environ.pop(k, None)


class S3Falso:
    """Lo minimo de boto3.client('s3') que usa DestinoS3, en memoria."""
    def __init__(self, fallar=False):
        self.obj = {}; self.fallar = fallar

    def upload_file(self, ruta, bucket, key):
        if self.fallar:
            raise RuntimeError('AccessDenied: el bucket no deja escribir')
        self.obj[key] = (open(ruta, 'rb').read(), datetime.now(timezone.utc))

    def list_objects_v2(self, Bucket, Prefix='', ContinuationToken=None):
        cont = [{'Key': k, 'Size': len(v[0]), 'LastModified': v[1]} for k, v in sorted(self.obj.items()) if k.startswith(Prefix)]
        return {'Contents': cont, 'IsTruncated': False}

    def download_file(self, bucket, key, ruta):
        if key not in self.obj:
            raise FileNotFoundError(key)
        open(ruta, 'wb').write(self.obj[key][0])

    def delete_object(self, Bucket, Key):
        self.obj.pop(Key, None)


def huellas(raiz, carpetas):
    out = {}
    for c in carpetas:
        for dp, _, fs in os.walk(os.path.join(raiz, c)):
            for f in fs:
                p = os.path.join(dp, f)
                if not f.endswith('.tmp'):
                    out[os.path.relpath(p, raiz)] = hashlib.md5(open(p, 'rb').read()).hexdigest()
    return out


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D          # ANTES de tocar YVE_DATA_DIR (almacen_persistente lo leeria)
    import copia_seguridad as CS
    if SABOTAJE:
        _orig = CS.crear_zip

        def _roto(destino, raiz=None, carpetas=None):    # el zip sale sin ficheros
            return _orig(destino, raiz=raiz, carpetas=['no-existe'])
        CS.crear_zip = _roto

    tmp = tempfile.mkdtemp(prefix='cs_')
    raiz = os.path.join(tmp, 'disco'); dest = os.path.join(tmp, 'copias')
    carpetas = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones', 'tenants']
    for c in carpetas:
        os.makedirs(os.path.join(raiz, c))
    open(os.path.join(raiz, 'datos-referencia', 'hoteles.json'), 'w').write('[{"id":"H1","nombre":"Hotel Uno"}]')
    open(os.path.join(raiz, 'facturas-procesadas', 'facturas_ap_20260901.xlsx'), 'wb').write(os.urandom(3000))
    os.makedirs(os.path.join(raiz, 'tenants', 'cliente-a', 'datos-referencia'))
    open(os.path.join(raiz, 'tenants', 'cliente-a', 'datos-referencia', 'config.json'), 'w').write('{"a":1}')
    open(os.path.join(raiz, 'reportes', 'x.tmp'), 'w').write('temporal')   # no se copia
    os.environ['YVE_DATA_DIR'] = raiz
    os.environ['YVE_BACKUP_DIR'] = dest
    os.environ['YVE_BACKUP_RETENCION'] = '30'
    enviados = []
    import notificaciones as N
    N.enviar_email = lambda dest_, asunto, cuerpo, tipo='general': (enviados.append((dest_, asunto)) or True)
    s3 = S3Falso()
    _dest_orig = CS.destinos

    def _dest(env=None):
        ds = _dest_orig(env)
        if (env or os.environ).get('YVE_BACKUP_S3_BUCKET'):
            ds = [CS.DestinoS3('yve-copias', 'k', 's', endpoint='https://x.r2.cloudflarestorage.com', prefijo='yve01/', cliente=s3)] + [d for d in ds if d.tipo != 's3']
        return ds
    CS.destinos = _dest
    try:
        ok(CS.raiz_datos() == raiz and CS.configurado(), 'raiz de datos = el disco (YVE_DATA_DIR) y hay destino')
        antes = huellas(raiz, carpetas)
        # 1. copia a la carpeta local
        r = CS.hacer_copia(motivo='test')
        ok(r['ok'] and r['ficheros'] == 3 and r['nombre'].startswith('yve01_') and r['destinos'][0]['ok'], f"hacer_copia: ok, 3 ficheros (sin el .tmp) → {r.get('ficheros')} {r['errores']}")
        z = os.path.join(tmp, 'guardado_' + r['nombre']); shutil.copy2(os.path.join(dest, r['nombre']), z)
        man = CS.comprobar_zip(z)
        ok(man['ficheros'] == 3 and set(man['carpetas']) == set(carpetas) and 'tenants/cliente-a/datos-referencia/config.json' in zipfile.ZipFile(z).namelist(),
           f"zip con MANIFIESTO (carpetas {man['carpetas']}) y los tenants dentro")
        est = CS.estado()
        ok(est.get('ultima_ok', {}).get('nombre') == r['nombre'] and CS.resumen()['al_dia'], 'estado: ultima copia buena apuntada y "al dia"')
        # 2. se rompe todo y se restaura
        os.remove(os.path.join(raiz, 'datos-referencia', 'hoteles.json'))
        open(os.path.join(raiz, 'facturas-procesadas', 'facturas_ap_20260901.xlsx'), 'wb').write(b'corrupto')
        open(os.path.join(raiz, 'reportes', 'intruso.xlsx'), 'wb').write(b'nuevo')
        rr = CS.restaurar(r['nombre'])
        despues = huellas(raiz, carpetas)
        ok(rr['ok'] and rr['ficheros'] == 3, f"restaurar: ok, 3 ficheros ({rr.get('error')})")
        ok(despues == antes, f"tras restaurar, los datos son IDENTICOS a la copia (hash a hash): {sorted(set(despues) ^ set(antes))}")
        ok(not os.path.exists(os.path.join(raiz, 'reportes', 'intruso.xlsx')) and not os.path.exists(os.path.join(raiz, 'reportes', 'x.tmp')), 'lo que no estaba en la copia desaparece (el intruso y el .tmp)')
        previas = [c for c in CS.listar() if c['nombre'].startswith('antes_de_restaurar_')]
        ok(len(previas) == 1 and CS.comprobar_zip(os.path.join(dest, previas[0]['nombre']))['ficheros'] == 3, 'antes de restaurar se guardo una copia de lo que habia (con el intruso dentro)')
        ok(CS.estado().get('ultima_restauracion', {}).get('ok') is True, 'la restauracion queda apuntada en el estado')
        # 3. retencion: copias viejas fuera, las 3 ultimas siempre
        loc = CS.DestinoLocal(dest)
        for d_ in (2, 10, 40, 50, 90):
            n = f"yve01_{(datetime.now(timezone.utc) - timedelta(days=d_)).strftime('%Y%m%d_%H%M%S')}.zip"
            shutil.copy2(z, os.path.join(dest, n))
        borradas = CS.aplicar_retencion(loc, 30)
        quedan = [c['nombre'] for c in loc.listar() if c['nombre'].startswith('yve01_')]
        ok(len(borradas) == 3 and len(quedan) == 3, f"retencion 30 dias: fuera las de 40, 50 y 90 dias ({len(borradas)} borradas, quedan {len(quedan)})")
        for f in os.listdir(dest):
            os.remove(os.path.join(dest, f))
        loc2 = CS.DestinoLocal(dest)
        for d_ in (40, 50):
            shutil.copy2(z, os.path.join(dest, f"yve01_{(datetime.now(timezone.utc) - timedelta(days=d_)).strftime('%Y%m%d_%H%M%S')}.zip"))
        ok(CS.aplicar_retencion(loc2, 30) == [] and len(loc2.listar()) == 2, 'con solo 2 copias (viejas) no se borra ninguna: siempre quedan las 3 ultimas')
        # 4. S3 (cliente falso) + local a la vez
        os.environ['YVE_BACKUP_S3_BUCKET'] = 'yve-copias'
        r2 = CS.hacer_copia(motivo='test')
        ok(r2['ok'] and len(r2['destinos']) == 2 and all(d['ok'] for d in r2['destinos']) and ('yve01/' + r2['nombre']) in s3.obj,
           f"con S3 configurado sube a los dos destinos ({[d['tipo'] for d in r2['destinos']]})")
        lst = CS.listar()
        ok(any(c['nombre'] == r2['nombre'] and c['destino'] == 's3' for c in lst) and any(c['nombre'] == r2['nombre'] and c['destino'] == 'local' for c in lst), 'listar: la misma copia aparece en s3 y en local')
        os.remove(os.path.join(raiz, 'datos-referencia', 'hoteles.json'))
        rr2 = CS.restaurar(r2['nombre'], tipo='s3')
        ok(rr2['ok'] and 'bucket' in rr2['desde'] and huellas(raiz, carpetas) == antes, f"restaurar desde S3: identico ({rr2.get('desde')})")
        # 5. el destino falla → rojo + email
        s3.fallar = True
        r3 = CS.hacer_copia(motivo='test')
        ok(not r3['ok'] and any('AccessDenied' in e for e in r3['errores']) and CS.estado()['ultima']['ok'] is False and CS.estado()['ultima_ok']['nombre'] == r2['nombre'],
           f"si S3 falla: la copia queda en rojo, la ultima BUENA sigue siendo la anterior ({r3['errores'][:1]})")
        ok(not enviados, 'sin YVE_BACKUP_EMAIL no se manda nada')
        os.environ['YVE_BACKUP_EMAIL'] = 'jordi@example.com'
        CS.hacer_copia(motivo='test')
        ok(enviados and enviados[-1][0] == 'jordi@example.com' and 'FALLADO' in enviados[-1][1], f"con YVE_BACKUP_EMAIL avisa por email del fallo ({enviados[-1:]})")
        s3.fallar = False
        # 6. endpoints
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        d = cl.get('/admin/api/copias').get_json()
        ok(d.get('ok') and d.get('configurado') and len(d.get('copias', [])) >= 2 and d.get('ultima_ok', {}).get('nombre') == r2['nombre'], '/admin/api/copias: configurado, lista y ultima buena')
        ra = cl.post('/admin/api/copias/ahora')
        ok(ra.status_code == 200 and ra.get_json().get('ok'), f"POST /admin/api/copias/ahora hace una copia ({ra.status_code})")
        nombre_web = ra.get_json()['nombre']
        rb = cl.post('/admin/api/copias/restaurar', json={'nombre': nombre_web})
        ok(rb.status_code == 400 and 'RESTAURAR' in rb.get_json().get('error', ''), 'restaurar sin escribir RESTAURAR: 400, no toca nada')
        os.remove(os.path.join(raiz, 'datos-referencia', 'hoteles.json'))
        rc = cl.post('/admin/api/copias/restaurar', json={'nombre': nombre_web, 'confirmar': 'RESTAURAR'})
        ok(rc.status_code == 200 and rc.get_json().get('ok') and os.path.exists(os.path.join(raiz, 'datos-referencia', 'hoteles.json')), f"restaurar con RESTAURAR: vuelve hoteles.json ({rc.status_code})")
        rd = cl.get('/admin/api/copias/descargar?nombre=' + nombre_web)
        ok(rd.status_code == 200 and rd.data[:2] == b'PK' and CS.comprobar_zip(io.BytesIO(rd.data))['ficheros'] == 3, f"descargar: llega el zip ({rd.status_code})")
        re_ = cl.post('/admin/api/copias/restaurar', json={'nombre': '../x.zip', 'confirmar': 'RESTAURAR'})
        ok(re_.status_code == 400, 'nombre con ../ → 400')
        fc = app.test_client(); assert fc.post('/api/login', json={'username': 'fc_user', 'password': 'hotel2024'}).status_code == 200
        ok(fc.get('/admin/api/copias').status_code == 403 and fc.post('/admin/api/copias/ahora').status_code == 403, 'fc_user (no admin): 403')
        h = cl.get('/api/health/detalle').get_json()
        ok(h['components'].get('copias', {}).get('configurado') is True and h['components']['copias']['ok'] is True, f"/api/health/detalle lleva las copias ({h['components'].get('copias')})")
        html = cl.get('/admin/').get_data(as_text=True)
        ok('Copias de seguridad' in html and 'loadCopias' in html and 'RESTAURAR' in html, 'el panel de administracion tiene la tarjeta de copias')
        # 7. la hora de la siguiente copia
        ahora = datetime(2026, 9, 12, 10, 0, tzinfo=timezone.utc)
        ok(CS._proxima('03:30', ahora) == datetime(2026, 9, 13, 3, 30, tzinfo=timezone.utc) and CS._proxima('11:00', ahora) == datetime(2026, 9, 12, 11, 0, tzinfo=timezone.utc),
           'siguiente copia: mañana a las 03:30 si ya paso; hoy a las 11:00 si no')
        ok(CS.arrancar_planificador({'YVE_BACKUP_HORA': 'off', 'YVE_BACKUP_DIR': dest}) is None, "YVE_BACKUP_HORA=off: sin hilo")
        # sin destino: no hace nada y lo dice
        os.environ.pop('YVE_BACKUP_S3_BUCKET'); os.environ.pop('YVE_BACKUP_DIR')
        r0 = CS.hacer_copia(motivo='test')
        ok(not r0['ok'] and 'sin destino' in r0['errores'][0] and not CS.configurado(), 'sin variables: no configurado, no copia, lo dice')
    finally:
        CS.destinos = _dest_orig
        for k in ('YVE_DATA_DIR', 'YVE_BACKUP_DIR', 'YVE_BACKUP_S3_BUCKET', 'YVE_BACKUP_EMAIL', 'YVE_BACKUP_RETENCION'):
            os.environ.pop(k, None)
        shutil.rmtree(tmp, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
