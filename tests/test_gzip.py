# -*- coding: utf-8 -*-
"""gzip (b66): Render no comprime, asi que lo hace la app.
  - el dashboard, el CSS/JS/JSON propios viajan en gzip (y descomprimidos son
    IDENTICOS a lo que se servia)
  - lo que no es texto (png, xlsx) y lo que se emite en streaming (SSE) no se toca
  - sin Accept-Encoding: gzip, nada cambia

  python3.12 tests/test_gzip.py
  python3.12 tests/test_gzip.py --sabotaje
"""
import gzip
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    if SABOTAJE:
        D._GZIP_TIPOS = ("application/nunca",)
    app = D.app; app.config['TESTING'] = True
    c = app.test_client()
    assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    GZ = {'Accept-Encoding': 'gzip, deflate, br'}
    for u, minimo in [('/', 3.0), ('/static/yve-guia.css', 3.0), ('/static/yve-icons.js', 2.5), ('/static/i18n/ca.json', 2.5), ('/sw.js', 2.0)]:
        r = c.get(u, headers=GZ); raw = c.get(u)
        es_gz = r.headers.get('Content-Encoding') == 'gzip'
        identico = es_gz and gzip.decompress(r.data) == raw.data
        ratio = (len(raw.data) / max(1, len(r.data)))
        ok(es_gz and identico and ratio >= minimo and r.headers.get('Content-Length') == str(len(r.data)),
           f"{u}: gzip, identico al descomprimir, {len(raw.data)//1024} KB -> {len(r.data)//1024} KB (x{ratio:.1f}), Content-Length bien")
    r = c.get('/', headers=GZ)
    ok('Accept-Encoding' in (r.headers.get('Vary') or ''), 'Vary: Accept-Encoding (para que ningun proxy mezcle las dos versiones)')
    raw = c.get('/')
    ok(raw.headers.get('Content-Encoding') is None and len(raw.data) > 500000, 'sin Accept-Encoding gzip se sirve tal cual')
    for u, nombre in [('/static/icons/yve-logo-192.png', 'png'), ('/api/exportar/ap', 'xlsx')]:
        r = c.get(u, headers=GZ)
        ok(r.status_code == 200 and r.headers.get('Content-Encoding') is None, f"{nombre}: no se toca")
    r = c.get('/api/procesar_batch_stream?archivos=%5B%5D', headers=GZ)
    ok(r.mimetype == 'text/event-stream' and r.headers.get('Content-Encoding') is None, 'SSE (streaming) no se comprime')
    r = c.get('/api/csrf_token', headers=GZ)
    ok(r.headers.get('Content-Encoding') is None and r.status_code == 200, 'respuestas pequeñas (<1 KB) tal cual')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
