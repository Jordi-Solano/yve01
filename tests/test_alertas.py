# -*- coding: utf-8 -*-
"""b82 — aviso por email cuando la app devuelve errores, y /health para el vigilante.

  - una excepcion sin capturar → 500 al usuario + email con ruta, usuario, hotel y traza
  - un endpoint que responde 500 a mano (jsonify) tambien avisa
  - dentro del margen (YVE_ALERTAS_MIN) no se repite el email: se acumula y sale
    UN resumen cuando pasa el margen
  - sin YVE_ALERTAS_EMAIL: se apunta, no se envia
  - /health: 200 + JSON con status ok, copias y errores; 503 si no puede leer datos
  - /admin/api/alertas (admin) y /admin/api/alertas/probar; fc_user 403
  - aviso de arranque (con y sin YVE_ALERTAS_ARRANQUE=off)

  python3.12 tests/test_alertas.py
  python3.12 tests/test_alertas.py --sabotaje
"""
import os
import sys
from datetime import timedelta

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'
os.environ.pop('YVE_ALERTAS_EMAIL', None)


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    import alertas as AL
    import notificaciones as N
    enviados = []
    N.enviar_email = lambda dest, asunto, cuerpo, tipo='general': (enviados.append((dest, asunto, cuerpo)) or True)
    AL._SINCRONO = True
    AL._reset()
    if SABOTAJE:
        AL.registrar = lambda *a, **k: 'sin_email'      # se olvida de apuntar
    app = D.app; app.config['TESTING'] = True; app.config['PROPAGATE_EXCEPTIONS'] = False

    @app.route('/_prueba_alertas/explota')
    def _explota():
        raise RuntimeError('boom de prueba')

    @app.route('/_prueba_alertas/500')
    def _cinco():
        from flask import jsonify
        return jsonify({'ok': False, 'error': 'falla a mano'}), 500

    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200

    # 1. sin email configurado: se apunta, no se envia
    r = cl.get('/_prueba_alertas/explota')
    ok(r.status_code == 500 and AL.resumen()['total'] == 1 and not enviados, f"sin YVE_ALERTAS_EMAIL: 500 al usuario, apuntado, ningun email ({r.status_code}, total {AL.resumen()['total']})")
    # 2. con email: la excepcion avisa al momento, con traza
    os.environ['YVE_ALERTAS_EMAIL'] = 'jordi@example.com'
    AL._reset()
    r = cl.get('/_prueba_alertas/explota')
    ok(r.status_code == 500 and len(enviados) == 1 and enviados[0][0] == 'jordi@example.com' and 'error 500' in enviados[0][1] and 'boom de prueba' in enviados[0][2] and 'RuntimeError' in enviados[0][2] and 'usuario <i>admin</i>' in enviados[0][2],
       f"excepcion → 1 email al momento con ruta, usuario y traza ({[e[1] for e in enviados]})")
    if not enviados:
        ok(False, 'sin email no hay nada mas que comprobar (sabotaje)')
        return _fin(fallos)
    ok('_prueba_alertas/explota' in enviados[0][2] and 'Traceback' in enviados[0][2], 'el email lleva la ruta y el traceback')
    # 3. un 500 devuelto a mano dentro del margen: se acumula, no se repite el email
    r = cl.get('/_prueba_alertas/500')
    res = AL.resumen()
    ok(r.status_code == 500 and len(enviados) == 1 and res['total'] == 2 and res['pendientes'] == 1, f"500 a mano dentro del margen: apuntado y acumulado, sin segundo email (pendientes {res['pendientes']})")
    ok(res['ultimos'][0]['error'].startswith("{'error': 'falla a mano'") or 'falla a mano' in res['ultimos'][0]['error'], 'el cuerpo del 500 queda en el apunte')
    # 4. pasa el margen → resumen con lo acumulado (en la siguiente peticion, aunque vaya bien)
    with AL._lock:
        AL._estado['ultimo_envio'] -= timedelta(minutes=AL.margen_min() + 1)
    r = cl.get('/health')
    ok(r.status_code == 200 and len(enviados) == 2 and '1 errores mas' in enviados[1][1] and '_prueba_alertas/500' in enviados[1][2] and AL.resumen()['pendientes'] == 0,
       f"pasado el margen: UN email resumen con lo acumulado ({enviados[-1][1] if len(enviados) > 1 else None})")
    # 5. /health
    h = r.get_json()
    ok(h['status'] == 'ok' and h['disco'] is True and 'copias' in h and h['errores']['total_desde_arranque'] == 2 and h['errores']['configurado'] is True,
       f"/health: status ok, disco, copias, errores ({h.get('status')}, {h.get('errores')})")
    ok(cl.get('/api/oracle/status').get_json().get('mode') in ('simulation', 'real'), '/api/oracle/status sigue vivo aparte')
    _ddir = D._ddir
    D._ddir = lambda: '/no/existe/nada'
    try:
        r = cl.get('/health')
        ok(r.status_code == 503 and r.get_json()['status'] == 'error' and r.get_json()['disco'] is False, f"/health con la carpeta de datos ilegible: 503 ({r.status_code})")
    finally:
        D._ddir = _ddir
    # 6. panel
    d = cl.get('/admin/api/alertas').get_json()
    ok(d.get('ok') and d['configurado'] and d['email'] == 'jordi@example.com' and d['total'] == 3 and d['enviados'] == 2 and len(d['ultimos']) == 3 and d['ultimos'][0]['ruta'] == '/health' and d['ultimos'][0]['status'] == 503,
       f"/admin/api/alertas: resumen; el 503 de /health tambien se apunta ({d.get('total')}, {d.get('enviados')})")
    n0 = len(enviados)
    rp = cl.post('/admin/api/alertas/probar')
    ok(rp.status_code == 200 and rp.get_json().get('ok') and len(enviados) == n0 + 1 and 'prueba' in enviados[-1][1], 'probar: manda el email de prueba')
    html = cl.get('/admin/').get_data(as_text=True)
    ok('Avisos de error por email' in html and 'loadAlertas' in html and 'probarAlertas' in html, 'el panel tiene la tarjeta de avisos')
    fc = app.test_client(); assert fc.post('/api/login', json={'username': 'fc_user', 'password': 'hotel2024'}).status_code == 200
    ok(fc.get('/admin/api/alertas').status_code == 403 and fc.post('/admin/api/alertas/probar').status_code == 403, 'fc_user: 403')
    # 7. arranque
    n0 = len(enviados)
    ok(AL.avisar_arranque({'YVE_ALERTAS_EMAIL': 'jordi@example.com'}) and len(enviados) == n0 + 1 and 'arrancado' in enviados[-1][1], 'aviso de arranque')
    ok(AL.avisar_arranque({'YVE_ALERTAS_EMAIL': 'jordi@example.com', 'YVE_ALERTAS_ARRANQUE': 'off'}) is False and len(enviados) == n0 + 1, 'YVE_ALERTAS_ARRANQUE=off: sin aviso')
    ok(AL.avisar_arranque({}) is False, 'sin email: sin aviso de arranque')
    os.environ.pop('YVE_ALERTAS_EMAIL', None)
    return _fin(fallos)


def _fin(fallos):
    os.environ.pop('YVE_ALERTAS_EMAIL', None)
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
