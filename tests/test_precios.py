# -*- coding: utf-8 -*-
"""OLA A · /precios daba 404.

`pricing_bp` estaba importado en dashboard.py pero no registrado, mientras la
landing (`/precios`), el blog y "Quienes somos" enlazan a la pagina.

  python3.12 tests/test_precios.py
  python3.12 tests/test_precios.py --sabotaje
"""
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

    import dashboard
    app = dashboard.app
    app.config['TESTING'] = True
    if SABOTAJE:
        # Como estaba: la vista responde como si la regla no existiera.
        from flask import abort
        app.view_functions['pricing.pricing_page'] = lambda: abort(404)
    c = app.test_client()          # SIN login: es una pagina publica
    for ruta in ('/precios', '/pricing'):
        r = c.get(ruta)
        ok(r.status_code == 200, f'{ruta} responde 200 sin login (da {r.status_code})')
        html = r.get_data(as_text=True)
        ok('€' in html and 'Precios' in html, f'{ruta} es la pagina de precios')
    landing = c.get('/').get_data(as_text=True)
    ok('href="/precios"' in landing, 'la landing enlaza a /precios')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else
              '*** SABOTAJE SIN EFECTO: la prueba no protege nada ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
