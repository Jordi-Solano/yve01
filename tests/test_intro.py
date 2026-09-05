# -*- coding: utf-8 -*-
"""Pieza 11 (Jordi, sep 2026): la intro.
Fuera la rueda de "cargando"; primero aparece la bola con su animacion y
despues el nombre; y si el usuario cambio el color en Personalizacion, la
bola sale de ese color desde el primer frame (igual que el nombre).
Con Playwright si esta instalado: se mide en Chromium.

  python3.12 tests/test_intro.py
  python3.12 tests/test_intro.py --sabotaje
"""
import os
import re
import subprocess
import sys
import threading

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
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace("var acc=(localStorage.getItem('yve_accent')||'').trim();", "var acc='';")
        html = html.replace('<div class="sp-ball"></div>', '<div class="sp-ball"></div><div class="sp-loader"></div>')
    sp = html[html.index('<div id="yve-splash"'):html.index('<nav class="nav"')]
    ok('sp-loader' not in sp and 'spSpin' not in html, "sin rueda de 'cargando' en la intro")
    ok('<div class="sp-ball"></div>' in sp and 'sp-logo' not in html, "la intro es una bola, no el logo cuadrado")
    css = html[html.index('#yve-splash{'):html.index('/* ── Arreglos responsive')]
    m_ball = re.search(r"\.sp-ball\{[^}]*animation:spBall ([\d.]+)s", css)
    m_name = re.search(r"\.sp-brand\{[^}]*animation:spFade [\d.]+s [^ ]+ ([\d.]+)s both", css)
    ok(m_ball and m_name and float(m_name.group(1)) >= float(m_ball.group(1)) * 0.75,
       f"primero la bola ({m_ball and m_ball.group(1)} s) y luego el nombre (empieza a los {m_name and m_name.group(1)} s)")
    ok("localStorage.getItem('yve_accent')" in sp and "sp.style.setProperty('--sp-acc',acc)" in sp, "la bola lee el color de Personalizacion antes del primer frame")
    ok("var(--sp-acc2)" in css and "var(--sp-acc)" in css, "bola y nombre usan el mismo color de acento")

    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide en el navegador)")
    else:
        import logging; logging.getLogger('werkzeug').setLevel(logging.ERROR)
        from werkzeug.serving import make_server
        # la pagina que se mide es la que se acaba de leer (con o sin sabotaje):
        # un envoltorio WSGI la sirve en /__intro_test sin tocar la app
        _cuerpo = html.encode('utf-8')
        def _wsgi(environ, start_response):
            if environ.get('PATH_INFO') == '/__intro_test':
                start_response('200 OK', [('Content-Type', 'text/html; charset=utf-8')])
                return [_cuerpo]
            return app(environ, start_response)
        srv = make_server('127.0.0.1', 5097, _wsgi, threaded=True)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        try:
            with sync_playwright() as p:
                _exe = '/opt/pw-browsers/chromium' if os.path.exists('/opt/pw-browsers/chromium') else None
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path=_exe)
                ctx = br.new_context(viewport={'width': 400, 'height': 740}); pg = ctx.new_page()
                pg.goto('http://127.0.0.1:5097/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                pg.click('#btn-login'); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(1500)
                pg.evaluate("localStorage.setItem('yve_accent','#e11d48'); sessionStorage.removeItem('yve_splash_shown')")
                pg.goto('http://127.0.0.1:5097/__intro_test', wait_until='commit')
                # en cuanto el splash existe (el HTML son 600 KB: 'commit' llega antes de parsearlo)
                # ...y se mide en cuanto el script del splash ha corrido (deja la marca en sessionStorage)
                pg.wait_for_function("sessionStorage.getItem('yve_splash_shown')==='1'", timeout=15000)
                pg.wait_for_timeout(200)
                m = pg.evaluate("""(function(){
                  var s=document.getElementById('yve-splash'); if(!s) return null;
                  var b=s.querySelector('.sp-ball'), n=s.querySelector('.sp-brand');
                  return {acc:getComputedStyle(s).getPropertyValue('--sp-acc').trim(), bola:b?parseFloat(getComputedStyle(b).opacity):-1,
                          nombre:n?parseFloat(getComputedStyle(n).opacity):-1, rueda:!!s.querySelector('.sp-loader')};
                })()""")
                ok(m and m['acc'].lower() == '#e11d48', f"en Chromium la bola sale del color elegido desde el principio ({m and m['acc']})")
                ok(m and m['bola'] > 0.05 and m['nombre'] < 0.05, f"a los 200 ms ya se ve la bola ({m and m['bola']}) y el nombre aun no ({m and m['nombre']})")
                pg.wait_for_timeout(1500)
                m2 = pg.evaluate("(function(){var n=document.querySelector('#yve-splash .sp-brand');return n?parseFloat(getComputedStyle(n).opacity):-1})()")
                ok(m2 is not None and m2 > 0.9, f"a los 1,7 s el nombre ya esta ({m2})")
                ok(m and not m['rueda'], "y no hay rueda")
                br.close()
        finally:
            srv.shutdown()
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_in.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_in.js'], capture_output=True, text=True)
        if rc.returncode:
            ok(False, f"JS roto: {rc.stderr[:100]}"); break
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
