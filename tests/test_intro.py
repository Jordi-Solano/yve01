# -*- coding: utf-8 -*-
"""Pieza 11 (Jordi, sep 2026), rehecha el mismo dia: la intro.
Sin rueda de "cargando". Primero se ENCIENDE el logo de verdad (anillos + punto),
luego "Yve" letra a letra a su derecha (cada letra desplaza el logo un poco a la
izquierda), luego ".01" y la frase. Lento. Acento Y fondo de Personalizacion
desde el primer frame.
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
        html = html.replace("var bg=(localStorage.getItem('yve_bg')||'').trim();", "var bg='';")
        html = html.replace('<div class="sp-sub">', '<div class="sp-loader"></div><div class="sp-sub">')
    sp = html[html.index('<div id="yve-splash"'):html.index('<nav class="nav"')]
    ok('sp-loader' not in sp and 'spSpin' not in html, "sin rueda de 'cargando' en la intro")
    ok('<svg class="sp-logo"' in sp and sp.count('<circle class="r r') == 4 and '<circle class="dot"' in sp and 'sp-ball' not in html,
       "la intro es el logo de verdad (4 anillos + punto), no una esfera")
    css = html[html.index('#yve-splash{'):html.index('/* ── Arreglos responsive')]
    ok(re.search(r"\.sp-logo \.r4\{animation-delay:1\.1s", css) and re.search(r"\.sp-logo \.dot\{[^}]*1\.35s both", css), "el logo se enciende anillo a anillo y el punto al final")
    delays = [float(x) for x in re.findall(r"\.sp-l:nth-child\(\d\)\{animation-delay:([\d.]+)s", css)]
    ok(len(delays) == 3 and delays == sorted(delays) and delays[0] >= 2.3 and min(b - a for a, b in zip(delays, delays[1:])) >= 0.6,
       f"las letras Y, v, e entran una a una y despacio (a los {delays} s)")
    ok("max-width:var(--w)" in css and 'class="sp-l">Y</span><span class="sp-l">v</span><span class="sp-l">e</span>' in sp, "cada letra abre su hueco (la fila centrada desplaza el logo a la izquierda)")
    ok("localStorage.getItem('yve_accent')" in sp and "sp.style.setProperty('--sp-acc',acc)" in sp, "el logo lee el acento de Personalizacion antes del primer frame")
    ok("localStorage.getItem('yve_bg')" in sp and "sp.style.setProperty('--sp-bg1'" in sp and "var(--sp-bg1)" in css, "y el FONDO de la intro tambien sale de Personalizacion")
    ok(re.search(r"MIN=(\d+)", sp) and int(re.search(r"MIN=(\d+)", sp).group(1)) >= 6000, "dura lo que dura la secuencia (MIN >= 6 s)")

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
                pg.evaluate("localStorage.setItem('yve_accent','#e11d48'); localStorage.setItem('yve_bg','#1a1020'); sessionStorage.removeItem('yve_splash_shown')")
                pg.goto('http://127.0.0.1:5097/__intro_test', wait_until='commit')
                # se mide en cuanto el script del splash ha corrido (deja la marca en sessionStorage)
                pg.wait_for_function("sessionStorage.getItem('yve_splash_shown')==='1'", timeout=15000)
                pg.wait_for_timeout(300)
                mide = """(function(){
                  var s=document.getElementById('yve-splash'); if(!s) return null;
                  var cs=getComputedStyle(s), L=[...s.querySelectorAll('.sp-l')];
                  var dot=s.querySelector('.sp-logo .dot'), logo=s.querySelector('.sp-logo');
                  return {acc:cs.getPropertyValue('--sp-acc').trim(), bg1:cs.getPropertyValue('--sp-bg1').trim(),
                          dot:dot?parseFloat(getComputedStyle(dot).opacity):-1,
                          letras:L.map(function(l){return Math.round(l.getBoundingClientRect().width)}),
                          logoX:logo?Math.round(logo.getBoundingClientRect().left):-1, rueda:!!s.querySelector('.sp-loader')};
                })()"""
                m0 = pg.evaluate(mide)
                ok(m0 and m0['acc'].lower() == '#e11d48' and m0['bg1'].lower() != '#101a2e', f"en Chromium acento y fondo son los elegidos desde el principio ({m0 and (m0['acc'], m0['bg1'])})")
                ok(m0 and m0['dot'] < 0.05 and sum(m0['letras']) == 0, f"a los 300 ms: logo apagado ({m0 and m0['dot']}) y sin letras ({m0 and m0['letras']})")
                pg.wait_for_timeout(2500)
                m1 = pg.evaluate(mide)
                ok(m1 and m1['dot'] > 0.9 and m1['letras'][0] > 5 and m1['letras'][1] == 0, f"a los 2,8 s: punto encendido y solo la Y ({m1 and m1['letras']})")
                pg.wait_for_timeout(1700)
                m2 = pg.evaluate(mide)
                ok(m2 and m2['letras'][1] > 5 and m2['logoX'] < m1['logoX'], f"a los 4,5 s: la v ya esta y el logo se ha movido a la izquierda ({m1 and m1['logoX']} → {m2 and m2['logoX']})")
                ok(m0 and not m0['rueda'], "y no hay rueda")
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
