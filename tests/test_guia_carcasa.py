# -*- coding: utf-8 -*-
"""Guia de estilo (b54): la CARCASA — barra de arriba, pestañas, menu ⚙️,
modales genericos y burbuja de Yve — y las piezas viejas con el aspecto de
la guia. Mide en Chromium TODOS los apartados a 370, 770, 800, 850 y 1280 px:
nada desborda, la rueda cabe, las pestañas son pildora en movil y subrayado
en PC, el menu ⚙️ abre dentro de la pantalla.

  python3.12 tests/test_guia_carcasa.py
  python3.12 tests/test_guia_carcasa.py --sabotaje
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
TABS = ['ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'multi_hotel', 'cierre']


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
    css = cl.get('/static/yve-guia.css').get_data(as_text=True)
    if SABOTAJE:
        css = css.replace('body .tabs .tab{border-radius:999px;border-bottom:none;margin:0;padding:8px 13px;font-size:12px}', 'body .tabs .tab{padding:8px 40px;font-size:12px}')
        css = css.replace('body .nav .btn-run{background:var(--acc);', 'body .nav .btn-run{background:var(--acc);min-width:420px;')
    ok('CARCASA' in css and 'body .tabs .tab.active{background:var(--acc);color:#fff}' in css, "la guia trae la carcasa: pestañas pildora en movil")
    ok('body .card{' in css and 'body .sc{' in css and 'body .btn-ref{' in css and 'body .badge.ok{' in css, "las piezas viejas toman el aspecto de la guia mientras se convierten")

    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide en el navegador)")
    else:
        import logging; logging.getLogger('werkzeug').setLevel(logging.ERROR)
        from werkzeug.serving import make_server
        _css = css.encode('utf-8')
        def _wsgi(environ, start_response):
            if environ.get('PATH_INFO') == '/static/yve-guia.css':
                start_response('200 OK', [('Content-Type', 'text/css; charset=utf-8')]); return [_css]
            return app(environ, start_response)
        srv = make_server('127.0.0.1', 5091, _wsgi, threaded=True)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        try:
            with sync_playwright() as p:
                _exe = '/opt/pw-browsers/chromium' if os.path.exists('/opt/pw-browsers/chromium') else None
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path=_exe)
                res = {}
                for w in (370, 770, 800, 850, 1280):
                    ctx = br.new_context(viewport={'width': w, 'height': 800}, is_mobile=(w < 500)); pg = ctx.new_page()
                    pg.goto('http://127.0.0.1:5091/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                    pg.click('#btn-login')
                    try:
                        pg.wait_for_url(lambda u: '/login' not in u, timeout=20000)
                    except Exception:
                        pg.goto('http://127.0.0.1:5091/')
                    pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(500)
                    pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_done','1')")
                    pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(800)
                    anchos = {}
                    for tab in TABS:
                        try:
                            pg.evaluate(f"typeof switchTab==='function' && switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))")
                            pg.wait_for_timeout(350)
                        except Exception:
                            pass
                        anchos[tab] = pg.evaluate("Math.max(document.body.scrollWidth, document.documentElement.scrollWidth)")
                    m = pg.evaluate("""(function(){
                      var t=document.querySelector('.tabs .tab.active'); var cs=t?getComputedStyle(t):null;
                      var d=document.querySelector('.dropdown'); var dr=d.getBoundingClientRect();
                      document.querySelector('.dropdown > .btn-ref').click();
                      var mn=document.getElementById('main-menu'); var mr=mn.getBoundingClientRect();
                      document.querySelector('.dropdown > .btn-ref').click();
                      var run=document.getElementById('btn-run'); var rr=run.getBoundingClientRect();
                      return {tabRadius:cs?cs.borderRadius:'', tabBg:cs?cs.backgroundColor:'', ruedaDer:Math.round(dr.right), menuIzq:Math.round(mr.left), menuDer:Math.round(mr.right), runRadius:getComputedStyle(run).borderRadius, runDer:Math.round(rr.right)};
                    })()""")
                    res[w] = (anchos, m); ctx.close()
                br.close()
            for w, (anchos, m) in res.items():
                peor = max(anchos.values())
                ok(peor <= w, f"{w} px: ningun apartado desborda (peor {peor}) · {anchos}")
                ok(m['ruedaDer'] <= w and m['runDer'] <= w and m['menuIzq'] >= 0 and m['menuDer'] <= w, f"{w} px: barra y menu ⚙️ dentro de pantalla (rueda {m['ruedaDer']}, menu {m['menuIzq']}–{m['menuDer']})")
            ok(res[370][1]['tabRadius'] == '999px' and res[370][1]['tabBg'] != 'rgba(0, 0, 0, 0)', f"movil: pestaña activa en pildora rellena ({res[370][1]['tabRadius']}, {res[370][1]['tabBg']})")
            ok(res[1280][1]['tabRadius'] == '0px' and res[770][1]['tabRadius'] == '0px', "PC (y 770): pestañas subrayadas, sin mezclas")
            ok(res[370][1]['runRadius'] == '999px' and res[1280][1]['runRadius'] == '8px', f"boton Procesar: pildora en movil ({res[370][1]['runRadius']}), 8 px en PC ({res[1280][1]['runRadius']})")
        finally:
            srv.shutdown()
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_gc.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_gc.js'], capture_output=True, text=True)
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
