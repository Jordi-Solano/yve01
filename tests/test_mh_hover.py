# -*- coding: utf-8 -*-
"""b104 (Jordi, 25 sep 2026): "Multi-Hotel: las tarjetas de hotel se pulsan, asi que al
pasar el raton tienen que subir un poco (elevacion y sombra), como señal de que llevan a
algun sitio. Solo esas: los tiles no clicables siguen sin hover."

- CSS: `.mh-card.is-hotel:hover` sube (translateY negativo) y gana `--g-shadow-up`,
  dentro de @media(hover:hover) (en el movil no se queda levantada tras tocarla).
- Solo la tarjeta de hotel lleva `is-hotel` y `onclick`; "sin asignar" y "desconocido" no.
- En Chromium: la de hotel sube y cambia la sombra; la de "sin asignar" y los tiles no.

  python3.12 tests/test_mh_hover.py
  python3.12 tests/test_mh_hover.py --sabotaje
"""
import json
import os
import re
import subprocess
import sys
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv
F = {'hotel_id': 'h1', 'nombre': 'Hotel Uno', 'drr': {'estado': 'sin_drr'}, 'censo': {},
     'ap': {'importe': 1200, 'facturas': 3, 'discrepancias': 0},
     'ar_ota': {'importe_reclamable': 40, 'facturas': 2, 'importe_bruto': 900},
     'ar_real': {'pendiente': 500, 'facturas': 1, 'vencido': 0},
     'fb': {'ventas': 800, 'food_cost_pct': 28}}
# el sabotaje: todas las tarjetas suben y la de hotel no
SAB_CSS = '.mh-card:hover{transform:translateY(-3px)!important}.mh-card.is-hotel:hover{transform:none!important;box-shadow:none!important}'


def _func_js(html, nombre):
    i = html.index('function ' + nombre + '(')
    return html[i:html.index('\n}\n', i) + 3]


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
        css = css.replace('.mh-card.is-hotel:hover{', '.mh-card:hover{')

    # 1. la regla
    m = re.search(r'@media\(hover:hover\)\{\.mh-card\.is-hotel:hover\{([^}]*)\}\}', css)
    regla = m.group(1) if m else ''
    ok(re.search(r'transform:translateY\(-[1-6]px\)', regla) and 'box-shadow:var(--g-shadow-up)' in regla,
       f"la tarjeta de hotel sube y gana sombra al pasar el raton, solo con raton ({regla[:90]})")
    ok(css.count('--g-shadow-up:') >= 3 and 'body.light-mode{' in css and '--g-shadow-up:' in css.split('body.light-mode{')[1].split('}')[0],
       "--g-shadow-up en oscuro, movil y claro")
    hovers = re.findall(r'([^{}]*mh-card[^{}]*:hover)\{', css)
    ok(hovers and all('.is-hotel:hover' in h for h in hovers), f"ninguna otra tarjeta de Multi-Hotel tiene hover ({hovers})")
    kh = re.search(r'\.g-kpi:hover\{([^}]*)\}', css)
    ok(kh and 'transform:none' in kh.group(1), "los tiles siguen sin moverse al pasar el raton")
    ok('.mh-card.is-hotel{cursor:pointer;transition:transform' in css and 'prefers-reduced-motion:reduce' in css, "transicion suave y sin movimiento si el sistema lo pide")

    # 2. solo la de hotel se pulsa (node)
    prog = ("function t(k,d){return d;}\nfunction gBadge(c,t){return '<span class=\"g-badge '+c+'\">'+t+'</span>';}\n"
            "function _fmtEurES(v){return String(v);}\n" +
            '\n'.join(_func_js(html, n) for n in ('_mhK', '_mhBloque', '_mhEur', '_mhFilaHotelera', '_mhTarjeta')) +
            f"\nvar f={json.dumps(F)};\nconsole.log(JSON.stringify([_mhTarjeta(f,'hotel'),_mhTarjeta(f,'sin_asignar'),_mhTarjeta(f,'desconocido')]));")
    open('/tmp/_mhh.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_mhh.js'], capture_output=True, text=True)
    outs = json.loads(rc.stdout) if rc.returncode == 0 else ['', '', '']
    ok(rc.returncode == 0, f"_mhTarjeta corre en node ({rc.stderr[:120]})")
    ok('mh-card is-hotel' in outs[0] and 'onclick="seleccionarHotelActivo(' in outs[0], "la tarjeta de hotel lleva is-hotel y se pulsa")
    ok(all('is-hotel' not in o and 'onclick' not in o for o in outs[1:]), "'sin asignar' y 'desconocido' no se pulsan ni llevan is-hotel")

    # 3. en Chromium
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide en el navegador)")
    else:
        import logging; logging.getLogger('werkzeug').setLevel(logging.ERROR)
        from werkzeug.serving import make_server
        srv = make_server('127.0.0.1', 5095, app, threaded=True)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        try:
            with sync_playwright() as p:
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path='/opt/pw-browsers/chromium')
                ctx = br.new_context(viewport={'width': 1280, 'height': 900}); pg = ctx.new_page()
                pg.goto('http://127.0.0.1:5095/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                pg.click('#btn-login'); pg.wait_for_url(lambda u: '/login' not in u, timeout=20000)
                pg.wait_for_load_state('networkidle')
                # sin intro por medio: aqui se mide el hover, no la intro
                pg.evaluate("(function(){var s=document.getElementById('yve-splash'); if(s) s.remove();})()")
                pg.evaluate("typeof switchTab==='function' && switchTab('multi_hotel')"); pg.wait_for_timeout(600)
                # las tarjetas se pintan con la funcion de verdad (sin depender de los datos)
                pg.evaluate("f => { var c=document.getElementById('mh-hotel-cards'); c.innerHTML=_mhTarjeta(f,'hotel')+_mhTarjeta(f,'sin_asignar');"
                            " var v=document.getElementById('mh-view-cards'); if(v) v.style.display='block'; c.scrollIntoView(); }", F)
                if SABOTAJE:
                    pg.add_style_tag(content=SAB_CSS)
                pg.mouse.move(5, 5); pg.wait_for_timeout(400)
                mide = """sel => { var e=document.querySelector(sel); if(!e) return null; var cs=getComputedStyle(e);
                                   return {tr:cs.transform, sh:cs.boxShadow, top:Math.round(e.getBoundingClientRect().top*10)/10}; }"""
                H, S = '#mh-hotel-cards .mh-card.is-hotel', '#mh-hotel-cards .mh-card.is-sin-asignar'
                K = S + ' .g-kpi'
                h0, s0, k0 = pg.evaluate(mide, H), pg.evaluate(mide, S), pg.evaluate(mide, K)
                ok(h0 and s0 and k0 and h0['tr'] == 'none', f"quietas sin raton ({h0 and h0['tr']})")
                pg.hover(H + ' .mh-card-title'); pg.wait_for_timeout(450)
                h1 = pg.evaluate(mide, H)
                ty = float(h1['tr'].split(',')[-1].strip(' )')) if h1 and h1['tr'].startswith('matrix') else 0
                ok(-6 <= ty <= -1 and h1['top'] < h0['top'], f"raton encima: la de hotel sube ({h0['top']} → {h1 and h1['top']}, {h1 and h1['tr']})")
                ok(h1 and h1['sh'] != h0['sh'] and h1['sh'] != 'none', "y gana sombra")
                pg.hover(S + ' .mh-card-title'); pg.wait_for_timeout(450)
                s1, h2 = pg.evaluate(mide, S), pg.evaluate(mide, H)
                ok(s1 and s1['tr'] == 'none' and s1['sh'] == s0['sh'] and s1['top'] == s0['top'], f"'sin asignar' no se mueve ({s1 and s1['tr']})")
                ok(h2 and h2['tr'] == 'none', "al salir, la de hotel vuelve a su sitio")
                pg.hover(K); pg.wait_for_timeout(450)
                k1 = pg.evaluate(mide, K)
                ok(k1 and k1['tr'] == 'none' and k1['sh'] == k0['sh'], f"un tile no se mueve ni cambia de sombra ({k1 and k1['tr']})")
                br.close()
        finally:
            srv.shutdown()

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
