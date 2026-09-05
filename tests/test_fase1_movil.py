# -*- coding: utf-8 -*-
"""Fase 1 del movil, rehecha (Jordi, sep 2026; el commit original 504e428 se perdio):
  1 la fila de acciones de F&B ya no desborda (flex-wrap)
  2 el selector de hotel manda (no capado a 92 px)
  3 procesar mas pequeño (solo el rayo)   4 ajustes mas pequeño
  5 el canal Push se ve siempre, apagado con motivo si no hay soporte
  6 las burbujas de canales pasan por el iconizador (_pintarYa)
  7 fotos: `_esFotoSubida` no se fia del MIME
Con Playwright (Chromium a 370 px) si esta instalado: mide `body.scrollWidth`
en TODOS los apartados (regla 57). Sin Playwright, solo las comprobaciones
sobre el HTML servido.

  python3.12 tests/test_fase1_movil.py
  python3.12 tests/test_fase1_movil.py --sabotaje
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
        html = html.replace('class="g-actions fb-acciones"', 'style="display:flex;gap:8px;align-items:center;flex-wrap:nowrap"')
        html = html.replace('#hotel-activo-sel{max-width:none!important;flex:1 1 auto;min-width:118px', '#hotel-activo-sel{max-width:92px!important')
        html = html.replace("NOTIF_CHANNELS.map(ch => {", "NOTIF_CHANNELS.filter(ch => ch.key !== 'push' || yvePushSupported()).map(ch => {")
    i = html.find('@media(max-width:768px)')
    movil = html[i:]
    # b59: la fila de acciones es `g-actions fb-acciones` y el flex-wrap:wrap lo pone la guia (.g-head .g-actions)
    _css = cl.get('/static/yve-guia.css').get_data(as_text=True)
    ok('fb-acciones' in html and ('flex-wrap:wrap' in html.split('fb-acciones')[1][:120] or ('.g-head .g-actions{' in _css and 'flex-wrap:wrap' in _css.split('.g-head .g-actions{')[1][:120])), "1 · fila de acciones de F&B con flex-wrap:wrap")
    ok('#hotel-activo-sel{max-width:none!important;flex:1 1 auto;min-width:118px' in movil, "2 · el selector de hotel manda en movil")
    ok("#run-lbl::after{content:'⚡'" in movil and '.btn-run{font-size:12px;padding:5px 9px;min-width:0}' in movil, "3 · procesar: solo el rayo y mas pequeño")
    ok('.dropdown>.btn-ref{font-size:13px!important;padding:4px 7px!important}' in movil, "4 · ajustes mas pequeño")
    ok("NOTIF_CHANNELS.map(ch => {" in html and "ch.key !== 'push' || yvePushSupported()" not in html and "notif.pushNo" in html, "5 · Push siempre visible, apagado con motivo")
    ok("_pintarYa(cont)" in html and "_pintarYa(fields)" in html, "6 · burbujas y campos de canales pasan por _pintarYa")
    ok("if (/\\.(pdf|xlsx?|xlsm|csv|docx?|txt)$/i.test(n)) return false;" in html, "7 · _esFotoSubida: la extension manda")
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_f1.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_f1.js'], capture_output=True, text=True)
        if rc.returncode:
            ok(False, f"JS roto: {rc.stderr[:100]}"); break

    # ── medida real a 370 px con Chromium, si hay Playwright ────────────
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide el desbordamiento en el navegador)")
    else:
        from werkzeug.serving import make_server
        srv = make_server('127.0.0.1', 5099, app, threaded=True)
        th = threading.Thread(target=srv.serve_forever, daemon=True); th.start()
        try:
            with sync_playwright() as p:
                _exe = '/opt/pw-browsers/chromium' if os.path.exists('/opt/pw-browsers/chromium') else None
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path=_exe) if _exe else None
                if br is None:
                    raise RuntimeError('sin chromium')
                ctx = br.new_context(viewport={'width': 370, 'height': 740}, is_mobile=True, has_touch=True)
                pg = ctx.new_page()
                pg.goto('http://127.0.0.1:5099/login')
                pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                pg.click('#btn-login'); pg.wait_for_timeout(1500)
                pg.goto('http://127.0.0.1:5099/'); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(800)
                if SABOTAJE:
                    pg.add_style_tag(content='.fb-acciones{flex-wrap:nowrap!important;min-width:420px}')
                anchos = {}
                for tab in ['ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'cierre']:
                    try:
                        pg.evaluate(f"typeof switchTab==='function' && switchTab('{tab}')")
                        pg.wait_for_timeout(500)
                    except Exception:
                        pass
                    anchos[tab] = pg.evaluate("document.body.scrollWidth")
                peor = max(anchos.values())
                ok(peor <= 370, f"ningun apartado desborda a 370 px (peor: {peor}) · {anchos}")
                # el selector solo se pinta con hoteles en el censo: se fuerza con dos opciones
                sel = pg.evaluate("(function(){var e=document.getElementById('hotel-activo-sel');if(!e)return null;e.innerHTML='<option>Hotel Els Pins</option><option>Hotel Prueba Mar</option>';e.style.display='';return e.getBoundingClientRect().width})()")
                anchoBarra = pg.evaluate("document.body.scrollWidth")
                ok(sel is not None and sel >= 118, f"selector de hotel legible: {sel and round(sel)} px (>= 118)")
                ok(anchoBarra <= 370, f"y con el selector visible la barra sigue cabiendo ({anchoBarra})")
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
