# -*- coding: utf-8 -*-
"""Tiles iguales y quietos (b74). Con datos (demo), en cada apartado:
  - todos los tiles de estadisticas tienen el mismo fondo y el mismo borde
    (ninguno "mas claro" que el resto), con y sin "acentuar todo"
  - pasar el raton por un tile que no hace nada no cambia nada (ni cursor)
  - los tiles-boton (cuadre de banco) si reaccionan al raton

  python3.12 tests/test_tiles.py
  python3.12 tests/test_tiles.py --sabotaje
"""
import logging
import os
import shutil
import sys
import tempfile
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
logging.getLogger('werkzeug').setLevel(logging.ERROR)
SABOTAJE = '--sabotaje' in sys.argv
PORT = 5107
DIRS_DATOS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']
TABS = ['ar', 'ap', 'banco', 'fb', 'ar_real', 'cierre', 'multi_hotel']

JS_ESTILO = "e => { const c = getComputedStyle(e); return [c.backgroundColor, c.borderColor, c.cursor, c.transform, c.boxShadow]; }"


def main():
    import dashboard as D
    from werkzeug.serving import make_server
    from playwright.sync_api import sync_playwright

    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    copia = tempfile.mkdtemp(prefix='yve_tiles_')
    for d in DIRS_DATOS:
        if os.path.isdir(d): shutil.copytree(d, os.path.join(copia, d))
    srv = make_server('127.0.0.1', PORT, D.app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        with sync_playwright() as p:
            br = p.chromium.launch(executable_path='/opt/pw-browsers/chromium')
            ctx = br.new_context(viewport={'width': 1280, 'height': 900}); pg = ctx.new_page()
            pg.goto(f'http://127.0.0.1:{PORT}/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123'); pg.click('#btn-login')
            pg.wait_for_url(lambda u: '/login' not in u, timeout=20000); pg.wait_for_load_state('networkidle')
            pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_skipped','1'); localStorage.setItem('yve_bancomodo_visto','1')")
            pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(800)
            if SABOTAJE:
                pg.add_style_tag(content="#panel-ap .g-kpi:nth-child(2){background:#243044} #panel-ar .g-kpi:hover{border-color:#fff;cursor:pointer}")
            gen = pg.evaluate("_postJson('/api/demo/generar', {cadenas:[{nombre:'Cadena Prueba', hoteles:['Hotel Uno','Hotel Dos']}]}).then(r=>r.json())")
            ok(gen.get('ok'), 'demo generado')
            pg.evaluate("_invalidarPaneles && _invalidarPaneles(); loadAll();")
            # el modal de modo de banco tapa el raton: se cierra si sale
            for acento in (False, True):
                pg.evaluate(f"_customColors.hlAll={'true' if acento else 'false'}; _applyCustomColors();")
                for tab in TABS:
                    if tab == 'multi_hotel': pg.evaluate("_mh_loaded=false")
                    pg.evaluate(f"switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))"); pg.wait_for_timeout(2200)
                    pg.evaluate("var m=document.getElementById('modal-banco-config'); if (m) m.style.display='none';")
                    tiles = [t for t in pg.query_selector_all(f'#panel-{tab} .g-kpi') if t.is_visible()]
                    quietos = [t for t in tiles if not t.evaluate("e => e.classList.contains('g-kpi-btn') || !!e.closest('.mh-card.is-hotel, [onclick]')")]
                    estilos = [t.evaluate(JS_ESTILO) for t in quietos]
                    fondos = set(e[0] for e in estilos); bordes = set(e[1] for e in estilos)
                    ok(len(quietos) >= 3 and len(fondos) == 1 and len(bordes) == 1, f"{'acento' if acento else 'normal'} · {tab}: {len(quietos)} tiles con el mismo fondo y borde ({fondos} / {bordes})")
                    cambios = []
                    for t in quietos[:6]:
                        antes = t.evaluate(JS_ESTILO); t.hover(); pg.wait_for_timeout(200); despues = t.evaluate(JS_ESTILO)
                        if antes != despues or despues[2] == 'pointer': cambios.append((antes, despues))
                    pg.mouse.move(0, 0)
                    ok(not cambios, f"{'acento' if acento else 'normal'} · {tab}: los tiles sin accion no reaccionan al raton ({cambios[:1]})")
            # los tiles-boton del cuadre de banco si reaccionan
            pg.evaluate("_customColors.hlAll=false; _applyCustomColors();")
            pg.evaluate("switchTab('cierre', document.querySelector('.tab[onclick*=\"\\'cierre\\'\"]'))"); pg.wait_for_timeout(2500)
            btns = [t for t in pg.query_selector_all('#cbanco-pestanas .g-kpi-btn') if t.is_visible()]
            if btns:
                antes = btns[0].evaluate(JS_ESTILO); btns[0].hover(); pg.wait_for_timeout(300); despues = btns[0].evaluate(JS_ESTILO)
                ok(despues[2] == 'pointer' and antes[1] != despues[1], f"los tiles-boton del cuadre de banco si reaccionan (cursor {despues[2]}, borde {antes[1]} -> {despues[1]})")
            else:
                ok(False, 'no hay tiles-boton en el cuadre de banco para probar')
            pg.evaluate("_customColors.hlAll=false; _applyCustomColors(); _postJson('/api/demo/toggle', {}).then(r=>r.json())"); pg.wait_for_timeout(400)
            ctx.close(); br.close()
    finally:
        srv.shutdown()
        for d in DIRS_DATOS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)): shutil.copytree(os.path.join(copia, d), d)
        shutil.rmtree(copia, ignore_errors=True)
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
