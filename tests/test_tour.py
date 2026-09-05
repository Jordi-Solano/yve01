# -*- coding: utf-8 -*-
"""Tour guiado (b71): cubre los 9 apartados, el paso de Cierre existe con el
mismo formato que los demas, y cada paso apunta a un elemento que EXISTE
(un id viejo dejaria un paso sin foco). Se recorre entero en Chromium.

  python3.12 tests/test_tour.py
  python3.12 tests/test_tour.py --sabotaje
"""
import logging
import os
import re
import sys
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
logging.getLogger('werkzeug').setLevel(logging.ERROR)
SABOTAJE = '--sabotaje' in sys.argv
PORT = 5103
TABS = ['ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'cierre', 'multi_hotel']


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

    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace("el: '#cierre-stats', tab: 'cierre'", "el: '#cierre-stats-viejo', tab: 'cierre'")
    # traducciones del paso nuevo en los 6 diccionarios inline
    n_tit = html.count('"🧾 Cierre — El mes contable":')
    ok(n_tit == 6, f"el titulo del paso de Cierre esta traducido en los 6 idiomas ({n_tit})")

    srv = make_server('127.0.0.1', PORT, D.app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        with sync_playwright() as p:
            br = p.chromium.launch(executable_path='/opt/pw-browsers/chromium')
            ctx = br.new_context(viewport={'width': 1280, 'height': 900}); pg = ctx.new_page()
            pg.goto(f'http://127.0.0.1:{PORT}/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123'); pg.click('#btn-login')
            pg.wait_for_url(lambda u: '/login' not in u, timeout=20000); pg.wait_for_load_state('networkidle')
            pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_skipped','1')")
            pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(800)
            if SABOTAJE:
                pg.evaluate("_tourSteps.filter(s=>s.tab==='cierre').forEach(s=>{ s.el='#cierre-stats-viejo'; })")
            pasos = pg.evaluate("_tourSteps.map(s=>({el:s.el, tab:s.tab, title:s.title||'', text:s.text||'', existe: !s.el || !!document.querySelector(s.el)}))")
            tabs = [s['tab'] for s in pasos if s['tab']]
            ok(all(t in tabs for t in TABS), f"el tour pasa por los 9 apartados (faltan {[t for t in TABS if t not in tabs]})")
            faltan = [s['el'] for s in pasos if not s['existe']]
            ok(not faltan, f"cada paso apunta a un elemento que existe (rotos: {faltan})")
            c = [s for s in pasos if s['tab'] == 'cierre']
            ok(len(c) == 1 and c[0]['title'].startswith('🧾 Cierre') and 'Debe' in c[0]['text'] and '⚙️' in c[0]['text'], 'el paso de Cierre: que es, que se mira (Debe/Haber, reconciliacion, 303) y de donde sale el archivo')
            ok(all(120 <= len(s['text']) <= 320 and re.match(r'^[^\w\s]\S* \S', s['title']) for s in pasos if s['tab']), 'todos los pasos con el mismo formato: emoji + nombre, y un parrafo de 120-320 caracteres')
            # recorrerlo entero
            pg.evaluate("startTour()"); pg.wait_for_timeout(600)
            vistos = []
            for i in range(len(pasos)):
                st = pg.evaluate("({i:_tourStep, tab:_currentTab, titulo:_tourSteps[_tourStep].title, enCaja:(document.getElementById('tour-box')||{textContent:''}).textContent.indexOf(_tourSteps[_tourStep].title.slice(3))>=0, box: !!document.getElementById('tour-box')})")
                vistos.append(st)
                if i < len(pasos) - 1:
                    pg.evaluate("nextTourStep()"); pg.wait_for_timeout(900)
            paso_c = [v for v in vistos if 'Cierre' in v['titulo']]
            ok(len(paso_c) == 1 and paso_c[0]['tab'] == 'cierre' and paso_c[0]['box'] and paso_c[0]['enCaja'], f"al llegar al paso de Cierre, la pestaña activa es Cierre y la tarjeta lo enseña ({paso_c})")
            ok(len(set(v['titulo'] for v in vistos)) == len(pasos), f"se recorren los {len(pasos)} pasos sin repetir ninguno")
            ctx.close(); br.close()
    finally:
        srv.shutdown()
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
