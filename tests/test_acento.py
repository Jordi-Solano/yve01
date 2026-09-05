# -*- coding: utf-8 -*-
"""Personalizacion · "acentuar todo" con la guia (b70). Antes las reglas
apuntaban a .sc/.card (piezas viejas) y las burbujas de la guia no cambiaban.
  - con el acento puesto, tarjetas, tiles y vacios llevan el color elegido
  - los colores de significado (verde/rojo/ambar) NO cambian
  - sin "acentuar todo", nada de eso pasa

  python3.12 tests/test_acento.py
  python3.12 tests/test_acento.py --sabotaje
"""
import logging
import os
import sys
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
logging.getLogger('werkzeug').setLevel(logging.ERROR)
SABOTAJE = '--sabotaje' in sys.argv
PORT = 5101
ACC = '#e11d48'   # rojo frambuesa: rgb(225, 29, 72)

JS = """
() => {
  const c = getComputedStyle(document.querySelector('#panel-ap .g-card')).borderColor;
  const k = getComputedStyle(document.querySelector('#panel-ap .g-kpi')).borderColor;
  const v = getComputedStyle(document.querySelector('#panel-ap .g-empty')).borderColor;
  const t = getComputedStyle(document.querySelector('#panel-ap .g-card-title')).color;
  const btn = getComputedStyle(document.documentElement).getPropertyValue('--acc').trim() + ' ' + getComputedStyle(document.querySelector('#panel-ap .g-primary')).backgroundColor;
  const ok = document.querySelector('.g-badge.g-ok') ? getComputedStyle(document.querySelector('.g-badge.g-ok')).color : '';
  return {card: c, kpi: k, vacio: v, titulo: t, btn, badgeOk: ok, clase: document.body.className};
}
"""


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
                pg.add_style_tag(content="body.acentuar-todo .g-card{border-color:#334155!important}")
            pg.evaluate("switchTab('ap', document.querySelector('.tab[onclick*=\"\\'ap\\'\"]'))"); pg.wait_for_timeout(1500)
            pg.evaluate("document.body.insertAdjacentHTML('beforeend', gBadge('g-ok','x'))")
            # sin acentuar todo: nada
            pg.evaluate(f"_customColors.accent='{ACC}'; _customColors.hlAll=false; _applyCustomColors();")
            r0 = pg.evaluate(JS)
            ok('#e11d48' in r0['btn'] and '225, 29, 72' not in r0['card'], f"acento cambiado pero sin 'acentuar todo': --acc lo lleva, las tarjetas no ({r0['btn']} / {r0['card']})")
            # con acentuar todo
            pg.evaluate("_customColors.hlAll=true; _applyCustomColors();")
            r1 = pg.evaluate(JS)
            ok('acentuar-todo' in r1['clase'], "body lleva la clase acentuar-todo")
            ok(all('225, 29, 72' in r1[k] for k in ('card', 'kpi', 'vacio')), f"tarjetas, tiles y vacios con el acento en el borde ({r1['card']} / {r1['kpi']} / {r1['vacio']})")
            ok('225, 29, 72' not in r1['badgeOk'] and '34, 197, 94' in r1['badgeOk'], f"los colores de significado no cambian (badge OK sigue verde: {r1['badgeOk']})")
            # y al quitarlo, vuelve
            pg.evaluate("_customColors.hlAll=false; _applyCustomColors();")
            r2 = pg.evaluate(JS)
            ok('225, 29, 72' not in r2['card'], "al desactivarlo, las tarjetas vuelven a su color")
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
