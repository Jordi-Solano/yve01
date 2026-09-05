# -*- coding: utf-8 -*-
"""Sin textos explicativos fijos (b72). Los usuarios son de finanzas: cada
apartado es titulo, cifras, tablas y botones. Lo imprescindible va al tour o a
un tooltip (title, traducible con data-i18n-title), no en pantalla.
  - en el HTML de los 9 apartados no queda ningun `g-sub`
  - los tooltips estan y se traducen
  - con datos (demo), en pantalla no hay frases de ayuda largas fuera de los
    vacios, alertas y tablas

  python3.12 tests/test_textos.py
  python3.12 tests/test_textos.py --sabotaje
"""
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
logging.getLogger('werkzeug').setLevel(logging.ERROR)
SABOTAJE = '--sabotaje' in sys.argv
PORT = 5105
DIRS_DATOS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']
TABS = ['multi_hotel', 'ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'cierre']

# frases de ayuda: texto largo, en un bloque "de ayuda" (g-sub/g-note/g-small/p), que
# no es un dato (sin cifras), fuera de vacios, alertas, tablas, tiles y avisos SMTP
JS = r"""
(tab) => {
  const p = document.getElementById('panel-' + tab);
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 2 && r.height > 2; };
  const out = [];
  p.querySelectorAll('.g-sub, .g-note, .g-small, p').forEach(e => {
    if (!vis(e)) return;
    if (e.closest('.g-empty, .g-alert, table, .g-kpi, .g-row, #notif-smtp-banner, .g-modal, .g-overlay, .mh-card, .rd-day, .rd-daywrap')) return;
    if (e.querySelector('.g-empty, .g-alert, table, .g-kpi')) return;   // un contenedor, no una frase
    const txt = (e.innerText || '').trim();
    if (txt.length < 70) return;
    if (/\d/.test(txt)) return;                 // con cifras es un dato, no una explicacion
    out.push(txt.slice(0, 70));
  });
  return out;
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

    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace('<div id="panel-ap" class="panel g-panel">', '<div id="panel-ap" class="panel g-panel"><div class="g-sub">Facturas recibidas, cruce con albarán y PO, aprobación y contabilización de cada una.</div>')
    for pid in TABS:
        pa = html[html.index('<div id="panel-%s"' % pid):html.index('<!-- /panel-%s -->' % pid)]
        ok(not re.search(r'class="g-sub(?:"| )', pa), f"{pid}: sin g-sub en el HTML del apartado")
    tips = re.findall(r'data-i18n-title="([^"]+)"', html)
    ok(len(tips) >= 20, f"las frases de ayuda pasaron a tooltips traducibles ({len(tips)} data-i18n-title)")
    js = "\n".join(re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S))
    ok("'<div class=\"g-sub\">'" not in js, 'el JS tampoco pinta subtitulos de ayuda')
    # _applyI18nBase traduce el title
    i = js.index('function _applyI18nBase('); j = js.index('\n}\n', i) + 3
    prog = ("var titulos={}; var document={querySelectorAll:function(sel){ return sel.indexOf('title')>=0 ? [{getAttribute:function(){return 'ap.tablaAyuda';}, set title(v){ titulos.t=v; }}] : []; }};\n"
            + js[i:j] + "\n_applyI18nBase({'ap.tablaAyuda':'Click a row'}); console.log(JSON.stringify(titulos));")
    open('/tmp/_tit.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_tit.js'], capture_output=True, text=True)
    ok(rc.returncode == 0 and '"t":"Click a row"' in rc.stdout, f"_applyI18nBase traduce los tooltips ({rc.stderr[:60] or rc.stdout.strip()})")
    ok('ap.tablaAyuda' in json.load(open('static/i18n/en.json', encoding='utf-8')), 'las claves de los tooltips siguen en los json de idioma')

    copia = tempfile.mkdtemp(prefix='yve_textos_')
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
            pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_skipped','1')")
            pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(800)
            gen = pg.evaluate("_postJson('/api/demo/generar', {cadenas:[{nombre:'Cadena Prueba', hoteles:['Hotel Uno','Hotel Dos']}]}).then(r=>r.json())")
            ok(gen.get('ok'), 'demo generado')
            pg.evaluate("_invalidarPaneles && _invalidarPaneles(); loadAll();")
            if SABOTAJE:
                pg.evaluate("document.querySelector('#panel-ap .g-card').insertAdjacentHTML('afterbegin', '<div class=\"g-note\">Cada entrega con sus líneas y la factura con la que ha cruzado; pulsa una fila para ver las líneas.</div>')")
            for tab in TABS:
                if tab == 'multi_hotel': pg.evaluate("_mh_loaded=false")
                pg.evaluate(f"switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))"); pg.wait_for_timeout(2500)
                r = pg.evaluate(JS, tab)
                ok(not r, f"{tab}: sin frases de ayuda en pantalla ({r[:3]})")
            pg.evaluate("_postJson('/api/demo/toggle', {}).then(r=>r.json())"); pg.wait_for_timeout(400)
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
