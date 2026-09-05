# -*- coding: utf-8 -*-
"""Estado vacio UNICO (b67). Con la app SIN datos, en los 9 apartados:
  - cada "sin datos" es el mismo bloque: icono, "Sin datos todavía", frase y
    (salvo lo que no viene de ficheros) el boton ⚡ Procesar Archivos
  - los tiles no pintan ceros: "—", sin color, sin subtitulo
  - los graficos no pintan barras: el lienzo se esconde y sale el vacio
  - en el JS no queda ningun vacio escrito a mano
Y con datos (demo) los tiles vuelven a tener numero.

  python3.12 tests/test_vacios.py
  python3.12 tests/test_vacios.py --sabotaje
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
PORT = 5099
DIRS_DATOS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']
TABS = ['multi_hotel', 'ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'cierre']
TITULO = 'Sin datos todavía'

JS_MEDIR = r"""
(tab) => {
  const p = document.getElementById('panel-' + tab);
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 2 && r.height > 2 && getComputedStyle(e).display !== 'none'; };
  const vacios = [...p.querySelectorAll('.g-empty')].filter(e => vis(e) && !e.classList.contains('g-cargando') && !e.classList.contains('is-ok'));
  const malos = vacios.filter(e => !(e.querySelector('.g-empty-ico') && (e.querySelector('b')||{}).textContent === 'Sin datos todavía' && e.querySelector('.g-empty-sub')));
  const conCta = vacios.filter(e => e.querySelector('.g-empty-cta button')).length;
  const tiles = [...p.querySelectorAll('.g-kpi')].filter(vis);
  const tilesConNumero = tiles.filter(k => !k.classList.contains('is-empty')).map(k => (k.querySelector('.g-kpi-val')||{}).textContent);
  const tilesConCero = tiles.filter(k => /^\s*0([,.]00)?\s*(€)?\s*$/.test((k.querySelector('.g-kpi-val')||{}).textContent||'')).length;
  const lineaColor = tiles.filter(k => k.classList.contains('is-empty') && getComputedStyle(k, '::before').opacity !== '0').length;
  const canvases = [...p.querySelectorAll('canvas')].filter(vis).length;
  const cargando = [...p.querySelectorAll('.g-cargando')].filter(vis).length;
  const barras = [...p.querySelectorAll('.g-progress-fill')].filter(vis).length;
  return {nVacios: vacios.length, malos: malos.map(e => e.textContent.trim().slice(0, 40)), conCta, nTiles: tiles.length, tilesConNumero, tilesConCero, lineaColor, canvases, cargando, barras};
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
        html = html.replace("function _todoOk(texto){", "function _todoOk(texto){ return '<div class=\"g-empty\">' + texto + '</div>';")
    # 1) en el JS no queda ningun vacio a mano
    js = "\n".join(re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S))
    sueltos = [m.strip() for m in re.findall(r"[^\n]{0,50}class=\\?\"g-empty\\?\"[^\n]{0,50}", js)
               if 'function _vacio' not in m and 'function _todoOk' not in m and 'function _cargando' not in m and 'data-vacio-de' not in m]
    ok(not sueltos, f"en el JS ningun vacio escrito a mano: todos por _vacio/_todoOk ({sueltos[:2]})")
    # 2) las funciones producen el MISMO bloque
    prog = "function t(k,d){return d;}\nfunction openUploadModal(){}\n" + "\n".join(_func(html, n) for n in ('gBadge', '_vacio', '_gVacio', '_emptyState', '_todoOk')) + "\nconsole.log(JSON.stringify([_vacio('x'), _gVacio('x'), _emptyState('🍽️','titulo','x'), _vacio('x',{cta:false}), _todoOk('nada')]));"
    open('/tmp/_vac.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_vac.js'], capture_output=True, text=True)
    outs = json.loads(rc.stdout) if rc.returncode == 0 else ['', '', '', '', '']
    ok(rc.returncode == 0, f"las funciones del vacio se ejecutan ({rc.stderr[:60]})")
    ok(all('<b>' + TITULO + '</b>' in o and '📂' in o and 'g-empty-sub' in o for o in outs[:4]), '_vacio, _gVacio y _emptyState: mismo icono, mismo titulo, misma frase')
    ok(all('⚡ Procesar Archivos' in o and 'openUploadModal()' in o for o in outs[:3]) and 'Procesar' not in outs[3], 'el boton ⚡ Procesar Archivos sale por defecto y se quita con cta:false')
    ok(outs[0] == outs[1] and outs[2] == outs[0].replace('class="g-empty"', 'class="g-empty g-empty-lg"'), 'los tres caminos dan literalmente el mismo HTML (el grande solo añade g-empty-lg)')
    ok('g-badge g-ok' in outs[4] and 'is-ok' in outs[4] and TITULO not in outs[4], '_todoOk: verde y sin "Sin datos" (no hay nada que hacer ≠ no hay datos)')
    # 2b) Multi-Hotel: un hotel del censo sin documentos -> sus 4 bloques en "—" (nada de "0 €")
    prog = "function t(k,d){return d;}\nfunction _fmtEurES(v){return String(v)+' €';}\nfunction gBadge(c,x){return '<span class=\"g-badge '+c+'\">'+x+'</span>';}\n" + "\n".join(_func(html, n) for n in ('_mhEur', '_mhK', '_mhBloque', '_mhTarjeta', '_mhFilaHotelera')) + \
        "\nvar vacio={hotel_id:'h',nombre:'H',censo:{},drr:{estado:'sin_drr'},ap:{importe:0,facturas:0,discrepancias:0},ar_ota:{importe_reclamable:0,importe_bruto:0,facturas:0},ar_real:{pendiente:0,vencido:0,facturas:0},fb:{ventas:0,food_cost_pct:0}};" + \
        "\nvar lleno=JSON.parse(JSON.stringify(vacio)); lleno.ap={importe:1200,facturas:3,discrepancias:0};" + \
        "\nconsole.log(JSON.stringify([_mhTarjeta(vacio,'hotel'), _mhTarjeta(lleno,'hotel')]));"
    open('/tmp/_mh.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_mh.js'], capture_output=True, text=True)
    mh = json.loads(rc.stdout) if rc.returncode == 0 else ['', '']
    ok(rc.returncode == 0 and mh[0].count('is-empty') == 4 and '0 €' not in mh[0], f"Multi-Hotel: hotel sin documentos -> 4 bloques en '—' ({rc.stderr[:60]})")
    ok(mh[1].count('is-empty') == 0 and '1200 €' in mh[1], 'Multi-Hotel: hotel con documentos -> numeros')
    # 3) i18n: el titulo existe en los 6 idiomas
    faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if 'vacio.titulo' not in json.load(open(f'static/i18n/{l}.json', encoding='utf-8'))]
    ok(not faltan, f"'vacio.titulo' en los 6 idiomas (faltan {faltan})")

    # 4) en el navegador, sin datos y con datos
    copia = tempfile.mkdtemp(prefix='yve_vacios_')
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
                pg.add_style_tag(content=".g-kpi.is-empty::before{opacity:1!important}")
            # vacio: limpiar por si el demo quedo activo
            pg.evaluate("_postJson('/api/demo/toggle',{}).then(r=>r.json()).then(d=>{ if (d.demo_mode) return _postJson('/api/demo/toggle',{}); })"); pg.wait_for_timeout(500)
            pg.evaluate("_invalidarPaneles && _invalidarPaneles()")
            for tab in TABS:
                pg.evaluate(f"switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))"); pg.wait_for_timeout(2500)
                r = pg.evaluate(JS_MEDIR, tab)
                ok(r['nVacios'] >= 1 and not r['malos'], f"sin datos · {tab}: {r['nVacios']} vacio(s) y todos con el patron ({r['malos'][:2]})")
                ok(r['tilesConNumero'] == [] and r['tilesConCero'] == 0, f"sin datos · {tab}: {r['nTiles']} tiles en '—', ninguno con cero ({r['tilesConNumero'][:3]})")
                ok(r['lineaColor'] == 0 and r['canvases'] == 0 and r['barras'] == 0, f"sin datos · {tab}: sin linea de color, sin graficos, sin barras ({r['lineaColor']}/{r['canvases']}/{r['barras']})")
                ok(r['cargando'] == 0, f"sin datos · {tab}: no queda ningun 'Cargando…' colgado")
            # con datos
            gen = pg.evaluate("_postJson('/api/demo/generar', {cadenas:[{nombre:'Cadena Prueba', hoteles:['Hotel Uno','Hotel Dos']}]}).then(r=>r.json())")
            ok(gen.get('ok'), 'demo generado para la prueba con datos')
            pg.evaluate("_invalidarPaneles && _invalidarPaneles(); loadAll();"); pg.wait_for_timeout(1500)
            for tab in ['ar', 'ap', 'multi_hotel', 'fb', 'cierre']:
                if tab == 'multi_hotel': pg.evaluate("_mh_loaded=false")
                pg.evaluate(f"switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))"); pg.wait_for_timeout(2500)
                r = pg.evaluate(JS_MEDIR, tab)
                ok(len(r['tilesConNumero']) >= 3 and all(v != '—' for v in r['tilesConNumero']), f"con datos · {tab}: los tiles vuelven a tener numero ({r['tilesConNumero'][:4]})")
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


def _func(html, nombre):
    i = html.index('function ' + nombre + '(')
    j = html.index('\n}\n', i) + 3
    return html[i:j]


if __name__ == '__main__':
    main()
