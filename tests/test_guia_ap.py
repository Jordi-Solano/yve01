# -*- coding: utf-8 -*-
"""Guia de estilo aplicada a AP — Proveedores (Jordi, 5 sep 2026).
Variante A en PC, variante B en movil (un solo corte: 768 px), badges y textos
identicos. Comprueba: (1) el HTML del panel usa solo piezas de la guia, (2) en
Chromium nada desborda a 370, 770, 800, 850 y 1280 px, (3) las funciones siguen:
tiles con datos, filtro por estado, fila → modal de detalle, botones con sus
onclick, provisiones/aging/albaranes/reclamaciones pintados con badges g-*.

  python3.12 tests/test_guia_ap.py
  python3.12 tests/test_guia_ap.py --sabotaje
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
    css = cl.get('/static/yve-guia.css').get_data(as_text=True)
    if SABOTAJE:
        # vuelve un boton viejo con estilo en linea, y una fila de tiles que no cabe en movil
        html = html.replace('<div class="g-tiles" id="stats-ap-grid">', '<a href="/api/exportar/ap" class="btn-ref" style="background:var(--acc);color:white">Excel</a><div class="g-tiles" id="stats-ap-grid" style="grid-template-columns:repeat(6,200px)">')
        css = css.replace('@media(max-width:768px){\n  :root{', '@media(max-width:700px){\n  :root{')
    ok(css.startswith('/* ====') and '.g-badge' in css and '.g-kpi' in css and '.g-btn' in css, "la hoja /static/yve-guia.css se sirve y trae las piezas")
    ok('href="/static/yve-guia.css?v=' in html, "el dashboard la carga (con sello de version)")
    ok(css.count('@media(max-width:768px)') >= 3 and '@media(max-width:700px)' not in css and not re.search(r'@media\(max-width:(7[0-5]\d|76[0-79]|8\d\d)px\)', css),
       "un solo punto de corte en la guia: 768 px (el de la app)")
    ok('--g-r-btn:999px' in css and "--g-f-head:'Space Grotesk'" in css, "variante B en movil: botones pildora y Space Grotesk en cifras")

    pa = html[html.index('<div id="panel-ap"'):html.index('<!-- /panel-ap -->')]
    ok('class="btn-ref"' not in pa and 'class="btn-run"' not in pa and 'class="sc' not in pa and 'class="card"' not in pa,
       "el panel AP ya no usa las piezas viejas (btn-ref, btn-run, sc, card)")
    estilos = re.findall(r'style="([^"]*)"', pa)
    feos = [e for e in estilos if re.search(r'color:|background:|border-radius:|font-size:|padding:', e)]
    ok(not feos, f"sin estilos de color/tamaño en linea en el panel ({len(feos)}: {feos[:2]})")
    ok(all(x in pa for x in ('id="ap-total"', 'id="ap-importe"', 'id="ap-matches"', 'id="ap-disc"', 'id="ap-sinpo"', 'id="ap-aprobadas"', 'id="ap-tbody"', 'id="ap-count"', 'id="btnOracle"', 'id="oracle-modo-chip"', 'id="ap-estado-filter"', 'id="prov-mes"', 'id="prov-body"', 'id="aging-tramos"', 'id="aging-body"', 'id="alb-list"', 'id="alb-resumen"', 'id="ap-recl-list"', 'id="ap-recl-resumen"')),
       "todos los ids que lee el JS siguen ahi")
    ok(all(x in pa for x in ('onclick="aprobarMatchOK()"', 'onclick="procesarOracle()"', 'onchange="filtrarAPPorEstado(this.value)"', 'onchange="loadProvisiones()"', 'href="/aprobaciones-ap/"')), "y los botones con sus funciones")
    ok(pa.count('g-primary') == 1, f"UN solo boton primario en el apartado ({pa.count('g-primary')})")
    ok("class=\"g-badge ' + cls" in html and "t('est.matchOk', 'Match OK')" in html and "'MATCH_3WAY_OK':    ['g-ok'" in html, "los estados del matching pasan por gBadge con nombre unificado")
    ok("ALBARAN_FACTURADO:['g-ok'" in html and "gBadge('g-ok', t('reclap.enviada'" in html, "albaranes y reclamaciones usan los MISMOS badges")
    ok("@media(max-width:900px){.nav .pill{display:none}}" in html, "entre 769 y 900 px la barra esconde las pastillas (la rueda cabia fuera)")

    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide en el navegador)")
    else:
        import logging; logging.getLogger('werkzeug').setLevel(logging.ERROR)
        from werkzeug.serving import make_server
        _html = html.encode('utf-8'); _css = css.encode('utf-8')
        def _wsgi(environ, start_response):
            p = environ.get('PATH_INFO')
            if p == '/__ap':
                start_response('200 OK', [('Content-Type', 'text/html; charset=utf-8')]); return [_html]
            if p == '/static/yve-guia.css':
                start_response('200 OK', [('Content-Type', 'text/css; charset=utf-8')]); return [_css]
            return app(environ, start_response)
        srv = make_server('127.0.0.1', 5093, _wsgi, threaded=True)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        # datos: el demo (en carpetas temporales para no tocar el checkout)
        import tempfile, shutil
        tmp = tempfile.mkdtemp(prefix='gap_')
        import demo_generator as DG
        for d in ('datos', 'reportes', 'proc'):
            os.makedirs(os.path.join(tmp, d), exist_ok=True)
        _bk = (DG.DATOS, DG.REPORTES, DG.PROCESADAS)
        try:
            with sync_playwright() as p:
                _exe = '/opt/pw-browsers/chromium' if os.path.exists('/opt/pw-browsers/chromium') else None
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path=_exe)
                anchos = {}
                for w in (370, 770, 800, 850, 1280):
                    ctx = br.new_context(viewport={'width': w, 'height': 800}, is_mobile=(w < 500)); pg = ctx.new_page()
                    pg.goto('http://127.0.0.1:5093/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                    pg.click('#btn-login')
                    try:
                        pg.wait_for_url(lambda u: '/login' not in u, timeout=20000)
                    except Exception:
                        pass
                    pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(500)
                    pg.evaluate("sessionStorage.setItem('yve_splash_shown','1')")
                    pg.goto('http://127.0.0.1:5093/__ap'); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(600)
                    # filas de prueba en la tabla AP, pintadas por el MISMO codigo que en produccion
                    pg.evaluate("""(function(){
                      var tb=document.getElementById('ap-tbody'); tb.innerHTML='';
                      var facts=[{numero_factura:'FM-2026-0412',proveedor:'Frutas del Maresme SL',tipo:'FB',total:1284.6,cuenta_contable:'600',estado:'MATCH_3WAY_OK',accion:'APROBADA'},
                                 {numero_factura:'LV-88213',proveedor:'Lavanderia Valles',tipo:'OTRAS',total:2310,cuenta_contable:'629',estado:'DISCREPANCIA_PO',accion:'',detalle_alerta:'PO 2100 vs factura 2310',importes_cuadran:'NO',aviso_importes:'base+iva 2290'},
                                 {numero_factura:'EN-1',proveedor:'Endesa',tipo:'OTRAS',total:4871.22,cuenta_contable:'628',estado:'SIN_PO',accion:'RECHAZADA',duplicados:2,duplicado_de:'EN-1.pdf'}];
                      window.__facts=facts;
                    })()""")
                    pg.evaluate("""(async function(){
                      var orig=window.fetch; window.fetch=function(u,o){ if(String(u).indexOf('/api/facturas_ap')===0) return Promise.resolve(new Response(JSON.stringify(window.__facts),{headers:{'Content-Type':'application/json'}}));
                        if(String(u).indexOf('/api/stats_ap')===0) return Promise.resolve(new Response(JSON.stringify({total:3,importe:8465.82,matches:1,discrepancias:1,sin_po:1,aprobadas:1}),{headers:{'Content-Type':'application/json'}}));
                        return orig(u,o); };
                      await loadAP();
                    })()""")
                    pg.wait_for_timeout(700)
                    pg.evaluate("switchTab('ap', document.getElementById('tab-ap'))"); pg.wait_for_timeout(400)
                    m = pg.evaluate("""(function(){
                      var body=document.body.scrollWidth, pa=document.getElementById('panel-ap');
                      var peor=0; pa.querySelectorAll('*').forEach(function(e){ if(e.closest('.g-tbl-wrap')) return; var r=e.getBoundingClientRect(); if(r.width>0 && r.right>peor) peor=r.right; });
                      var badges=[...pa.querySelectorAll('#ap-tbody .g-badge')].map(function(b){return b.textContent.trim()});
                      var tiles=[...pa.querySelectorAll('.g-kpi-val')].slice(0,6).map(function(e){return e.textContent.trim()});
                      var sel=document.getElementById('ap-estado-filter'); sel.value='DISCREPANCIA_PO'; sel.dispatchEvent(new Event('change'));
                      var visibles=[...document.querySelectorAll('#ap-tbody tr[data-estado]')].filter(function(r){return r.style.display!=='none'}).length;
                      sel.value=''; sel.dispatchEvent(new Event('change'));
                      var fila=document.querySelector('#ap-tbody tr[data-estado]'); fila && fila.click();
                      var modal=document.getElementById('invoice-modal'); var abierto=!!(modal && getComputedStyle(modal).display!=='none' && modal.querySelectorAll('.g-badge').length>=2);
                      if (typeof closeInvoiceModal==='function') closeInvoiceModal();
                      var btn=document.querySelector('#panel-ap .g-primary'); var r=btn.getBoundingClientRect();
                      var tabRadius=getComputedStyle(btn).borderRadius; var kpiFont=getComputedStyle(pa.querySelector('.g-kpi-val')).fontFamily;
                      var chipOk=getComputedStyle(document.getElementById('oracle-modo-chip')).display;
                      return {body:body, peor:Math.round(peor), badges:badges, tiles:tiles, visibles:visibles, abierto:abierto, radius:tabRadius, kpiFont:kpiFont, nav:document.querySelector('.dropdown').getBoundingClientRect().right};
                    })()""")
                    anchos[w] = m
                    ctx.close()
                br.close()
            for w, m in anchos.items():
                ok(m['body'] <= w and m['peor'] <= w + 1 and m['nav'] <= w, f"{w} px: nada desborda (body {m['body']}, elemento mas a la derecha {m['peor']}, rueda {round(m['nav'])})")
            m = anchos[1280]
            ok(m['tiles'][:6] == ['3', '8.465,82 €', '1', '1', '1', '1'] or (m['tiles'][0] == '3' and '8.465' in m['tiles'][1]), f"los tiles se rellenan con los datos ({m['tiles']})")
            ok('Match OK' in m['badges'] and 'Discrepancia PO' in m['badges'] and 'Sin PO' in m['badges'] and 'Aprobada' in m['badges'] and 'Rechazada' in m['badges'] and 'Sin decisión' in m['badges'] and 'base + IVA ≠ total' in m['badges'],
               f"la tabla pinta los badges unificados ({m['badges']})")
            ok(m['visibles'] == 1, f"el filtro por estado sigue funcionando (1 visible con DISCREPANCIA_PO → {m['visibles']})")
            ok(m['abierto'], "pulsar una fila abre el detalle de la factura (antes #invoice-modal no existia y no se abria nada)")
            ok(m['radius'] == '8px' and 'Space Grotesk' not in m['kpiFont'], f"PC = variante A (radio {m['radius']}, cifras {m['kpiFont'][:20]})")
            mm = anchos[370]
            ok(mm['radius'] == '999px' and 'Space Grotesk' in mm['kpiFont'], f"movil = variante B (radio {mm['radius']}, cifras {mm['kpiFont'][:22]})")
            ok(anchos[370]['badges'] == anchos[1280]['badges'], "los badges (texto y orden) son identicos en movil y PC")
            ok(anchos[770]['radius'] == '8px' and anchos[850]['radius'] == '8px', "770 y 850 px son PC: sin mezclas intermedias")
        finally:
            srv.shutdown(); shutil.rmtree(tmp, ignore_errors=True)
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_ga.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_ga.js'], capture_output=True, text=True)
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
