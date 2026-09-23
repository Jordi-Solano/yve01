# -*- coding: utf-8 -*-
"""b85 — las comisiones OTA viven dentro de AP (decision de Jordi, 23 sep 2026).

  - ya no hay pestaña "AR — OTAs"; la seccion #panel-ar esta DENTRO de #panel-ap con
    los mismos ids (el JS de siempre sigue pintando ahi)
  - AP es la pestaña que arranca activa
  - switchTab('ar') abre AP y baja a la seccion (tour, avisos y atajos siguen vivos)
  - las descargas de OTAs cuelgan del grupo AP del menu ⚙️
  - el badge rojo de discrepancias/DI cuelga de la pestaña AP
  - los roles que no veian AR — OTAs (F&B, jefe de servicios) no ven la seccion
  - i18n de la seccion en 6 idiomas; el DRR sigue siendo pestaña

  python3.12 tests/test_otas_en_ap.py
  python3.12 tests/test_otas_en_ap.py --sabotaje
"""
import json
import os
import re
import shutil
import sys
import tempfile
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv
os.environ['YVE_BACKUP_HORA'] = 'off'
PORT = 5093


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    from werkzeug.serving import make_server
    from playwright.sync_api import sync_playwright
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace('<div id="panel-ar" class="ap-seccion">', '</div><div id="panel-ar" class="panel g-panel">')
    ia = html.index('id="panel-ap"'); fa = html.index('<!-- /panel-ap -->'); ir = html.index('id="panel-ar"')
    ok('id="tab-ar"' not in html and 'id="tab-ap"' in html and 'id="tab-drr"' in html, 'sin pestaña AR — OTAs; AP y DRR siguen')
    ok(ia < ir < fa and 'class="ap-seccion"' in html, 'la seccion de comisiones esta dentro de AP')
    ok('class="tab active" id="tab-ap"' in html and 'id="panel-ap" class="panel active' in html, 'AP arranca activa')
    for i in ('s-tot', 's-disc', 's-di', 'ota-chart', 'tbl-body', 'ar-recl-list', 'ar-select-all'):
        if 'id="' + i + '"' not in html[ir:fa]:
            ok(False, f'falta el id {i} en la seccion'); break
    else:
        ok(True, 'los ids de siempre siguen en la seccion (el JS no cambia)')
    ok("if (tab === 'ar') { tab = 'ap'; _irA = 'panel-ar'; }" in html and "var _currentTab = 'ap';" in html, "switchTab('ar') es un alias de AP")
    cat = html[html.index('var _DESCARGAS = ['):html.index('function _mesDescarga')]
    ok("{tab: 'ar'," not in cat and "'/api/exportar/ar'" in cat.split("{tab: 'drr'")[0] and "ap.dlOtas" in cat, 'las descargas de OTAs cuelgan del grupo AP')
    ok("_setTabBadge('ap', (stats.discrepancias||0)" in html and "_setTabBadge('ar_otas'" not in html, 'el badge rojo de discrepancias/DI va en la pestaña AP')
    ok("!_rolVeApartado('ar')) _secOta.style.display = 'none'" in html, 'los roles que no veian AR — OTAs no ven la seccion')
    ok("title: '📥 Comisiones OTA (en AP)'" in html, 'el tour la llama Comisiones OTA (en AP)')
    faltan = [l for l in ('en', 'ca', 'fr', 'de', 'it', 'pt') if not all(k in json.load(open(f'static/i18n/{l}.json', encoding='utf-8')) for k in ('ap.comisionesOta', 'ap.dlOtas', 'btn.aprobarARp'))]
    ok(not faltan, f'i18n en los 6 idiomas (faltan {faltan})')
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_ota.js', 'w', encoding='utf-8').write(b)
        import subprocess
        if subprocess.run(['node', '--check', '/tmp/_ota.js'], capture_output=True, text=True).returncode:
            ok(False, 'JS roto'); break

    # en el navegador: switchTab('ar') abre AP y la seccion se ve; DRR sigue siendo pestaña
    srv = make_server('127.0.0.1', PORT, app, threaded=True)
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
                pg.evaluate("document.getElementById('panel-ar').style.display='none'")
            r = pg.evaluate("""(function(){
              var activa = document.querySelector('.tab.active'); var ap = document.getElementById('panel-ap');
              var sec = document.getElementById('panel-ar');
              return {activa: activa && activa.id, apVisible: getComputedStyle(ap).display !== 'none', secDentro: !!(sec && ap.contains(sec)),
                      secVisible: sec && getComputedStyle(sec).display !== 'none', tabs: [...document.querySelectorAll('.tab')].map(function(t){ return t.id; })};
            })()""")
            ok(r['activa'] == 'tab-ap' and r['apVisible'] and r['secDentro'] and r['secVisible'], f"al entrar: AP activa y la seccion de comisiones visible dentro ({r['activa']}, dentro {r['secDentro']}, visible {r['secVisible']})")
            ok('tab-drr' in r['tabs'] and 'tab-ar' not in r['tabs'], f"pestañas: {r['tabs']}")
            pg.evaluate("switchTab('banco', document.getElementById('tab-banco'))"); pg.wait_for_timeout(300)
            pg.evaluate("switchTab('ar')"); pg.wait_for_timeout(900)
            r2 = pg.evaluate("""(function(){ var sec=document.getElementById('panel-ar'); var rc=sec.getBoundingClientRect();
              return {activa: document.querySelector('.tab.active').id, top: Math.round(rc.top), y: Math.round(window.scrollY)}; })()""")
            ok(r2['activa'] == 'tab-ap' and r2['y'] > 100 and r2['top'] < 200, f"switchTab('ar') desde Banco: abre AP y baja a la seccion (scroll {r2['y']}, top {r2['top']})")
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
