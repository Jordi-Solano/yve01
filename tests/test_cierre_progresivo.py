# -*- coding: utf-8 -*-
"""b106 (Jordi, 25 sep 2026): "Cierre: que las secciones se pinten una a una segun se
calculan, no todas juntas al final."

Con 8 hilos en Render las seis peticiones del Cierre iban a la vez, se repartian la CPU del
servidor y llegaban todas juntas (~2,3 s). Ahora `loadCierre` las pide UNA A UNA en el orden
de la pantalla (asientos → paquete → cuadre banco → inventarios → fiscal → inmovilizado) y
cada seccion se pinta en cuanto llega. Si se cambia de mes a medias, la cadena vieja se para.

En Chromium, con un `fetch` envuelto que apunta cuando empieza y acaba cada peticion y que
secciones estan ya pintadas en ese momento.

  python3.12 tests/test_cierre_progresivo.py
  python3.12 tests/test_cierre_progresivo.py --sabotaje
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
PORT = 5092
ORDEN = ['/api/cierre/asientos', '/api/cierre/paquete', '/api/cuadre_banco', '/api/inventarios', '/api/fiscal', '/api/inmovilizado']
# el cuerpo de cada seccion: pintada = ya no tiene el "Montando el mes…" (o, fiscal, ya tiene algo)
TRAZA = r"""
(function(){
  window.__traza = [];
  var SEC = {'/api/cierre/asientos':'cierre-recon-body', '/api/cierre/paquete':'paq-body', '/api/cuadre_banco':'cbanco-body',
             '/api/inventarios':'inv-body', '/api/fiscal':'fis-body', '/api/inmovilizado':'inm-body'};
  var pintada = function(id){ var e = document.getElementById(id); if (!e) return false;
    return id === 'fis-body' ? e.innerHTML.trim() !== '' : !e.querySelector('.g-cargando'); };
  var f0 = window.fetch;
  window.fetch = function(u, o){
    var s = String(u && u.url ? u.url : u), ruta = s.split('?')[0].replace(/^https?:\/\/[^\/]+/, '');
    if (!SEC[ruta]) return f0.apply(this, arguments);
    var rec = {u: ruta, mes: (s.match(/mes=([0-9-]+)/) || [])[1] || '', t0: performance.now(), t1: null, pintadas: {}};
    Object.keys(SEC).forEach(function(k){ rec.pintadas[k] = pintada(SEC[k]); });
    window.__traza.push(rec);
    var self = this, args = arguments;
    var espera = (ruta === '/api/cierre/paquete' && window.__lentoPaquete) ? window.__lentoPaquete : 0;
    return new Promise(function(res){ setTimeout(res, espera); })
      .then(function(){ return f0.apply(self, args); })
      .then(function(r){ rec.t1 = performance.now(); return r; });
  };
})();
"""


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
        # vuelve el "todas a la vez" y sin guarda de mes
        html = html.replace('if (sigue()) { try { await load', 'if (true) { try { load')

    # 1. el codigo
    i = html.index('async function loadCierre(forzar){')
    lc = html[i:html.index('\n}\n', i)]
    pos = [lc.find(x) for x in ("fetch('/api/cierre/asientos", 'await loadPaquete()', 'await loadCuadreBanco()', 'await loadInventarios()', 'await loadFiscal()', 'await loadInmovilizado()')]
    ok(all(p > 0 for p in pos) and pos == sorted(pos), f"loadCierre espera cada seccion antes de pedir la siguiente, en el orden de la pantalla ({pos})")
    ok('var gen = ++_cierreGen;' in lc and lc.count('if (sigue())') == 5, "si se pide otro mes, la cadena vieja se para")
    ok('function _pintarAsientosCierre(d, rb, mb, db, av)' in html and 'return;' not in lc.split("fetch('/api/cierre/asientos")[1].split('await loadPaquete')[0],
       "un error en los asientos no corta la cadena")

    # 2. en Chromium
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        sync_playwright = None
    if sync_playwright is None:
        print("  (sin Playwright: no se mide en el navegador)")
    else:
        import logging; logging.getLogger('werkzeug').setLevel(logging.ERROR)
        from werkzeug.serving import make_server
        _cuerpo = html.encode('utf-8')

        def _wsgi(environ, start_response):
            # el panel es el HTML leido (con o sin sabotaje)
            if environ.get('PATH_INFO') == '/' and environ.get('REQUEST_METHOD') == 'GET':
                start_response('200 OK', [('Content-Type', 'text/html; charset=utf-8')])
                return [_cuerpo]
            return app(environ, start_response)
        srv = make_server('127.0.0.1', PORT, _wsgi, threaded=True)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        try:
            with sync_playwright() as p:
                try:
                    br = p.chromium.launch()
                except Exception:
                    br = p.chromium.launch(executable_path='/opt/pw-browsers/chromium')
                ctx = br.new_context(viewport={'width': 1280, 'height': 900})
                ctx.add_init_script(TRAZA)
                pg = ctx.new_page()
                pg.goto(f'http://127.0.0.1:{PORT}/login'); pg.fill('#username', 'admin'); pg.fill('#password', 'admin123')
                pg.click('#btn-login'); pg.wait_for_url(lambda u: '/login' not in u, timeout=20000); pg.wait_for_load_state('networkidle')
                pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_skipped','1'); localStorage.setItem('yve_bancomodo_visto','1')")
                pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(500)
                pg.evaluate("window.__traza = []")
                # 2a. primera entrada en Cierre
                pg.evaluate("switchTab('cierre')")
                pg.wait_for_function("window.__traza.length >= 6 && window.__traza.every(function(r){ return r.t1 !== null; })", timeout=30000)
                pg.wait_for_timeout(600)
                tr = pg.evaluate("window.__traza")
                urls = [r['u'] for r in tr]
                ok(urls == ORDEN, f"se piden en el orden de la pantalla ({[u.split('/')[-1] for u in urls]})")
                solapes = [urls[k] for k in range(1, len(tr)) if tr[k]['t0'] < (tr[k - 1]['t1'] or 1e12)]
                ok(not solapes, f"cada una empieza cuando ha llegado la anterior (se solapan: {solapes})")
                malos = []
                for k, r in enumerate(tr[:len(ORDEN)]):
                    antes = [u for u in ORDEN[:k] if not r['pintadas'].get(u)]
                    despues = [u for u in ORDEN[k:] if r['pintadas'].get(u)]
                    if antes or despues:
                        malos.append((r['u'].split('/')[-1], antes, despues))
                ok(not malos, f"al pedir cada una, las de antes ya estan pintadas y las de despues no ({malos[:2]})")
                # 2b. cambiar de mes a medias: la cadena vieja no sigue
                pg.evaluate("""() => { window.__traza = []; window.__lentoPaquete = 1200;
                    var inp = document.getElementById('cierre-mes'); inp.value = '2026-07'; loadCierre(true);
                    setTimeout(function(){ inp.value = '2026-06'; loadCierre(true); }, 500); }""")
                pg.wait_for_function("window.__traza.filter(function(r){ return r.u === '/api/inmovilizado'; }).length >= 1 && window.__traza.every(function(r){ return r.t1 !== null; })", timeout=30000)
                pg.wait_for_timeout(2500)
                tr = pg.evaluate("window.__traza")
                cuenta = {u.split('/')[-1]: sum(1 for r in tr if r['u'] == u) for u in ORDEN}
                tras = [r for r in tr if r['u'] not in ('/api/cierre/asientos', '/api/cierre/paquete')]
                # (la cadena vieja se para donde la pille el cambio: tras los asientos o tras el paquete)
                ok(all(v == 1 for k, v in cuenta.items() if k not in ('asientos', 'paquete')) and cuenta['asientos'] == 2 and cuenta['paquete'] in (1, 2),
                   f"cambiar de mes a medias: la cadena vieja se para tras su paso ({cuenta})")
                ok(tras and all(r['mes'] == '2026-06' for r in tras), f"y lo que queda se pide del mes nuevo ({sorted(set(r['mes'] for r in tras))})")
                br.close()
        finally:
            srv.shutdown()

    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_cprog.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_cprog.js'], capture_output=True, text=True)
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
