# -*- coding: utf-8 -*-
"""Movil DE VERDAD (b64): Chromium emulando un iPhone y un Android corrientes —
tamaño, densidad, tactil, user agent movil— y ademas la barra de estado y la
zona segura (notch arriba, barra de gestos abajo), que Chromium no simula:
se fuerzan `--sa-top`/`--sa-bottom` (las variables que la app usa en vez de
`env(safe-area-inset-*)` a pelo) y se comprueba que nada fijo se mete debajo.

Por cada dispositivo y cada apartado (+ el chat abierto):
  - nada desborda a lo ancho
  - ningun elemento fijo/pegado pisa la barra de estado ni la de gestos
  - ningun par de textos/botones visibles se solapa (lo que Jordi veia)

  python3.12 tests/test_movil_real.py
  python3.12 tests/test_movil_real.py --sabotaje
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
PORT = 5093
SA_TOP, SA_BOTTOM = 47, 34          # iPhone con notch: 47 px arriba, 34 abajo
TABS = ['multi_hotel', 'ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'cierre']
DISPOSITIVOS = ['iPhone 13', 'Pixel 5']

# Que se comprueba en el navegador. Devuelve {ancho, fijos_en_zona, solapes}.
JS_MEDIR = r"""
({saTop, saBottom}) => {
  const W = window.innerWidth, H = window.innerHeight;
  const vis = e => { const cs = getComputedStyle(e); if (cs.display==='none'||cs.visibility==='hidden'||cs.opacity==='0') return false; const r = e.getBoundingClientRect(); return r.width>2 && r.height>2; };
  const fixedAnc = e => { for (let x=e; x && x!==document.body; x=x.parentElement){ const p=getComputedStyle(x).position; if (p==='fixed'||p==='sticky') return x; } return null; };
  const ignora = e => e.closest('#yve-splash,#tour-overlay,#tour-card,#top-bar,.g-tbl-wrap,.tbl-wrap,#drr-chart-wrap,canvas,svg');
  // hojas: elementos visibles con texto propio o botones/inputs, sin hijos con texto
  const hojas = [];
  document.querySelectorAll('body *').forEach(e => {
    if (ignora(e) || !vis(e)) return;
    const tag = e.tagName;
    const propio = [...e.childNodes].some(n => n.nodeType===3 && n.textContent.trim());
    const esCtl = /^(BUTTON|INPUT|SELECT|TEXTAREA|A)$/.test(tag);
    if (!propio && !esCtl) return;
    if (e.closest('#yve-splash')) return;
    const r = e.getBoundingClientRect();
    // un inline que salta de linea tiene una caja envolvente que pisa a sus vecinos sin pisarlos: se miran sus fragmentos
    const rs = getComputedStyle(e).display==='inline' ? [...e.getClientRects()] : [r];
    hojas.push({e, r, rs, fixed: fixedAnc(e), txt: (e.innerText||e.value||e.tagName).trim().slice(0,30)});
  });
  // 1) fijos dentro de la zona de la barra de estado / gestos
  const fijos = [];
  hojas.forEach(h => { if (!h.fixed) return; if (h.r.top < saTop-1 && h.r.bottom > 0) fijos.push('arriba: '+h.txt+' @'+Math.round(h.r.top)); if (h.r.bottom > H-saBottom+1 && h.r.top < H) fijos.push('abajo: '+h.txt+' @'+Math.round(h.r.bottom)); });
  // 2) solapes entre hojas que no son pariente/hijo y no estan en capas distintas (fijo vs flujo)
  const sol = [];
  for (let i=0;i<hojas.length;i++) for (let j=i+1;j<hojas.length;j++){
    const a=hojas[i], b=hojas[j];
    if (a.e.contains(b.e)||b.e.contains(a.e)) continue;
    if (!!a.fixed !== !!b.fixed) continue;           // el fab flotando sobre el contenido es normal
    if (a.fixed && b.fixed && a.fixed!==b.fixed) continue;
    let hit = null;
    for (const ra of a.rs) for (const rb of b.rs) { const x = Math.min(ra.right,rb.right)-Math.max(ra.left,rb.left), y = Math.min(ra.bottom,rb.bottom)-Math.max(ra.top,rb.top); if (x>3 && y>3) hit = Math.round(x)+'x'+Math.round(y); }
    if (hit) sol.push(a.txt+' | '+b.txt+' ('+hit+')');
  }
  return {ancho: document.documentElement.scrollWidth <= W+1, fijos, solapes: sol.slice(0,8), nHojas: hojas.length};
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
    css_sab = "#chat-fab{bottom:4px!important} #ap-tbody{margin-left:-40px}" if SABOTAJE else ""
    with sync_playwright() as p:
        br = p.chromium.launch(executable_path='/opt/pw-browsers/chromium')
        for nombre in DISPOSITIVOS:
            dev = p.devices[nombre]
            ctx = br.new_context(**dev)
            pg = ctx.new_page()
            pg.goto(f'http://127.0.0.1:{PORT}/login')
            pg.fill('#username', 'admin'); pg.fill('#password', 'admin123'); pg.click('#btn-login')
            pg.wait_for_url(lambda u: '/login' not in u, timeout=20000); pg.wait_for_load_state('networkidle')
            pg.evaluate("sessionStorage.setItem('yve_splash_shown','1'); localStorage.setItem('tour_done', (typeof _TOUR_VER!=='undefined'?_TOUR_VER:'1')); localStorage.setItem('tour_skipped','1')")
            pg.reload(); pg.wait_for_load_state('networkidle'); pg.wait_for_timeout(600)
            pg.add_style_tag(content=f":root{{--sa-top:{SA_TOP}px;--sa-bottom:{SA_BOTTOM}px}} {css_sab}")
            ua = pg.evaluate("navigator.userAgent"); dpr = pg.evaluate("devicePixelRatio"); w = pg.evaluate("innerWidth")
            ok('Mobile' in ua and dpr >= 2 and w < 500, f"{nombre}: emulado (UA movil, dpr {dpr}, {w} px)")
            html = pg.content()
            ok('var(--sa-top)' in html and 'env(safe-area-inset-top' in html, f"{nombre}: la app usa --sa-top (zona segura simulable)")
            for tab in TABS:
                pg.evaluate(f"switchTab('{tab}', document.querySelector('.tab[onclick*=\"\\'{tab}\\'\"]'))"); pg.wait_for_timeout(700)
                r = pg.evaluate(JS_MEDIR, {'saTop': SA_TOP, 'saBottom': SA_BOTTOM})
                ok(r['ancho'], f"{nombre} · {tab}: no desborda a lo ancho")
                ok(not r['fijos'], f"{nombre} · {tab}: nada fijo pisa barra de estado/gestos {r['fijos'][:3]}")
                ok(not r['solapes'], f"{nombre} · {tab}: sin solapes ({len(r['solapes'])}: {r['solapes'][:3]})")
            # el chat abierto a pantalla completa
            pg.evaluate("toggleChat()"); pg.wait_for_timeout(500)
            r = pg.evaluate(JS_MEDIR, {'saTop': SA_TOP, 'saBottom': SA_BOTTOM})
            ok(not r['fijos'], f"{nombre} · chat abierto: cabecera y entrada fuera de la barra de estado/gestos {r['fijos'][:3]}")
            ok(not r['solapes'], f"{nombre} · chat abierto: sin solapes {r['solapes'][:3]}")
            pg.evaluate("toggleChat()")
            ctx.close()
        br.close()
    srv.shutdown()
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
