# -*- coding: utf-8 -*-
"""Guia de estilo, apartado a apartado (b55+). Para cada panel YA convertido:
  - solo piezas de la guia (nada de btn-ref/btn-run/sc/card/badge viejos en el HTML)
  - sin estilos de color/tamaño en linea en el HTML del panel
  - los ids que lee el JS y los onclick siguen ahi
  - UN solo boton primario
  - las funciones de badges devuelven g-badge con los textos unificados (node)
El desbordamiento en el navegador lo mide tests/test_guia_carcasa.py (todos
los apartados, 5 anchos).

  python3.12 tests/test_guia_apartados.py
  python3.12 tests/test_guia_apartados.py --sabotaje
"""
import json
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv

# panel -> (ids obligatorios, trozos de onclick/onchange obligatorios, nº de primarios)
PANELES = {
    'panel-ap': (['ap-total', 'ap-importe', 'ap-matches', 'ap-disc', 'ap-sinpo', 'ap-aprobadas', 'ap-tbody', 'ap-count', 'btnOracle', 'ap-estado-filter', 'prov-mes', 'prov-body', 'aging-tramos', 'aging-body', 'alb-list', 'ap-recl-list'],
                 ['aprobarMatchOK()', 'procesarOracle()', 'filtrarAPPorEstado(this.value)', 'loadProvisiones()', 'href="/aprobaciones-ap/"'], 1),
    'panel-ar': (['s-tot', 's-imp', 's-ok', 's-disc', 's-disc-sub', 's-di', 's-pend', 's-pend-sub', 'ota-chart', 'activity', 'tbl-count', 'ar-select-all', 'tbl-body', 'ar-recl-section', 'ar-recl-resumen', 'ar-recl-list', 'btn-export-selected'],
                 ["toggleSelectAll(this,'ar-row-cb')", 'exportarSeleccionados()', 'href="/aprobaciones-ar/"'], 1),
    'panel-banco': (['banco-modo-chip', 'banco-modo-cambiar', 'banco-progress-bar', 'bk-total', 'bk-conc', 'bk-pend', 'bk-imp-pend', 'bk-diff', 'bk-alertas', 'modal-banco-config', 'banco-modal-cancelar'],
                    ['runConciliacion()', 'abrirModoBanco()', "elegirModoBanco('grupo')", "elegirModoBanco('por_hotel')", 'cerrarModoBanco()', 'href="/conciliacion/"'], 1),
    'panel-drr': (['drr-status', 'drr-oob-badge', 'drr-body', 'drr-metrics', 'drr-drop-zone'], ['openUploadModal()', '_recibirEnProcesar(event.dataTransfer.files)'], 1),
}
FIN = {'panel-ap': '<!-- /panel-ap -->', 'panel-ar': '<!-- /panel-ar -->', 'panel-banco': '<!-- /panel-banco -->', 'panel-drr': '<!-- /panel-drr -->'}

# funciones de badges que se ejecutan con node: (nombre, [(argumento, texto esperado, clase esperada)])
BADGES = [
    ('estadoBadgeAP', [('MATCH_3WAY_OK', 'Match OK', 'g-ok'), ('DISCREPANCIA_PO', 'Discrepancia PO', 'g-err'), ('SIN_PO', 'Sin PO', 'g-warn'), ('PENDIENTE', 'Pendiente', 'g-mute')]),
    ('bEstado', [('CORRECTO', 'Correcto', 'g-ok'), ('DISCREPANCIA', 'Discrepancia', 'g-err'), ('SIN_TARIFA_PACTADA', 'Sin tarifa pactada', 'g-warn'), ('OTA_DESCONOCIDA', 'OTA no reconocida', 'g-mute')]),
    ('bDI', [('CERTIFICADO_OK', 'Cert. DI OK', 'g-ok'), ('FALTA_CERTIFICADO_DI', 'Falta DI', 'g-err'), ('NO_APLICA', 'No aplica', 'g-mute')]),
    ('bApro', [('APROBADA', 'Aprobada', 'g-pur'), ('RECHAZADA', 'Rechazada', 'g-err'), ('', 'Sin decisión', 'g-mute')]),
]


def _func_js(html, nombre):
    i = html.index('function ' + nombre + '(')
    j = html.index('\n}\n', i) + 3
    return html[i:j]


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
        html = html.replace('<div class="g-tiles" id="ar-stats-section">', '<button class="btn-ref" style="color:red">viejo</button><div class="g-tiles" id="ar-stats-section">')
        html = html.replace("CORRECTO:           ['g-ok',   t('est.correcto', 'Correcto')],", "CORRECTO:           ['g-mute',   '✓ CORRECTO'],")
    for pid, (ids, clicks, n_prim) in PANELES.items():
        pa = html[html.index('<div id="' + pid + '"'):html.index(FIN[pid])]
        viejos = [v for v in ('class="btn-ref"', 'class="btn-run"', 'class="sc ', 'class="sc"', 'class="card"', 'class="badge', 'class="tbl-wrap"', 'class="stats"') if v in pa]
        ok(not viejos, f"{pid}: sin piezas viejas ({viejos})")
        feos = [e for e in re.findall(r'style="([^"]*)"', pa) if re.search(r'color:|background:|border-radius:|font-size:|padding:', e)]
        ok(not feos, f"{pid}: sin estilos de color/tamaño en linea ({len(feos)}: {feos[:2]})")
        faltan = [i for i in ids if 'id="' + i + '"' not in pa]
        ok(not faltan, f"{pid}: los ids que lee el JS siguen ahi (faltan {faltan})")
        faltan = [c for c in clicks if c not in pa]
        ok(not faltan, f"{pid}: botones con sus funciones (faltan {faltan})")
        ok(pa.count('g-primary') == n_prim, f"{pid}: {n_prim} boton primario ({pa.count('g-primary')})")
    # badges por node
    prog = "function t(k,d){return d;}\n" + _func_js(html, 'gBadge') + "\n" + "\n".join(_func_js(html, n) for n, _ in BADGES) + "\nvar out={};\n"
    for n, casos in BADGES:
        for arg, _, _ in casos:
            prog += f"out[{json.dumps(n + '|' + arg)}]={n}({json.dumps(arg)});\n"
    prog += "console.log(JSON.stringify(out));"
    open('/tmp/_gb.js', 'w', encoding='utf-8').write(prog)
    rc = subprocess.run(['node', '/tmp/_gb.js'], capture_output=True, text=True)
    out = json.loads(rc.stdout) if rc.returncode == 0 else {}
    ok(rc.returncode == 0, f"las funciones de badges se ejecutan ({rc.stderr[:80]})")
    for n, casos in BADGES:
        malos = [(arg, out.get(n + '|' + arg, '')[:60]) for arg, txt, cls in casos if not (out.get(n + '|' + arg, '').startswith('<span class="g-badge ' + cls + '"') and ('>' + txt + '<') in out.get(n + '|' + arg, ''))]
        ok(not malos, f"{n}: badges unificados de la guia ({malos})")
    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
        open('/tmp/_gap.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_gap.js'], capture_output=True, text=True)
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
