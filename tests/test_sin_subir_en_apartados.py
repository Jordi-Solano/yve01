# -*- coding: utf-8 -*-
"""b103 (Jordi, 25 sep 2026): "Quedamos en que todo entra por Procesar archivos.
Quita los que queden en F&B y revisa que no quede ninguno en ningun apartado."

Ningun apartado tiene boton, zona de arrastre ni <input type=file> para subir
documentos. La entrada es ⚡ Procesar archivos: el boton de la cabecera (#btn-run)
y, cuando un apartado esta vacio, el boton del vacio (`_vacio`, b67), que abre lo
mismo. Excepcion a proposito: "📎 Adjuntar informe" en la peticion de credito
(b90: el informe de Informa se adjunta a ESA peticion, o se anota).

  python3.12 tests/test_sin_subir_en_apartados.py
  python3.12 tests/test_sin_subir_en_apartados.py --sabotaje
"""
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv
PALABRAS = re.compile(r'\b(subir|sube|súbelo|importar|arrastra|adjuntar|cargar fichero|upload)\b', re.I)


def _texto(h):
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', h)).strip()


def _func_js(html, nombre):
    i = html.index('function ' + nombre + '(')
    return html[i:html.index('\n}\n', i) + 3]


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
        # vuelve el boton de F&B y aparece un input suelto en el Cierre
        html = html.replace('<div class="g-subtabs" id="fb-subtabs">',
                            '<button class="g-btn g-primary g-sm" onclick="openUploadModal()">📤 Subir mermas</button><div class="g-subtabs" id="fb-subtabs">')
        html = html.replace('<div id="inv-resumen"', '<input type="file" onchange="x(this)"><div id="inv-resumen"')

    # 1. apartado a apartado (el HTML fijo de cada panel)
    paneles = re.findall(r'<div id="(panel-[a-z_]+)"', html)
    ok(len(paneles) >= 10, f"{len(paneles)} apartados en la pagina")
    for pid in paneles:
        i = html.index('<div id="' + pid + '"')
        fin = html.find('<!-- /' + pid + ' -->', i)
        pa = html[i:fin if fin > 0 else i + 20000]
        malos = []
        if 'openUploadModal(' in pa:
            malos.append('abre Procesar archivos desde el apartado')
        if re.search(r'type=["\']?file', pa):
            malos.append('input de fichero')
        if 'ondrop=' in pa or 'ondragover=' in pa:
            malos.append('zona de arrastre')
        for tag, dentro in re.findall(r'<(button|label|a)\b[^>]*>(.*?)</\1>', pa, re.S):
            if PALABRAS.search(_texto(dentro)):
                malos.append(f'{tag} «{_texto(dentro)[:50]}»')
        ok(not malos, f"{pid}: sin botones de subir ({malos})")

    # 2. la pagina entera: los unicos <input type=file> son los de Procesar archivos y el de Informa
    inputs = re.findall(r'<input\b[^>]*type=["\']?file[^>]*>', html)
    onchange = sorted(re.search(r'onchange="([^"]*)"', x).group(1) if 'onchange="' in x else '?' for x in inputs)
    esperado = sorted(['handleUploadFiles(this.files, this)'] * 3 + ['subirInformaCredito(this)'])
    ok(onchange == esperado, f"inputs de fichero: los 3 de Procesar archivos + Informa ({onchange})")
    ok(all(x in html for x in ('id="upload-file-input"', 'id="upload-photo-input"', 'id="upload-folder-input"')), "los 3 de Procesar archivos siguen")
    # quien abre Procesar archivos: la cabecera, la funcion y el boton del vacio (b67)
    vacio = _func_js(html, '_vacio')
    n = html.count('openUploadModal(')
    ok('id="btn-run" onclick="openUploadModal()"' in html and 'async function openUploadModal()' in html and 'onclick="openUploadModal()"' in vacio,
       "⚡ Procesar archivos: boton de la cabecera y boton del vacio")
    ok(n == 3, f"nadie mas abre Procesar archivos ({n} apariciones: cabecera, funcion, vacio)")
    drops = re.findall(r'ondrop="([^"]*)"', html)
    ok(drops == ['handleUploadDrop(event)'], f"una sola zona de arrastre: la de Procesar archivos ({drops})")
    ok(all(x not in html for x in ('drr-drop-zone', '_recibirEnProcesar', 'g-drop', 'drr.subirBoton', 'inv.subir"', 'btn.importarFB')), "fuera la zona del DRR y los botones de DRR, Cierre y F&B")
    # el tour del DRR ya no dice "arrastra aqui"
    paso = html[html.index("el: '#drr-metrics'"):][:600]
    ok('Arrastra' not in paso and 'Procesar Archivos' in paso, "tour del DRR: sube con ⚡ Procesar Archivos")
    # textos de ayuda del F&B y del Cierre apuntan a Procesar archivos
    ok("Sube un inventario con ⚡ Procesar Archivos" in html and "Súbelo con ⚡ Procesar archivos" in html, "los vacios de F&B dicen por donde se sube")

    # 3. las otras paginas
    for url in ('/conciliacion/', '/aprobaciones-ap/', '/aprobaciones-ar/', '/admin/', '/onboarding/'):
        r = cl.get(url)
        if r.status_code != 200:
            ok(False, f"{url}: {r.status_code}"); continue
        pg = r.get_data(as_text=True); r.close()
        malos = [m for m in re.findall(r'<input\b[^>]*type=["\']?file[^>]*>', pg)]
        malos += [_texto(d)[:50] for t_, d in re.findall(r'<(button|a)\b[^>]*>(.*?)</\1>', pg, re.S) if PALABRAS.search(_texto(d))]
        ok(not malos, f"{url}: sin botones de subir ({malos})")
    conc = cl.get('/conciliacion/').get_data(as_text=True)
    ok('Los extractos entran por ⚡ Procesar archivos' in conc, "/conciliacion dice por donde entra el extracto (texto, no boton)")

    # 4. i18n: fuera las claves de los botones quitados
    import json
    for lang in ('en', 'ca', 'fr', 'de', 'it', 'pt'):
        d = json.load(open(os.path.join(BASE, 'static', 'i18n', f'{lang}.json'), encoding='utf-8'))
        sobran = [k for k in ('btn.importarFB', 'drr.arrastra', 'drr.hazClic', 'drr.subirBoton', 'inv.subir') if k in d]
        if sobran:
            ok(False, f"i18n {lang}: sobran {sobran}"); break
    else:
        ok(True, "i18n: sin las claves de los botones quitados (6 idiomas)")

    for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html + conc, re.S):
        open('/tmp/_ssa.js', 'w', encoding='utf-8').write(b)
        rc = subprocess.run(['node', '--check', '/tmp/_ssa.js'], capture_output=True, text=True)
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
