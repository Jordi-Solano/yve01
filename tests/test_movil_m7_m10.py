# -*- coding: utf-8 -*-
"""M7 · M9 · M10 — las tres del móvil, sobre el HTML que recibe el navegador.

M7  el fondo se movía con el modal abierto: `overflow:hidden` no basta en iOS,
    hay que FIJAR el body y devolver el desplazamiento al cerrar.
M9  el historial de procesados era una tabla de 4 columnas ilegible en móvil;
    ahora cada fila es una tarjeta con etiquetas (`data-r`).
M10 la nav derecha medía 409 px en una pantalla de 370 y arrastraba TODA la
    página 162 px hacia la derecha al cambiar de apartado.

Regla 26: se mide el HTML SERVIDO, no el fuente. La función de M7 se saca del
HTML servido y se corre de verdad en node contra un `document` de mentira.

  python3.12 tests/test_movil_m7_m10.py
  python3.12 tests/test_movil_m7_m10.py --sabotaje
"""
import os
import re
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import dashboard                                       # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv

app = dashboard.app
app.config['TESTING'] = True
c = app.test_client()
assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
HTML = c.get('/').get_data(as_text=True)
assert 'upload-file-input' in HTML, 'esto no es el panel'


def sacar(nombre):
    """Extrae una función del HTML servido, contando llaves."""
    m = re.search(r'\n(?:async\s+)?function\s+' + re.escape(nombre) + r'\s*\(', HTML)
    assert m, f'no encuentro {nombre}() en el HTML servido'
    i = HTML.index('{', m.end() - 1)
    prof, j = 0, i
    while j < len(HTML):
        if HTML[j] == '{':
            prof += 1
        elif HTML[j] == '}':
            prof -= 1
            if prof == 0:
                break
        j += 1
    return HTML[m.start():j + 1]


# ── el CSS servido, sólo el bloque de móvil ─────────────────────────────
def bloque_media(ancho):
    i = HTML.find(f'@media(max-width:{ancho}px)' + '{')
    assert i != -1, f'no encuentro @media(max-width:{ancho}px)'
    j = HTML.index('{', i)
    prof, k = 0, j
    while k < len(HTML):
        if HTML[k] == '{':
            prof += 1
        elif HTML[k] == '}':
            prof -= 1
            if prof == 0:
                break
        k += 1
    return HTML[j:k + 1]


MOVIL = bloque_media(768)

JS_M7 = sacar('_bloquearFondo')
if SABOTAJE:
    # como estaba antes: nadie tocaba el body
    JS_M7 = 'function _bloquearFondo(si){ /* nada */ }'
    MOVIL = MOVIL.replace('flex-shrink:1', 'flex-shrink:0').replace('.hist-t td[data-r]', '.nope')

PRUEBA = r'''
// ── un document/window de mentira, lo justo ──────────────────────────
var _est = {};
var body = {
  dataset: {},
  style: new Proxy({}, {
    set: function(o, k, v) { o[k] = v; _est[k] = v; return true; },
    get: function(o, k) { return o[k] === undefined ? '' : o[k]; }
  })
};
var document = { body: body };
var _scrollA = null;
var window = {
  scrollY: 0, pageYOffset: 0,
  scrollTo: function(x, y) { _scrollA = y; window.scrollY = y; }
};
var scrollY = 0;

__JS__

var res = [];
function chk(nombre, ok, extra) { res.push([nombre, !!ok, extra === undefined ? '' : String(extra)]); }

// 1 · abrir con la página desplazada: el body se fija en el sitio
window.scrollY = 640;
_bloquearFondo(true);
chk('el fondo se fija al abrir', body.style.position === 'fixed', 'position=' + body.style.position);
chk('compensa el desplazamiento (no salta al principio)', body.style.top === '-640px',
    'top=' + body.style.top);

// 2 · abrir DOS veces (modal sobre modal) no pierde el sitio
window.scrollY = 0;                       // el body fijo hace que scrollY sea 0
_bloquearFondo(true);
chk('abrir dos veces no borra el sitio guardado', body.style.top === '-640px',
    'top=' + body.style.top);

// 3 · cerrar: se sueltan los estilos y se vuelve al mismo sitio
_bloquearFondo(false);
chk('al cerrar se suelta el body', body.style.position === '' && body.style.top === '',
    'position=' + JSON.stringify(body.style.position) + ' top=' + JSON.stringify(body.style.top));
chk('al cerrar se vuelve al MISMO punto', _scrollA === 640, 'scrollTo(' + _scrollA + ')');

// 4 · cerrar sin haber abierto no hace nada raro
_scrollA = null;
_bloquearFondo(false);
chk('cerrar sin abrir no mueve la página', _scrollA === null, 'scrollTo=' + _scrollA);

console.log('###RES###' + JSON.stringify(res));
'''.replace('__JS__', JS_M7)

with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as fh:
    fh.write(PRUEBA)
    p = fh.name
out = subprocess.run(['node', p], capture_output=True, text=True)
os.unlink(p)
if '###RES###' not in out.stdout:
    print('node ha fallado:\n' + (out.stderr or out.stdout)[-1200:])
    sys.exit(1)
import json                                             # noqa: E402
res = json.loads(out.stdout.split('###RES###')[1].strip())

fallos = 0
if SABOTAJE:
    print('*** MODO SABOTAJE: el fondo no se fija, la nav no encoge y el '
          'historial no lleva etiquetas ***')

print('M7 · el fondo no se mueve con el modal abierto')
for nombre, ok, extra in res:
    print(f"  {'OK ' if ok else 'FALLA'}  {nombre}" + (f'  ({extra})' if extra else ''))
    if not ok:
        fallos += 1

# M7 · y está enchufada donde toca
for fn, esperado in [('openUploadModal', '_bloquearFondo(true)'),
                     ('closeUploadModal', '_bloquearFondo(false)')]:
    cuerpo = sacar(fn) if not SABOTAJE else ''
    ok = esperado in cuerpo
    print(f"  {'OK ' if ok else 'FALLA'}  {fn}() llama a {esperado}")
    if not ok:
        fallos += 1

print('\nM9 · el historial, legible en el móvil')
comprob_m9 = [
    ('la tabla lleva la clase que la convierte en tarjetas',
     'class="hist-t"' in HTML),
    ('las celdas llevan su etiqueta', 'data-r="Archivo"' in HTML
     and 'data-r="Dónde"' in HTML and 'data-r="Fecha"' in HTML),
    ('en móvil la cabecera se esconde', '.hist-t thead{display:none}' in MOVIL),
    ('en móvil cada fila es un bloque', '.hist-t tr{' in MOVIL),
    ('la etiqueta se pinta delante del dato', '.hist-t td[data-r]::before' in MOVIL),
    ('el nombre del archivo ya no se corta',
     'white-space:normal!important' in MOVIL and 'max-width:none!important' in MOVIL),
]
for nombre, ok in comprob_m9:
    print(f"  {'OK ' if ok else 'FALLA'}  {nombre}")
    if not ok:
        fallos += 1

print('\nM10 · la pantalla ya no se desplaza de lado')
# el orden importa: la regla de 768 va DESPUES de la de 640, y las dos
# despues de la base con flex-shrink:0. Si se colara antes, no ganaría.
i_base = HTML.find('.nav-right{display:flex')
i_640 = HTML.find('.nav-right{gap:6px}')
i_768 = HTML.find('.nav-right{gap:4px')
comprob_m10 = [
    ('la nav derecha puede encoger', 'flex-shrink:1' in MOVIL and 'min-width:0' in MOVIL),
    ('si no cabe, se desliza DENTRO de la nav', 'overflow-x:auto' in MOVIL),
    # Medido en el navegador: el selector lleva `max-width:190px` EN LINEA, y
    # el estilo inline gana a cualquier hoja. Sin `!important` la regla existe
    # y no hace NADA — que es peor que no ponerla, porque parece hecha.
    ('el selector de hotel tiene tope', '#hotel-activo-sel{max-width:104px!important}' in MOVIL),
    ('...y ese tope gana al estilo en línea del propio elemento',
     re.search(r'id="hotel-activo-sel"[^>]*style="[^"]*max-width:\s*190px', HTML) is not None
     and 'max-width:104px!important' in MOVIL),
    ('red de seguridad: nada puede desplazar la página',
     'html,body{max-width:100%;overflow-x:hidden}' in MOVIL),
    ('la regla nueva va DESPUÉS de las que anula (si no, no ganaría)',
     i_base != -1 and i_640 != -1 and i_768 != -1 and i_base < i_640 < i_768),
]
for nombre, ok in comprob_m10:
    print(f"  {'OK ' if ok else 'FALLA'}  {nombre}")
    if not ok:
        fallos += 1

print()
if SABOTAJE:
    if fallos:
        print(f'SABOTAJE OK: {fallos} en rojo.')
        sys.exit(0)
    print('SABOTAJE MAL.')
    sys.exit(1)
if fallos:
    print(f'{fallos} en rojo')
    sys.exit(1)
print('Todo OK. El fondo se queda quieto, el historial se lee en el móvil y la '
      'página no se va de lado.')
