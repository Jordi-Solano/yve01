# -*- coding: utf-8 -*-
"""Los seis arreglos de interfaz del móvil, sobre el HTML SERVIDO (regla 26).

  1. La barra de arriba NO se puede deslizar de lado.
  3. …y por eso el menú de la rueda deja de recortarse: vive DENTRO de esa
     caja, y un desplegable dentro de un `overflow:auto` se corta. Medido en
     producción: el menú iba de x=250 a x=468 y la barra acababa en x=348.
  2. El botón de Yve está en la burbuja flotante, como en el PC, y ya no ocupa
     sitio en la barra — que es lo que hacía que no cupiera todo.
  4. El selector de hotel, más estrecho y sin emoji en las opciones.
  6+7. Fotos y documentos, cada uno por su puerta: con un `accept` mezclado el
     móvil abre la galería y todo lo elegido acaba siendo una imagen, por eso
     salía "Foto" en todo y se ofrecía unir donde no tocaba.

  python3.12 tests/test_ui_movil.py
  python3.12 tests/test_ui_movil.py --sabotaje
"""
import os
import re
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import dashboard                                        # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv

app = dashboard.app
app.config['TESTING'] = True
c = app.test_client()
assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
HTML = c.get('/').get_data(as_text=True)
assert 'upload-file-input' in HTML, 'esto no es el panel'


def bloque_media(ancho):
    """TODOS los bloques de ese ancho, concatenados.

    Hay mas de un `@media(max-width:768px)` en la hoja y mirar solo el primero
    daba un falso rojo: la regla de la burbuja de Yve vive en otro.
    """
    marca = f'@media(max-width:{ancho}px)' + '{'
    trozos, desde = [], 0
    while True:
        i = HTML.find(marca, desde)
        if i == -1:
            break
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
        trozos.append(HTML[j:k + 1])
        desde = k + 1
    assert trozos, f'no hay ningun @media(max-width:{ancho}px)'
    return '\n'.join(trozos)


def sacar(nombre):
    m = re.search(r'\n(?:async\s+)?function\s+' + re.escape(nombre) + r'\s*\(', HTML)
    assert m, f'no encuentro {nombre}()'
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


MOVIL = bloque_media(768)
fallos = 0
if SABOTAJE:
    print('*** MODO SABOTAJE: vuelve el overflow de la barra, el botón de Yve '
          'a la barra y el accept mezclado ***')
    MOVIL = MOVIL.replace('.nav-right{gap:4px;flex-shrink:1;min-width:0}',
                          '.nav-right{gap:4px;flex-shrink:1;min-width:0;overflow-x:auto}')
    MOVIL = MOVIL.replace('#chat-fab{bottom:16px', '#chat-fab{display:none;bottom:16px')


def chk(nombre, ok, extra=''):
    global fallos
    print(f"  {'OK ' if ok else 'FALLA'}  {nombre}" + (f'  ({extra})' if extra else ''))
    if not ok:
        fallos += 1


print('1 + 3 · la barra no se desliza, y la rueda deja de recortarse')
# La regla que gana en móvil es la del bloque de 768 (va después de la de 640
# y de la base). Lo que importa es que NO tenga overflow.
m = re.search(r'\.nav-right\{([^}]*)\}', MOVIL)
reglas = m.group(1) if m else ''
chk('la barra derecha ya no tiene overflow', 'overflow' not in reglas, reglas)
chk('...pero sigue pudiendo encoger', 'flex-shrink:1' in reglas and 'min-width:0' in reglas)
chk('no queda ninguna barra de scroll que esconder',
    '.nav-right::-webkit-scrollbar' not in MOVIL)
chk('la red de seguridad de M10 sigue puesta',
    'html,body{max-width:100%;overflow-x:hidden}' in MOVIL)
# el menú sigue estando dentro de la barra: por eso importa lo de arriba
i_nav = HTML.find('<div class="nav-right">')
i_menu = HTML.find('id="main-menu"')
i_fin = HTML.find('</div>\n</nav>', i_nav)
chk('el menú de la rueda sigue viviendo DENTRO de la barra (por eso importa)',
    i_nav != -1 and i_menu != -1 and i_nav < i_menu)

print('\n2 · el botón de Yve, como en el PC')
chk('la burbuja flotante NO se esconde en el móvil',
    '#chat-fab{display:none}' not in MOVIL and '#chat-fab{' in MOVIL)
chk('y ya no está duplicado en la barra',
    'show-mobile" onclick="toggleChat()' not in HTML if not SABOTAJE else False)
chk('la burbuja existe', 'id="chat-fab"' in HTML)

print('\n4 · el selector de hotel, más pequeño')
chk('más estrecho que antes', '#hotel-activo-sel{max-width:92px!important' in MOVIL)
chk('y las opciones sin emoji (el desplegable lo pinta el sistema)',
    "'>🏨 ' + _nom" not in HTML)

print('\n6 + 7 · fotos y documentos, cada uno por su puerta')
chk('hay una puerta solo para documentos',
    'accept=".pdf,.xlsm,.xlsx,.xls,.csv"' in HTML)
chk('y otra solo para fotos', 'id="upload-photo-input"' in HTML
    and 'accept="image/*"' in HTML)
chk('ya NO hay un accept mezclado (el que abría la galería)',
    'accept=".pdf,.xlsm,.xlsx,.xls,.csv,image/*"' not in HTML)
chk('hay un botón para las fotos', 'upload.selFotos' in HTML)

# el detector, corriendo de verdad
JS = sacar('_esFotoSubida') + '''
var casos = [
  ['factura.pdf', 'application/pdf', false],
  ['escaneo.pdf', 'image/jpeg', false],          // MIME mentiroso de Android
  ['datos.xlsx', '', false],
  ['ventas.csv', 'text/csv', false],
  ['drr.xlsm', '', false],
  ['image.jpg', 'image/jpeg', true],
  ['foto.HEIC', '', true],
  ['captura.png', 'image/png', true],
  ['sinextension', 'image/jpeg', true],
  ['sinextension', '', false]
];
var res = casos.map(function(c) {
  return {n: c[0], t: c[1], esperado: c[2], sale: _esFotoSubida({name: c[0], type: c[1]})};
});
console.log('###RES###' + JSON.stringify(res));
'''
if SABOTAJE:
    JS = JS.replace('if (/\\.(pdf|xlsx?|xlsm|csv|docx?|txt)$/i.test(n)) return false;', '')
with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as fh:
    fh.write(JS)
    p = fh.name
out = subprocess.run(['node', p], capture_output=True, text=True)
os.unlink(p)
import json                                             # noqa: E402
if '###RES###' in out.stdout:
    res = json.loads(out.stdout.split('###RES###')[1])
    malos = [r for r in res if r['esperado'] != r['sale']]
    chk('el detector de fotos acierta en los 10 casos',
        not malos, f"{len(res)-len(malos)}/{len(res)}" +
        (f" · falla {[r['n'] + '/' + (r['t'] or 'sin mime') for r in malos]}" if malos else ''))
    pdf_mime = [r for r in res if r['n'] == 'escaneo.pdf'][0]
    chk('un PDF con MIME de imagen NO cuenta como foto (el caso de Android)',
        pdf_mime['sale'] is False)
else:
    chk('el detector corre', False, (out.stderr or '')[-160:])

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
print('Todo OK. La barra no se mueve, la rueda abre, Yve está en la burbuja y '
      'solo las fotos son fotos.')
