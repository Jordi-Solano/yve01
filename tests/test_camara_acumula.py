# -*- coding: utf-8 -*-
"""PASO 3 — el nombre repetido deja de tirar fotos (y la camara funciona).

Causa real del bug de la camara, confirmada por Jordi en su movil: el telefono
llama `image.jpg` a TODAS las fotos de camara, y el dedup por nombre
descartaba la 2a y siguientes.

Propiedades que protege, sobre las funciones REALES sacadas del HTML servido
(regla 26), no sobre una copia escrita a mano:

  1. Dos fotos DISTINTAS con el mismo nombre entran LAS DOS, con nombre unico.
  2. El MISMO fichero exacto dos veces entra una sola vez, y se dice.
  3. Nada se pierde: recibidas = en la lista + repetidas + ilegibles.
  4. El orden DA IGUAL: ninguna permutacion pierde un documento.
  5. `recibo.png` y `recibo.jpeg` no chocan al comprimirse a `.jpg`.
  6. El input se limpia despues de cada seleccion.

  python3.12 tests/test_camara_acumula.py
  python3.12 tests/test_camara_acumula.py --sabotaje
"""
import itertools
import json
import os
import re
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv
FUNCIONES = ('_pareceDocumento', '_claveNombre', '_huellaFichero', '_nombreLibre',
             '_esImagen', 'handleUploadFiles', '_addFilesToList')

# Que avisos se ven y cuales no. Renombrar es SILENCIOSO a proposito (con la
# camara todas las fotos se llaman igual, asi que avisar en cada tanda seria
# ruido); los otros dos si se dicen.
#   (titulo, [recibidos, anadidos, ilegibles, repetidos, renombrados], se_ve, texto_que_debe_salir)
AVISOS = [
    ('Solo renombrados: NO se avisa', [3, 3, 0, 0, 2], False, None),
    ('Muchos renombrados (25 fotos de camara): NO se avisa', [25, 25, 0, 0, 24], False, None),
    ('El mismo fichero otra vez: SI se avisa', [2, 1, 0, 1, 0], True, 'ya estaba en la lista'),
    ('Fichero ilegible: SI se avisa', [1, 0, 1, 0, 0], True, 'no identifica'),
    ('Renombrado + ilegible: se avisa SOLO de lo ilegible', [3, 2, 1, 0, 1], True, 'no identifica'),
    ('Nada raro: NO se avisa', [2, 2, 0, 0, 0], False, None),
]


def html_servido():
    import dashboard
    app = dashboard.app
    app.config['TESTING'] = True
    c = app.test_client()
    r = c.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
    assert r.status_code == 200, 'login fallido: se serviria la landing'
    r = c.get('/')
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    assert 'upload-file-input' in html, 'esto no es el panel'
    return html


def extraer_funcion(js, nombre):
    m = re.search(r'function\s+' + re.escape(nombre) + r'\s*\(', js)
    assert m, f'no encuentro {nombre} en el HTML servido'
    i = js.index('{', m.end())
    prof, j = 0, i
    while j < len(js):
        if js[j] == '{':
            prof += 1
        elif js[j] == '}':
            prof -= 1
            if prof == 0:
                return js[m.start():j + 1]
        j += 1
    raise AssertionError(f'llaves sin cerrar en {nombre}')


# El descarte mudo de antes: la 2a foto de camara se iba sin dejar rastro.
VIEJO = """
function _addFilesToList(newFiles) {
  var existing = new Set(_uploadFiles.map(function(f){ return f.name; }));
  (newFiles || []).forEach(function(f) { if (!existing.has(f.name)) { _uploadFiles.push(f); } });
  _renderFileList();
  return { anadidos: 0, repetidos: [], renombrados: [] };
}
function handleUploadFiles(fileList, input) {
  _addFilesToList(Array.from(fileList).filter(_pareceDocumento));
}
"""

# (nombre, tipo, tamaño, fecha) — tamaño y fecha distinguen "otra foto" de
# "el mismo fichero otra vez".
CASOS = [
    ('CAMARA — 3 capturas, todas `image.jpg` (SU movil)',
     [[('image.jpg', 'image/jpeg', 1000, 111)],
      [('image.jpg', 'image/jpeg', 2000, 222)],
      [('image.jpg', 'image/jpeg', 3000, 333)]],
     {'en_lista': 3, 'repetidos': 0, 'desconocidos': 0}),

    ('El MISMO fichero elegido dos veces: entra una vez',
     [[('IMG_0101.jpg', 'image/jpeg', 5000, 999)],
      [('IMG_0101.jpg', 'image/jpeg', 5000, 999)]],
     {'en_lista': 1, 'repetidos': 1, 'desconocidos': 0}),

    ('GALERIA — nombres unicos, como siempre',
     [[('IMG_0101.jpg', 'image/jpeg', 1, 1), ('IMG_0102.jpg', 'image/jpeg', 2, 2),
       ('IMG_0103.jpg', 'image/jpeg', 3, 3)]],
     {'en_lista': 3, 'repetidos': 0, 'desconocidos': 0}),

    ('Tres `image.jpg` en la MISMA tanda (camara en rafaga)',
     [[('image.jpg', 'image/jpeg', 10, 1), ('image.jpg', 'image/jpeg', 20, 2),
       ('image.jpg', 'image/jpeg', 30, 3)]],
     {'en_lista': 3, 'repetidos': 0, 'desconocidos': 0}),

    ('recibo.png y recibo.jpeg: chocarian AL COMPRIMIR, no aqui',
     [[('recibo.png', 'image/png', 10, 1)], [('recibo.jpeg', 'image/jpeg', 20, 2)]],
     {'en_lista': 2, 'repetidos': 0, 'desconocidos': 0}),

    ('25 fotos de camara seguidas, todas `image.jpg`',
     [[('image.jpg', 'image/jpeg', 100 + i, i)] for i in range(25)],
     {'en_lista': 25, 'repetidos': 0, 'desconocidos': 0}),

    ('Sin extension y sin tipo: no entra, pero se DICE',
     [[('captura', '', 5, 5)]],
     {'en_lista': 0, 'repetidos': 0, 'desconocidos': 1}),

    ('Mezcla: camara repetida + el mismo fichero + uno ilegible',
     [[('image.jpg', 'image/jpeg', 10, 1)],
      [('image.jpg', 'image/jpeg', 20, 2), ('image.jpg', 'image/jpeg', 10, 1),
       ('raro', '', 1, 1)]],
     {'en_lista': 2, 'repetidos': 1, 'desconocidos': 1}),
]

# El orden no puede decidir quien sobrevive: todas las permutaciones de esta
# tanda tienen que dejar los mismos 4 documentos.
TANDA_ORDEN = [('image.jpg', 'image/jpeg', 10, 1), ('image.jpg', 'image/jpeg', 20, 2),
               ('recibo.png', 'image/png', 30, 3), ('recibo.jpeg', 'image/jpeg', 40, 4)]

PRELUDIO = '''
// Shim minimo de File: el codigo solo usa name, type, size y lastModified.
class File {
  constructor(partes, nombre, opts) {
    const o = (partes && partes[0]) || {};
    this.name = nombre;
    this.type = (opts && opts.type) || o.type || '';
    this.size = o.size || 0;
    this.lastModified = (opts && opts.lastModified) || o.lastModified || 0;
  }
}
let _uploadFiles = [];
let _processedNames = new Set();
function _renderFileList() {}
let ULTIMO = null;
function _avisoDescartes(recibidos, anadidos, desconocidos, repetidos, renombrados) {
  ULTIMO = { recibidos, anadidos, desconocidos: (desconocidos||[]).length,
             repetidos: (repetidos||[]).length, renombrados: (renombrados||[]).length };
}
function correr(tandas) {
  _uploadFiles = [];
  let rep = 0, desc = 0, ren = 0, recib = 0, inputLimpio = true;
  const input = { value: '' };
  tandas.forEach(function(tanda) {
    ULTIMO = null;
    const files = tanda.map(function(x) {
      return { name: x[0], type: x[1], size: x[2], lastModified: x[3] };
    });
    recib += files.length;
    input.value = 'C:\\\\fakepath\\\\' + files[0].name;
    handleUploadFiles(files, input);
    if (input.value !== '') inputLimpio = false;
    if (ULTIMO) { rep += ULTIMO.repetidos; desc += ULTIMO.desconocidos; ren += ULTIMO.renombrados; }
  });
  const nombres = _uploadFiles.map(function(f) { return f.name; });
  const claves = nombres.map(_claveNombre);
  return { en_lista: _uploadFiles.length, nombres: nombres,
           claves_unicas: new Set(claves).size === claves.length,
           repetidos: rep, desconocidos: desc, renombrados: ren,
           recibidos: recib, input_limpio: inputLimpio };
}
'''


def avisos_resultado(html):
    """Corre el _avisoDescartes REAL contra una caja de mentira y mira que sale."""
    fn = extraer_funcion(html, '_avisoDescartes')
    if SABOTAJE:
        # Vuelve el aviso de renombrado que Jordi pidio quitar.
        fn = fn.replace('caja.style.borderColor = \'rgba(245,158,11,.35)\';', '''
  if (!fuera && renombrados.length) {
    caja.innerHTML = '<div>' + renombrados.length + ' foto se llamaba igual que otra.</div>';
    caja.style.display = 'block'; return;
  }
  caja.style.borderColor = 'rgba(245,158,11,.35)';''', 1)
    guion = '''
const caja = { style: {}, innerHTML: '' };
Object.defineProperty(caja, 'innerText', { get() {
  return String(this.innerHTML).replace(/<[^>]*>/g, ' ').replace(/\\s+/g, ' ').trim();
} });
const document = { getElementById: function() { return caja; } };
''' + fn + '''
const CASOS = ''' + json.dumps([a[1] for a in AVISOS]) + ''';
console.log(JSON.stringify(CASOS.map(function(a) {
  caja.style = {}; caja.innerHTML = '';
  _avisoDescartes(a[0], a[1], new Array(a[2]).fill({name:'raro'}),
                  new Array(a[3]).fill('dup.jpg'),
                  new Array(a[4]).fill('image.jpg → image_2.jpg'));
  return { visible: caja.style.display === 'block', texto: caja.innerText };
})));
'''
    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as fh:
        fh.write(guion)
        p = fh.name
    out = subprocess.run(['node', p], capture_output=True, text=True)
    os.unlink(p)
    assert out.returncode == 0, f'node fallo en los avisos:\n{out.stderr[:600]}'
    return json.loads(out.stdout.strip().splitlines()[-1])


def main():
    html = html_servido()
    fns = '\n'.join(extraer_funcion(html, n) for n in FUNCIONES)
    print(f'funciones extraidas del HTML servido ({len(fns)} chars)')

    if SABOTAJE:
        fns = '\n'.join(extraer_funcion(html, n) for n in
                        ('_pareceDocumento', '_claveNombre')) + VIEJO
        print('*** MODO SABOTAJE: vuelve el descarte por nombre ***')

    permutaciones = [list(p) for p in itertools.permutations(TANDA_ORDEN)]
    guion = (PRELUDIO + fns + '\nconst CASOS = ' + json.dumps([c[1] for c in CASOS]) +
             ';\nconst PERMS = ' + json.dumps([[[x] for x in p] for p in permutaciones]) + ''';
console.log(JSON.stringify({
  casos: CASOS.map(correr),
  orden: PERMS.map(function(p) { const r = correr(p);
    return { n: r.en_lista, unicas: r.claves_unicas }; })
}));
''')
    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as fh:
        fh.write(guion)
        p = fh.name
    out = subprocess.run(['node', p], capture_output=True, text=True)
    os.unlink(p)
    if out.returncode != 0:
        print('node fallo:\n' + out.stderr[:900])
        return 1
    got = json.loads(out.stdout.strip().splitlines()[-1])

    fallos = 0
    for (titulo, _t, esp), r in zip(CASOS, got['casos']):
        ok_cuentas = all(r[k] == v for k, v in esp.items())
        ok_cuadre = r['recibidos'] == r['en_lista'] + r['repetidos'] + r['desconocidos']
        ok = ok_cuentas and ok_cuadre and r['claves_unicas'] and r['input_limpio']
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        print(f"          {r['en_lista']} en la lista · {r['renombrados']} renombradas · "
              f"{r['repetidos']} el mismo fichero · {r['desconocidos']} ilegibles "
              f"· de {r['recibidos']} recibidas")
        if not ok:
            fallos += 1
            if not ok_cuentas:
                print(f'          esperaba {esp}')
            if not ok_cuadre:
                print('          NO CUADRA: se ha perdido alguna sin decir nada')
            if not r['claves_unicas']:
                print(f"          NOMBRES QUE CHOCAN al comprimir: {r['nombres']}")
            if not r['input_limpio']:
                print('          el input NO se limpia')

    # ── Que se avisa y que no ────────────────────────────────────────
    for (titulo, args, se_ve, texto), r in zip(AVISOS, avisos_resultado(html)):
        ok = (r['visible'] == se_ve)
        if ok and texto:
            ok = texto in r['texto']
        if ok and not se_ve:
            ok = not r['texto'].strip()
        # renombrar nunca se nombra, ni cuando la caja se enseña por otra cosa
        if ok and 'llamaba igual' in r['texto']:
            ok = False
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        if not ok:
            fallos += 1
            print(f"          visible={r['visible']} (esperaba {se_ve}) · texto: {r['texto'][:90]!r}")

    # El orden no decide
    ns = sorted(set(x['n'] for x in got['orden']))
    todas_unicas = all(x['unicas'] for x in got['orden'])
    ok_orden = ns == [4] and todas_unicas
    print(f"  {'OK ' if ok_orden else 'FALLA'}  EL ORDEN DA IGUAL — "
          f"{len(got['orden'])} permutaciones, documentos que quedan: {ns}")
    if not ok_orden:
        fallos += 1
        print('          alguna permutacion pierde documentos o deja nombres que chocan')

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: el test canta ({fallos} casos en rojo). Protege de verdad.')
            return 0
        print('SABOTAJE MAL: el test no se entera del descarte por nombre.')
        return 1
    if fallos:
        print(f'{fallos} caso(s) en rojo')
        return 1
    print(f'{len(CASOS)} casos + {len(got["orden"])} permutaciones OK. '
          'Ninguna foto se pierde y el orden no decide.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
