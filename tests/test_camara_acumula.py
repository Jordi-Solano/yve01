# -*- coding: utf-8 -*-
"""El bug de la camara: la 2a foto no se añadia (movil).

Dos propiedades, sobre las funciones REALES sacadas del HTML servido
(regla 26), no sobre una copia escrita a mano:

  1. **El input se limpia** despues de cada seleccion. Si no, el navegador
     puede no volver a disparar `change` y la siguiente foto no llega.
  2. **Nada se descarta en silencio**: recibidos = añadidos + repetidos +
     no reconocidos. Siempre.

  python3.12 tests/test_camara_acumula.py
  python3.12 tests/test_camara_acumula.py --sabotaje
"""
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


VIEJO_SILENCIOSO = """
function handleUploadFiles(fileList, input) {
  var files = Array.from(fileList).filter(_pareceDocumento);
  _addFilesToList(files);
}
function _addFilesToList(newFiles) {
  var existing = new Set(_uploadFiles.map(function(f){ return f.name; }));
  newFiles.forEach(function(f) { if (!existing.has(f.name)) { _uploadFiles.push(f); } });
  _renderFileList();
  return { anadidos: 0, repetidos: [] };
}
"""

CASOS = [
    # (titulo, [(nombre, tipo), ...] por tandas, esperado)
    ('CAMARA — dos capturas con nombres de uuid distintos (su movil)',
     [[('6e7dfcd4-951b-44ff-9a68-9b6c0eefd3fa.jpeg', 'image/jpeg')],
      [('1d08bbd4-fa36-4478-9c6c-5ef2e6f35e29.jpeg', 'image/jpeg')]],
     {'en_lista': 2, 'repetidos': 0, 'desconocidos': 0}),

    ('CAMARA — dos capturas que el movil llama IGUAL (iOS: image.jpg)',
     [[('image.jpg', 'image/jpeg')], [('image.jpg', 'image/jpeg')]],
     {'en_lista': 1, 'repetidos': 1, 'desconocidos': 0}),

    ('GALERIA — varias de golpe, nombres distintos',
     [[('IMG_0101.jpg', 'image/jpeg'), ('IMG_0102.jpg', 'image/jpeg'),
       ('IMG_0103.jpg', 'image/jpeg')]],
     {'en_lista': 3, 'repetidos': 0, 'desconocidos': 0}),

    ('Repetidas dentro de la MISMA tanda',
     [[('foto.jpg', 'image/jpeg'), ('foto.jpg', 'image/jpeg')]],
     {'en_lista': 1, 'repetidos': 1, 'desconocidos': 0}),

    ('Sin extension y sin tipo: no entra, pero se DICE',
     [[('captura', '')]],
     {'en_lista': 0, 'repetidos': 0, 'desconocidos': 1}),

    ('Sin extension pero con tipo image/*: entra',
     [[('captura', 'image/jpeg')]],
     {'en_lista': 1, 'repetidos': 0, 'desconocidos': 0}),

    ('Mezcla: una buena, una repetida y una ilegible',
     [[('a.jpg', 'image/jpeg')],
      [('b.jpg', 'image/jpeg'), ('a.jpg', 'image/jpeg'), ('raro', '')]],
     {'en_lista': 2, 'repetidos': 1, 'desconocidos': 1}),
]


def main():
    html = html_servido()
    fns = '\n'.join(extraer_funcion(html, n)
                    for n in ('_pareceDocumento', 'handleUploadFiles', '_addFilesToList'))
    print(f'funciones extraidas del HTML servido ({len(fns)} chars)')

    if SABOTAJE:
        fns = extraer_funcion(html, '_pareceDocumento') + VIEJO_SILENCIOSO
        print('*** MODO SABOTAJE: se vuelve al descarte mudo y sin limpiar el input ***')

    guion = '''
let _uploadFiles = [];
function _renderFileList() {}
let ULTIMO = null;
function _avisoDescartes(recibidos, anadidos, desconocidos, repetidos) {
  ULTIMO = { recibidos, anadidos,
             desconocidos: (desconocidos||[]).length, repetidos: (repetidos||[]).length };
}
''' + fns + '''
const CASOS = ''' + json.dumps([c[1] for c in CASOS]) + ''';
const salida = CASOS.map(function(tandas) {
  _uploadFiles = [];
  let rep = 0, desc = 0, recib = 0, avisos = 0;
  const input = { value: '' };
  let inputLimpio = true;
  tandas.forEach(function(tanda) {
    ULTIMO = null;
    const files = tanda.map(function(x) { return { name: x[0], type: x[1] }; });
    recib += files.length;
    input.value = 'C:\\\\fakepath\\\\' + files[0][0];   // lo que deja el navegador
    handleUploadFiles(files, input);
    if (input.value !== '') inputLimpio = false;
    if (ULTIMO) { avisos++; rep += ULTIMO.repetidos; desc += ULTIMO.desconocidos; }
  });
  return { en_lista: _uploadFiles.length, repetidos: rep, desconocidos: desc,
           recibidos: recib, avisos: avisos, input_limpio: inputLimpio };
});
console.log(JSON.stringify(salida));
'''
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
    for (titulo, _tandas, esp), r in zip(CASOS, got):
        ok_cuentas = all(r[k] == v for k, v in esp.items())
        # La invariante de verdad: nada desaparece sin que se diga.
        ok_cuadre = (r['recibidos'] == r['en_lista'] + r['repetidos'] + r['desconocidos'])
        ok_input = r['input_limpio']
        ok = ok_cuentas and ok_cuadre and ok_input
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        print(f"          en lista {r['en_lista']} · repetidas {r['repetidos']} · "
              f"ilegibles {r['desconocidos']} · de {r['recibidos']} recibidas · "
              f"input limpio: {'si' if ok_input else 'NO'}")
        if not ok:
            fallos += 1
            if not ok_cuentas:
                print(f'          esperaba {esp}')
            if not ok_cuadre:
                print('          NO CUADRA: se ha perdido alguna sin decir nada')
            if not ok_input:
                print('          el input NO se limpia: la camara puede no volver a disparar change')

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: el test canta ({fallos} casos en rojo). Protege de verdad.')
            return 0
        print('SABOTAJE MAL: el test no se entera del descarte mudo.')
        return 1
    if fallos:
        print(f'{fallos} caso(s) en rojo')
        return 1
    print(f'{len(CASOS)} casos OK. El input se limpia y nada se descarta en silencio.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
