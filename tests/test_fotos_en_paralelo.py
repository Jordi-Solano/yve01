# -*- coding: utf-8 -*-
"""PASO 4 — tres fotos a la vez, sin perder ninguna.

Lo que importa NO es que vaya rapido: es que lleguen TODAS. El paralelismo
tiene dos formas tipicas de fallar y las dos se comprueban aqui:

  · que una foto se quede sin procesar, o se procese dos veces;
  · que se suelten todas a la vez y se ahogue el servidor (con el candado
    puesto, 25 peticiones simultaneas serian 25 hilos haciendo cola).

Y la que no se puede mover: **con UNA foto el log sale identico** al de
siempre, sin barra de progreso y sin concurrencia.

Se prueba la funcion REAL sacada del HTML servido (regla 26).

  python3.12 tests/test_fotos_en_paralelo.py
  python3.12 tests/test_fotos_en_paralelo.py --sabotaje
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
    m = re.search(r'(?:async\s+)?function\s+' + re.escape(nombre) + r'\s*\(', js)
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


# Sin limite: se sueltan TODAS a la vez. Es el cambio peligroso que el test
# tiene que cazar — con el candado del servidor, 25 fotos serian 25 hilos
# haciendo cola.
SIN_LIMITE = """
async function _procesarImagenes(imgs, addLine, acc) {
  var lista = imgs || []; var total = lista.length;
  if (!total) return 0;
  var errs = 0;
  await Promise.all(lista.map(async function(f, i) {
    errs += await _unaFoto(f, i, total, addLine, acc);
  }));
  return errs;
}"""

CASOS = [
    # (titulo, n_fotos, fallan)
    ('1 foto: sin concurrencia ni barra (no se puede mover)', 1, []),
    ('2 fotos', 2, []),
    ('6 fotos', 6, []),
    ('12 fotos, que es lo que sube Jordi', 12, []),
    ('25 fotos', 25, []),
    ('12 fotos, dos dan error del servidor', 12, [3, 7]),
]


def main():
    html = html_servido()
    partes = [extraer_funcion(html, n) for n in ('_progresoFotos', '_unaFoto', '_procesarImagenes')]
    fn = '\n'.join(partes)
    fn = 'var _FOTOS_A_LA_VEZ = ' + re.search(
        r'var _FOTOS_A_LA_VEZ = (\d+);', html).group(1) + ';\n' + fn
    if SABOTAJE:
        fn = fn[:fn.index('async function _procesarImagenes(')] + SIN_LIMITE.lstrip('\n')
        print('*** MODO SABOTAJE: sin limite de fotos a la vez ***')
    print(f'funciones extraidas del HTML servido ({len(fn)} chars)')

    guion = '''
let vivas = 0, pico = 0;
const vistas = [];          // que fotos se han llegado a procesar
const lineas = [];
let FALLAN = [];
const _csrfToken = 'x';
const document = { getElementById: () => null, createElement: () => ({ style:{}, remove(){} }) };
function _tSSE(t) { return t; }
function _mb(n) { return n > 950000 ? (n/1048576).toFixed(1) + 'MB' : Math.round(n/1024) + 'KB'; }
async function _comprimirImagen(f) { return f; }
function addLine(t) { lineas.push(String(t)); }
class FormData {
  constructor() { this.fichero = null; }
  append(k, v) { if (k === 'image') this.fichero = v; }
}
async function fetch(u, o) {
  // el indice sale del NOMBRE del fichero que va en el FormData, que es por
  // donde viaja de verdad. Antes se metia en las opciones desde un envoltorio
  // global de `fetch`, y con tres llamadas a la vez se pisaban entre ellas:
  // el fallo era del test, no del codigo.
  const nombre = (o && o.body && o.body.fichero && o.body.fichero.name) || '';
  const idx = parseInt(String(nombre).replace(/\D/g, ''), 10);
  vivas++; if (vivas > pico) pico = vivas;
  await new Promise(r => setTimeout(r, 12 + (idx % 5) * 4));
  vivas--;
  return { json: async () => (FALLAN.indexOf(idx) !== -1
      ? { ok:false, error:'error del servidor' }
      : { ok:true, tipo:'FACTURA', mensaje:'Proveedor ' + idx + ' — €100', cierre:['ap'] }) };
}
''' + fn + '''
// se marca cada foto que entra, para saber si alguna se queda fuera o se repite
const _unaFotoReal = _unaFoto;
_unaFoto = async function(original, fi, total, addLine, acc) {
  vistas.push(fi);
  return await _unaFotoReal(original, fi, total, addLine, acc);
};
const CASOS = ''' + json.dumps([[c[1], c[2]] for c in CASOS]) + ''';
const salida = [];
for (const [n, fallan] of CASOS) {
  vivas = 0; pico = 0; vistas.length = 0; lineas.length = 0; FALLAN = fallan;
  const acc = {};
  const fotos = Array.from({length:n}, (_,k) => ({name:'f'+k+'.jpg', size:1000+k}));
  const errs = await _procesarImagenes(fotos, addLine, acc);
  salida.push({
    n, errs, pico,
    procesadas: vistas.length,
    unicas: new Set(vistas).size,
    lineas: lineas.slice(),
    cierre: Object.keys(acc)
  });
}
console.log(JSON.stringify(salida));
'''
    with tempfile.NamedTemporaryFile('w', suffix='.mjs', delete=False, encoding='utf-8') as fh:
        fh.write(guion)
        p = fh.name
    out = subprocess.run(['node', p], capture_output=True, text=True)
    os.unlink(p)
    if out.returncode != 0:
        print('node fallo:\n' + out.stderr[:900])
        return 1
    got = json.loads(out.stdout.strip().splitlines()[-1])

    fallos = 0
    for (titulo, n, fallan), r in zip(CASOS, got):
        todas = (r['procesadas'] == n and r['unicas'] == n)
        tope = (r['pico'] <= 3)
        usa_paralelo = (r['pico'] == min(3, n))
        errs_ok = (r['errs'] == len(fallan))
        cierre_ok = (r['cierre'] == ['ap'] if n > len(fallan) else True)
        ok = todas and tope and usa_paralelo and errs_ok and cierre_ok
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        print(f"          {r['procesadas']}/{n} procesadas ({r['unicas']} distintas) · "
              f"pico de {r['pico']} a la vez · {r['errs']} error(es)")
        if not ok:
            fallos += 1
            if not todas:
                print('          SE HA PERDIDO O REPETIDO ALGUNA')
            if not tope:
                print(f"          SE SUELTAN {r['pico']} A LA VEZ — el servidor se ahoga")
            if not usa_paralelo:
                print(f"          no esta usando el paralelo (pico {r['pico']}, esperaba {min(3, n)})")
            if not errs_ok:
                print(f"          errores {r['errs']}, esperaba {len(fallan)}")

    # Con UNA foto, el log tiene que salir EXACTAMENTE como siempre
    l1 = got[0]['lineas']
    ok_una = (len(l1) == 2
              and l1[0] == '🔍 [1/1] f0.jpg (1KB)...'
              and l1[1] == '✓ f0.jpg: FACTURA — Proveedor 0 — €100'
              and got[0]['pico'] == 1)
    print(f"  {'OK ' if ok_una else 'FALLA'}  con 1 foto el log sale IDENTICO (2 lineas, sin progreso)")
    if not ok_una:
        fallos += 1
        print(f'          recibi: {l1}')

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: sin limite, {fallos} en rojo.')
            return 0
        print('SABOTAJE MAL: el test no se entera de que se sueltan todas a la vez.')
        return 1
    if fallos:
        print(f'{fallos} en rojo')
        return 1
    print(f'{len(CASOS)} casos OK. Ninguna foto se queda fuera y nunca hay mas de 3 a la vez.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
