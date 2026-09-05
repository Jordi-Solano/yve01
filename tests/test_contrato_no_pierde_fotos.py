# -*- coding: utf-8 -*-
"""PASO 2 — cuando el contrato de grupo falla, las fotos NO se tiran.

`_procesarGrupoFotos` decidia por el TEXTO del error, asi que solo rescataba
las fotos en uno de los cuatro caminos de fallo. El servidor ya lo dice con
una MARCA (`reprocesar`) y la pone en dos casos; los otros dos no la traen
pero las fotos siguen siendo documentos.

Dos propiedades:
  · **ningun camino de fallo se queda sin reprocesar**;
  · cada uno de los dos casos que el servidor marca dice **lo suyo** — antes
    los dos enseñaban "no son un contrato", incluso cuando el sistema si creia
    que lo era y lo que fallo fue leerlo.
Y la que no se puede mover: el contrato bueno, con sus cuatro lineas de dinero.

Se prueba la funcion REAL sacada del HTML servido (regla 26).

  python3.12 tests/test_contrato_no_pierde_fotos.py
  python3.12 tests/test_contrato_no_pierde_fotos.py --sabotaje
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


# Las CUATRO respuestas reales de `procesar_contrato_grupo` + la caida de red.
# (titulo, respuesta_del_servidor | 'RED', reprocesa_esperado)
CASOS = [
    ('El contrato bueno: NO se reprocesa nada (no se puede mover)',
     {'ok': True, 'contrato': 'CG-2026-0417', 'cliente': 'Viajes Meridiano S.A.',
      'total_receivable': 23100, 'beo_lineas': 4,
      'distribucion': {'ap': 1735.53, 'banco': 6930, 'fb': 7200}}, False),

    ('"las fotos no parecen un contrato de grupo" -> reprocesa (ya iba)',
     {'ok': False, 'needs_review': True, 'reprocesar': True,
      'error': 'las fotos no parecen un contrato de grupo',
      'message': 'Las imágenes no parecen un contrato de grupo/BEO.'}, True),

    ('"no se ha podido leer ni importe, ni habitaciones..." -> reprocesa (ERA EL BUG)',
     {'ok': False, 'needs_review': True, 'reprocesar': True,
      'error': 'no se ha podido leer ni importe, ni habitaciones, ni el nombre del contrato',
      'message': 'Parece un contrato de grupo, pero no se ha extraído ningún dato '
                 'aprovechable — revisar manualmente.'}, True),

    ('La IA fallo, sin marca -> reprocesa igual',
     {'ok': False, 'needs_review': True,
      'error': 'sin ANTHROPIC_API_KEY (extracción no disponible en este entorno)'}, True),

    ('Se cayo la red navegador<->Render -> reprocesa igual',
     'RED', True),

    ('Respuesta rara del servidor (500 con texto) -> reprocesa igual',
     {'ok': False, 'error': 'Internal Server Error'}, True),
]

VIEJO = """
    } else if (_dc && /no parecen un contrato/i.test(_dc.error || '')) {
      addLine('Las fotos no son un contrato — proceso cada una como documento suelto', 'l-info');
      errs += await _procesarImagenes(grupo, addLine, cierre);
    } else {
      addLine('no se pudo leer', 'l-warn');
      errs++;
    }
  } catch(e) {
    addLine('error de red', 'l-err');
    errs++;
  }
  return errs;
}"""


def main():
    html = html_servido()
    fn = extraer_funcion(html, '_procesarGrupoFotos')
    if SABOTAJE:
        corte = fn.index("    } else if (")
        fn = fn[:corte] + VIEJO.lstrip('\n')
        print('*** MODO SABOTAJE: vuelve el descarte por el texto del error ***')
    print(f'_procesarGrupoFotos extraida del HTML servido ({len(fn)} chars)')

    guion = '''
let RESPUESTA = null;
let reprocesadas = 0;
let lineas = [];
const _csrfToken = 'x';
class FormData { append() {} }
async function _comprimirImagen(f) { return f; }
async function _procesarImagenes(grupo, addLine, cierre) { reprocesadas += grupo.length; return 0; }
function addLine(t) { lineas.push(String(t)); }
// desde b24 el resumen del contrato formatea con el formateador unico del panel
function _fmtEurES(v, dec) { return Number(v).toFixed(dec === undefined ? 2 : dec).replace('.', ',') + ' €'; }
async function fetch() {
  if (RESPUESTA === 'RED') throw new Error('Failed to fetch');
  return { json: async () => RESPUESTA };
}
''' + fn + '''
const CASOS = ''' + json.dumps([c[1] for c in CASOS]) + ''';
const salida = [];
for (const r of CASOS) {
  RESPUESTA = r; reprocesadas = 0; lineas = [];
  const grupo = [{name:'a.jpg'}, {name:'b.jpg'}];
  const errs = await _procesarGrupoFotos(grupo, '', addLine, {});
  salida.push({ reprocesadas, errs, lineas });
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
    for (titulo, _resp, esperado), r in zip(CASOS, got):
        reproceso = r['reprocesadas'] == 2
        ok = (reproceso == esperado)
        # nada puede perderse en silencio: si no reprocesa, tiene que ser el exito
        if ok and not esperado:
            ok = any('Contrato CG-2026-0417' in l for l in r['lineas'])
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        print(f"          fotos reprocesadas: {r['reprocesadas']}/2 · "
              f"lineas: {len(r['lineas'])}")
        if not ok:
            fallos += 1
            print(f"          esperaba reproceso={esperado}; lineas={r['lineas']}")

    # Los dos casos de `reprocesar` tienen que decir COSAS DISTINTAS: el
    # servidor manda un `message` para cada uno y antes se enseñaba siempre el
    # mismo texto ("no son un contrato") aunque el sistema si creyera que lo era.
    l_no_contrato = ' | '.join(got[1]['lineas'])
    l_sin_datos = ' | '.join(got[2]['lineas'])
    ok_a = ('no parecen un contrato de grupo/BEO' in l_no_contrato
            and 'proceso cada una como documento suelto' in l_no_contrato)
    ok_b = ('no se ha extraído ningún dato aprovechable' in l_sin_datos
            and 'proceso cada una como documento suelto' in l_sin_datos
            and 'no parecen un contrato' not in l_sin_datos
            and 'revisar manualmente' not in l_sin_datos)
    ok_dist = l_no_contrato != l_sin_datos
    for etiqueta, ok in (('«no parecen un contrato» dice lo suyo', ok_a),
                         ('«no se ha leido ningun dato» dice lo SUYO, no lo otro', ok_b),
                         ('los dos mensajes son DISTINTOS', ok_dist)):
        print(f"  {'OK ' if ok else 'FALLA'}  {etiqueta}")
        if not ok:
            fallos += 1
    if not (ok_a and ok_b):
        print(f'          A: {l_no_contrato}')
        print(f'          B: {l_sin_datos}')

    # Y el camino del contrato bueno, con sus cuatro lineas de dinero
    l0 = ' | '.join(got[0]['lineas'])
    ok_ok = ('Contrato CG-2026-0417' in l0 and 'AP comisión agencia' in l0
             and 'Banco depósito previsto' in l0 and 'F&B evento' in l0)
    print(f"  {'OK ' if ok_ok else 'FALLA'}  El contrato bueno sigue sacando sus cuatro lineas")
    if not ok_ok:
        fallos += 1
        print(f'          recibi: {l0}')

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: el test canta ({fallos} en rojo). Protege de verdad.')
            return 0
        print('SABOTAJE MAL: el test no se entera de que las fotos se tiran.')
        return 1
    if fallos:
        print(f'{fallos} caso(s) en rojo')
        return 1
    print(f'{len(CASOS)} casos OK. Ningun camino de fallo tira las fotos.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
