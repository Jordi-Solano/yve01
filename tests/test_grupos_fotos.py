# -*- coding: utf-8 -*-
"""PASO 1 de M1 — el reparto de fotos en documentos.

Prueba la funcion REAL, sacada del HTML que recibe el navegador (regla 26),
no una copia escrita a mano aqui: una copia se queda vieja y da verde sobre
codigo que ya no existe.

La propiedad que protege: **el numero de fotos NO decide nada**. Lo unico que
manda es lo que el usuario haya unido a mano en el modal.

  python3.12 tests/test_grupos_fotos.py
  python3.12 tests/test_grupos_fotos.py --sabotaje
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
    """El texto de `function <nombre>(...) {...}`, casando llaves."""
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


CASOS = [
    # (titulo, [(nombre, grupo)], grupos_esperados, sueltas_esperadas)
    ('CASO 1 base — 1 foto suelta',
     [('img_a.jpg', 0)], [], ['img_a.jpg']),

    ('CASO 2 — 2 facturas distintas, sin agrupar (LO QUE CAMBIA)',
     [('img_b.jpg', 0), ('img_c.jpg', 0)], [], ['img_b.jpg', 'img_c.jpg']),

    ('CASO 3 base — 2 fotos unidas = un contrato',
     [('pag_1.jpg', 1), ('pag_2.jpg', 1)], [['pag_1.jpg', 'pag_2.jpg']], []),

    ('MIXTO — contrato A + factura suelta + contrato B',
     [('a1.jpg', 1), ('a2.jpg', 1), ('fact.jpg', 0), ('b1.jpg', 2), ('b2.jpg', 2)],
     [['a1.jpg', 'a2.jpg'], ['b1.jpg', 'b2.jpg']], ['fact.jpg']),

    ('Un grupo que se quedo con UNA foto vuelve a ser suelta',
     [('sola.jpg', 7), ('otra.jpg', 0)], [], ['sola.jpg', 'otra.jpg']),

    ('25 paginas unidas = UN documento, no 25 llamadas',
     [(f'p{i}.jpg', 3) for i in range(25)],
     [[f'p{i}.jpg' for i in range(25)]], []),

    ('25 fotos sueltas = 25 documentos, ni un contrato',
     [(f'q{i}.jpg', 0) for i in range(25)],
     [], [f'q{i}.jpg' for i in range(25)]),

    ('El orden de las sueltas es el de la lista',
     [('z.jpg', 0), ('g1.jpg', 4), ('a.jpg', 0), ('g2.jpg', 4), ('m.jpg', 0)],
     [['g1.jpg', 'g2.jpg']], ['z.jpg', 'a.jpg', 'm.jpg']),
]


def main():
    html = html_servido()
    fn = extraer_funcion(html, '_repartirFotos')
    print(f'_repartirFotos extraida del HTML servido ({len(fn)} chars)')

    if SABOTAJE:
        # Se vuelve a la regla vieja: contar fotos. Los casos 2 y 7 tienen que
        # ponerse en ROJO; si no, el test no protege de nada.
        fn = ('function _repartirFotos(imgs) {\n'
              '  return imgs.length >= 2 ? {grupos:[imgs], sueltas:[]} : {grupos:[], sueltas:imgs};\n'
              '}')
        print('*** MODO SABOTAJE: decidiendo por el numero de fotos ***')

    guion = fn + '\nconst CASOS = ' + json.dumps(
        [[[n, g] for n, g in c[1]] for c in CASOS]) + ''';
const salida = CASOS.map(function(caso) {
  const imgs = caso.map(function(x) { return { name: x[0], _grp: x[1] }; });
  const r = _repartirFotos(imgs);
  return {
    grupos: r.grupos.map(function(g) { return g.map(function(f) { return f.name; }); }),
    sueltas: r.sueltas.map(function(f) { return f.name; })
  };
});
console.log(JSON.stringify(salida));
'''
    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as fh:
        fh.write(guion)
        p = fh.name
    out = subprocess.run(['node', p], capture_output=True, text=True)
    os.unlink(p)
    assert out.returncode == 0, f'node fallo:\n{out.stderr}'
    got = json.loads(out.stdout.strip().splitlines()[-1])

    fallos = 0
    for (titulo, _entrada, gr_esp, su_esp), r in zip(CASOS, got):
        ok = (r['grupos'] == gr_esp and r['sueltas'] == su_esp)
        # lo que de verdad importa: cuantas llamadas a la IA se hacen
        llamadas_contrato = len(r['grupos'])
        llamadas_sueltas = len(r['sueltas'])
        print(f"  {'OK ' if ok else 'FALLA'}  {titulo}")
        print(f"          -> {llamadas_contrato} llamada(s) al lector de contratos"
              f" + {llamadas_sueltas} foto(s) sueltas")
        if not ok:
            fallos += 1
            print(f'          esperaba grupos={gr_esp} sueltas={su_esp}')
            print(f'          recibi   grupos={r["grupos"]} sueltas={r["sueltas"]}')

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: el test canta ({fallos} casos en rojo). Protege de verdad.')
            return 0
        print('SABOTAJE MAL: el test NO se entera de que se vuelve a contar fotos.')
        return 1
    if fallos:
        print(f'{fallos} caso(s) en rojo')
        return 1
    print(f'{len(CASOS)} casos OK. El numero de fotos no decide nada.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
