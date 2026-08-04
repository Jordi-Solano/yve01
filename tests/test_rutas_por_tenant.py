# -*- coding: utf-8 -*-
"""Los tres scripts que clavaban las rutas de la RAÍZ, al árbol del tenant.

`detector_doble_imposicion.py`, `lector_ota.py` y `lector_facturas_ap.py`
calculaban sus rutas con `os.path.join(BASE_DIR, ...)`, sin pasar por
`tenant_dirs`. Con un solo cliente no se notaba —para el tenant `default`,
`tenant_dirs` devuelve exactamente esas mismas rutas—, pero el día que hubiera
un segundo, sus facturas OTA, sus facturas AP, su `proveedores.xlsx` y sus
informes de doble imposición se escribían y se leían del árbol del primero.

Lo que se comprueba, en este orden de importancia:

  1. **Para `default` no cambia NADA.** Es la condición que puso Jordi: esto
     prepara el multi-tenant, no cambia su instalación.
  2. Con otro tenant, todo va a `tenants/<id>/...`.
  3. El clasificador NO congela sus rutas al importar. `dashboard` lo importa
     EN PROCESO y un proceso sirve a varios tenants: si la ruta se resolviera
     una sola vez, el segundo cliente leería el `proveedores.xlsx` del primero.
  4. Los envoltorios funcionan de verdad donde se usan (`os.path.join`,
     `os.path.exists`, `glob`, `pandas`), no solo al imprimirlos.
  5. La lógica de clasificación no se ha tocado.

  python3.12 tests/test_rutas_por_tenant.py
  python3.12 tests/test_rutas_por_tenant.py --sabotaje
"""
import glob as _glob
import json
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv

MODULOS = {
    'detector_doble_imposicion': ['REPORTES_DIR', 'ENTRADA_DIR', 'REPORTE_DI_SALIDA'],
    'lector_ota': ['ENTRADA_DIR', 'SALIDA_DIR', 'SALIDA_EXCEL'],
    'lector_facturas_ap': ['ENTRADA_DIR', 'SALIDA_DIR', 'REFERENCIA_DIR',
                           'SALIDA_EXCEL', 'PROV_FILE'],
}

_CAP = '''
import os, sys, json
sys.path.insert(0, %(base)r); os.chdir(%(base)r)
os.environ["YVE_TENANT"] = %(tenant)r
out = {}
for mod, campos in %(mods)r.items():
    m = __import__(mod)
    out[mod] = {c: str(getattr(m, c)) for c in campos}
print("###J###" + json.dumps(out, sort_keys=True))
'''


def constantes(tenant):
    cod = _CAP % {'base': BASE, 'tenant': tenant, 'mods': MODULOS}
    r = subprocess.run([sys.executable, '-c', cod], capture_output=True, text=True, cwd=BASE)
    if '###J###' not in r.stdout:
        return {'ERROR': (r.stderr or r.stdout)[-300:]}
    return json.loads(r.stdout.split('###J###')[1])


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: se comparan las rutas del tenant contra la RAÍZ, '
              'que es como estaban antes ***')
    fallos = 0

    # ── 1 · para `default`, EXACTAMENTE lo de siempre ────────────────
    raiz = {
        'detector_doble_imposicion': {
            'REPORTES_DIR': os.path.join(BASE, 'reportes'),
            'ENTRADA_DIR': os.path.join(BASE, 'facturas-entrada')},
        'lector_ota': {
            'ENTRADA_DIR': os.path.join(BASE, 'facturas-entrada'),
            'SALIDA_DIR': os.path.join(BASE, 'facturas-procesadas')},
        'lector_facturas_ap': {
            'ENTRADA_DIR': os.path.join(BASE, 'facturas-entrada'),
            'SALIDA_DIR': os.path.join(BASE, 'facturas-procesadas'),
            'REFERENCIA_DIR': os.path.join(BASE, 'datos-referencia'),
            'PROV_FILE': os.path.join(BASE, 'datos-referencia', 'proveedores.xlsx')},
    }
    dflt = constantes('default')
    malas = []
    for mod, esperado in raiz.items():
        for c, v in esperado.items():
            got = (dflt.get(mod) or {}).get(c)
            if got != v:
                malas.append(f'{mod}.{c}: {got} != {v}')
    ok1 = not malas
    print(f"  {'OK ' if ok1 else 'FALLA'}  con `default`, las {sum(len(v) for v in raiz.values())} "
          f"rutas son EXACTAMENTE las de siempre" + ('' if ok1 else f' — {malas[:3]}'))
    if not ok1:
        fallos += 1

    # ── 2 · con otro tenant, a su árbol ──────────────────────────────
    otro = constantes('acme')
    try:
        esperada = os.path.join(BASE, 'tenants', 'acme')
        todas = [v for cs in otro.values() for v in cs.values()]
        if SABOTAJE:
            # como estaban antes: clavadas a la raíz, ignorando el tenant.
            # Se comprueba que CONTIENEN el árbol del cliente, no que empiecen
            # por la raíz — `<repo>/tenants/acme/...` también empieza por la
            # raíz, así que ese sabotaje no mordía.
            todas = [v.replace(os.path.join('tenants', 'acme') + os.sep, '') for v in todas]
        ok2 = bool(todas) and all(v.startswith(esperada) for v in todas)
        print(f"  {'OK ' if ok2 else 'FALLA'}  con otro tenant, las {len(todas)} rutas van a "
              f"tenants/acme/…" + ('' if ok2 else f' — {[v for v in todas if not v.startswith(esperada)][:2]}'))
        if not ok2:
            fallos += 1

        # ── 3 · el clasificador NO congela la ruta al importar ───────
        cod = f'''
import os, sys, json
sys.path.insert(0, {BASE!r}); os.chdir({BASE!r})
os.environ["YVE_TENANT"] = "default"
import lector_facturas_ap as L
antes = str(L.PROV_FILE)
os.environ["YVE_TENANT"] = "acme"          # el proceso sigue vivo, cambia el cliente
despues = str(L.PROV_FILE)
print("###J###" + json.dumps({{"antes": antes, "despues": despues}}))
'''
        r = subprocess.run([sys.executable, '-c', cod], capture_output=True, text=True, cwd=BASE)
        d = json.loads(r.stdout.split('###J###')[1]) if '###J###' in r.stdout else {}
        if SABOTAJE:
            d = {'antes': d.get('antes', ''), 'despues': d.get('antes', '')}
        ok3 = (d.get('antes', '').endswith('datos-referencia/proveedores.xlsx')
               and 'tenants/acme' in d.get('despues', ''))
        print(f"  {'OK ' if ok3 else 'FALLA'}  el clasificador NO congela la ruta al importar: "
              f"…{d.get('antes','')[-34:]} → …{d.get('despues','')[-42:]}")
        if not ok3:
            fallos += 1

        # ── 4 · los envoltorios funcionan donde se usan de verdad ────
        import lector_facturas_ap as L
        pruebas = {
            'os.path.join': os.path.join(L.SALIDA_DIR, 'x.xlsx').endswith('facturas-procesadas/x.xlsx'),
            'os.path.exists': isinstance(os.path.exists(L.PROV_FILE), bool),
            'os.fspath': os.fspath(L.SALIDA_DIR) == os.path.join(BASE, 'facturas-procesadas'),
            'glob': isinstance(_glob.glob(os.path.join(L.ENTRADA_DIR, '*.pdf')), list),
            'str()': str(L.SALIDA_EXCEL).endswith('.xlsx'),
            'f-string': f'{L.REFERENCIA_DIR}'.endswith('datos-referencia'),
        }
        if SABOTAJE:
            pruebas = {k: False for k in pruebas}
        ok4 = all(pruebas.values())
        print(f"  {'OK ' if ok4 else 'FALLA'}  los envoltorios funcionan donde se usan: "
              f"{', '.join(k for k, v in pruebas.items() if v) or 'ninguno'}")
        if not ok4:
            fallos += 1

        # y `os.makedirs(SALIDA_DIR)`, que corre al importar
        ok4b = os.path.isdir(os.fspath(L.SALIDA_DIR))
        print(f"  {'OK ' if ok4b else 'FALLA'}  os.makedirs() sobre el envoltorio ha creado "
              f"la carpeta")
        if not ok4b:
            fallos += 1
    finally:
        shutil.rmtree(os.path.join(BASE, 'tenants', 'acme'), ignore_errors=True)

    # ── 5 · la lógica de clasificación, sin tocar ────────────────────
    r = subprocess.run(['git', 'diff', 'HEAD', '--', 'lector_facturas_ap.py'],
                       capture_output=True, text=True, cwd=BASE)
    tocadas = [l for l in r.stdout.split('\n')
               if (l.startswith('+') or l.startswith('-'))
               and not l.startswith('+++') and not l.startswith('---')]
    # lo único que puede aparecer son rutas y sus comentarios; jamás el prompt
    sospechosas = [l for l in tocadas
                   if 'PROMPT' in l or 'CLASIFICA' in l.upper() or 'prompt_foto' in l]
    ok5 = not sospechosas if not SABOTAJE else False
    print(f"  {'OK ' if ok5 else 'FALLA'}  el diff del clasificador no roza el prompt "
          f"({len(tocadas)} líneas cambiadas, {len(sospechosas)} sospechosas)")
    if not ok5:
        fallos += 1

    # ── 6 · ninguna llamada a subproceso sin el entorno del tenant ───
    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read().split('\n')
    sin_env = []
    for i, l in enumerate(src):
        if '.run([' in l and ('python3' in l or 'sys.executable' in l):
            blob = ' '.join(x.strip() for x in src[i:i + 4])
            if 'env=_env_tenant()' not in blob:
                sin_env.append(i + 1)
    ok6 = not sin_env if not SABOTAJE else False
    print(f"  {'OK ' if ok6 else 'FALLA'}  ninguna llamada a subproceso se queda sin el "
          f"tenant" + ('' if ok6 else f' — líneas {sin_env}'))
    if not ok6:
        fallos += 1

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: {fallos} en rojo.')
            return 0
        print('SABOTAJE MAL.')
        return 1
    if fallos:
        print(f'{fallos} en rojo')
        return 1
    print('Todo OK. Para `default` no cambia nada; el multi-tenant queda preparado.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
