# -*- coding: utf-8 -*-
"""BOMBA 3 · la factura de la luz iba a la cuenta de las comisiones OTA.

`_auto_cuenta_pgc` devolvia 628 para energia/suministros, y en
`datos-referencia/plan_cuentas.xlsx` la 628 es "Comisiones de agencias y
OTAs": Endesa y Booking en la misma cuenta. En este plan la energia va a la
629 ("Otros servicios (telecom, energia)"), que es tambien lo que dice
`asignador_cuentas`.

Matriz: la energia CAMBIA (628 -> 629) y todo lo demas sale IDENTICO. Y la
cuenta que devuelve el lector tiene que existir en el plan de cuentas.

  python3.12 tests/test_cuenta_suministros.py
  python3.12 tests/test_cuenta_suministros.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                  # noqa: E402
import lector_facturas_ap as L                       # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv

# (concepto, proveedor) -> cuenta esperada. Los IDENTICOS son la linea base
# tomada con el codigo viejo; solo las dos primeras filas cambian a proposito.
CASOS = [
    ('Suministro electrico', 'Endesa Energia SA', '629'),      # CAMBIA: era 628
    ('Consumo de agua ', 'Aigues de Barcelona', '629'),         # CAMBIA: era 628
    ('Comision reservas', 'Booking.com', '628'),                # identico
    ('Commission', 'Expedia', '628'),                           # identico
    ('Alimentacion cocina', 'Makro', '600'),
    ('Servicio de limpieza', 'Limpiezas Sur', '629'),
    ('Mantenimiento ascensor', 'Otis', '622'),
    ('Poliza seguro', 'Mapfre', '625'),
    ('Alquiler local', 'Inmo SL', '621'),
    ('Asesoria contable', 'Gestoria Nord', '623'),
    ('Campaña marketing', 'Agencia', '627'),
    ('Fibra internet', 'Movistar', '629'),
    ('Transporte mercancia', 'Seur', '624'),
    ('Decoracion evento', 'Flores', '629'),
    ('Cosa rara', 'Nadie', '629'),
]


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        orig = L._auto_cuenta_pgc

        def sab(concepto, proveedor=None):
            r = orig(concepto, proveedor)
            return '628' if 'suministro' in (concepto or '').lower() else r
        L._auto_cuenta_pgc = sab

    plan = pd.read_excel(os.path.join(BASE, 'datos-referencia', 'plan_cuentas.xlsx'))
    codigos = {str(c) for c in plan['codigo_cuenta'].astype(str)}
    desc = dict(zip(plan['codigo_cuenta'].astype(str), plan['descripcion'].astype(str)))

    for concepto, prov, esperada in CASOS:
        got = L._auto_cuenta_pgc(concepto, prov)
        ok(got == esperada, f"{concepto!r} / {prov}: {got} (esperada {esperada})")
        if concepto.startswith(('Suministro', 'Consumo')):      # las dos que cambian
            ok(got in codigos, f"    la {got} existe en plan_cuentas.xlsx")
        elif got not in codigos:
            # APUNTADO, no arreglado aqui: 627 (publicidad) y 624 (transportes)
            # salen del lector y NO estan en plan_cuentas.xlsx. Preexistente.
            print(f"  nota  la {got} de {concepto!r} no esta en plan_cuentas.xlsx (preexistente)")

    luz = L._auto_cuenta_pgc('Suministro electrico', 'Endesa Energia SA')
    ok('omision' not in desc.get(luz, '').lower() and 'comision' not in desc.get(luz, '').lower(),
       f"la luz NO va a una cuenta de comisiones: {luz} = {desc.get(luz)}")
    ok(L._auto_cuenta_pgc('Suministro electrico', 'Endesa') !=
       L._auto_cuenta_pgc('Comision reservas', 'Booking.com'),
       'Endesa y Booking ya no comparten cuenta')

    # El clasificador es intocable salvo con OK explicito (BONO, 3 sep). Lo que
    # se exige: los tipos y esquemas de la linea base siguen identicos.
    import json as _json, re as _re
    _base = _json.load(open(os.path.join(BASE, 'tests', 'baseline_tipos_clasificador.json'), encoding='utf-8'))
    _lineas = L.PROMPT_CLASIFICACION.split('\n')
    _tipos = set(_re.findall(r'→ ([A-Z_]+)', L.PROMPT_CLASIFICACION))
    ok(set(_base['tipos']) <= _tipos and all(e in _lineas for es in _base['esquemas'].values() for e in es),
       'PROMPT_CLASIFICACION: tipos y esquemas de la linea base intactos')
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True,
                          text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_')], 'ningun oracle_* tocado')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else
              '*** SABOTAJE SIN EFECTO: la prueba no protege nada ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
