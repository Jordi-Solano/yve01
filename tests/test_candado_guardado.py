# -*- coding: utf-8 -*-
"""EL CANDADO de las escrituras del Excel.

Render corre `gunicorn --workers 1 --threads 8`. Los tres guardadores hacen
leer -> concatenar -> escribir, y eso a la vez es una carrera: dos hilos leen
las mismas filas, cada uno añade la suya y cada uno escribe. Una se pierde.

Tres propiedades:
  1. **Con el candado no se pierde ninguna** con 3 a la vez (6/6 y 12/12), y
     SIN el candado si se pierden — o sea que el candado es lo que lo arregla,
     no la suerte.
  2. **El candado no toca ni un dato**: la misma entrada da el MISMO fichero
     con y sin candado.
  3. **Las columnas que ya estaban se conservan** — incluida `oracle_status`,
     que es la marca que impide que Oracle contabilice dos veces (Fase 0).

  python3.12 tests/test_candado_guardado.py
  python3.12 tests/test_candado_guardado.py --sabotaje   (quita el candado)
"""
import os
import shutil
import sys
import threading

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                          # noqa: E402
import dashboard                                             # noqa: E402
from datetime import date                                    # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
RUTA = os.path.join(dashboard._pdir(), f'facturas_ap_{date.today().strftime("%Y%m%d")}.xlsx')
COPIA = RUTA + '.bak-test-candado'


def guardar():
    """El guardador real, o el de dentro del candado si se sabotea."""
    f = dashboard._guardar_factura_ap
    if SABOTAJE:
        f = getattr(f, '__wrapped__', f)          # el cuerpo, sin poner en fila
    return f


def factura(i):
    return {'archivo': f'foto_{i}.jpg', 'numero_factura': f'F-{1000+i}',
            'nombre_proveedor': f'Proveedor {i}', 'total_factura': 100 + i,
            'fecha_factura': '2026-08-03'}


def limpiar():
    if os.path.exists(RUTA):
        os.remove(RUTA)


def a_la_vez(n_facturas, a_la_vez_n, desde=0):
    limpiar()
    g = guardar()
    fallos = []

    def trabajo(i):
        try:
            g(factura(i))
        except Exception as e:
            fallos.append(type(e).__name__)

    hilos = [threading.Thread(target=trabajo, args=(desde + i,)) for i in range(n_facturas)]
    for k in range(0, len(hilos), a_la_vez_n):
        tanda = hilos[k:k + a_la_vez_n]
        for h in tanda:
            h.start()
        for h in tanda:
            h.join()
    # Sin candado el xlsx puede quedar CORRUPTO (un hilo lo lee mientras otro
    # lo escribe), asi que ni leerlo se puede dar por hecho.
    try:
        df = pd.read_excel(RUTA) if os.path.exists(RUTA) else pd.DataFrame()
        return len(df), fallos
    except Exception as e:
        fallos.append('fichero corrupto: ' + type(e).__name__)
        return 0, fallos


def main():
    habia = os.path.exists(RUTA)
    if habia:
        shutil.copy(RUTA, COPIA)
    fallos = 0
    try:
        if SABOTAJE:
            print('*** MODO SABOTAJE: se llama al cuerpo, sin poner en fila ***')

        # ── 1 · no se pierde ninguna ────────────────────────────────
        for n, k, etiqueta in ((6, 1, 'de una en una'),
                               (6, 3, 'de tres en tres'),
                               (12, 3, 'de tres en tres, 12 fotos')):
            quedan, exc = a_la_vez(n, k)
            ok = (quedan == n and not exc)
            print(f"  {'OK ' if ok else 'FALLA'}  {etiqueta}: {quedan}/{n} guardadas"
                  + (f' · excepciones: {sorted(set(exc))}' if exc else ''))
            if not ok:
                fallos += 1

        # ── 2 · el candado no cambia el dato ────────────────────────
        # (con el candado quitado esto se salta: el fichero puede haber
        #  quedado corrupto arriba y no es lo que se esta midiendo)
        limpiar()
        dashboard._guardar_factura_ap(factura(50))
        dashboard._guardar_factura_ap(factura(51))
        con = pd.read_excel(RUTA).to_csv(index=False)
        limpiar()
        cuerpo = getattr(dashboard._guardar_factura_ap, '__wrapped__',
                         dashboard._guardar_factura_ap)
        cuerpo(factura(50))
        cuerpo(factura(51))
        sin = pd.read_excel(RUTA).to_csv(index=False)
        ok_igual = (con == sin)
        print(f"  {'OK ' if ok_igual else 'FALLA'}  el fichero sale IDENTICO con y sin candado")
        if not ok_igual:
            fallos += 1

        # ── 3 · las columnas que ya estaban se conservan ────────────
        # `oracle_status` es la marca que impide contabilizar dos veces
        # (Fase 0). El candado no la toca, pero mas vale comprobarlo.
        limpiar()
        dashboard._guardar_factura_ap(factura(60))
        df = pd.read_excel(RUTA)
        df['oracle_status'] = 'CONTABILIZADA'
        df['journal_id'] = 'JRN-9001'
        df.to_excel(RUTA, index=False)
        dashboard._guardar_factura_ap(factura(61))          # entra otra al lado
        df2 = pd.read_excel(RUTA)
        fila_vieja = df2[df2['numero_factura'] == 'F-1060']
        ok_marca = (len(df2) == 2
                    and 'oracle_status' in df2.columns
                    and not fila_vieja.empty
                    and str(fila_vieja.iloc[0]['oracle_status']) == 'CONTABILIZADA'
                    and str(fila_vieja.iloc[0]['journal_id']) == 'JRN-9001')
        print(f"  {'OK ' if ok_marca else 'FALLA'}  oracle_status y el Journal ID sobreviven al guardado siguiente")
        if not ok_marca:
            fallos += 1
            print(f'          columnas: {list(df2.columns)}')
            print(f'          filas: {len(df2)}')

        # ── 4 · los tres guardadores estan puestos en fila ──────────
        envueltos = [n for n in ('_guardar_factura_ap', '_guardar_albaran', '_guardar_orden_compra')
                     if hasattr(getattr(dashboard, n), '__wrapped__')]
        ok_tres = len(envueltos) == 3
        print(f"  {'OK ' if ok_tres else 'FALLA'}  los tres guardadores llevan candado: {envueltos}")
        if not ok_tres:
            fallos += 1
    finally:
        limpiar()
        if habia:
            shutil.move(COPIA, RUTA)

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: sin el candado, {fallos} en rojo. El candado es lo que lo arregla.')
            return 0
        print('SABOTAJE MAL: sin candado tambien pasa — la prueba no demuestra nada.')
        return 1
    if fallos:
        print(f'{fallos} en rojo')
        return 1
    print('Todo OK. Con 3 a la vez no se pierde ninguna, y el dato no se mueve.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
