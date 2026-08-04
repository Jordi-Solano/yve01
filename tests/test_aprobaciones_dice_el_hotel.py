# -*- coding: utf-8 -*-
"""El panel de Aprobaciones dice QUÉ hotel enseña, y no miente al vaciarse.

El bug que vivió Jordi: tenía una factura del Hotel Costa Azul a la vista,
cambió de hotel siguiendo la guía, y la factura desapareció. Cambió el filtro
de departamento y volvió a "Todos" — no volvió. Salió y entró — tampoco.

El filtro por hotel es CORRECTO y no se toca: aprobar la factura de otro hotel
acaba en el libro mayor. Lo que estaba mal es que el panel filtraba **en
silencio** y, al quedarse vacío, decía "No queda nada por aprobar" — falso, y
peligroso justo aquí, porque es la puerta de Oracle: quien se lo cree y cierra
deja facturas sin aprobar.

  python3.12 tests/test_aprobaciones_dice_el_hotel.py
  python3.12 tests/test_aprobaciones_dice_el_hotel.py --sabotaje
"""
import os
import re
import shutil
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_TENANT'] = 'default'

import pandas as pd                                    # noqa: E402
import censo_hoteles                                   # noqa: E402
import app_aprobacion_ap as PANEL                      # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv

HOTELES = [{'id': 'HTEST01', 'nombre': 'Hotel de Prueba'},
           {'id': 'HTEST02', 'nombre': 'Otro Hotel'}]


def fila(n, prov, tot, hotel, dept='General'):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'base_imponible': tot, 'cuota_iva': 0, 'hotel_id': hotel,
            'fecha_factura': '01/08/2026', 'cuenta_contable': '600',
            'estado_asignacion': 'ASIGNADA', 'departamento_po': dept}


def con_hotel(activo, fn):
    _h, _a = censo_hoteles.hoteles, censo_hoteles.activo
    # `por_id` llama a `hoteles(solo_activos=False)`: el stub TIENE que
    # aceptar argumentos o `_nombre_hotel` cae al id por su try/except.
    censo_hoteles.hoteles = lambda *a, **kw: HOTELES
    censo_hoteles.activo = lambda: activo
    try:
        return fn()
    finally:
        censo_hoteles.hoteles, censo_hoteles.activo = _h, _a


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: el panel vuelve a callar el hotel ***')
    fallos = 0
    raiz = tempfile.mkdtemp(prefix='aph_')
    try:
        pdir = os.path.join(raiz, 'facturas-procesadas')
        apdir = os.path.join(raiz, 'aprobaciones')
        rdir = os.path.join(raiz, 'reportes')
        for d in (pdir, apdir, rdir):
            os.makedirs(d)
        # 2 facturas del hotel 1, 1 del hotel 2
        pd.DataFrame([fila('F-001', 'Makro', 121, 'HTEST01'),
                      fila('F-002', 'Lavanderia', 242, 'HTEST01'),
                      fila('F-003', 'Bebidas', 363, 'HTEST02')]
                     ).to_excel(os.path.join(pdir, 'facturas_ap_20260803.xlsx'), index=False)
        PANEL.PROCESADAS_DIR = pdir
        PANEL.REPORTES_DIR = rdir
        PANEL.APROBACIONES_DIR = apdir
        PANEL.APRO_FILE = os.path.join(apdir, 'aprobaciones_ap.xlsx')

        import flask
        app = flask.Flask(__name__)
        app.register_blueprint(PANEL.bp)
        app.config['TESTING'] = True
        # el blueprint exige login; en la prueba se mira el endpoint a pelo
        c = app.test_client()

        def stats(activo):
            return con_hotel(activo, lambda: PANEL.api_stats.__wrapped__()
                             if hasattr(PANEL.api_stats, '__wrapped__') else None)

        # se llaman las funciones directamente: lo que se prueba es el dato,
        # no el enrutado (que ya está probado en test_aprobaciones_por_hotel)
        def datos(activo):
            def _f():
                filas = PANEL.facturas_a_lista(PANEL.cargar_facturas_ap())
                return {'visibles': sorted(r['numero_factura'] for r in filas),
                        'hotel_id': PANEL._hotel_activo_id(),
                        'hotel_nombre': PANEL._nombre_hotel(PANEL._hotel_activo_id()),
                        'en_otros': PANEL.pendientes_en_otros_hoteles()}
            return con_hotel(activo, _f)

        # ── 1 · el hotel 1 ve lo suyo y sabe que hay más fuera ───────
        d1 = datos('HTEST01')
        if SABOTAJE:
            d1 = dict(d1, hotel_nombre='', en_otros=0)
        ok1 = d1['visibles'] == ['F-001', 'F-002']
        print(f"  {'OK ' if ok1 else 'FALLA'}  el filtro por hotel NO cambia: {d1['visibles']}")
        if not ok1:
            fallos += 1

        ok2 = d1['hotel_nombre'] == 'Hotel de Prueba'
        print(f"  {'OK ' if ok2 else 'FALLA'}  el panel sabe qué hotel enseña: "
              f"{d1['hotel_nombre']!r}")
        if not ok2:
            fallos += 1

        ok3 = d1['en_otros'] == 1
        print(f"  {'OK ' if ok3 else 'FALLA'}  y cuántas quedan en los otros: {d1['en_otros']} "
              f"(esperaba 1)")
        if not ok3:
            fallos += 1

        # ── 2 · EL CASO DE JORDI · un hotel sin nada, con trabajo fuera
        raiz2 = tempfile.mkdtemp(prefix='aph2_')
        pdir2 = os.path.join(raiz2, 'facturas-procesadas')
        os.makedirs(pdir2)
        os.makedirs(os.path.join(raiz2, 'reportes'))
        pd.DataFrame([fila('F-001', 'Makro', 121, 'HTEST01')]
                     ).to_excel(os.path.join(pdir2, 'facturas_ap_20260803.xlsx'), index=False)
        PANEL.PROCESADAS_DIR = pdir2
        PANEL.REPORTES_DIR = os.path.join(raiz2, 'reportes')
        try:
            d2 = datos('HTEST02')          # el hotel donde NO está la factura
            if SABOTAJE:
                d2 = dict(d2, hotel_nombre='', en_otros=0)
            ok4 = d2['visibles'] == []
            print(f"  {'OK ' if ok4 else 'FALLA'}  con el otro hotel, la lista sale vacía "
                  f"(el filtro sigue igual): {d2['visibles']}")
            if not ok4:
                fallos += 1

            ok5 = d2['en_otros'] == 1 and d2['hotel_nombre'] == 'Otro Hotel'
            print(f"  {'OK ' if ok5 else 'FALLA'}  ...pero AHORA se puede decir la verdad: "
                  f"vacío en {d2['hotel_nombre']!r}, {d2['en_otros']} esperando fuera")
            if not ok5:
                fallos += 1
        finally:
            shutil.rmtree(raiz2, ignore_errors=True)
            PANEL.PROCESADAS_DIR = pdir
            PANEL.REPORTES_DIR = rdir

        # ── 3 · vista de grupo: se ve todo y no falta nada fuera ─────
        d3 = datos('')
        ok6 = d3['visibles'] == ['F-001', 'F-002', 'F-003'] and d3['en_otros'] == 0
        print(f"  {'OK ' if ok6 else 'FALLA'}  en vista de grupo se ve todo y no hay "
              f"«otros»: {d3['visibles']} · fuera {d3['en_otros']}")
        if not ok6:
            fallos += 1

        # ── 4 · la pantalla: chip + mensaje que no miente ────────────
        html = PANEL.HTML
        if SABOTAJE:
            html = html.replace('No queda nada por aprobar en ', 'No queda nada por aprobar')
            html = html.replace('id="hotel-chip"', 'id="otra-cosa"')
            html = html.replace('pista-otros', 'nada')
        comprob = [
            ('la barra tiene el distintivo del hotel', 'id="hotel-chip"' in html),
            ('el mensaje vacío nombra el hotel', 'No queda nada por aprobar en ' in html),
            ('y avisa de las que esperan fuera', 'pista-otros' in html
             and 'esperando en' in html),
            ('el nombre del hotel se escapa (no se inyecta)',
             "esc(hn)" in html or "txt(hn)" in html),
        ]
        for nombre, ok in comprob:
            print(f"  {'OK ' if ok else 'FALLA'}  {nombre}")
            if not ok:
                fallos += 1

        # ── 5 · nada de lo que decide QUÉ se aprueba se ha movido ────
        import subprocess
        dif = subprocess.run(['git', 'diff', 'HEAD', '--', 'app_aprobacion_ap.py'],
                             capture_output=True, text=True, cwd=BASE).stdout
        tocadas = [l for l in dif.split('\n')
                   if (l.startswith('+') or l.startswith('-'))
                   and not l.startswith('+++') and not l.startswith('---')]
        peligrosas = [l for l in tocadas
                      if re.search(r'solo_del_hotel_activo|def api_accion|APROBADA|'
                                   r'RECHAZADA|dept in r\.get', l)]
        ok7 = not peligrosas if not SABOTAJE else False
        print(f"  {'OK ' if ok7 else 'FALLA'}  el diff no toca el filtro ni la aprobación "
              f"({len(tocadas)} líneas, {len(peligrosas)} peligrosas)")
        if not ok7:
            fallos += 1
    finally:
        shutil.rmtree(raiz, ignore_errors=True)

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
    print('Todo OK. El panel dice qué hotel enseña y, si está vacío, dice la verdad.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
