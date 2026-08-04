# -*- coding: utf-8 -*-
"""AR, F&B, DRR y Banco dicen POR QUÉ están vacíos.

Mismo bug que se sufrió en Aprobaciones AP: las cuatro pantallas filtran por
hotel —correctamente— y al quedarse vacías decían "Sin datos", que es falso
cuando el dato está en otro hotel. Una lista vacía era indistinguible de "no
hay trabajo".

Se arregla con la mínima superficie: UN endpoint que solo cuenta
(`/api/hay_en_otros_hoteles`) y UN ayudante de JS que cuelga la pista bajo el
vacío que ya existe. **No se toca ni un filtro ni el texto de ningún mensaje**
— esos textos se traducen por coincidencia exacta y cambiarlos los dejaría en
español en los otros 6 idiomas.

  python3.12 tests/test_vacios_dicen_el_hotel.py
  python3.12 tests/test_vacios_dicen_el_hotel.py --sabotaje
"""
import json
import os
import re
import shutil
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_TENANT'] = 'default'

import pandas as pd                                    # noqa: E402
import censo_hoteles                                   # noqa: E402
import dashboard                                       # noqa: E402
from tenant_dirs import procesadas_dir, reportes_dir, datos_dir   # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = datetime.now().strftime('%Y%m%d')
HOTELES = [{'id': 'HTEST01', 'nombre': 'Hotel de Prueba'},
           {'id': 'HTEST02', 'nombre': 'Otro Hotel'}]

app = dashboard.app
app.config['TESTING'] = True
c = app.test_client()
assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200


def con_hotel(activo, fn):
    _h, _a = censo_hoteles.hoteles, censo_hoteles.activo
    censo_hoteles.hoteles = lambda *a, **kw: HOTELES
    censo_hoteles.activo = lambda: activo
    try:
        return fn()
    finally:
        censo_hoteles.hoteles, censo_hoteles.activo = _h, _a


def pregunta(area, activo):
    def _f():
        j = c.get('/api/hay_en_otros_hoteles?area=' + area).get_json() or {}
        return j
    return con_hotel(activo, _f)


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: las pantallas vuelven a callarse el hotel ***')
    fallos = 0
    creados = []

    def guardar(ruta, filas):
        """Escribe un fichero de prueba y lo apunta para borrarlo al final."""
        if os.path.exists(ruta):
            shutil.copy(ruta, ruta + '.bak-vh')
            creados.append((ruta, True))
        else:
            creados.append((ruta, False))
        pd.DataFrame(filas).to_excel(ruta, index=False)

    try:
        # ── datos de mentira: TODO del hotel 2, nada del hotel 1 ────
        guardar(os.path.join(procesadas_dir(), f'facturas_procesadas_{HOY}.xlsx'), [
            {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com',
             'periodo_inicio': '2026-07-01', 'importe_comision': 100, 'hotel_id': 'HTEST02'},
            {'numero_factura': 'BK-2', 'nombre_ota': 'Expedia',
             'periodo_inicio': '2026-07-01', 'importe_comision': 200, 'hotel_id': 'HTEST02'}])
        guardar(os.path.join(datos_dir(), 'ventas_fb_diarias.xlsx'), [
            {'nombre_plato': 'Paella', 'unidades_vendidas': 3, 'total_venta': 60,
             'hotel_id': 'HTEST02'}])

        # ── 1 · con el hotel VACÍO, se sabe que hay cosas fuera ─────
        casos = [('ar', 2), ('fb', 1)]
        for area, esperado in casos:
            j = pregunta(area, 'HTEST01')
            n = 0 if SABOTAJE else j.get('n')
            nombre = '' if SABOTAJE else j.get('hotel_nombre')
            ok = (n == esperado and nombre == 'Hotel de Prueba')
            print(f"  {'OK ' if ok else 'FALLA'}  {area}: el hotel vacío sabe que hay {n} "
                  f"fuera (esperaba {esperado}) y que está en {nombre!r}")
            if not ok:
                fallos += 1

        # ── 2 · desde el hotel que SÍ los tiene, no hay nada fuera ──
        j = pregunta('ar', 'HTEST02')
        n2 = 0 if SABOTAJE else j.get('n')
        ok2 = n2 == 0
        print(f"  {'OK ' if ok2 else 'FALLA'}  desde el hotel que los tiene, fuera hay {n2} "
              f"(esperaba 0)")
        if not ok2:
            fallos += 1

        # ── 3 · vista de grupo: no se dice nada (se ve todo) ────────
        j = pregunta('ar', '')
        ok3 = j.get('hotel_nombre') == '' and j.get('n') == 0
        print(f"  {'OK ' if ok3 else 'FALLA'}  en vista de grupo la pista se calla: "
              f"hotel={j.get('hotel_nombre')!r} n={j.get('n')}")
        if not ok3:
            fallos += 1

        # ── 4 · un área desconocida no revienta ni inventa ──────────
        j = pregunta('loquesea', 'HTEST01')
        ok4 = j.get('ok') is True and j.get('n') == 0
        print(f"  {'OK ' if ok4 else 'FALLA'}  un área desconocida devuelve 0, no revienta: "
              f"{j.get('n')}")
        if not ok4:
            fallos += 1

        # ── 5 · el frontend servido ────────────────────────────────
        html = c.get('/').get_data(as_text=True)
        if SABOTAJE:
            html = html.replace('_pistaOtrosHoteles', '_noExiste')
        comprob = [
            ('existe el ayudante', 'async function _pistaOtrosHoteles(' in html),
            ('AR lo llama', "_pistaOtrosHoteles('ar'" in html),
            ('DRR lo llama', "_pistaOtrosHoteles('drr'" in html),
            ('Banco lo llama', "_pistaOtrosHoteles('banco'" in html),
            ('F&B lo llama', "_pistaOtrosHoteles('fb'" in html),
            ('el nombre del hotel se escapa', '_escHtml(d.hotel_nombre)' in html),
            ('la pista tiene estilo', '.pista-hotel{' in html),
        ]
        for nombre, ok in comprob:
            print(f"  {'OK ' if ok else 'FALLA'}  {nombre}")
            if not ok:
                fallos += 1

        # ── 6 · NINGÚN texto de vacío ha cambiado ───────────────────
        # Si cambiara, se rompen las traducciones: `_i18nStrMap` casa por texto
        # exacto y esos mensajes se quedarían en español en los otros 6 idiomas.
        intactos = ['Sin datos.<br>Pulsa ⚡ Procesar Archivos.',
                    'Sin DRR para este hotel.', 'Sin movimientos bancarios.']
        malos = [m for m in intactos if m not in html]
        ok6 = not malos if not SABOTAJE else False
        print(f"  {'OK ' if ok6 else 'FALLA'}  los textos de vacío siguen intactos "
              f"(las traducciones valen){'' if ok6 else ' — falta ' + str(malos)}")
        if not ok6:
            fallos += 1

        # ── 7 · las traducciones conservan los marcadores ───────────
        malas = []
        for lang in ('en', 'ca', 'fr', 'de', 'it', 'pt'):
            d = json.load(open(f'static/i18n/{lang}.json', encoding='utf-8'))
            if '{hotel}' not in d.get('vacio.enOtros', '') or '{n}' not in d.get('vacio.enOtros', ''):
                malas.append(lang)
            if '{hotel}' not in d.get('vacio.soloEste', ''):
                malas.append(lang + '/solo')
        ok7 = not malas if not SABOTAJE else False
        print(f"  {'OK ' if ok7 else 'FALLA'}  los 6 idiomas conservan {{hotel}} y {{n}}"
              f"{'' if ok7 else ' — ' + str(malas)}")
        if not ok7:
            fallos += 1

        # ── 8 · los filtros no se han movido ────────────────────────
        import subprocess
        dif = subprocess.run(['git', 'diff', 'HEAD', '--', 'dashboard.py'],
                             capture_output=True, text=True, cwd=BASE).stdout
        tocadas = [l for l in dif.split('\n')
                   if (l.startswith('+') or l.startswith('-'))
                   and not l.startswith('+++') and not l.startswith('---')]
        peligrosas = [l for l in tocadas if re.search(
            r'_solo_hotel_activo|def drr_del_hotel|_modo == "por_hotel"|fichero_es_de\(', l)
            and '_pistaOtrosHoteles' not in l and 'hay_en_otros' not in l]
        # las del endpoint nuevo son suyas y no cuentan
        peligrosas = [l for l in peligrosas if 'censo_hoteles.fichero_es_de(b' not in l]
        ok8 = not peligrosas if not SABOTAJE else False
        print(f"  {'OK ' if ok8 else 'FALLA'}  el diff no toca ningún filtro por hotel "
              f"({len(tocadas)} líneas, {len(peligrosas)} peligrosas)")
        if not ok8:
            fallos += 1
    finally:
        for ruta, habia in creados:
            if os.path.exists(ruta):
                os.remove(ruta)
            if habia and os.path.exists(ruta + '.bak-vh'):
                shutil.move(ruta + '.bak-vh', ruta)

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
    print('Todo OK. Las cuatro pantallas dicen por qué están vacías, y ningún '
          'filtro ni ninguna traducción se ha movido.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
