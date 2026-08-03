# -*- coding: utf-8 -*-
"""M8 — el historial de archivos procesados, separado por hotel.

Era otra puerta que no filtraba: enseñaba los archivos de TODOS los hoteles
mezclados. Dos mitades:

  · al ESCRIBIR — las cuatro puertas que apuntan en el log (el lote, el
    escaneo por foto, el camino no-stream y el reproceso) tienen que estampar
    el hotel. Se hace en un único sitio, `_entrada_proc()`, para que ninguna
    pueda volver a olvidarse.
  · al LEER — igualdad estricta, como el cruce factura↔albarán: el vacío NO es
    comodín. Con un hotel elegido se ve sólo lo suyo; en vista de grupo, todo.

  python3.12 tests/test_historial_por_hotel.py
  python3.12 tests/test_historial_por_hotel.py --sabotaje
"""
import json
import os
import re
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import censo_hoteles                                    # noqa: E402
import dashboard                                        # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv

LOG = {
    'factura_a.pdf': {'fecha': '2026-08-01 10:00', 'resultado': 'AP_OK', 'hotel_id': 'HTEST01'},
    'factura_b.pdf': {'fecha': '2026-08-01 11:00', 'resultado': 'AP_OK', 'hotel_id': 'HTEST01'},
    'factura_c.pdf': {'fecha': '2026-08-01 12:00', 'resultado': 'AP_OK', 'hotel_id': 'HTEST02'},
    'viejo_1.pdf':   {'fecha': '2026-07-01 09:00', 'resultado': 'AP_OK'},
    'viejo_2.pdf':   {'fecha': '2026-07-01 08:00', 'resultado': 'ALBARAN_OK'},
}
HOTELES = [{'id': 'HTEST01', 'nombre': 'Hotel de Prueba'},
           {'id': 'HTEST02', 'nombre': 'Otro Hotel'}]

app = dashboard.app
app.config['TESTING'] = True
c = app.test_client()
assert c.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200


def ver(hoteles, activo):
    _h, _a = censo_hoteles.hoteles, censo_hoteles.activo
    censo_hoteles.hoteles = lambda: hoteles
    censo_hoteles.activo = lambda: activo
    try:
        j = c.get('/api/historial_procesado').get_json() or {}
        return sorted(x['archivo'] for x in (j.get('items') or []))
    finally:
        censo_hoteles.hoteles, censo_hoteles.activo = _h, _a


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: se quita el filtro y el sello del hotel ***')
    fallos = 0
    RUTA = dashboard.PROC_LOG_PATH
    habia = os.path.exists(RUTA)
    copia = RUTA + '.bak-m8'
    if habia:
        shutil.copy(RUTA, copia)
    try:
        json.dump(LOG, open(RUTA, 'w', encoding='utf-8'), ensure_ascii=False)
        TODO = sorted(LOG)

        casos = [
            ('vista de grupo: se ve TODO, como siempre', ([], ''), TODO),
            ('2 hoteles pero ninguno elegido: se ve TODO', (HOTELES, ''), TODO),
            ('Hotel 1 elegido: sólo lo suyo', (HOTELES, 'HTEST01'),
             ['factura_a.pdf', 'factura_b.pdf']),
            ('Hotel 2 elegido: sólo lo suyo', (HOTELES, 'HTEST02'), ['factura_c.pdf']),
            ('un hotel que no tiene nada: vacío, no todo', (HOTELES, 'HTEST99'), []),
        ]
        for nombre, (hs, act), esperado in casos:
            got = TODO if SABOTAJE else ver(hs, act)
            ok = got == esperado
            print(f"  {'OK ' if ok else 'FALLA'}  {nombre}: {got}")
            if not ok:
                fallos += 1

        # el vacío NO es comodín: lo viejo sin hotel no se cuela en un hotel
        got = TODO if SABOTAJE else ver(HOTELES, 'HTEST01')
        ok = 'viejo_1.pdf' not in got and 'viejo_2.pdf' not in got
        print(f"  {'OK ' if ok else 'FALLA'}  lo viejo sin hotel NO se cuela en un hotel "
              f"(el vacío no es comodín)")
        if not ok:
            fallos += 1

        # ── al escribir: las CUATRO puertas pasan por el mismo sitio ─────
        src = open('dashboard.py', encoding='utf-8').read()
        if SABOTAJE:
            src = src.replace('_entrada_proc(', '_NO(')
        n_usos = len(re.findall(r'=\s*_entrada_proc\(', src))
        ok_usos = n_usos == 4
        print(f"  {'OK ' if ok_usos else 'FALLA'}  las 4 puertas que escriben el log usan "
              f"_entrada_proc(): {n_usos}")
        if not ok_usos:
            fallos += 1

        # y ninguna se ha quedado escribiendo la entrada a mano
        n_mano = len(re.findall(r"log\[[^\]]+\]\s*=\s*\{'fecha'", src))
        ok_mano = n_mano == 0 if not SABOTAJE else False
        print(f"  {'OK ' if ok_mano else 'FALLA'}  ninguna escribe la entrada a mano: "
              f"{n_mano} sitios")
        if not ok_mano:
            fallos += 1

        # la entrada lleva el hotel de verdad
        _h, _a = censo_hoteles.hoteles, censo_hoteles.activo
        censo_hoteles.hoteles = lambda: HOTELES
        censo_hoteles.activo = lambda: 'HTEST02'
        try:
            e = dashboard._entrada_proc('AP_OK')
        finally:
            censo_hoteles.hoteles, censo_hoteles.activo = _h, _a
        if SABOTAJE:
            e = {'fecha': e['fecha'], 'resultado': 'AP_OK'}
        ok_e = (e.get('hotel_id') == 'HTEST02' and e.get('resultado') == 'AP_OK'
                and re.match(r'^\d{4}-\d\d-\d\d \d\d:\d\d$', str(e.get('fecha', ''))))
        print(f"  {'OK ' if ok_e else 'FALLA'}  la entrada nueva lleva hotel, fecha y "
              f"resultado: {e}")
        if not ok_e:
            fallos += 1
    finally:
        if os.path.exists(RUTA):
            os.remove(RUTA)
        if habia:
            shutil.move(copia, RUTA)

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
    print('Todo OK. El historial se separa por hotel y la vista de grupo sigue '
          'enseñándolo todo.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
