"""
tab_multi_hotel.py — Yve.01 Multi-Hotel

FASE C: aqui vivian cuatro endpoints (/overview, /rankings, /alertas y
/hotel/<id>) que leian `kpis_hoteles.xlsx`, el fichero del demo. Ya no existen:
desde la fase B el panel lee `/api/multi_hotel/agregado`, que sale de los
documentos de verdad, y desde la fase C ese fichero ni se siembra ni se
resiembra.

Se borran en vez de dejarlos apagados a proposito. Un endpoint que sigue en pie
leyendo un fichero que ya no existe es peor que no tenerlo: el dia que alguien
lo encuentre creera que es la fuente buena, y contestara 404 "Sin datos KPI"
haciendo pensar que faltan datos cuando lo que falta es el fichero del demo.
"""
import os as _os
from flask import Blueprint, jsonify

multi_hotel_bp = Blueprint('multi_hotel', __name__)
BASE_DIR = _os.path.dirname(_os.path.abspath(__file__))


@multi_hotel_bp.route('/api/multi_hotel/agregado')
def api_multi_agregado():
    """La ficha real de cada hotel, leyendo los documentos.

    UNICA fuente del panel Multi-Hotel desde la fase B. Nacio en la fase A como
    endpoint de diagnostico, para poder comparar contra los paneles de cada
    hotel antes de enseñar nada: si el agregador da los mismos numeros por otro
    camino, esta bien; si no, uno de los dos miente. Sigue valiendo para eso.

    Devuelve una caja por hotel, mas `sin_asignar` y `desconocido`, mas el total
    del grupo calculado por separado y el bloque de `cuadre` que compara los
    dos. Solo lee: ver `agregador_grupo` y sus tests.
    """
    try:
        from agregador_grupo import agregado
        return jsonify(agregado())
    except Exception as e:
        import traceback
        return jsonify({'ok': False, 'error': str(e),
                        'traza': traceback.format_exc()[-1500:]}), 500
