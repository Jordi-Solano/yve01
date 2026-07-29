"""FASE F · El codigo muerto esta muerto, y el bug de las estrellas no vuelve.

Dos de los tres puntos de la fase F ya los habian matado las fases B y C sin
que nos dieramos cuenta: la tabla que adivinaba las estrellas del nombre se fue
con `api_multi_overview` (fase C), y el panel que la pintaba se reescribio
entero (fase B). Lo que quedaba era el cadaver: siete funciones que no llamaba
nadie, con el bug dentro.

Esto comprueba que no vuelven. Un test de "esto ya no existe" parece tonto
hasta que alguien reintroduce la funcion copiando de una version vieja.
"""
import ast
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)


def _js_servido(sin_comentarios=True):
    """Los bloques <script> del HTML que se SIRVE, no el fuente de Python.

    Por defecto SIN comentarios, y esa es la parte importante. El codigo
    explica en un comentario cual era el bug —`'★'.repeat(h.stars)`— y ese
    comentario viaja al navegador dentro del <script>, asi que un test que
    busque el texto se encuentra a si mismo y falla con el arreglo puesto.

    Es la trampa del assert que cuenta sus propios comentarios. En este
    proyecto ya ha mordido cinco veces, esta incluida.
    """
    src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    html = re.search(r'^HTML\s*=\s*r?"""(.*?)"""', src, re.S | re.M).group(1)
    js = "\n".join(re.findall(r"<script[^>]*>(.*?)</script>", html, re.S))
    if not sin_comentarios:
        return js
    js = re.sub(r"/\*.*?\*/", "", js, flags=re.S)          # bloques
    # Comentarios de linea, pero NO el '//' de dentro de una URL ('https://').
    js = re.sub(r"(?<![:'\"])//[^\n]*", "", js)
    return js


def test_los_modulos_muertos_no_estan():
    for m in ("multi_hotel_data.py", "dashboard_multihotel.py"):
        assert not os.path.exists(os.path.join(BASE, m)), f"{m} sigue en el repo"


def test_las_funciones_huerfanas_no_vuelven():
    js = _js_servido()
    for f in ("renderMHTableFull", "renderMHStatus", "renderMHRankings",
              "renderMHAlertasClasica", "renderMHGrupos", "filtrarMHGrupo",
              "_calSparkline", "renderMHMap", "openHotelDetail"):
        assert f"function {f}(" not in js, f"{f} ha vuelto al JS servido"


def test_las_estrellas_no_se_adivinan():
    js = _js_servido()

    # El bug: '★'.repeat(x) con x siendo el TEXTO '4★'. repeat() convierte su
    # argumento a numero, '4★' da NaN, NaN se convierte a 0, y la columna sale
    # SIEMPRE vacia. Un fallo que no peta y no se ve: lo peor de los dos mundos.
    assert "'★'.repeat(" not in js, "ha vuelto el repeat sobre las estrellas"

    # Y la adivinanza por el nombre, que hacia que un "Hotel 5 de Mayo" saliera
    # de cinco estrellas.
    for veneno in ("'5★' :", "'5★':", "boutique"):
        assert veneno not in js, f"ha vuelto la adivinanza de categoria ({veneno!r})"


def test_la_categoria_sale_del_censo():
    """El agregador lee categoria/habitaciones del censo, que es quien las sabe."""
    arbol = ast.parse(open(os.path.join(BASE, "agregador_grupo.py"), encoding="utf-8").read())
    literales = {n.value for n in ast.walk(arbol)
                 if isinstance(n, ast.Constant) and isinstance(n.value, str)}
    for campo in ("categoria", "habitaciones", "censo"):
        assert campo in literales, f"el agregador ya no expone {campo!r}"


def test_multihotel_solo_en_vista_de_grupo():
    js = _js_servido()
    assert "_ordenarPestanaMultiHotel" in js, "falta la regla de visibilidad de Multi-Hotel"
    # Y que se llama al resolver el hotel activo, no solo que exista.
    assert re.search(r"_ordenarPestanaMultiHotel\(!!\(?d", js), \
        "la regla existe pero no se aplica al cambiar de hotel"


if __name__ == "__main__":
    test_los_modulos_muertos_no_estan()
    test_las_funciones_huerfanas_no_vuelven()
    test_las_estrellas_no_se_adivinan()
    test_la_categoria_sale_del_censo()
    test_multihotel_solo_en_vista_de_grupo()
    print("OK · codigo muerto fuera, estrellas del censo, Multi-Hotel solo en grupo")
