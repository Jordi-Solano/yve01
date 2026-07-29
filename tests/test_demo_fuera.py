"""FASE C · El demo no vuelve, ni al entrar ni al crear un tenant.

Comprueba las dos cosas que el usuario pidio ver, y una tercera que es la
trampa por la que se podia colar todo:

  1. Entrando con `solmar` dos veces seguidas no se regenera nada.
  2. Un tenant nuevo NO nace con `kpis_hoteles.xlsx`.
  3. La trampa del orden: la condicion de la resiembra era
     `si no existe O esta vacio`. Quitando el fichero de las semillas SIN
     quitar antes la resiembra, `no existe` pasa a ser siempre cierto y el demo
     se habria regenerado EN CADA LOGIN, recreando el fichero recien borrado.
     Aqui se comprueba que ese codigo ya no esta en ningun sitio.
"""
import ast
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "fase-c-test"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)

import tenant_dirs                                    # noqa: E402


def test_tenant_nuevo_sin_fichero_de_demo():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()                          # crea el arbol
    datos = tenant_dirs.datos_dir()

    kpis = os.path.join(datos, "kpis_hoteles.xlsx")
    assert not os.path.exists(kpis), \
        "un tenant nuevo sigue naciendo con kpis_hoteles.xlsx"
    assert "kpis_hoteles.xlsx" not in tenant_dirs._SEED_FILES, \
        "kpis_hoteles.xlsx sigue en _SEED_FILES"

    # Los que SI deben seguir sembrandose, para no pasarse de frenada: si al
    # quitar uno se llevara por delante los demas, el tenant nuevo nacería roto
    # y no lo notariamos hasta que alguien subiera un fichero.
    for f in ("extracto_banco.xlsx", "reservas_credito.xlsx", "recetas.xlsx",
              "inventario.xlsx", "ventas_fb_diarias.xlsx", "mermas.xlsx"):
        assert os.path.exists(os.path.join(datos, f)), f"falta la semilla {f}"

    shutil.rmtree(base, ignore_errors=True)


def test_login_no_resiembra():
    """El codigo de la resiembra ya no existe en login.py.

    Se mira el AST y no el texto: el fichero EXPLICA en un comentario lo que
    hacia antes y como se llamaba la funcion, asi que un grep se encontraria a
    si mismo. Es la trampa del assert que cuenta sus propios comentarios, que en
    este proyecto ya ha mordido cuatro veces.
    """
    arbol = ast.parse(open(os.path.join(BASE, "login.py"), encoding="utf-8").read())

    llamadas, importes = [], []
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Call):
            f = nodo.func
            nombre = f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", "")
            if nombre in ("generar_demo", "limpiar_demo"):
                llamadas.append(f"linea {nodo.lineno}: {nombre}()")
        elif isinstance(nodo, ast.ImportFrom):
            if (nodo.module or "") == "demo_generator":
                importes.append(f"linea {nodo.lineno}")
        elif isinstance(nodo, ast.Import):
            for a in nodo.names:
                if a.name == "demo_generator":
                    importes.append(f"linea {nodo.lineno}")

    assert not llamadas, f"login.py sigue llamando al generador de demo: {llamadas}"
    assert not importes, f"login.py sigue importando demo_generator: {importes}"

    # Y que no quede ninguna constante con las cuentas de ejemplo, que era el
    # disparador. Sin esto, alguien podria volver a colgar la resiembra de ella.
    literales = [n.value for n in ast.walk(arbol)
                 if isinstance(n, ast.Constant) and isinstance(n.value, str)]
    for cuenta in ("Cadena Sol", "Gestoría Nord", "Gestoria Nord"):
        assert cuenta not in literales, \
            f"login.py sigue llevando dentro la cadena de demo {cuenta!r}"


def test_el_panel_no_lee_el_fichero_del_demo():
    """Ni el agregador ni el blueprint de Multi-Hotel tocan kpis_hoteles.xlsx."""
    for modulo in ("agregador_grupo.py", "tab_multi_hotel.py"):
        arbol = ast.parse(open(os.path.join(BASE, modulo), encoding="utf-8").read())
        literales = [n.value for n in ast.walk(arbol)
                     if isinstance(n, ast.Constant) and isinstance(n.value, str)]
        # El nombre puede aparecer en un docstring explicando que ya no se usa;
        # lo que no puede aparecer es como cadena dentro del CODIGO. Se separa
        # mirando si el literal es el nombre a secas o una frase que lo menciona.
        malos = [s for s in literales if s.strip() == "kpis_hoteles.xlsx"]
        assert not malos, f"{modulo} sigue abriendo kpis_hoteles.xlsx"


if __name__ == "__main__":
    test_tenant_nuevo_sin_fichero_de_demo()
    test_login_no_resiembra()
    test_el_panel_no_lee_el_fichero_del_demo()
    print("OK · el demo esta fuera: no se siembra, no se resiembra, nadie lo lee")
