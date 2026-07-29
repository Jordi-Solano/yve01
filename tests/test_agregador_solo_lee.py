"""El agregador SOLO LEE. Esta es la comprobacion, y mira codigo, no texto.

Con `grep` no vale: el propio modulo explica en su docstring como comprobarlo,
asi que el grep se encuentra a si mismo y da un falso positivo. Es la trampa de
siempre —el assert que cuenta sus propios comentarios— y ya nos ha mordido
cuatro veces en este proyecto. Aqui se recorre el AST, donde los docstrings y
los comentarios no existen.
"""
import ast
import os

RUTA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "agregador_grupo.py")

# Dos listas, y la diferencia importa.
#
# Estas dejan huella en el disco se llamen como se llamen:
ESCRITURA = {
    "to_excel", "to_csv", "to_json", "to_pickle", "to_parquet",
    "save", "savefig", "write", "writelines", "writerow", "writerows",
    "truncate",
}

# Y estas SOLO si van sobre `os`, `shutil`, `json` o `pickle`. Sin el receptor
# darian falsos positivos y el test se volveria ruido que hay que silenciar:
# `df.copy()` copia en memoria y `"a".replace()` es texto — ninguna de las dos
# toca el disco. Un test que grita por cosas correctas acaba desactivado, y
# entonces no protege de nada.
ESCRITURA_DE_MODULO = {
    ("os", "remove"), ("os", "unlink"), ("os", "rename"), ("os", "replace"),
    ("os", "mkdir"), ("os", "makedirs"), ("os", "rmdir"), ("os", "truncate"),
    ("shutil", "rmtree"), ("shutil", "copy"), ("shutil", "copy2"),
    ("shutil", "copyfile"), ("shutil", "move"),
    ("json", "dump"), ("pickle", "dump"),
}
MODOS_ESCRITURA = ("w", "a", "x", "+")


def _nombre(nodo):
    if isinstance(nodo, ast.Attribute):
        return nodo.attr
    if isinstance(nodo, ast.Name):
        return nodo.id
    return ""


def _receptor(nodo):
    """`shutil` en `shutil.copy(...)`. Cadena vacia si no es un nombre suelto."""
    if isinstance(nodo, ast.Attribute) and isinstance(nodo.value, ast.Name):
        return nodo.value.id
    return ""


def test_agregador_no_escribe():
    arbol = ast.parse(open(RUTA, encoding="utf-8").read())
    delitos = []

    for nodo in ast.walk(arbol):
        if not isinstance(nodo, ast.Call):
            continue
        nombre = _nombre(nodo.func)

        if nombre in ESCRITURA:
            delitos.append(f"linea {nodo.lineno}: llamada a {nombre}()")

        if (_receptor(nodo.func), nombre) in ESCRITURA_DE_MODULO:
            delitos.append(f"linea {nodo.lineno}: {_receptor(nodo.func)}.{nombre}()")

        # open(..., 'w') y compañia. El modo puede ir posicional o por nombre.
        if nombre == "open":
            modos = [a for a in nodo.args[1:2]]
            modos += [k.value for k in nodo.keywords if k.arg == "mode"]
            for m in modos:
                if isinstance(m, ast.Constant) and isinstance(m.value, str) \
                        and any(c in m.value for c in MODOS_ESCRITURA):
                    delitos.append(f"linea {nodo.lineno}: open(..., {m.value!r})")

    assert not delitos, (
        "agregador_grupo.py tiene que ser de SOLO LECTURA y escribe:\n  "
        + "\n  ".join(delitos))


def test_no_toca_oracle_ni_el_clasificador():
    """Ni lo importa. Corre despues de todo, sobre lo ya guardado."""
    arbol = ast.parse(open(RUTA, encoding="utf-8").read())
    importados = []
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Import):
            importados += [a.name for a in nodo.names]
        elif isinstance(nodo, ast.ImportFrom):
            importados.append(nodo.module or "")

    prohibidos = [m for m in importados
                  if m.startswith("oracle") or "clasificador" in m]
    assert not prohibidos, f"importa modulos intocables: {prohibidos}"


if __name__ == "__main__":
    test_agregador_no_escribe()
    test_no_toca_oracle_ni_el_clasificador()
    print("OK · agregador_grupo.py solo lee, y no toca Oracle ni el clasificador")
