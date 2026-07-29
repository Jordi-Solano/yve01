"""El apaño de `YVE_HOTELES_SEED`: siembra el censo, pero NUNCA pisa nada.

Es un apaño de conveniencia, asi que lo que hay que demostrar no es tanto que
funcione como que **no puede hacer daño**:

  - con hoteles ya dentro, no toca nada (jamas pisa los de un cliente)
  - con la variable sin poner, no hace nada
  - con JSON roto, no revienta: avisa y sigue
  - con hoteles a medias (sin id o sin nombre), los descarta
"""
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "semilla-test"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)

import tenant_dirs                                     # noqa: E402
import censo_hoteles                                   # noqa: E402

SEMILLA = [
    {"id": "HAAA111", "nombre": "Hotel Uno",  "ciudad": "Cadiz",  "categoria": "4★",
     "habitaciones": 120},
    {"id": "HBBB222", "nombre": "Hotel Dos",  "ciudad": "Segovia", "categoria": "5★",
     "habitaciones": 20},
]


def _tenant_limpio():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()
    return base


def _escribir_censo(datos):
    json.dump(datos, open(os.path.join(tenant_dirs.datos_dir(), "hoteles.json"),
                          "w", encoding="utf-8"))


def main():
    fallos = []
    base = _tenant_limpio()

    def caso(etiqueta, valor_env, censo_previo, esperado_n, esperado_nombres=None):
        _escribir_censo(censo_previo)
        if valor_env is None:
            os.environ.pop("YVE_HOTELES_SEED", None)
        else:
            os.environ["YVE_HOTELES_SEED"] = valor_env
        censo_hoteles.sembrar_desde_entorno()
        despues = censo_hoteles.hoteles(solo_activos=False)
        ok = len(despues) == esperado_n
        if ok and esperado_nombres:
            ok = [h["nombre"] for h in despues] == esperado_nombres
        print(f"  {'OK ' if ok else 'MAL'} {etiqueta:<52} -> {len(despues)} hoteles")
        if not ok:
            fallos.append(f"{etiqueta}: {len(despues)} hoteles, esperaba {esperado_n} "
                          f"{[h['nombre'] for h in despues]}")

    print("\n── Semilla del censo desde la variable de entorno ─────────────")

    # 1. Censo vacio + semilla -> siembra
    caso("censo vacio + semilla buena", json.dumps(SEMILLA), [],
         2, ["Hotel Uno", "Hotel Dos"])

    # 2. LO IMPORTANTE: con hoteles dentro NO toca nada.
    caso("YA HAY hoteles -> no los pisa",
         json.dumps(SEMILLA),
         [{"id": "HREAL99", "nombre": "Hotel De Un Cliente", "activo": True}],
         1, ["Hotel De Un Cliente"])

    # 3. Sin variable, no hace nada
    caso("sin la variable puesta", None, [], 0)

    # 4. JSON roto: no revienta
    caso("JSON roto", "{esto no es json", [], 0)

    # 5. No es una lista
    caso("no es una lista", json.dumps({"id": "X", "nombre": "Y"}), [], 0)

    # 6. Lista vacia
    caso("lista vacia", "[]", [], 0)

    # 7. Hoteles a medias: se descartan, los buenos entran
    caso("descarta los que no traen id o nombre",
         json.dumps([{"nombre": "Sin id"}, {"id": "HZZZ"}, SEMILLA[0]]),
         [], 1, ["Hotel Uno"])

    # 8. Rellena los campos que faltan sin inventar el nombre
    _escribir_censo([])
    os.environ["YVE_HOTELES_SEED"] = json.dumps([{"id": "HMIN01", "nombre": "Minimo"}])
    censo_hoteles.sembrar_desde_entorno()
    h = censo_hoteles.hoteles(solo_activos=False)[0]
    completo = (h["categoria"] == "4★" and h["grupo"] == "Principal"
                and h["activo"] is True and h["habitaciones"] == 0)
    print(f"  {'OK ' if completo else 'MAL'} {'rellena los campos que faltan':<52} -> {h}")
    if not completo:
        fallos.append(f"defaults mal: {h}")

    os.environ.pop("YVE_HOTELES_SEED", None)
    shutil.rmtree(base, ignore_errors=True)

    if fallos:
        print("\n FALLOS:")
        for f in fallos:
            print("   ·", f)
        return 1
    print("\n OK · siembra cuando toca y no pisa nunca lo que ya hay\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
