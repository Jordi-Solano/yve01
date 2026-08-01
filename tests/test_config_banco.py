"""La config del banco: 'grupo' vs 'por_hotel', elegida por el usuario una vez.

Es la base de todo el frente del banco, así que lo que hay que demostrar es:

  - de fábrica no hay elección (modo=='', elegido()==False): el banco va junto,
    como siempre; nadie nota nada hasta que el usuario elige.
  - elegir('grupo'|'por_hotel') se guarda y se relee (sobrevive al proceso).
  - un valor inválido NO se guarda y NO corrompe una elección ya hecha.
  - la semilla YVE_BANCO_MODO siembra cuando no hay nada, pero NUNCA pisa una
    elección existente (mismo apaño y misma garantía que el censo de hoteles).
  - vive en el servidor (config_banco.json del tenant), no en el navegador.
"""
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "config-banco-test"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)
os.environ.pop("YVE_BANCO_MODO", None)

import tenant_dirs                                     # noqa: E402
import config_banco                                    # noqa: E402


def _tenant_limpio():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()
    return base


def _borrar_config():
    try:
        os.remove(os.path.join(tenant_dirs.datos_dir(), "config_banco.json"))
    except OSError:
        pass


def main():
    fallos = []
    base = _tenant_limpio()

    def chk(etiqueta, cond):
        print(f"  {'OK ' if cond else 'MAL'} {etiqueta}")
        if not cond:
            fallos.append(etiqueta)

    print("\n── Config del banco: grupo vs por hotel ───────────────────────")

    # 1. De fábrica: sin elección. El banco va junto, como siempre.
    _borrar_config()
    chk("de fábrica modo()==''", config_banco.modo() == "")
    chk("de fábrica elegido()==False", config_banco.elegido() is False)
    chk("de fábrica por_hotel()==False", config_banco.por_hotel() is False)

    # 2. Elegir 'grupo' se guarda y se relee.
    r = config_banco.elegir("grupo")
    chk("elegir('grupo') devuelve 'grupo'", r == "grupo")
    chk("modo()=='grupo' tras elegir", config_banco.modo() == "grupo")
    chk("elegido()==True tras elegir", config_banco.elegido() is True)
    chk("por_hotel()==False en modo grupo", config_banco.por_hotel() is False)
    # Está en un fichero del servidor, no en el navegador.
    ruta = os.path.join(tenant_dirs.datos_dir(), "config_banco.json")
    chk("existe config_banco.json en el tenant", os.path.exists(ruta))
    chk("el JSON guarda {'modo':'grupo'}",
        json.load(open(ruta, encoding="utf-8")).get("modo") == "grupo")

    # 3. Cambiar a 'por_hotel' (elección directa: sí puede sobrescribir).
    r = config_banco.elegir("por_hotel")
    chk("elegir('por_hotel') devuelve 'por_hotel'", r == "por_hotel")
    chk("por_hotel()==True en modo por_hotel", config_banco.por_hotel() is True)

    # 4. El payload entra sin pulir (viene del navegador): se normaliza mayúsculas
    #    y espacios, no se rechaza por eso.
    chk("elegir('GRUPO ') normaliza a 'grupo'", config_banco.elegir("GRUPO ") == "grupo")
    chk("elegir('  Por_Hotel ') normaliza a 'por_hotel'",
        config_banco.elegir("  Por_Hotel ") == "por_hotel")

    # 5. Un valor de verdad inválido NO se guarda y NO corrompe lo que ya había.
    config_banco.elegir("por_hotel")   # estado conocido de partida
    for malo in ("basura", "", None, 3, "cuenta-unica"):
        r = config_banco.elegir(malo)
        ok = (r == "" and config_banco.modo() == "por_hotel")
        chk(f"elegir({malo!r}) se rechaza y no corrompe", ok)

    # 6. La semilla NO pisa una elección existente (la propiedad de seguridad).
    os.environ["YVE_BANCO_MODO"] = "grupo"
    config_banco.sembrar_desde_entorno()
    chk("con elección hecha, la semilla NO pisa", config_banco.modo() == "por_hotel")
    os.environ.pop("YVE_BANCO_MODO", None)

    # 7. La semilla siembra cuando NO hay nada.
    _borrar_config()
    os.environ["YVE_BANCO_MODO"] = "por_hotel"
    config_banco.sembrar_desde_entorno()
    chk("sin elección, la semilla siembra 'por_hotel'", config_banco.modo() == "por_hotel")
    os.environ.pop("YVE_BANCO_MODO", None)

    # 8. Semilla con valor inválido: no revienta, no siembra.
    _borrar_config()
    os.environ["YVE_BANCO_MODO"] = "cuenta-unica"
    config_banco.sembrar_desde_entorno()
    chk("semilla inválida se ignora (modo sigue '')", config_banco.modo() == "")
    os.environ.pop("YVE_BANCO_MODO", None)

    # 9. Semilla sin variable: no hace nada.
    _borrar_config()
    config_banco.sembrar_desde_entorno()
    chk("sin la variable, la semilla no hace nada", config_banco.modo() == "")

    shutil.rmtree(base, ignore_errors=True)

    if fallos:
        print("\n FALLOS:")
        for f in fallos:
            print("   ·", f)
        return 1
    print("\n OK · elige, persiste en el servidor, y la semilla nunca pisa\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
