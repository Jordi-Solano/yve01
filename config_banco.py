#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cómo funciona el banco de la empresa: una cuenta del GRUPO o una POR HOTEL.

Por qué existe
--------------
Un extracto bancario no dice a qué hotel es cada movimiento. Para una empresa
con UNA cuenta de sociedad, repartirlo por hotel sería inventar; para una con
una cuenta POR hotel, mostrarlo junto es una fuga. Las dos son ciertas según la
empresa, así que NO lo decidimos nosotros: lo elige el usuario la primera vez
que abre el banco, y esa elección gobierna cómo se muestra y se filtra a partir
de ahí.

  'grupo'     -> el banco se muestra junto (etiquetado "del grupo"); que aparezca
                 en cualquier hotel es correcto, no fuga.
  'por_hotel' -> el extracto se etiqueta con el hotel activo al subirlo (como
                 AP/AR) y se filtra por hotel; lo no etiquetado queda "sin
                 asignar", visible, no escondido.

Dónde vive (y por qué así)
--------------------------
`config_banco.json` en el árbol del tenant (`datos-referencia/`), junto a
`hoteles.json`. Es del SERVIDOR, no del navegador: así funciona igual desde el
móvil y desde el PC (una decisión de empresa no puede depender de un navegador).

Render (plan gratuito) borra el disco en cada despliegue, así que —igual que el
censo de hoteles— se puede SEMBRAR desde una variable de entorno,
`YVE_BANCO_MODO` (`grupo` | `por_hotel`), que sobrevive a los despliegues. Se
siembra solo si aún no hay elección; nunca pisa una ya hecha.
"""
import json
import os

MODOS = ("grupo", "por_hotel")


def _ruta():
    from tenant_dirs import datos_dir
    return os.path.join(datos_dir(), "config_banco.json")


def modo():
    """El modo del banco del tenant: 'grupo', 'por_hotel', o '' si no se ha elegido."""
    try:
        d = json.load(open(_ruta(), encoding="utf-8"))
    except Exception:
        return ""
    m = str((d or {}).get("modo", "")).strip()
    return m if m in MODOS else ""


def elegido():
    """¿Ya ha elegido el usuario cómo funciona su banco?"""
    return modo() != ""


def por_hotel():
    """Atajo: ¿el banco se separa por hotel? (False también cuando no hay elección)."""
    return modo() == "por_hotel"


def elegir(m):
    """Guarda el modo. Devuelve el modo guardado, o '' si el valor no es válido
    o no se pudo escribir. Solo acepta 'grupo' / 'por_hotel'."""
    m = str(m or "").strip().lower()
    if m not in MODOS:
        return ""
    try:
        from tenant_dirs import datos_dir
        os.makedirs(datos_dir(), exist_ok=True)
        tmp = _ruta() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"modo": m}, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, _ruta())
        return m
    except Exception as e:
        print(f"[config_banco] no se pudo guardar la elección: {e}")
        return ""


def sembrar_desde_entorno():
    """Recrea la config desde `YVE_BANCO_MODO` si aún no hay elección. Apaño igual
    que el de los hoteles: el disco de Render es efímero, la variable de entorno
    no, así que la decisión de empresa sobrevive a los despliegues.

    Las dos propiedades que lo hacen seguro (idénticas al censo):
      1. Solo siembra si NO hay elección. Con una hecha, no toca nada.
      2. Si algo falla (valor mal, disco de solo lectura), avisa y sigue: un apaño
         de conveniencia no puede tumbar el arranque.
    """
    if elegido():
        return None
    crudo = os.environ.get("YVE_BANCO_MODO", "").strip().lower()
    if not crudo:
        return None
    if crudo not in MODOS:
        print(f"[config_banco] YVE_BANCO_MODO='{crudo}' no válido (grupo|por_hotel); se ignora")
        return None
    r = elegir(crudo)
    if r:
        print(f"[config_banco] sembrado modo '{r}' desde YVE_BANCO_MODO")
    return r
