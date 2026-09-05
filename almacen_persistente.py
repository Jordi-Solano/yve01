# -*- coding: utf-8 -*-
"""almacen_persistente.py — que un deploy no borre lo subido (Jordi, sep 2026).

Render (free y de pago) rehace el sistema de ficheros en cada deploy: todo lo
que Yve guarda en `datos-referencia/`, `facturas-procesadas/`, `reportes/`,
`aprobaciones/`, `facturas-entrada/`, `tenants/`… desaparece. Con un disco
persistente montado (Render → Disks, p. ej. en /var/data) y la variable de
entorno `YVE_DATA_DIR=/var/data`, al arrancar:

  1. si el disco esta vacio, se SIEMBRA con lo que trae el repo (proveedores,
     plan de cuentas, usuarios…); si ya tiene datos, solo se copian los
     ficheros de referencia que el repo traiga NUEVOS y en el disco no esten
     (nunca se pisa lo que hay en el disco);
  2. cada carpeta de datos del repo se sustituye por un ENLACE al disco.

Asi ningun modulo cambia: todos siguen abriendo `BASE_DIR/datos-referencia/…`
y en realidad escriben en el disco. Sin `YVE_DATA_DIR` no hace NADA (el
comportamiento de siempre). Se llama desde la primera linea de dashboard.py,
antes de que nadie abra un fichero.
"""
import os
import shutil
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETAS = ["datos-referencia", "facturas-entrada", "facturas-procesadas", "reportes",
            "aprobaciones", "tenants", "ar_real_data"]
MARCA = ".yve_montado"


def _copiar_faltantes(src, dst):
    """Copia a `dst` lo que hay en `src` y en `dst` no existe. No pisa nada. Devuelve cuantos."""
    n = 0
    for raiz, dirs, files in os.walk(src):
        rel = os.path.relpath(raiz, src)
        destino = dst if rel == "." else os.path.join(dst, rel)
        os.makedirs(destino, exist_ok=True)
        for f in files:
            d = os.path.join(destino, f)
            if not os.path.exists(d):
                shutil.copy2(os.path.join(raiz, f), d)
                n += 1
    return n


def montar(base=None, data_dir=None):
    """Enlaza las carpetas de datos de `base` al disco `data_dir`. Devuelve un dict con lo hecho, o None si no hay disco."""
    base = base or BASE_DIR
    data_dir = data_dir if data_dir is not None else os.environ.get("YVE_DATA_DIR", "")
    data_dir = str(data_dir or "").strip()
    if not data_dir:
        return None
    os.makedirs(data_dir, exist_ok=True)
    hecho = {"data_dir": data_dir, "sembradas": [], "enlazadas": [], "copiados": 0}
    for nombre in CARPETAS:
        src = os.path.join(base, nombre)
        dst = os.path.join(data_dir, nombre)
        if os.path.islink(src):
            # ya enlazada (segundo arranque en el mismo contenedor)
            if os.path.realpath(src) == os.path.realpath(dst):
                continue
            os.unlink(src)
        if not os.path.isdir(dst):
            os.makedirs(dst, exist_ok=True)
            if os.path.isdir(src):
                hecho["copiados"] += _copiar_faltantes(src, dst)
            hecho["sembradas"].append(nombre)
        elif os.path.isdir(src):
            hecho["copiados"] += _copiar_faltantes(src, dst)
        if os.path.isdir(src) and not os.path.islink(src):
            shutil.rmtree(src)
        elif os.path.lexists(src):
            os.unlink(src)
        os.symlink(dst, src, target_is_directory=True)
        hecho["enlazadas"].append(nombre)
    try:
        with open(os.path.join(data_dir, MARCA), "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now().isoformat(timespec='seconds')} montado desde {base}\n")
    except Exception:
        pass
    return hecho


def estado(base=None):
    """Para /health y el panel: si los datos viven en un disco persistente."""
    base = base or BASE_DIR
    data_dir = os.environ.get("YVE_DATA_DIR", "").strip()
    enlazadas = [c for c in CARPETAS if os.path.islink(os.path.join(base, c))]
    return {"persistente": bool(data_dir) and len(enlazadas) == len(CARPETAS), "data_dir": data_dir, "enlazadas": enlazadas}
