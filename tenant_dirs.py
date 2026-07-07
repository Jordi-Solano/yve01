"""
tenant_dirs.py — Yve.01
Aislamiento de datos por cliente (multi-tenant).
Cada tenant tiene su propio árbol de datos en tenants/<id>/.
El tenant 'default' usa los directorios raíz de siempre (compatibilidad total).
Resolución del tenant: sesión Flask → variable de entorno YVE_TENANT (para
scripts lanzados por subprocess) → 'default'.
"""
import os, json, re, shutil
from pathlib import Path

BASE_DIR = Path(__file__).parent

SUBDIRS = ["datos-referencia", "reportes", "facturas-entrada", "facturas-procesadas",
           "aprobaciones", "reportes/emails_pendientes", "reportes/emails_pendientes_ap"]

# Plantillas: xlsx/json de referencia que cada tenant nuevo recibe vacíos (solo cabeceras)
_SEED_FILES = ["extracto_banco.xlsx", "ventas_fb_diarias.xlsx", "inventario.xlsx",
               "mermas.xlsx", "recetas.xlsx", "clientes_credito.xlsx",
               "reservas_credito.xlsx", "kpis_hoteles.xlsx", "comisiones_pactadas.xlsx",
               "proveedores.xlsx", "plan_cuentas.xlsx"]


def slug(nombre):
    s = re.sub(r"[^a-z0-9]+", "-", str(nombre).lower()).strip("-")
    return s[:40] or "cliente"


def tenant_id():
    try:
        from flask import session, has_request_context
        if has_request_context():
            t = session.get("tenant_id")
            if t:
                return t
    except Exception:
        pass
    return os.environ.get("YVE_TENANT", "default") or "default"


def tenant_base():
    t = tenant_id()
    if t == "default":
        return BASE_DIR
    base = BASE_DIR / "tenants" / slug(t)
    if not (base / "datos-referencia").exists():
        _crear_tenant(base)
    return base


def _crear_tenant(base):
    for sub in SUBDIRS:
        (base / sub).mkdir(parents=True, exist_ok=True)
    datos = base / "datos-referencia"
    for f in _SEED_FILES:
        src = BASE_DIR / "datos-referencia" / f
        dst = datos / f
        if src.exists() and not dst.exists():
            try:
                import pandas as pd
                pd.read_excel(src).iloc[0:0].to_excel(dst, index=False)
            except Exception:
                shutil.copy2(src, dst)
    for jf, contenido in [("hoteles.json", []), ("archivos_procesados.json", {}),
                          ("notificaciones_historial.json", [])]:
        dst = datos / jf
        if not dst.exists():
            json.dump(contenido, open(dst, "w", encoding="utf-8"), indent=2)


def datos_dir():        return str(tenant_base() / "datos-referencia")
def reportes_dir():     return str(tenant_base() / "reportes")
def entrada_dir():      return str(tenant_base() / "facturas-entrada")
def procesadas_dir():   return str(tenant_base() / "facturas-procesadas")
def aprobaciones_dir(): return str(tenant_base() / "aprobaciones")
