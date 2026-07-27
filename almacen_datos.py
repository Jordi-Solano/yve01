"""
almacen_datos.py — Yve.01

PUNTO UNICO de lectura consolidada de facturas del tenant.

Hoy los datos viven en ficheros .xlsx partidos POR DIA
(facturas_ap_YYYYMMDD.xlsx, facturas_procesadas_YYYYMMDD.xlsx, ...) y ademas
cada factura aparece repetida en varias ETAPAS del pipeline:

    facturas_procesadas_*  ->  verificacion_*  ->  doble_imposicion_*      (AR)
    facturas_ap_*          ->  facturas_contabilizadas_*                   (AP)

Antes, cada consumidor hacia su propio glob + "coge el mas reciente", con dos
consecuencias: al cambiar de dia los paneles se vaciaban (los datos seguian en
disco, invisibles) y el verificador de comisiones cruzaba solo una fraccion de
las facturas.

Este modulo centraliza TODA la logica de "que ficheros hay, como se juntan y
como se deduplican". El resto del codigo llama a facturas_ap() / facturas_ar()
/ reporte_verificacion() y NO sabe de donde salen los datos.

    ┌─────────────────────────────────────────────────────────────────┐
    │  EL DIA QUE HAYA PERSISTENCIA Y UN ALMACEN POR HOTEL, SOLO HAY  │
    │  QUE CAMBIAR ESTE FICHERO. Se reimplementan las funciones       │
    │  publicas para que lean del almacen y nadie mas se entera.      │
    └─────────────────────────────────────────────────────────────────┘

NO añadir aqui logica de negocio (calculos, filtros por hotel, merges con
aprobaciones): esto es solo la capa de acceso a datos.
"""

import os
import glob

import pandas as pd

# ── Etapas del pipeline, de MAS avanzada a MENOS ──────────────────────────
# El orden importa: al deduplicar gana la etapa mas avanzada, que es la que
# lleva los campos enriquecidos (estado, discrepancia, cuenta contable...).
_ETAPAS_AR = [
    ("doble_imposicion_*.xlsx",    "reportes"),
    ("verificacion_*.xlsx",        "reportes"),
    ("facturas_procesadas_*.xlsx", "procesadas"),
]
_ETAPAS_AP = [
    ("facturas_contabilizadas_*.xlsx", "procesadas"),
    ("facturas_ap_*.xlsx",             "procesadas"),
]

# Campos que identifican un documento. El PRIMERO es obligatorio: si viene
# vacio, la fila NO se deduplica (ver _clave_doc).
_ID_AR = ("numero_factura", "nombre_ota", "periodo_inicio")
_ID_AP = ("numero_factura", "nombre_proveedor")

_VACIOS = ("", "nan", "none", "nat", "<na>", "no_encontrado", "null")


def _txt(v):
    """Normaliza un valor a texto comparable. Devuelve '' si es un vacio.

    A PROPOSITO en Python plano y no con el accesor .str de pandas: en pandas 3
    astype(str) NO convierte los nulos a la cadena 'nan' (los deja como NaN) y
    .str.lower() los propaga, con lo que todas las filas sin dato acabarian
    compartiendo la MISMA clave y drop_duplicates se llevaria por delante
    facturas distintas. Ya nos paso con la tabla de comisiones pactadas.
    """
    s = "" if v is None else str(v)
    s = " ".join(s.split())
    return "" if s.lower() in _VACIOS else s.lower()


def _clave_doc(fila, campos_id):
    """Identidad de un documento, o '' si no se puede establecer.

    Clave vacia = la fila NUNCA se deduplica. Es deliberado: preferimos una
    fila repetida a una factura desaparecida. Fusionar dos facturas sin numero
    porque vengan del mismo fichero seria justo el fallo que estamos quitando.
    """
    principal = _txt(fila.get(campos_id[0]))
    if not principal:
        return ""
    resto = [_txt(fila.get(c)) for c in campos_id[1:]]
    return "|".join([principal] + resto)


def _dirs(procesadas_dir=None, reportes_dir=None):
    """Resuelve los directorios del tenant activo si no se pasan explicitos.

    Acepta cadenas, Path o los envoltorios multi-tenant _TDir/_TFile (que
    resuelven la ruta en cada uso). Se normaliza a str aqui, en el modulo, para
    que ningun consumidor tenga que preocuparse de que tipo esta pasando.
    """
    def _s(v):
        if v is None:
            return None
        return v if isinstance(v, str) else str(os.fspath(v) if hasattr(v, "__fspath__") else v)

    procesadas_dir, reportes_dir = _s(procesadas_dir), _s(reportes_dir)
    if procesadas_dir is None or reportes_dir is None:
        try:
            from tenant_dirs import procesadas_dir as _p, reportes_dir as _r
            procesadas_dir = procesadas_dir or _s(_p())
            reportes_dir = reportes_dir or _s(_r())
        except Exception:
            base = os.path.dirname(os.path.abspath(__file__))
            procesadas_dir = procesadas_dir or os.path.join(base, "facturas-procesadas")
            reportes_dir = reportes_dir or os.path.join(base, "reportes")
    return procesadas_dir, reportes_dir


def _leer_etapas(etapas, procesadas_dir, reportes_dir, hoja=None):
    """Lee TODOS los ficheros de TODAS las etapas y TODOS los dias.

    Devuelve (df_concatenado, rutas). Cada fila lleva '_etapa' con la prioridad
    (0 = mas avanzada) para poder deduplicar quedandose con la mejor version.
    """
    trozos, rutas = [], []
    for prio, (patron, cual) in enumerate(etapas):
        directorio = reportes_dir if cual == "reportes" else procesadas_dir
        for ruta in sorted(glob.glob(os.path.join(directorio, patron))):
            try:
                # Algunos informes tienen varias hojas (Detalle / Resumen)
                df = pd.read_excel(ruta, sheet_name=hoja) if hoja else pd.read_excel(ruta)
            except Exception:
                try:
                    df = pd.read_excel(ruta)
                except Exception as e:
                    print(f"[almacen_datos] no se pudo leer {os.path.basename(ruta)}: {e}")
                    continue
            if df is None or df.empty:
                continue
            df = df.copy()
            df["_etapa"] = prio
            trozos.append(df)
            rutas.append(ruta)
    if not trozos:
        return pd.DataFrame(), []
    return pd.concat(trozos, ignore_index=True), rutas


def _consolidar(df, campos_id):
    """Deduplica quedandose con la version de la etapa mas avanzada.

    Las filas sin identidad ('' como clave) se conservan TODAS.
    """
    if df.empty:
        return df
    claves = [_clave_doc(fila, campos_id) for fila in df.to_dict("records")]
    df = df.assign(_clave=claves)

    sin_id = df[df["_clave"] == ""]
    con_id = df[df["_clave"] != ""]
    if not con_id.empty:
        # estable: ordenar por etapa (0 primero) y quedarse con la primera
        con_id = (con_id.sort_values("_etapa", kind="stable")
                        .drop_duplicates(subset=["_clave"], keep="first"))
    out = pd.concat([con_id, sin_id], ignore_index=True)
    return out.drop(columns=[c for c in ("_clave", "_etapa") if c in out.columns])


# ── API publica ───────────────────────────────────────────────────────────

def facturas_ap(procesadas_dir=None, reportes_dir=None):
    """Todas las facturas AP del tenant, de TODOS los dias, deduplicadas."""
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, rutas = _leer_etapas(_ETAPAS_AP, p, r)
    return _consolidar(df, _ID_AP)


def facturas_ar(procesadas_dir=None, reportes_dir=None):
    """Todas las facturas AR/OTA del tenant, de TODOS los dias, deduplicadas."""
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, rutas = _leer_etapas(_ETAPAS_AR, p, r)
    return _consolidar(df, _ID_AR)


def facturas_ota_para_verificar(procesadas_dir=None):
    """Facturas OTA en crudo (solo facturas_procesadas_*), todos los dias.

    Para el verificador de comisiones: necesita las facturas SIN enriquecer,
    no el informe ya verificado, o se realimentaria a si mismo.
    """
    p, _ = _dirs(procesadas_dir, None)
    df, rutas = _leer_etapas([("facturas_procesadas_*.xlsx", "procesadas")], p, p)
    return _consolidar(df, _ID_AR), rutas


def reporte_verificacion(reportes_dir=None):
    """Informe de verificacion consolidado de TODOS los dias.

    Lo consume el panel de reclamaciones OTA: una reclamacion pendiente de ayer
    tiene que seguir viendose hoy.
    """
    _, r = _dirs(None, reportes_dir)
    df, rutas = _leer_etapas([("verificacion_*.xlsx", "reportes")], r, r, hoja="Detalle")
    return _consolidar(df, _ID_AR)


def resumen_fuentes(procesadas_dir=None, reportes_dir=None):
    """Que ficheros se estan leyendo. Solo para diagnostico."""
    p, r = _dirs(procesadas_dir, reportes_dir)
    _, rap = _leer_etapas(_ETAPAS_AP, p, r)
    _, rar = _leer_etapas(_ETAPAS_AR, p, r)
    return {"ap": [os.path.basename(x) for x in rap],
            "ar": [os.path.basename(x) for x in rar]}
