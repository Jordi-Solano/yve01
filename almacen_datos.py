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
import warnings

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

# Albaranes (notas de entrega). El informe del cruce factura-albaran va DELANTE
# y gana, igual que facturas_contabilizadas gana sobre facturas_ap: es el que
# lleva el estado (ALBARAN_FACTURADO / ALBARAN_SIN_FACTURAR). Su hoja se llama
# "Albaranes" a proposito, para que las dos etapas se lean con la misma hoja.
_ETAPAS_ALB = [
    ("matching_albaran_*.xlsx", "reportes"),
    ("albaranes_*.xlsx",        "procesadas"),
]

# Las LINEAS solo viven en albaranes_*: el informe del cruce no las repite. Si
# se reutilizara _ETAPAS_ALB, al no encontrar la hoja "Lineas" en el informe
# _leer_etapas caeria a leer su PRIMERA hoja (Facturas) y colaria facturas en
# la lista de lineas.
_ETAPAS_ALB_LIN = [
    ("albaranes_*.xlsx", "procesadas"),
]

# Lineas de FACTURA (Fase 3c). Mismo razonamiento que _ETAPAS_ALB_LIN: solo
# viven en facturas_ap_*, porque facturas_contabilizadas_* no las repite. Si se
# reutilizara _ETAPAS_AP, al no encontrar la hoja "Lineas" en el informe del
# asignador _leer_etapas caeria a su PRIMERA hoja y colaria facturas enteras en
# la lista de lineas.
_ETAPAS_FAC_LIN = [
    ("facturas_ap_*.xlsx", "procesadas"),
]

# Ordenes de compra (Fase 4a). Hoy solo hay una etapa: lo que entra por el
# pipeline. Cuando exista el informe del cruce PO-factura ira DELANTE y ganara,
# igual que matching_albaran_* gana sobre albaranes_*. Su hoja se llama
# "Ordenes" a proposito para que las dos etapas se lean con la misma hoja.
_ETAPAS_PO = [
    ("ordenes_compra_*.xlsx", "procesadas"),
]

# Las LINEAS del pedido, en su propia lista de etapas por el mismo motivo que
# _ETAPAS_ALB_LIN y _ETAPAS_FAC_LIN.
_ETAPAS_PO_LIN = [
    ("ordenes_compra_*.xlsx", "procesadas"),
]

# Campos que identifican un documento. El PRIMERO es obligatorio: si viene
# vacio, la fila NO se deduplica (ver _clave_doc).
_ID_AR = ("numero_factura", "nombre_ota", "periodo_inicio")
_ID_AP = ("numero_factura", "nombre_proveedor")
_ID_ALB = ("numero_albaran", "nombre_proveedor")
_ID_PO = ("numero_po", "nombre_proveedor")

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


def albaranes(procesadas_dir=None, reportes_dir=None):
    """Cabeceras de TODOS los albaranes del tenant, de TODOS los dias.

    Un albaran es una cabecera con N lineas y se guarda en dos hojas
    (`Albaranes` / `Lineas`) unidas por la columna `clave`. Esto devuelve las
    cabeceras; para las lineas, `lineas_albaran()`.
    """
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, _rutas = _leer_etapas(_ETAPAS_ALB, p, r, hoja="Albaranes")
    return _consolidar(df, _ID_ALB)


def lineas_albaran(procesadas_dir=None, reportes_dir=None):
    """Lineas de TODOS los albaranes, de TODOS los dias.

    NO se deduplican por identidad de documento: un albaran puede repetir el
    mismo producto en dos lineas (dos lotes, dos precios) y fusionarlas seria
    perder mercancia. Se quitan solo los duplicados EXACTOS, que es lo que deja
    reprocesar dos veces el mismo fichero.
    """
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, _rutas = _leer_etapas(_ETAPAS_ALB_LIN, p, r, hoja="Lineas")
    if df.empty:
        return df
    cols = [c for c in df.columns if c != "_etapa"]
    return df.drop_duplicates(subset=cols, keep="last").reset_index(drop=True)


def lineas_factura(procesadas_dir=None, reportes_dir=None):
    """Lineas de TODAS las facturas AP, de TODOS los dias (Fase 3c).

    Igual que `lineas_albaran`: NO se deduplican por identidad de documento —una
    factura puede repetir el mismo producto en dos lineas, dos lotes y dos
    precios, y fusionarlas seria perder mercancia—, solo se quitan los
    duplicados EXACTOS, que es lo que deja reprocesar dos veces el mismo
    fichero.

    Un fichero de facturas sin hoja "Lineas" (los de antes de la Fase 3c, o un
    dia en el que solo entraron facturas de un concepto suelto) no es un error:
    devuelve vacio y el nivel 3 del cruce simplemente no se aplica.
    """
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, _rutas = _leer_etapas(_ETAPAS_FAC_LIN, p, r, hoja="Lineas")
    if df.empty:
        return df
    # OJO (reproducido): si el fichero no tiene hoja "Lineas" —todos los de
    # antes de la Fase 3c, y cualquier dia en el que solo entraran facturas de
    # un concepto suelto— `_leer_etapas` cae a leer su PRIMERA hoja, o sea las
    # FACTURAS, y volvian disfrazadas de linea. `n_linea` es la columna que solo
    # tiene una linea de verdad: sin ella, esto no son lineas.
    if "n_linea" not in df.columns:
        return df.iloc[0:0]
    df = df[df["n_linea"].notna()]
    if df.empty:
        return df
    cols = [c for c in df.columns if c != "_etapa"]
    return df.drop_duplicates(subset=cols, keep="last").reset_index(drop=True)


def ordenes_compra(procesadas_dir=None, reportes_dir=None):
    """Ordenes de compra del tenant, de TODOS los dias, deduplicadas (Fase 4a).

    Punto UNICO de lectura de POs, decision del usuario: el dia que haya
    persistencia solo se toca este fichero. Va por aqui desde el primer dia por
    el mismo motivo que los albaranes: **el pedido es ANTERIOR a la factura**,
    asi que el caso normal sera cruzar una factura de hoy con un pedido de hace
    semanas. Con el "coge el mas reciente" de siempre, la mitad de los cruces no
    encontraria nada.
    """
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, _rutas = _leer_etapas(_ETAPAS_PO, p, r, hoja="Ordenes")
    if df.empty:
        return df
    # Un fichero sin hoja "Ordenes" haria que _leer_etapas cayera a su primera
    # hoja; aqui la primera ES Ordenes, pero se comprueba igual para no repetir
    # el bug que reproduje en lineas_factura.
    if "numero_po" not in df.columns and "clave" not in df.columns:
        return df.iloc[0:0]
    return _consolidar(df, _ID_PO)


def lineas_po(procesadas_dir=None, reportes_dir=None):
    """Lineas de TODAS las ordenes de compra, de todos los dias.

    Igual que `lineas_albaran` y `lineas_factura`: NO se deduplican por identidad
    de documento —un pedido puede repetir el mismo producto en dos lineas— solo
    se quitan los duplicados EXACTOS.

    **El cruce por totales NO usa estas lineas.** Se guardan para el dia que se
    compare articulo por articulo.
    """
    p, r = _dirs(procesadas_dir, reportes_dir)
    df, _rutas = _leer_etapas(_ETAPAS_PO_LIN, p, r, hoja="Lineas")
    if df.empty:
        return df
    # La guarda que hizo falta en lineas_factura: si el fichero no tiene hoja
    # "Lineas", _leer_etapas cae a la PRIMERA hoja y los pedidos volverian
    # disfrazados de linea. `n_linea` es la columna que solo tiene una linea.
    if "n_linea" not in df.columns:
        return df.iloc[0:0]
    df = df[df["n_linea"].notna()]
    if df.empty:
        return df
    cols = [c for c in df.columns if c != "_etapa"]
    return df.drop_duplicates(subset=cols, keep="last").reset_index(drop=True)


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


# ── Banco ─────────────────────────────────────────────────────────────────
# El extracto manda: dice QUE movimientos existen y CUANTOS. El informe de
# conciliacion solo aporta el ESTADO de los que ya se cruzaron (incluidas las
# asignaciones manuales, que son trabajo humano y no se pueden perder).

_ESTADO_DEFECTO = "PENDIENTE"
_CAMPOS_INFORME = ("estado", "factura_ref", "origen", "match_proveedor", "diferencia")


def _fecha(v):
    """Fecha normalizada a 'YYYY-MM-DD'. Cadena vacia si no se puede leer.

    Hace falta porque el mismo movimiento llega como Timestamp desde un fichero
    y como '20/07/2026' desde otro: sin normalizar, no cruzarian nunca.
    """
    if v is None:
        return ""
    try:
        # dayfirst avisa cuando la fecha ya viene en ISO; es esperado y ensucia
        # los logs de produccion, asi que lo callamos solo aqui.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            t = pd.to_datetime(v, dayfirst=True, errors="coerce")
    except Exception:
        t = None
    if t is None or (hasattr(pd, "isna") and pd.isna(t)):
        return _txt(v)
    return t.strftime("%Y-%m-%d")


def _num(v):
    """Importe normalizado a float con 2 decimales, o None si no es un numero.

    Tolera '1.234,56', '1,234.56', '-450 EUR'. Se redondea porque el ida y
    vuelta por Excel puede dejar -450.00000000001 y romperia el cruce.
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            return None if pd.isna(v) else round(float(v), 2)
        except Exception:
            return None
    s = str(v).strip()
    for basura in ("EUR", "eur", "\u20ac", "$", " ", "\u00a0"):
        s = s.replace(basura, "")
    if not s:
        return None
    if "," in s and "." in s:
        # el separador decimal es el que aparece MAS a la derecha
        s = (s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".")
             else s.replace(",", ""))
    elif "," in s:
        ent, _, dec = s.rpartition(",")
        s = (ent + "." + dec) if len(dec) in (1, 2) else s.replace(",", "")
    try:
        return round(float(s), 2)
    except (TypeError, ValueError):
        return None


def _clave_mov(fila):
    """Identidad de un movimiento: fecha + concepto + importe.

    A diferencia de las facturas, aqui la clave NUNCA se deja vacia: un
    movimiento sin concepto sigue siendo un movimiento y tiene que cruzar.
    Dos movimientos identicos comparten clave a proposito; el cruce los trata
    como cola (uno consume una entrada del informe), asi que NO se fusionan.
    """
    imp = _num(fila.get("importe"))
    return "|".join([_fecha(fila.get("fecha")),
                     _txt(fila.get("concepto")),
                     "" if imp is None else f"{imp:.2f}"])


def clave_movimiento(fila):
    """Identidad publica de un movimiento bancario.

    La exponemos porque la pantalla de conciliacion necesita decir "este
    movimiento" al servidor sin usar su POSICION en la lista: la posicion cambia
    en cuanto el extracto crece o se vuelve a bajar con fechas anteriores, y
    escribir por posicion marcaria conciliado un movimiento distinto.
    Misma clave que usa el cruce, para que las dos cosas no se separen nunca.
    """
    return _clave_mov(fila)


def _ultimo_informe(reportes_dir):
    """Ruta del informe de conciliacion mas reciente, o None."""
    hits = glob.glob(os.path.join(reportes_dir, "conciliacion_*.xlsx"))
    if not hits:
        return None
    hits.sort(key=lambda p: (os.path.getmtime(p), p), reverse=True)
    return hits[0]


def movimientos_banco(datos_dir=None, reportes_dir=None):
    """Movimientos del extracto REAL con el estado de conciliacion del informe.

    Devuelve (df, info). El df lleva las columnas del extracto mas
    estado/factura_ref/origen/match_proveedor/diferencia. Todo movimiento que
    no aparezca en el informe queda PENDIENTE.

    Antes el panel leia SOLO el informe: al subir movimientos nuevos seguia
    enseñando la foto del dia que se concilio y no habia forma de ver lo nuevo.
    Leer solo el extracto habria sido peor: se perderia lo ya conciliado, y con
    ello las asignaciones manuales.
    """
    from collections import defaultdict, deque

    p, r = _dirs(datos_dir, reportes_dir)
    datos_dir = datos_dir if datos_dir is not None else _dirs(None, None)[0]
    # el extracto vive en datos-referencia, no en facturas-procesadas
    try:
        from tenant_dirs import datos_dir as _dd
        ruta_ext = os.path.join(str(os.fspath(_dd()) if hasattr(_dd(), "__fspath__") else _dd()),
                                "extracto_banco.xlsx")
    except Exception:
        ruta_ext = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "datos-referencia", "extracto_banco.xlsx")

    info = {"extracto": None, "informe": None, "movimientos_extracto": 0,
            "filas_informe": 0, "cruzados": 0}

    df_ext = pd.DataFrame()
    if os.path.exists(ruta_ext):
        try:
            df_ext = pd.read_excel(ruta_ext)
            info["extracto"] = os.path.basename(ruta_ext)
        except Exception as e:
            print(f"[almacen_datos] no se pudo leer el extracto: {e}")

    ruta_rep = _ultimo_informe(r)
    df_rep = pd.DataFrame()
    if ruta_rep:
        try:
            df_rep = pd.read_excel(ruta_rep)
            info["informe"] = os.path.basename(ruta_rep)
        except Exception as e:
            print(f"[almacen_datos] no se pudo leer {os.path.basename(ruta_rep)}: {e}")

    info["movimientos_extracto"] = len(df_ext)
    info["filas_informe"] = len(df_rep)

    # Sin extracto no hay nada que enseñar salvo el propio informe.
    if df_ext.empty:
        if not df_rep.empty:
            for c in _CAMPOS_INFORME:
                if c not in df_rep.columns:
                    df_rep[c] = "" if c != "diferencia" else 0.0
            info["cruzados"] = len(df_rep)
        return df_rep, info

    # Cola por clave: dos movimientos iguales consumen DOS filas del informe.
    cola = defaultdict(deque)
    for fila in df_rep.to_dict("records"):
        cola[_clave_mov(fila)].append(fila)

    filas = []
    for mov in df_ext.to_dict("records"):
        fila = dict(mov)
        q = cola.get(_clave_mov(mov))
        origen_info = q.popleft() if q else None
        if origen_info is not None:
            info["cruzados"] += 1
            for c in _CAMPOS_INFORME:
                fila[c] = origen_info.get(c, "")
            if not _txt(fila.get("estado")):
                fila["estado"] = _ESTADO_DEFECTO
        else:
            fila["estado"] = _ESTADO_DEFECTO
            for c in _CAMPOS_INFORME[1:]:
                fila[c] = 0.0 if c == "diferencia" else ""
        filas.append(fila)

    return pd.DataFrame(filas), info


def resumen_fuentes(procesadas_dir=None, reportes_dir=None):
    """Que ficheros se estan leyendo. Solo para diagnostico."""
    p, r = _dirs(procesadas_dir, reportes_dir)
    _, rap = _leer_etapas(_ETAPAS_AP, p, r)
    _, rar = _leer_etapas(_ETAPAS_AR, p, r)
    return {"ap": [os.path.basename(x) for x in rap],
            "ar": [os.path.basename(x) for x in rar]}
