# -*- coding: utf-8 -*-
"""cuentas_proveedor.py — UNA cuenta y UN tipo por proveedor.

Problema que arregla (ronda de pruebas de Jordi, fase 1): dos facturas del
mismo distribuidor de alimentacion salian una a 600 y otra a 629, y una
lavanderia caia en 600 porque la IA escribio "restaurant" en el concepto. La
cuenta la decidia una palabra suelta del concepto de CADA factura.

Reglas, en este orden:
  1. proveedores.xlsx (el maestro del cliente) manda si conoce al proveedor.
  2. Lo ya decidido para ese proveedor (proveedores_aprendidos.json): mismo
     proveedor → misma cuenta, siempre.
  3. Evidencia de mercancia F&B en las LINEAS, el concepto o el nombre
     (pollo, merluza, vino, distribucion alimentaria...) → tipo FB, cuenta 600.
  4. Servicios y suministros por palabras clave → 621/622/623/624/625/627/629.
     **Un servicio nunca va a 600.**
  5. Lo que trajera la factura si es una cuenta de servicio; si no, 629.
La decision se guarda en `datos-referencia/proveedores_aprendidos.json` para
que la siguiente factura del mismo proveedor salga igual. El fichero se puede
editar a mano (o borrar una entrada para que se vuelva a decidir).

`exige_albaran(fila)`: solo la mercancia (tipo FB o cuenta 60x) lleva albaran
de entrega. Un suministro o un servicio no tiene albaran y no se reclama por
"factura sin albaran".
"""
import json
import os
import re
import unicodedata

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FICHERO = "proveedores_aprendidos.json"
NF = "NO_ENCONTRADO"

# mercancia F&B (castellano, catalan, ingles): basta una palabra en una linea
FB_KW = ["pollo", "pollastre", "carne", "carn ", "ternera", "vedella", "cerdo", "porc", "cordero", "jamon", "jamón",
         "embutid", "pescado", "peix", "merluza", "lluç", "bacalao", "salmon", "salmón", "gamba", "marisco", "atun",
         "verdura", "hortaliza", "patata", "tomate", "tomàquet", "cebolla", "ceba", "lechuga", "fruta", "fruita",
         "limon", "limón", "naranja", "manzana", "platano", "plátano", "huevo", "ou ", "ous ", "leche", "llet",
         "queso", "formatge", "mantequilla", "yogur", "pan ", "panader", "harina", "arroz", "arròs", "pasta",
         "aceite", "oli ", "sal ", "azucar", "azúcar", "cafe", "café", "te ", "vino", "vi ", "cava", "cerveza",
         "cervesa", "refresco", "agua mineral", "zumo", "suc ", "licor", "ginebra", "gin ", "ron ", "whisky",
         "vodka", "bebida", "beguda", "aliment", "comida", "food", "beverage", "catering", "congelad", "conserva",
         "distribucion", "distribució", "mercabarna", "makro", "frutas", "fruites", "carnes", "carns", "pescados",
         "peixos", "vins", "vinos", "bodega", "celler"]
# servicios / suministros → cuenta. Primera coincidencia gana (orden = prioridad)
SERVICIO_KW = [
    ("625", ["seguro", "assegur", "insurance", "prima ", "poliza", "pòlissa", "mapfre", "axa", "allianz", "zurich"]),
    ("621", ["alquiler", "lloguer", "arrendamiento", "renting", "leasing", "rent "]),
    ("622", ["mantenimiento", "manteniment", "maintenance", "reparac", "repair", "climatiz", "instal·lac", "instalac", "instal.lac",
             "fontaner", "lampister", "electricista", "ascensor", "elevador", "otis", "schindler", "kone", "pintura", "obra "]),
    ("623", ["seguretat", "seguridad", "vigilan", "securitas", "prosegur", "alarma", "consultor", "asesor", "assessor",
             "abogado", "advocat", "legal", "audit", "gestoria", "notari"]),
    ("624", ["transport", "mensajer", "missatger", "courier", "envio", "envío", "shipping", "logist"]),
    ("627", ["publicidad", "publicitat", "marketing", "advertising", "diseño", "disseny", "promo"]),
    ("629", ["bugaderia", "lavander", "laundry", "roba ", "tovall", "toalla", "sabana", "llençol", "limpieza", "neteja",
             "neteges", "cleaning", "housekeeping", "energia", "energía", "electric", "endesa", "iberdrola", "naturgy",
             "gas ", "agua", "aigua", "aigües", "suministro", "subministr", "kwh", "telefon", "telèfon", "vodafone",
             "orange", "movistar", "internet", "wifi", "fibra", "software", "licencia", "llicència", "suscrip",
             "hosting", "informatic", "informàtic"]),
]
CUENTAS_SERVICIO = {"621", "622", "623", "624", "625", "626", "627", "628", "629"}


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).lower()
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()


def clave_proveedor(nombre):
    """'Distribucions Garraf, S.L.' y 'DISTRIBUCIONS GARRAF SL' son el mismo."""
    n = re.sub(r"\s+", " ", _norm(nombre).replace(",", " ").replace(".", " ")).strip()
    n = re.sub(r"\b(s l(?: u)?|s a(?: u)?|sccl|scp|cb|ltd|gmbh|sl|sa|slu|sau)\b", " ", n)
    return re.sub(r"\s+", " ", n).strip()


def _ruta(datos_dir=None):
    return os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), FICHERO)


def aprendidos(datos_dir=None):
    try:
        with open(_ruta(datos_dir), encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def guardar_aprendidos(d, datos_dir=None):
    ruta = _ruta(datos_dir)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)


def _texto_lineas(lineas):
    out = []
    for l in lineas or []:
        if isinstance(l, dict):
            out.append(_norm(l.get("descripcion") or l.get("concepto") or ""))
    return " | ".join(out)


def _hit(txt, k):
    """Palabra completa si la clave acaba en espacio ('vi ', 'sal '); si no, prefijo
    de palabra ('aliment' casa con alimentacion, no 'ali' con 'cali')."""
    entera = k.endswith(" ")
    k = _norm(k)
    if entera:
        return re.search(r"\b" + re.escape(k) + r"\b", txt) is not None
    return re.search(r"\b" + re.escape(k), txt) is not None


def evidencia_fb(nombre, concepto, lineas):
    """True si las lineas, el concepto o el nombre hablan de mercancia F&B."""
    txt = " ".join((_norm(nombre), _norm(concepto), _texto_lineas(lineas)))
    return any(_hit(txt, k) for k in FB_KW)


def cuenta_servicio(nombre, concepto, lineas):
    txt = " ".join((_norm(nombre), _norm(concepto), _texto_lineas(lineas)))
    for cuenta, kws in SERVICIO_KW:
        if any(_hit(txt, k) for k in kws):
            return cuenta
    return ""


def _maestro(nombre):
    """(tipo, cuenta) de proveedores.xlsx via el propio lector; ('OTRAS', NF) si no lo conoce."""
    try:
        from lector_facturas_ap import clasificar_proveedor, cargar_proveedores
        return clasificar_proveedor(nombre, cargar_proveedores())
    except Exception:
        return "OTRAS", NF


def decidir(fila, lineas=None, apr=None, maestro=None):
    """Devuelve (tipo, cuenta, origen). No escribe nada."""
    nombre = str(fila.get("nombre_proveedor") or "")
    concepto = str(fila.get("descripcion_concepto") or "")
    lineas = lineas if lineas is not None else fila.get("_lineas") or []
    tipo_m, cta_m = maestro if maestro is not None else _maestro(nombre)
    if str(cta_m) not in ("", NF, "nan", "None") or str(tipo_m).upper() == "FB":
        cta = str(cta_m) if str(cta_m) not in ("", NF, "nan", "None") else "600"
        return (str(tipo_m).upper() or "OTRAS"), cta, "proveedores.xlsx"
    k = clave_proveedor(nombre)
    if k and apr and k in apr:
        a = apr[k]
        return str(a.get("tipo") or "OTRAS"), str(a.get("cuenta") or "629"), "aprendido"
    if evidencia_fb(nombre, concepto, lineas):
        return "FB", "600", "mercancia F&B en lineas/concepto"
    cs = cuenta_servicio(nombre, concepto, lineas)
    if cs:
        return "OTRAS", cs, "palabra clave de servicio"
    traia = str(fila.get("cuenta_contable") or "").strip()
    if traia in CUENTAS_SERVICIO:
        return "OTRAS", traia, "cuenta de la factura"
    return "OTRAS", "629", "sin pista: otros servicios"


def normalizar(filas, datos_dir=None):
    """Aplica `decidir` a cada factura (dicts con `_lineas`), en sitio, y aprende.

    Dentro del mismo lote, dos facturas del mismo proveedor salen iguales: la
    primera decide y la segunda hereda. Devuelve el numero de filas cambiadas."""
    apr = aprendidos(datos_dir)
    cambios = 0
    nuevos = False
    for f in filas:
        if not isinstance(f, dict):
            continue
        nombre = str(f.get("nombre_proveedor") or "")
        if not nombre or nombre == NF:
            continue
        tipo, cuenta, origen = decidir(f, f.get("_lineas") or [], apr)
        if str(f.get("tipo_proveedor") or "").upper() != tipo or str(f.get("cuenta_contable") or "") != cuenta:
            cambios += 1
        f["tipo_proveedor"] = tipo
        f["cuenta_contable"] = cuenta
        f["origen_cuenta"] = origen
        k = clave_proveedor(nombre)
        if k and k not in apr and origen != "proveedores.xlsx":
            apr[k] = {"proveedor": nombre, "tipo": tipo, "cuenta": cuenta, "origen": origen,
                      "primera_factura": str(f.get("numero_factura") or "")}
            nuevos = True
    if nuevos:
        try:
            guardar_aprendidos(apr, datos_dir)
        except Exception:
            pass
    return cambios


def exige_albaran(fila):
    """Solo la mercancia lleva albaran: tipo FB o cuenta de compras (60x)."""
    tipo = str(fila.get("tipo_proveedor") or "").upper()
    cta = str(fila.get("cuenta_debe_gasto") or fila.get("cuenta_contable") or "").strip()
    if cta.endswith(".0"):
        cta = cta[:-2]
    return tipo == "FB" or cta.startswith("60")
