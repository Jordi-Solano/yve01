# -*- coding: utf-8 -*-
"""contratos_grupo.py — Yve.01 · el registro de los contratos de grupo de AR (b87).

Respuestas de finanzas (Jordi, 24 sep 2026) que manda este modulo:

  1. COMISION DE AGENCIA. La dice el contrato, en su campo de comisiones:
       - "tarifa neta": el hotel factura el NETO y NO hay factura de comision;
       - un % de comision: el hotel factura el BRUTO y la agencia manda su
         factura de comision APARTE (Yve no crea ninguna: guarda la comision
         ESPERADA y, cuando llega la de la agencia, la une al contrato — b88).
     Si el contrato no lo dice, se PREGUNTA al usuario: no se supone nada.
  2. QUIEN PAGA la factura del grupo (agencia o cliente final) decide si se
     puede compensar la comision contra esa factura (b89). Se lee del contrato;
     si no lo dice, tambien se pregunta. Sin agencia, paga el cliente: no hay
     nadie mas.

El registro vive en datos-referencia/contratos_grupo.json (una entrada por
contrato y hotel). Lo que decide una persona (origen "usuario") NO lo pisa un
reproceso del mismo contrato. Funciones puras donde se puede (entra el dict del
contrato, sale el dato), para probarlas sin Flask; las que tocan disco llevan
`datos_dir` para los tests.
"""
import json
import os
import re
import unicodedata
from datetime import datetime

FICHERO = "contratos_grupo.json"
MODOS = ("neta", "porcentaje")
PAGADORES = ("agencia", "cliente")
CONCEPTOS = ("alojamiento", "fb", "salas")

_NETA = ("neta", "netas", "neto", "netos", "tarifa neta", "tarifas netas", "net", "net rate", "net rates",
         "nett", "sin comision", "no comisionable", "no comisionables")
_PORC = ("porcentaje", "%", "comisionable", "comisionables", "comision", "commission", "percentage", "commissionable")
_P_AGENCIA = ("agencia", "agency", "la agencia", "agencia de viajes", "dmc", "opc", "intermediario")
_P_CLIENTE = ("cliente", "client", "customer", "empresa", "organizador", "cliente final", "end client")


# ── utilidades ───────────────────────────────────────────────────────────────
def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "nat", "no_encontrado") else s


def _norm(v):
    s = unicodedata.normalize("NFKD", _txt(v)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9%]+", " ", s)
    return " ".join(s.split())


def _f(v, defecto=0.0):
    if v is None or v == "" or (isinstance(v, float) and v != v):
        return defecto
    if isinstance(v, (int, float)):
        return float(v)
    s = _txt(v).replace("€", "").replace("EUR", "").replace("%", "").replace(" ", "")
    if not s:
        return defecto
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return defecto


def _r(x):
    return round(float(x or 0), 2)


def _dd(datos_dir=None):
    if datos_dir:
        return str(datos_dir)
    try:
        from tenant_dirs import datos_dir as _d
        return str(_d())
    except Exception:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos-referencia")


def ruta(datos_dir=None):
    return os.path.join(_dd(datos_dir), FICHERO)


def leer(datos_dir=None):
    try:
        with open(ruta(datos_dir), encoding="utf-8") as fh:
            d = json.load(fh)
        return [c for c in d if isinstance(c, dict)] if isinstance(d, list) else []
    except Exception:
        return []


def _escribir(lista, datos_dir=None):
    p = ruta(datos_dir)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(lista, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ── lo que dice el contrato ──────────────────────────────────────────────────
def _pcts(datos):
    com = datos.get("comisiones") or {}
    return {"alojamiento": _f(com.get("alojamiento_pct")), "fb": _f(com.get("fb_pct")), "salas": _f(com.get("salas_pct"))}


def modo_comision(datos):
    """{'modo', 'texto', 'origen'} leyendo el campo de comisiones del contrato.

    modo: 'neta' | 'porcentaje' | 'sin_agencia' | '' (el contrato no lo dice:
    hay que preguntarlo). Un porcentaje leido del campo ES el contrato diciendo
    que hay comision; un texto que habla de tarifa neta, que no la hay.
    """
    ag = _txt((datos.get("agencia") or {}).get("nombre"))
    com = datos.get("comisiones") or {}
    texto = _txt(com.get("texto"))[:300]
    if not ag:
        return {"modo": "sin_agencia", "texto": texto, "origen": "contrato"}
    m = _norm(com.get("modo"))
    if m in _NETA:
        return {"modo": "neta", "texto": texto, "origen": "contrato"}
    if m in _PORC or sum(_pcts(datos).values()) > 0:
        return {"modo": "porcentaje", "texto": texto, "origen": "contrato"}
    t = _norm(texto)
    if t and (re.search(r"\b(tarifas? netas?|net rates?|nett?|netos?)\b", t) or "sin comision" in t):
        return {"modo": "neta", "texto": texto, "origen": "contrato"}
    if t and re.search(r"\d+\s*%|% ?\d+", t):
        return {"modo": "porcentaje", "texto": texto, "origen": "contrato"}
    return {"modo": "", "texto": texto, "origen": ""}


def pagador_de(datos):
    """{'quien', 'texto', 'origen'}: quien paga la factura del grupo.

    quien: 'agencia' | 'cliente' | '' (no lo dice: se pregunta). Sin agencia en
    el contrato solo puede pagar el cliente."""
    ag = _txt((datos.get("agencia") or {}).get("nombre"))
    fac = datos.get("facturacion") or {}
    texto = _txt(fac.get("texto"))[:300]
    if not ag:
        return {"quien": "cliente", "texto": texto, "origen": "contrato"}
    p = _norm(fac.get("pagador"))
    if p in _P_AGENCIA:
        return {"quien": "agencia", "texto": texto, "origen": "contrato"}
    if p in _P_CLIENTE:
        return {"quien": "cliente", "texto": texto, "origen": "contrato"}
    return {"quien": "", "texto": texto, "origen": ""}


def bases_contrato(datos):
    """Bases SIN IVA por concepto (mismo calculo que la comision de siempre)."""
    aloj = datos.get("alojamiento") or {}
    fb = datos.get("fb") or {}
    salas = datos.get("salas") or {}

    def base(total, iva):
        total = _f(total); iva = _f(iva, 10)
        return total / (1 + iva / 100) if total else 0.0
    return {"alojamiento": _r(base(aloj.get("total_habitaciones"), aloj.get("iva_pct", 10))),
            "fb": _r(base(fb.get("total"), fb.get("iva_pct", 10))),
            "salas": _r(base(salas.get("total"), 21))}


def comision_esperada(c):
    """Comision (base sin IVA) que la agencia deberia facturar. 0 si no hay
    comision (neta / sin agencia); None si todavia no se sabe (pendiente)."""
    modo = (c.get("comision") or {}).get("modo", "")
    if modo in ("neta", "sin_agencia"):
        return 0.0
    if modo != "porcentaje":
        return None
    pct = c.get("pct") or {}
    if sum(_f(pct.get(k)) for k in CONCEPTOS) <= 0:
        return None
    b = c.get("bases") or {}
    return _r(sum(_f(b.get(k)) * _f(pct.get(k)) / 100 for k in CONCEPTOS))


def pendientes(c):
    """Lo que falta que decida una persona: 'modo' (neta o %), 'pct' (hay
    comision pero no se leyo el %), 'pagador'."""
    out = []
    modo = (c.get("comision") or {}).get("modo", "")
    if modo not in ("neta", "porcentaje", "sin_agencia"):
        out.append("modo")
    elif modo == "porcentaje" and sum(_f((c.get("pct") or {}).get(k)) for k in CONCEPTOS) <= 0:
        out.append("pct")
    if (c.get("pagador") or {}).get("quien") not in PAGADORES:
        out.append("pagador")
    return out


def deudor(c):
    """A nombre de quien va la factura del grupo (el cliente AR). '' si aun no
    se sabe quien paga."""
    q = (c.get("pagador") or {}).get("quien")
    if q == "agencia":
        return _txt(c.get("agencia"))
    if q == "cliente":
        return _txt(c.get("cliente"))
    return ""


def clave_contrato(numero, evento, fecha_entrada, hotel_id):
    base = _norm(numero) or (_norm(evento) + " " + _norm(fecha_entrada)).strip() or "sin nombre"
    return "CG|" + base + "|" + _txt(hotel_id)


# ── el registro ──────────────────────────────────────────────────────────────
def registrar(datos, transformado=None, hotel_id=None, archivo="", datos_dir=None):
    """Da de alta (o actualiza) el contrato en el registro. Un reproceso del mismo
    contrato refresca lo leido pero conserva lo que decidio una persona."""
    if hotel_id is None:
        try:
            import censo_hoteles as _censo
            hotel_id = _censo.para_guardar()
        except Exception:
            hotel_id = os.environ.get("YVE_HOTEL", "")
    ev = datos.get("evento") or {}
    cli = datos.get("cliente") or {}
    ag = datos.get("agencia") or {}
    aloj = datos.get("alojamiento") or {}
    numero = _txt(datos.get("contrato_numero"))
    evento = (_txt(ev.get("id")) + " " + _txt(ev.get("nombre"))).strip() or "Evento de grupo"
    cid = clave_contrato(numero, evento, aloj.get("fecha_entrada"), hotel_id)
    res = (transformado or {}).get("resumen") or {}
    nuevo = {
        "id": cid, "contrato": numero, "evento": evento,
        "cliente": _txt(cli.get("nombre")) or "Cliente grupo", "cliente_nif": _txt(cli.get("cif")),
        "cliente_pais": _txt(cli.get("pais")), "cliente_email": _txt(cli.get("email")),
        "agencia": _txt(ag.get("nombre")), "agencia_nif": _txt(ag.get("cif")), "agencia_email": _txt(ag.get("email")),
        "hotel_id": _txt(hotel_id), "archivo": _txt(archivo),
        "fecha_contrato": _txt(datos.get("fecha_contrato")),
        "fecha_entrada": _txt(aloj.get("fecha_entrada")), "fecha_salida": _txt(aloj.get("fecha_salida")),
        "importes": {"habitaciones": _r(res.get("habitaciones")), "fb": _r(res.get("fb")),
                     "salas": _r(res.get("salas")), "total": _r(res.get("total_receivable"))},
        "bases": bases_contrato(datos),
        "factura_grupo": _txt(res.get("numero")),
        "comision": modo_comision(datos),
        "pct": _pcts(datos),
        "pagador": pagador_de(datos),
        "datos_contrato": {k: v for k, v in datos.items() if not str(k).startswith("_")},
        "leido": _ahora(),
    }
    if nuevo["comision"]["modo"] in ("neta", "sin_agencia"):
        nuevo["pct"] = {k: 0.0 for k in CONCEPTOS}
    lista = leer(datos_dir)
    previo = next((c for c in lista if c.get("id") == cid), None)
    if previo:
        # lo que decidio una persona manda sobre una nueva lectura del papel
        for clave in ("comision", "pagador"):
            if (previo.get(clave) or {}).get("origen") == "usuario":
                nuevo[clave] = previo[clave]
                if clave == "comision":
                    nuevo["pct"] = previo.get("pct") or nuevo["pct"]
        for clave in ("factura_comision", "creado", "historial", "beos"):
            if previo.get(clave) is not None:
                nuevo[clave] = previo[clave]
        lista = [c for c in lista if c.get("id") != cid]
    nuevo.setdefault("creado", _ahora())
    nuevo.setdefault("historial", [])
    lista.append(nuevo)
    _escribir(lista, datos_dir)
    return nuevo


def buscar(cid, datos_dir=None):
    return next((c for c in leer(datos_dir) if c.get("id") == cid), None)


def del_hotel(lista, hotel):
    """Igualdad estricta, como el resto de AR: con hotel elegido, solo lo suyo;
    sin hotel (vista de grupo o 0 hoteles), todo."""
    if not hotel:
        return list(lista)
    return [c for c in lista if _txt(c.get("hotel_id")) == _txt(hotel)]


def decidir(cid, modo=None, pct=None, pagador=None, usuario="", datos_dir=None):
    """Lo que decide una persona cuando el contrato no lo dice (o para corregir
    una lectura). Lanza ValueError con el motivo si algo no vale."""
    lista = leer(datos_dir)
    c = next((x for x in lista if x.get("id") == cid), None)
    if c is None:
        raise KeyError("contrato no encontrado")
    hist = c.get("historial") or []
    if modo is not None:
        modo = _norm(modo)
        if modo not in MODOS:
            raise ValueError("el modo de comision es 'neta' o 'porcentaje'")
        if not _txt(c.get("agencia")):
            raise ValueError("este contrato no tiene agencia: no hay comision que decidir")
        if modo == "neta":
            nuevo_pct = {k: 0.0 for k in CONCEPTOS}
        else:
            if isinstance(pct, (int, float, str)) and not isinstance(pct, dict):
                pct = {"alojamiento": pct}
            pct = pct or {}
            nuevo_pct = {k: _f(pct.get(k)) for k in CONCEPTOS}
            if any(v < 0 or v > 100 for v in nuevo_pct.values()):
                raise ValueError("el % de comision tiene que estar entre 0 y 100")
            if sum(nuevo_pct.values()) <= 0:
                raise ValueError("con comision hay que decir el % (al menos sobre el alojamiento)")
        c["comision"] = {"modo": modo, "texto": (c.get("comision") or {}).get("texto", ""),
                         "origen": "usuario", "por": usuario, "cuando": _ahora()}
        c["pct"] = nuevo_pct
        hist.append({"fecha": _ahora(), "usuario": usuario, "campo": "comision",
                     "valor": "tarifa neta" if modo == "neta" else ", ".join(f"{k} {v:g}%" for k, v in nuevo_pct.items() if v)})
    if pagador is not None:
        pagador = _norm(pagador)
        if pagador not in PAGADORES:
            raise ValueError("quien paga es 'agencia' o 'cliente'")
        if pagador == "agencia" and not _txt(c.get("agencia")):
            raise ValueError("este contrato no tiene agencia")
        c["pagador"] = {"quien": pagador, "texto": (c.get("pagador") or {}).get("texto", ""),
                        "origen": "usuario", "por": usuario, "cuando": _ahora()}
        hist.append({"fecha": _ahora(), "usuario": usuario, "campo": "pagador", "valor": pagador})
    c["historial"] = hist[-30:]
    _escribir(lista, datos_dir)
    return c


def vista(c):
    """Lo que ve la pantalla de un contrato (sin los datos crudos)."""
    out = {k: v for k, v in c.items() if k not in ("datos_contrato",)}
    out["esperada"] = comision_esperada(c)
    out["pendientes"] = pendientes(c)
    out["deudor"] = deudor(c)
    return out


def sincronizar_factura(c, datos_dir=None):
    """Lleva a la factura del grupo (reservas_credito.xlsx) lo que se sabe del
    contrato —a nombre de quien paga, el modo de comision y la comision
    esperada— y da de alta la ficha AR de quien paga, SIN credito. No toca el
    estado de la factura (emitida, cobrada) ni sus fechas. Devuelve True si
    encontro la factura."""
    import pandas as pd
    num = _txt(c.get("factura_grupo"))
    dd = _dd(datos_dir)
    q = (c.get("pagador") or {}).get("quien")
    d = deudor(c)
    if q in PAGADORES and d:
        try:
            from tab_ar_real import alta_cliente_pendiente
            nif, mail = (c.get("agencia_nif"), c.get("agencia_email")) if q == "agencia" else (c.get("cliente_nif"), c.get("cliente_email"))
            alta_cliente_pendiente(d, _txt(nif), f"contrato {_txt(c.get('contrato')) or num}".strip(),
                                   datos_dir=dd, hotel_id=_txt(c.get("hotel_id")), email=_txt(mail))
        except Exception:
            pass
    pr = os.path.join(dd, "reservas_credito.xlsx")
    if not num or not os.path.exists(pr):
        return False
    try:
        df = pd.read_excel(pr)
    except Exception:
        return False
    col = "numero_reserva" if "numero_reserva" in df.columns else ("numero" if "numero" in df.columns else None)
    if not col:
        return False
    m = df[col].map(_txt) == num
    if _txt(c.get("hotel_id")) and "hotel_id" in df.columns:
        m = m & (df["hotel_id"].map(_txt) == _txt(c.get("hotel_id")))
    if not m.any():
        return False
    esp = comision_esperada(c)
    valores = {"cliente": d or _txt(c.get("cliente")), "pagador": q or "",
               "modo_comision": (c.get("comision") or {}).get("modo", ""),
               "comision_total": esp if esp is not None else 0.0,
               "agencia": _txt(c.get("agencia")), "cliente_final": _txt(c.get("cliente"))}
    for k, v in valores.items():
        if k not in df.columns:
            df[k] = ""
        df[k] = df[k].astype(object)
        df.loc[m, k] = v
    tmp = pr + ".tmp.xlsx"
    df.to_excel(tmp, index=False)
    os.replace(tmp, pr)
    return True
