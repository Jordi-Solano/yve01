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
import candados as _cand

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
    _cand.escribir_json(lista, ruta(datos_dir))      # b99: atomica, con temporal propio


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
@_cand.protegido(_dd)       # b99: leer-cambiar-escribir sin que otra peticion se cuele
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
        "evento_id": _txt(ev.get("id")), "evento_nombre": _txt(ev.get("nombre")),
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
        for clave in ("factura_comision", "creado", "historial", "beos", "factura_numero"):
            if previo.get(clave) is not None:
                nuevo[clave] = previo[clave]
        lista = [c for c in lista if c.get("id") != cid]
    nuevo.setdefault("creado", _ahora())
    nuevo.setdefault("historial", [])
    lista.append(nuevo)
    _escribir(lista, datos_dir)
    return nuevo


@_cand.protegido(_dd)
def anotar_numero_factura(numero_reserva, hotel, numero_factura, datos_dir=None):
    """b93: al emitir la factura del grupo, el contrato apunta su numero legal
    (FAC-<año>-CORP-<nnnn>): la agencia puede citarlo en su factura de comision."""
    lista = leer(datos_dir)
    tocado = False
    for c in lista:
        if _txt(c.get("factura_grupo")) == _txt(numero_reserva) and (not _txt(hotel) or _txt(c.get("hotel_id")) in ("", _txt(hotel))):
            c["factura_numero"] = _txt(numero_factura)
            tocado = True
    if tocado:
        _escribir(lista, datos_dir)
    return tocado


def buscar(cid, datos_dir=None):
    return next((c for c in leer(datos_dir) if c.get("id") == cid), None)


def del_hotel(lista, hotel):
    """Igualdad estricta, como el resto de AR: con hotel elegido, solo lo suyo;
    sin hotel (vista de grupo o 0 hoteles), todo."""
    if not hotel:
        return list(lista)
    return [c for c in lista if _txt(c.get("hotel_id")) == _txt(hotel)]


@_cand.protegido(_dd)
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


@_cand.protegido(_dd)
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
    _cand.escribir_excel(df, pr)
    return True


# ── b88 · la factura de comision de la agencia ──────────────────────────────
# La manda la agencia y entra por Procesar archivos como cualquier factura AP.
# Aqui se une a SU contrato: por la referencia que cite (numero de contrato,
# factura del grupo o el evento) o, si la agencia solo tiene una factura libre
# y un solo contrato con comision abierto, por eliminacion. Lo que sea dudoso
# NO se une solo: se enseña para que una persona elija. Unida a un contrato:
#   - se compara lo facturado (base sin IVA) con la comision ESPERADA;
#   - se imputa por su FECHA DE FACTURA (regla de finanzas, aunque el resto de
#     AP vaya por fecha de registro desde b84);
#   - su asiento es 628 (D) / 472 (D) / 410 (H), como las comisiones OTA.
TOL_ABS = 1.0
TOL_PCT = 1.0


def _nif(v):
    s = re.sub(r"[^A-Z0-9]", "", _txt(v).upper())
    return s[2:] if s.startswith("ES") and len(s) > 9 else s


def _prov(v):
    try:
        from cuentas_proveedor import clave_proveedor
        return clave_proveedor(v)
    except Exception:
        return _norm(v)


def _clave_ap(fila):
    try:
        from almacen_datos import clave_ap
        return clave_ap(fila)
    except Exception:
        return _txt(fila.get("numero_factura")) or _txt(fila.get("archivo"))


def es_de_la_agencia(c, fila):
    """La factura AP es de la agencia del contrato (por NIF o por nombre)."""
    ag = _prov(c.get("agencia"))
    if not ag:
        return False
    n1, n2 = _nif(c.get("agencia_nif")), _nif(fila.get("NIF_proveedor") or fila.get("nif_proveedor"))
    if n1 and n2 and n1 == n2:
        return True
    pv = _prov(fila.get("nombre_proveedor"))
    return bool(pv) and (pv == ag or (len(ag) >= 5 and len(pv) >= 5 and (ag in pv or pv in ag)))


def _mismo_hotel(c, fila):
    return _txt(c.get("hotel_id")) == _txt(fila.get("hotel_id"))


def referencia_fuerte(c, fila):
    """La factura cita el contrato: su numero, la factura del grupo o el evento."""
    t = " " + _norm(" ".join(_txt(fila.get(k)) for k in ("numero_factura", "descripcion_concepto", "concepto",
                                                             "descripcion", "archivo", "referencia"))) + " "
    claves = [_norm(c.get("contrato")), _norm(c.get("factura_grupo")), _norm(c.get("evento_id")),
              _norm(c.get("factura_numero"))]          # b93: el numero legal de la factura del grupo
    if any(k and len(k) >= 4 and (" " + k + " ") in t for k in claves):
        return True
    nombre = _norm(c.get("evento_nombre")) or _norm(c.get("evento"))
    fichas = [w for w in nombre.split() if len(w) >= 4]
    if nombre and len(nombre) >= 8 and (" " + nombre + " ") in t:
        return True
    return len(fichas) >= 2 and all((" " + w + " ") in t for w in fichas)


def base_factura(fila):
    """Lo facturado sin IVA (con eso se compara la comision esperada)."""
    b = _f(fila.get("base_imponible"))
    if b:
        return _r(b)
    t = _f(fila.get("total_factura"))
    iva = _f(fila.get("cuota_iva"))
    if t and iva:
        return _r(t - iva)
    pct = _f(fila.get("porcentaje_iva"), 21)
    return _r(t / (1 + pct / 100)) if t else 0.0


def enlazar(contratos, facturas):
    """Une cada contrato con la factura de comision de su agencia.

    facturas: filas AP (dicts). Devuelve {id_contrato: {clave, origen, candidatas}}.
    origen: 'manual' (una persona la unio), 'referencia' (la factura cita el
    contrato) o 'unica' (la agencia tiene UNA factura libre y UN contrato con
    comision abierto, y el importe cuadra). Una factura nunca va a dos contratos."""
    por_clave = {}
    for f in facturas:
        k = _clave_ap(f)
        if k and k not in por_clave:
            por_clave[k] = f
    res = {c["id"]: {"clave": "", "origen": "", "candidatas": []} for c in contratos}
    tomadas = set()
    for c in contratos:                                     # 1. lo que unio una persona
        fc = c.get("factura_comision") or {}
        k = _txt(fc.get("clave"))
        if fc.get("origen") == "manual" and k in por_clave and k not in tomadas:
            res[c["id"]].update(clave=k, origen="manual")
            tomadas.add(k)
    cand = {}
    for c in contratos:
        if res[c["id"]]["clave"] or not _txt(c.get("agencia")):
            continue
        fuera = set((c.get("factura_comision") or {}).get("excluidas") or [])
        cand[c["id"]] = [k for k, f in por_clave.items()
                         if k not in tomadas and k not in fuera and _mismo_hotel(c, f) and es_de_la_agencia(c, f)]
    for c in contratos:                                     # 2. la factura cita el contrato
        ks = [k for k in cand.get(c["id"], []) if k not in tomadas and referencia_fuerte(c, por_clave[k])]
        if len(ks) == 1:
            res[c["id"]].update(clave=ks[0], origen="referencia")
            tomadas.add(ks[0])
    abiertos = [c for c in contratos if not res[c["id"]]["clave"]
                and (c.get("comision") or {}).get("modo") == "porcentaje"]
    for c in abiertos:                                      # 3. por eliminacion
        ks = [k for k in cand.get(c["id"], []) if k not in tomadas]
        compite = [x for x in abiertos if x["id"] != c["id"] and not res[x["id"]]["clave"]
                   and set(ks) & set(cand.get(x["id"], []))]
        # sin referencia solo se une sola si ademas el importe cuadra con lo
        # esperado; si no cuadra, que la elija una persona (y vea la diferencia)
        if len(ks) == 1 and not compite and estado_comision(c, por_clave[ks[0]])["estado"] == "CUADRA":
            res[c["id"]].update(clave=ks[0], origen="unica")
            tomadas.add(ks[0])
    for c in contratos:
        if not res[c["id"]]["clave"]:
            res[c["id"]]["candidatas"] = [k for k in cand.get(c["id"], []) if k not in tomadas]
    return res


def estado_comision(c, fila):
    """Esperado contra facturado. CUADRA / DIFERENCIA / SIN_ESPERADA (aun no se
    sabe el %) / NO_DEBERIA (tarifa neta o sin agencia: no hay comision)."""
    fact = base_factura(fila)
    modo = (c.get("comision") or {}).get("modo")
    if modo in ("neta", "sin_agencia"):
        return {"estado": "NO_DEBERIA", "esperada": 0.0, "facturada": fact, "diferencia": fact}
    esp = comision_esperada(c)
    if esp is None:
        return {"estado": "SIN_ESPERADA", "esperada": None, "facturada": fact, "diferencia": None}
    dif = _r(fact - esp)
    tol = max(TOL_ABS, abs(esp) * TOL_PCT / 100)
    return {"estado": "CUADRA" if abs(dif) <= tol else "DIFERENCIA", "esperada": esp, "facturada": fact, "diferencia": dif}


def _iso(v):
    try:
        from almacen_datos import _iso as iso
        return iso(v)
    except Exception:
        return _txt(v)[:10]


def marcar_comisiones(df, contratos=None, datos_dir=None):
    """Marca en las facturas AP las que son la factura de comision de un
    contrato de grupo (columnas es_comision_agencia, comision_*, factura_grupo)
    y les aplica la regla de finanzas: fecha contable = fecha de la factura y
    gasto 628 (salvo que una persona haya corregido la cuenta)."""
    if df is None or getattr(df, "empty", True):
        return df
    contratos = leer(datos_dir) if contratos is None else contratos
    filas = df.to_dict("records")
    enl = enlazar(contratos, filas) if contratos else {}
    cid_de = {v["clave"]: cid for cid, v in enl.items() if v["clave"]}
    cmap = {c["id"]: c for c in contratos}
    cols = {k: [] for k in ("es_comision_agencia", "comision_contrato", "comision_evento", "comision_esperada",
                            "comision_facturada", "comision_estado", "comision_diferencia", "comision_pagador",
                            "factura_grupo", "comision_vinculo", "grupo_estado", "grupo_numero")}
    grupos = [None]         # b89: las facturas de grupo se leen solo si hay alguna comision

    def _estado_grupo(c):
        num = _txt(c.get("factura_grupo"))
        if not num:
            return "", ""
        if grupos[0] is None:
            try:
                import compensaciones as _cmp
                grupos[0] = _cmp.leer_facturas_grupo(datos_dir)
            except Exception:
                grupos[0] = {}
        fg = grupos[0].get(num + "|" + _txt(c.get("hotel_id"))) or grupos[0].get(num + "|") or {}
        return _txt(fg.get("estado")), (_txt(fg.get("numero_factura")) or num)
    f_con, ctas, ctas_g = [], [], []
    tiene_fc = "fecha_contable" in df.columns
    tiene_cdg = "cuenta_debe_gasto" in df.columns
    for f in filas:
        cid = cid_de.get(_clave_ap(f))
        legado = _txt(f.get("tipo")).upper() == "COMISION_AGENCIA"
        ajustada = bool(f.get("cuenta_ajustada")) and not (isinstance(f.get("cuenta_ajustada"), float) and f.get("cuenta_ajustada") != f.get("cuenta_ajustada"))
        if cid or legado:
            c = cmap.get(cid) or {}
            e = estado_comision(c, f) if c else {"estado": "SIN_ESPERADA", "esperada": None, "facturada": base_factura(f), "diferencia": None}
            cols["es_comision_agencia"].append(True)
            cols["comision_contrato"].append(cid or "")
            cols["comision_evento"].append(_txt(c.get("evento")))
            cols["comision_esperada"].append(e["esperada"])
            cols["comision_facturada"].append(e["facturada"])
            cols["comision_estado"].append(e["estado"])
            cols["comision_diferencia"].append(e["diferencia"])
            cols["comision_pagador"].append((c.get("pagador") or {}).get("quien", ""))
            cols["factura_grupo"].append(_txt(c.get("factura_grupo")))
            cols["comision_vinculo"].append((enl.get(cid) or {}).get("origen", "") if cid else "legado")
            _eg = _estado_grupo(c) if c else ("", "")
            _eg = _eg if isinstance(_eg, tuple) else ("", "")
            cols["grupo_estado"].append(_eg[0])
            cols["grupo_numero"].append(_eg[1])
            f_con.append(_iso(f.get("fecha_factura") if _txt(f.get("fecha_factura")) else f.get("fecha")) or (f.get("fecha_contable") if tiene_fc else ""))
            ctas.append(f.get("cuenta_contable") if ajustada else "628")
            ctas_g.append(f.get("cuenta_debe_gasto") if ajustada else "628")
        else:
            cols["es_comision_agencia"].append(False)
            for k in ("comision_contrato", "comision_evento", "comision_estado", "comision_pagador", "factura_grupo", "comision_vinculo", "grupo_estado", "grupo_numero"):
                cols[k].append("")
            for k in ("comision_esperada", "comision_facturada", "comision_diferencia"):
                cols[k].append(None)
            f_con.append(f.get("fecha_contable") if tiene_fc else "")
            ctas.append(f.get("cuenta_contable"))
            ctas_g.append(f.get("cuenta_debe_gasto") if tiene_cdg else None)
    df = df.copy()
    for k, v in cols.items():
        df[k] = v
    if any(cols["es_comision_agencia"]):
        df["fecha_contable"] = f_con
        df["cuenta_contable"] = [str(x) if x is not None else x for x in ctas]
        if tiene_cdg:
            df["cuenta_debe_gasto"] = [str(x) if x is not None else x for x in ctas_g]
    return df


@_cand.protegido(_dd)
def vincular(cid, clave, usuario="", datos_dir=None):
    """Una persona une (o elige) la factura de comision de un contrato."""
    lista = leer(datos_dir)
    c = next((x for x in lista if x.get("id") == cid), None)
    if c is None:
        raise KeyError("contrato no encontrado")
    clave = _txt(clave)
    if not clave:
        raise ValueError("falta la factura")
    for x in lista:                     # una factura no va a dos contratos
        fx = x.get("factura_comision") or {}
        if x.get("id") != cid and fx.get("origen") == "manual" and _txt(fx.get("clave")) == clave:
            x["factura_comision"] = {"origen": "", "excluidas": fx.get("excluidas") or []}
    prev = c.get("factura_comision") or {}
    c["factura_comision"] = {"clave": clave, "origen": "manual", "por": usuario, "cuando": _ahora(),
                             "excluidas": [k for k in (prev.get("excluidas") or []) if k != clave]}
    c.setdefault("historial", []).append({"fecha": _ahora(), "usuario": usuario, "campo": "factura_comision", "valor": clave})
    _escribir(lista, datos_dir)
    return c


@_cand.protegido(_dd)
def desvincular(cid, clave, usuario="", datos_dir=None):
    """"Esta factura no es de este contrato": se separa y no se vuelve a unir sola."""
    lista = leer(datos_dir)
    c = next((x for x in lista if x.get("id") == cid), None)
    if c is None:
        raise KeyError("contrato no encontrado")
    prev = c.get("factura_comision") or {}
    excl = list(prev.get("excluidas") or [])
    if _txt(clave) and _txt(clave) not in excl:
        excl.append(_txt(clave))
    c["factura_comision"] = {"origen": "separada", "por": usuario, "cuando": _ahora(), "excluidas": excl}
    c.setdefault("historial", []).append({"fecha": _ahora(), "usuario": usuario, "campo": "factura_comision", "valor": "separada " + _txt(clave)})
    _escribir(lista, datos_dir)
    return c
