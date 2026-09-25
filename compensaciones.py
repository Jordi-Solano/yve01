# -*- coding: utf-8 -*-
"""compensaciones.py — Yve.01 · compensar la comision de la agencia contra la factura del grupo (b89).

Regla de finanzas (Jordi, 24 sep 2026):
  - SOLO se compensa la factura de comision contra la factura de grupo si quien PAGA la
    factura de grupo es la AGENCIA (es a la vez nuestra deudora y nuestra acreedora).
  - Si paga otro (el cliente final) NO hay compensacion: primero se COBRA la factura de
    grupo y despues se PAGA la de comision.
  - Lo compensado se asienta 410 (D) / 430 (H): se cancela la deuda con la agencia (410,
    la comision) contra lo que la agencia nos debe (430, la factura del grupo).

La compensacion puede ser total o parcial, y queda registrada (quien, cuando, cuanto, nota)
en datos-referencia/compensaciones_ar.json. Lo compensado se refleja en:
  - la factura del grupo (reservas_credito.xlsx, columna `compensado`): su saldo baja;
  - la factura de comision (ajustes_ap.json, campo `compensado`): su saldo baja y, si
    llega a cero, queda COMPENSADA (sale del aging AP como lo pagado).
Funciones puras donde se puede; `datos_dir` para las pruebas.
"""
import json
import os
from datetime import date, datetime

FICHERO = "compensaciones_ar.json"
TOL = 0.005


class ReglaError(ValueError):
    """La compensacion no se puede hacer por una regla (409 en la API)."""


def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "nat") else s


def _f(v):
    try:
        x = float(v)
        return 0.0 if x != x else x
    except (TypeError, ValueError):
        s = _txt(v).replace("€", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0


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


def leer(datos_dir=None):
    try:
        with open(os.path.join(_dd(datos_dir), FICHERO), encoding="utf-8") as fh:
            d = json.load(fh)
        return [c for c in d if isinstance(c, dict)] if isinstance(d, list) else []
    except Exception:
        return []


def _escribir(lista, datos_dir=None):
    p = os.path.join(_dd(datos_dir), FICHERO)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(lista, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def _iso(v):
    s = _txt(v)
    if not s:
        return ""
    try:
        import pandas as pd
        t = pd.to_datetime(s[:10], dayfirst=("/" in s[:10]), errors="coerce")
        return "" if pd.isna(t) else t.strftime("%Y-%m-%d")
    except Exception:
        return ""


def de_grupo(numero, hotel="", comps=None, datos_dir=None):
    comps = leer(datos_dir) if comps is None else comps
    return [c for c in comps if _txt(c.get("factura_grupo")) == _txt(numero) and _txt(c.get("hotel_id")) == _txt(hotel)]


def de_comision(clave, comps=None, datos_dir=None):
    comps = leer(datos_dir) if comps is None else comps
    return [c for c in comps if _txt(c.get("factura_comision")) == _txt(clave)]


def total(comps):
    return _r(sum(_f(c.get("importe")) for c in comps))


def saldo_grupo(fila, comps=None):
    """Lo que queda por cobrar de la factura del grupo."""
    if _txt(fila.get("estado")).upper() in ("COBRADO", "COBRADA"):
        return 0.0
    tot = _f(fila.get("total")) or _f(fila.get("importe"))
    comp = total(comps) if comps is not None else _f(fila.get("compensado"))
    return _r(max(0.0, tot - comp))


def saldo_comision(fila, comps=None):
    """Lo que queda por pagar de la factura de comision."""
    tot = _f(fila.get("total_factura"))
    comp = total(comps) if comps is not None else _f(fila.get("compensado"))
    if fila.get("pagada") is True or _txt(fila.get("pagada")).lower() == "true":
        return 0.0
    return _r(max(0.0, tot - comp))


def puede_compensar(contrato, fila_grupo, fila_com, comps_g=None, comps_c=None):
    """(True, maximo) o (False, motivo). La regla de finanzas va primero."""
    quien = (contrato.get("pagador") or {}).get("quien")
    if quien == "cliente":
        return False, ("No se compensa: la factura del grupo la paga el cliente final. "
                       "Primero se cobra la factura del grupo y después se paga la de comisión.")
    if quien != "agencia":
        return False, "Falta decidir quién paga la factura del grupo (AR › Contratos)."
    if fila_grupo is None:
        return False, "No se encuentra la factura del grupo."
    est = _txt(fila_grupo.get("estado")).upper()
    if est in ("", "PENDIENTE_FACTURA"):
        return False, "Primero hay que emitir la factura del grupo."
    if est in ("COBRADO", "COBRADA"):
        return False, "La factura del grupo ya está cobrada: no queda nada que compensar."
    if fila_com is None:
        return False, "Falta la factura de comisión de la agencia (se une en AR › Contratos)."
    sg, sc = saldo_grupo(fila_grupo, comps_g), saldo_comision(fila_com, comps_c)
    maximo = _r(min(sg, sc))
    if maximo <= 0:
        return False, "No queda saldo que compensar."
    return True, maximo


def compensar(contrato, fila_grupo, fila_com, importe, fecha=None, usuario="", nota="", datos_dir=None):
    """Registra una compensacion total o parcial. Lanza ReglaError (regla de finanzas) o
    ValueError (dato mal). Devuelve el registro."""
    comps = leer(datos_dir)
    num = _txt((fila_grupo or {}).get("numero_reserva")) or _txt((fila_grupo or {}).get("numero"))
    hotel = _txt((fila_grupo or {}).get("hotel_id"))
    try:
        from almacen_datos import clave_ap
        clave = clave_ap(fila_com) if fila_com is not None else ""
    except Exception:
        clave = _txt((fila_com or {}).get("numero_factura"))
    ok, info = puede_compensar(contrato, fila_grupo, fila_com, de_grupo(num, hotel, comps), de_comision(clave, comps))
    if not ok:
        raise ReglaError(info)
    imp = _r(_f(importe))
    if imp <= 0:
        raise ValueError("el importe a compensar tiene que ser mayor que 0")
    if imp > info + TOL:
        raise ValueError(f"como máximo se pueden compensar {info:.2f} € (el menor de los dos saldos)")
    f = _iso(fecha) or date.today().isoformat()
    import uuid
    reg = {"id": f"CMP-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:4]}", "fecha": f, "importe": imp,
           "factura_grupo": num, "factura_comision": clave, "contrato_id": _txt(contrato.get("id")),
           "agencia": _txt(contrato.get("agencia")), "hotel_id": hotel, "usuario": usuario,
           "nota": _txt(nota)[:200], "cuando": datetime.now().strftime("%Y-%m-%d %H:%M")}
    comps.append(reg)
    _escribir(comps, datos_dir)
    sincronizar(num, hotel, clave, datos_dir, usuario)
    return reg


def anular(cid, usuario="", datos_dir=None):
    """Quita una compensacion (una equivocacion). Devuelve la quitada. No se puede si la
    factura del grupo ya esta cobrada (su cobro se asento por el neto) o la de comision
    ya esta pagada: primero se deshace eso."""
    comps = leer(datos_dir)
    c = next((x for x in comps if x.get("id") == cid), None)
    if c is None:
        raise KeyError("compensacion no encontrada")
    fg = fila_grupo(c.get("factura_grupo"), c.get("hotel_id"), datos_dir=datos_dir)
    if fg is not None and fg.get("estado") in COBRADOS:
        raise ReglaError("La factura del grupo ya está cobrada (por lo que quedaba después de compensar): no se puede anular.")
    try:
        from almacen_datos import ajustes_ap
        if (ajustes_ap(_dd(datos_dir)).get(_txt(c.get("factura_comision"))) or {}).get("pagada"):
            raise ReglaError("La factura de comisión ya está marcada pagada: deshaz primero el pago en su ficha.")
    except ReglaError:
        raise
    except Exception:
        pass
    _escribir([x for x in comps if x.get("id") != cid], datos_dir)
    sincronizar(c.get("factura_grupo"), c.get("hotel_id"), c.get("factura_comision"), datos_dir, usuario)
    return c


def sincronizar(num_grupo, hotel, clave_com, datos_dir=None, usuario=""):
    """Lleva lo compensado a la factura del grupo (columna `compensado`) y a la de comision
    (ajuste `compensado` de ajustes_ap.json). El registro json es la fuente."""
    import pandas as pd
    comps = leer(datos_dir)
    dd = _dd(datos_dir)
    pr = os.path.join(dd, "reservas_credito.xlsx")
    if _txt(num_grupo) and os.path.exists(pr):
        try:
            df = pd.read_excel(pr)
            col = "numero_reserva" if "numero_reserva" in df.columns else ("numero" if "numero" in df.columns else None)
            if col:
                m = df[col].map(_txt) == _txt(num_grupo)
                if "hotel_id" in df.columns:
                    m = m & (df["hotel_id"].map(_txt) == _txt(hotel))
                if m.any():
                    if "compensado" not in df.columns:
                        df["compensado"] = 0.0
                    df["compensado"] = df["compensado"].astype(object)
                    df.loc[m, "compensado"] = total(de_grupo(num_grupo, hotel, comps))
                    tmp = pr + ".tmp.xlsx"
                    df.to_excel(tmp, index=False)
                    os.replace(tmp, pr)
        except Exception:
            pass
    if _txt(clave_com):
        try:
            from almacen_datos import guardar_ajuste_ap
            guardar_ajuste_ap(clave_com, {"compensado": total(de_comision(clave_com, comps))}, usuario or "compensacion", datos_dir=dd)
        except Exception:
            pass


# ── la factura del grupo (reservas_credito.xlsx) ───────────────────────────
COBRADOS = ("COBRADO", "COBRADA")


def leer_facturas_grupo(datos_dir=None):
    """{"numero|hotel": fila} de reservas_credito.xlsx, con `compensado` (del registro,
    que es la fuente) y `saldo` = lo que queda por cobrar."""
    import pandas as pd
    ruta = os.path.join(_dd(datos_dir), "reservas_credito.xlsx")
    out = {}
    if not os.path.exists(ruta):
        return out
    try:
        df = pd.read_excel(ruta)
    except Exception:
        return out
    comps = leer(datos_dir)
    for f in df.to_dict("records"):
        num = _txt(f.get("numero_reserva")) or _txt(f.get("numero"))
        if not num:
            continue
        hotel = _txt(f.get("hotel_id"))
        cg = de_grupo(num, hotel, comps)
        f = dict(f, numero=num, hotel_id=hotel, estado=_txt(f.get("estado")).upper(),
                 total=_r(_f(f.get("total")) or _f(f.get("importe"))), compensado=total(cg))
        f["saldo"] = saldo_grupo(f, cg)
        out[num + "|" + hotel] = f
    return out


def fila_grupo(numero, hotel="", facturas=None, datos_dir=None):
    """La factura del grupo `numero` del hotel (o sin hotel, las de antes de fase 5)."""
    numero = _txt(numero)
    if not numero:
        return None
    facturas = leer_facturas_grupo(datos_dir) if facturas is None else facturas
    return facturas.get(numero + "|" + _txt(hotel)) or facturas.get(numero + "|")


def puede_pagar_comision(fila_ap, facturas=None, datos_dir=None):
    """Regla de orden (finanzas, 24 sep 2026): si la factura del grupo la paga el CLIENTE
    final, primero se cobra la factura del grupo y despues se paga la de comision.
    (True, "") o (False, motivo). Si la factura del grupo no esta en Yve no se bloquea:
    no hay nada con que comprobarlo (se avisa en la ficha)."""
    if _txt(fila_ap.get("comision_pagador")) != "cliente":
        return True, ""
    num = _txt(fila_ap.get("factura_grupo"))
    fg = fila_grupo(num, _txt(fila_ap.get("hotel_id")), facturas, datos_dir)
    if fg is None or fg.get("estado") in COBRADOS:
        return True, ""
    return False, (f"Primero se cobra la factura del grupo {num} (la paga el cliente final) "
                   "y después se paga la de comisión.")


def vista(contrato, fila_g, fila_com, comps=None, datos_dir=None):
    """Lo que la pantalla del contrato necesita saber de la compensacion."""
    comps = leer(datos_dir) if comps is None else comps
    quien = (contrato.get("pagador") or {}).get("quien") or ""
    num = _txt((fila_g or {}).get("numero")) or _txt(contrato.get("factura_grupo"))
    hotel = _txt((fila_g or {}).get("hotel_id"))
    clave = ""
    if fila_com is not None:
        try:
            from almacen_datos import clave_ap
            clave = clave_ap(fila_com)
        except Exception:
            clave = _txt(fila_com.get("numero_factura"))
    lista = [c for c in comps if _txt(c.get("contrato_id")) == _txt(contrato.get("id"))]
    cg = de_grupo(num, hotel, comps) if num else []
    cc = de_comision(clave, comps) if clave else []
    ok, info = puede_compensar(contrato, fila_g, fila_com, cg, cc)
    return {"regla": quien, "puede": bool(ok), "maximo": info if ok else None, "motivo": "" if ok else info,
            "lista": sorted(lista, key=lambda c: (c.get("fecha") or "", c.get("id") or "")), "total": total(lista),
            "factura_grupo": num, "grupo_estado": _txt((fila_g or {}).get("estado")),
            "saldo_grupo": saldo_grupo(fila_g, cg) if fila_g is not None else None,
            "saldo_comision": saldo_comision(fila_com, cc) if fila_com is not None else None,
            "grupo_cobrada": _txt((fila_g or {}).get("estado")) in COBRADOS}


def del_mes(mes_ini, mes_fin, hotel=None, comps=None, datos_dir=None):
    comps = leer(datos_dir) if comps is None else comps
    out = []
    for c in comps:
        f = _iso(c.get("fecha"))
        if not f or not (mes_ini <= f <= mes_fin):
            continue
        if hotel and _txt(c.get("hotel_id")) != _txt(hotel):
            continue
        out.append(c)
    return out
