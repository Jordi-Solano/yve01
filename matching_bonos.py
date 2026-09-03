# -*- coding: utf-8 -*-
"""matching_bonos.py — direct bill: la factura a crédito contra el bono (OLA A).

Un BONO (voucher) es lo que una agencia o empresa autoriza al hotel a
facturarle por la estancia de un huésped. La factura direct bill
(`reservas_credito.xlsx`, estado FACTURADO/COBRADO) tiene que cuadrar con él:
mismo pagador, mismas fechas y mismo importe. Y al revés: una factura a
crédito SIN bono es un cobro sin autorización.

Estados por bono:
  CUADRA              factura encontrada e importe dentro del margen
  DIFERENCIA_IMPORTE  factura encontrada, importe distinto
  DIFERENCIA_FECHAS   factura encontrada por pagador+importe, fechas distintas
  SIN_FACTURA         no hay factura para este bono (todavia no se ha facturado)
Y aparte: FACTURA_SIN_BONO = facturas a credito que ningun bono respalda.

Como se busca la factura de un bono, en orden:
  1. `referencia_reserva` del bono == numero de la factura
  2. mismo pagador (agencia ~ cliente) y mismas fechas de entrada/salida
  3. mismo pagador y mismo importe (±margen)

SOLO LEE. Funciones puras sobre DataFrames para poder probarlas sin Flask.
"""
import unicodedata
from io import BytesIO

import pandas as pd

MARGEN_EUR = 1.0     # hasta 1 EUR de redondeo cuadra
MARGEN_PCT = 1.0     # o hasta el 1 %


def _txt(v):
    if v is None:
        return ""
    try:
        if isinstance(v, float) and v != v:
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _num(v):
    try:
        if isinstance(v, str):
            v = v.replace("€", "").strip()
            if "," in v:
                v = v.replace(".", "").replace(",", ".")
        f = float(v)
        return None if f != f else round(f, 2)
    except Exception:
        return None


def _norm(x):
    x = unicodedata.normalize("NFKD", _txt(x).lower())
    x = "".join(c for c in x if c.isalnum() or c.isspace())
    palabras = [p for p in x.split() if p not in ("sl", "sa", "slu", "sau", "de", "la", "el", "the", "hotel", "viajes", "travel")]
    return " ".join(palabras)


def _fecha(v):
    s = _txt(v)
    if not s:
        return None
    try:
        if len(s) >= 10 and s[4] == "-":
            return pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce").date()
        return pd.to_datetime(s[:10], dayfirst=True, errors="coerce").date()
    except Exception:
        return None


def _mismo_pagador(agencia, cliente):
    a, c = _norm(agencia), _norm(cliente)
    if not a or not c:
        return False
    return a == c or a in c or c in a


def _cuadra(a, b):
    if a is None or b is None:
        return False
    d = abs(a - b)
    return d <= MARGEN_EUR or (b and d / abs(b) * 100 <= MARGEN_PCT)


def _facturas(df_reservas):
    """Las facturas a credito, normalizadas: numero, cliente, fechas, total."""
    out = []
    if df_reservas is None or df_reservas.empty:
        return out
    for _, r in df_reservas.iterrows():
        estado = _txt(r.get("estado")).upper()
        if estado not in ("FACTURADO", "COBRADO"):
            continue        # PENDIENTE_FACTURA no es una factura todavia
        num = _txt(r.get("numero_reserva")) or _txt(r.get("numero"))
        total = _num(r.get("total"))
        if total is None:
            total = _num(r.get("importe"))
        out.append({
            "numero":        num,
            "cliente":       _txt(r.get("cliente")),
            "fecha_entrada": _fecha(r.get("fecha_entrada")),
            "fecha_salida":  _fecha(r.get("fecha_salida")),
            "total":         total,
            "estado":        estado,
            "hotel_id":      _txt(r.get("hotel_id")),
        })
    return out


def cotejar(df_bonos, df_reservas):
    """Devuelve {bonos: [...], facturas_sin_bono: [...], resumen: {...}}."""
    facturas = _facturas(df_reservas)
    usadas = set()
    bonos = []
    if df_bonos is not None and not df_bonos.empty:
        for _, b in df_bonos.iterrows():
            agencia = _txt(b.get("agencia"))
            total_b = _num(b.get("importe_total"))
            fe, fs = _fecha(b.get("fecha_entrada")), _fecha(b.get("fecha_salida"))
            ref = _txt(b.get("referencia_reserva"))
            fila = None
            via = ""
            if ref:
                for f in facturas:
                    if f["numero"] and f["numero"].upper() == ref.upper() and id(f) not in usadas:
                        fila, via = f, "referencia"
                        break
            if fila is None and fe and fs:
                for f in facturas:
                    if id(f) in usadas:
                        continue
                    if _mismo_pagador(agencia, f["cliente"]) and f["fecha_entrada"] == fe and f["fecha_salida"] == fs:
                        fila, via = f, "pagador+fechas"
                        break
            if fila is None and total_b is not None:
                for f in facturas:
                    if id(f) in usadas:
                        continue
                    if _mismo_pagador(agencia, f["cliente"]) and _cuadra(f["total"], total_b):
                        fila, via = f, "pagador+importe"
                        break
            if fila is None:
                estado, detalle = "SIN_FACTURA", "todavia no se ha facturado esta estancia"
            else:
                usadas.add(id(fila))
                if via == "pagador+importe" and fe and fila["fecha_entrada"] and fila["fecha_entrada"] != fe:
                    estado = "DIFERENCIA_FECHAS"
                    detalle = f"bono {fe} vs factura {fila['fecha_entrada']}"
                elif total_b is None:
                    estado, detalle = "CUADRA", "el bono no trae importe: se coteja por pagador y fechas"
                elif _cuadra(fila["total"], total_b):
                    estado, detalle = "CUADRA", ""
                else:
                    estado = "DIFERENCIA_IMPORTE"
                    detalle = f"factura {fila['total']:.2f} vs bono {total_b:.2f} (dif {fila['total'] - total_b:+.2f})"
            bonos.append({
                "clave":            _txt(b.get("clave")),
                "numero_bono":      _txt(b.get("numero_bono")),
                "agencia":          agencia,
                "huesped":          _txt(b.get("huesped")),
                "fecha_entrada":    fe.isoformat() if fe else "",
                "fecha_salida":     fs.isoformat() if fs else "",
                "importe_bono":     total_b,
                "numero_factura":   fila["numero"] if fila else "",
                "importe_factura":  fila["total"] if fila else None,
                "estado_factura":   fila["estado"] if fila else "",
                "via":              via,
                "estado":           estado,
                "detalle":          detalle,
                "hotel_id":         _txt(b.get("hotel_id")),
            })
    sin_bono = [{"numero": f["numero"], "cliente": f["cliente"],
                 "fecha_entrada": f["fecha_entrada"].isoformat() if f["fecha_entrada"] else "",
                 "total": f["total"], "estado": f["estado"], "hotel_id": f["hotel_id"]}
                for f in facturas if id(f) not in usadas]
    n = {k: sum(1 for b in bonos if b["estado"] == k)
         for k in ("CUADRA", "DIFERENCIA_IMPORTE", "DIFERENCIA_FECHAS", "SIN_FACTURA")}
    n["FACTURA_SIN_BONO"] = len(sin_bono)
    return {
        "bonos": bonos,
        "facturas_sin_bono": sin_bono,
        "resumen": {**n, "total_bonos": len(bonos),
                    "importe_en_disputa": round(sum(abs((b["importe_factura"] or 0) - (b["importe_bono"] or 0))
                                                    for b in bonos if b["estado"] == "DIFERENCIA_IMPORTE"), 2),
                    "importe_sin_bono": round(sum(f["total"] or 0 for f in sin_bono), 2)},
        "nota": "Se cotejan facturas FACTURADO/COBRADO; una reserva PENDIENTE_FACTURA todavia no es factura.",
    }


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([res["resumen"]]).to_excel(w, index=False, sheet_name="Resumen")
        pd.DataFrame(res["bonos"] or [{}]).to_excel(w, index=False, sheet_name="Bonos")
        pd.DataFrame(res["facturas_sin_bono"] or [{}]).to_excel(w, index=False, sheet_name="Facturas sin bono")
    buf.seek(0)
    return buf, "bonos_vs_facturas.xlsx"
