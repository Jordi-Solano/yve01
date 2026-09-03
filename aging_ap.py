# -*- coding: utf-8 -*-
"""aging_ap.py — a quien debemos y desde cuando (OLA A).

Antiguedad de las facturas de proveedor y de las liquidaciones OTA que
todavia no se han PAGADO. "Pagada" es lo unico que Yve puede saber sin
conector: el extracto bancario la ha conciliado (`estado == CONCILIADO` con
`factura_ref` = su numero). Lo demas esta pendiente, aprobado o no.

Tramos por dias desde la fecha de factura: 0-30 · 31-60 · 61-90 · >90.
No se inventa vencimiento: sin fecha de factura la fila va a "sin fecha" y
se cuenta aparte, nunca a un tramo.

SOLO LEE. Funciones puras sobre DataFrames para poder probarlas sin Flask.
"""
from datetime import date
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt

TRAMOS = ("0-30", "31-60", "61-90", ">90")


def _tramo(dias):
    if dias is None:
        return "sin fecha"
    if dias <= 30:
        return "0-30"
    if dias <= 60:
        return "31-60"
    if dias <= 90:
        return "61-90"
    return ">90"


def _pagadas(df_banco):
    """Numeros de factura que el banco ya ha conciliado."""
    if df_banco is None or df_banco.empty or "factura_ref" not in df_banco.columns:
        return set()
    pag = set()
    for _, r in df_banco.iterrows():
        if _txt(r.get("estado")).upper() == "CONCILIADO":
            ref = _txt(r.get("factura_ref"))
            if ref:
                pag.add(ref.upper())
    return pag


def calcular_aging(df_ap, df_ar=None, df_banco=None, hoy=None):
    """Devuelve {filas, por_acreedor, tramos, total, ...}.

    df_ap: facturas de proveedor (del panel: ya con `accion` y filtradas por
    hotel). df_ar: liquidaciones OTA (importe_comision es lo que se debe).
    """
    hoy = hoy or date.today()
    pagadas = _pagadas(df_banco)
    filas = []

    def _add(origen, num, acreedor, fecha, importe, aprobacion, hotel):
        num = _txt(num)
        if num and num.upper() in pagadas:
            return
        f = _fecha(fecha)
        dias = (hoy - f).days if f else None
        filas.append({
            "origen":      origen,
            "numero_factura": num or "N/D",
            "acreedor":    _txt(acreedor) or "Desconocido",
            "fecha":       f.isoformat() if f else "",
            "dias":        dias,
            "tramo":       _tramo(dias),
            "importe":     round(importe, 2),
            "aprobacion":  _txt(aprobacion).upper() or "PENDIENTE",
            "hotel_id":    _txt(hotel),
        })

    if df_ap is not None and not df_ap.empty:
        for _, r in df_ap.iterrows():
            imp = _num(r.get("total_factura")) or _num(r.get("importe_total")) or _num(r.get("total"))
            _add("Proveedor", r.get("numero_factura"), r.get("nombre_proveedor"),
                 r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha"),
                 imp, r.get("accion"), r.get("hotel_id"))
    if df_ar is not None and not df_ar.empty:
        for _, r in df_ar.iterrows():
            imp = _num(r.get("importe_comision")) or _num(r.get("importe_comision_factura"))
            if imp <= 0:
                continue
            _add("OTA", r.get("numero_factura"), r.get("nombre_ota"), r.get("fecha"),
                 imp, r.get("accion"), r.get("hotel_id"))

    tramos = {t: 0.0 for t in TRAMOS}
    tramos["sin fecha"] = 0.0
    por = {}
    for f in filas:
        tramos[f["tramo"]] = round(tramos[f["tramo"]] + f["importe"], 2)
        k = (f["acreedor"], f["origen"])
        p = por.setdefault(k, {"acreedor": k[0], "origen": k[1], "n": 0, "importe": 0.0,
                               "mas_antigua": None, "dias_max": None,
                               **{t: 0.0 for t in TRAMOS}, "sin fecha": 0.0,
                               "sin_aprobar": 0})
        p["n"] += 1
        p["importe"] = round(p["importe"] + f["importe"], 2)
        p[f["tramo"]] = round(p[f["tramo"]] + f["importe"], 2)
        if f["aprobacion"] not in ("APROBADA",):
            p["sin_aprobar"] += 1
        if f["dias"] is not None and (p["dias_max"] is None or f["dias"] > p["dias_max"]):
            p["dias_max"] = f["dias"]
            p["mas_antigua"] = f["fecha"]
    por_acreedor = sorted(por.values(), key=lambda p: -(p["dias_max"] or -1))
    total = round(sum(f["importe"] for f in filas), 2)
    vencido = round(tramos["61-90"] + tramos[">90"], 2)
    return {
        "hoy": hoy.isoformat(),
        "filas": sorted(filas, key=lambda f: -(f["dias"] if f["dias"] is not None else -1)),
        "por_acreedor": por_acreedor,
        "tramos": tramos,
        "total": total,
        "n": len(filas),
        "n_pagadas_excluidas": len(pagadas),
        "mas_de_60": vencido,
        "sin_fecha": sum(1 for f in filas if f["dias"] is None),
        "nota": ("Pendiente = sin conciliar en el extracto bancario. Sin extracto subido, "
                 "todo cuenta como pendiente."),
    }


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([{"tramo": k, "importe": v} for k, v in res["tramos"].items()]
                     + [{"tramo": "TOTAL", "importe": res["total"]}]
                     ).to_excel(w, index=False, sheet_name="Resumen")
        pd.DataFrame(res["por_acreedor"] or [{}]).to_excel(w, index=False, sheet_name="Por acreedor")
        pd.DataFrame(res["filas"] or [{}]).to_excel(w, index=False, sheet_name="Facturas")
    buf.seek(0)
    return buf, f"aging_ap_{res['hoy']}.xlsx"
