# -*- coding: utf-8 -*-
"""aging_ap.py — a quien debemos y desde cuando (OLA A).

Antiguedad de las facturas de proveedor y de las liquidaciones OTA que
todavia no se han PAGADO. "Pagada" es lo unico que Yve puede saber sin
conector: el extracto bancario la ha conciliado (`estado == CONCILIADO` con
`factura_ref` = su numero). Lo demas esta pendiente, aprobado o no.

Tramos por dias desde el VENCIMIENTO (b84: fecha de factura + dias de pago del
proveedor, o el que una persona ponga en la ficha); si una fila no trae
vencimiento, desde la fecha de factura como antes. Dias negativos = todavia no
vence (tramo 0-30). Sin ninguna fecha la fila va a "sin fecha", nunca a un tramo.
Tambien se excluye lo marcado PAGADA a mano en la ficha (b84), ademas de lo
conciliado en el banco. b89: de la factura de comision de una agencia se resta lo
compensado contra la factura del grupo (si queda a cero, sale).

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
    if dias <= 30:           # incluye lo que aun no vence (dias negativos)
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
    pagadas_mano = [0]
    compensadas = [0]
    filas = []

    def _add(origen, num, acreedor, fecha, importe, aprobacion, hotel, vencimiento=None, pagada_mano=False, extra=None):
        num = _txt(num)
        if num and num.upper() in pagadas:
            return
        if pagada_mano:
            pagadas_mano[0] += 1
            return
        f = _fecha(fecha)
        v = _fecha(vencimiento) if vencimiento is not None and _txt(vencimiento) else None
        dias = (hoy - (v or f)).days if (v or f) else None
        filas.append({
            "origen":      origen,
            "numero_factura": num or "N/D",
            "acreedor":    _txt(acreedor) or "Desconocido",
            "fecha":       f.isoformat() if f else "",
            "vencimiento": v.isoformat() if v else "",
            "dias":        dias,
            "tramo":       _tramo(dias),
            "importe":     round(importe, 2),
            "aprobacion":  _txt(aprobacion).upper() or "PENDIENTE",
            "hotel_id":    _txt(hotel),
            **(extra or {}),
        })

    if df_ap is not None and not df_ap.empty:
        from cierre_mes import es_comision_agencia
        for _, r in df_ap.iterrows():
            imp = _num(r.get("total_factura")) or _num(r.get("importe_total")) or _num(r.get("total"))
            f_fac = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            # b89: lo compensado contra la factura del grupo ya no se debe (410/430)
            comp = _num(r.get("compensado")) if "compensado" in df_ap.columns else 0.0
            if comp and imp - comp <= 0.005:
                compensadas[0] += 1
                continue
            imp = round(imp - (comp or 0.0), 2)
            # b88: la comision de una agencia de grupos va aparte y dice en que mes se imputa
            com = es_comision_agencia(r)
            extra = None
            if com:
                extra = {"comision_de": _txt(r.get("comision_evento")), "imputacion": (_fecha(f_fac).isoformat()[:7] if _fecha(f_fac) else ""),
                         "compensado": round(comp or 0.0, 2)}
                # b89: si paga el cliente final, se paga cuando se cobre la factura del grupo
                if _txt(r.get("comision_pagador")) == "cliente" and _txt(r.get("grupo_estado")).upper() not in ("COBRADO", "COBRADA") \
                        and _txt(r.get("factura_grupo")):
                    extra["tras_cobro"] = _txt(r.get("factura_grupo"))
            _add("Comisión grupo" if com else "Proveedor", r.get("numero_factura"), r.get("nombre_proveedor"),
                 f_fac, imp, r.get("accion"), r.get("hotel_id"),
                 vencimiento=r.get("vencimiento"), pagada_mano=bool(r.get("pagada")) if "pagada" in df_ap.columns else False,
                 extra=extra)
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
        if f.get("imputacion"):
            p.setdefault("imputacion", [])
            if f["imputacion"] not in p["imputacion"]:
                p["imputacion"] = sorted(p["imputacion"] + [f["imputacion"]])
        if f.get("tras_cobro"):
            p.setdefault("tras_cobro", [])
            if f["tras_cobro"] not in p["tras_cobro"]:
                p["tras_cobro"] = sorted(p["tras_cobro"] + [f["tras_cobro"]])
        if f.get("compensado"):
            p["compensado"] = round(p.get("compensado", 0.0) + f["compensado"], 2)
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
        "n_pagadas_mano": pagadas_mano[0],
        "n_compensadas": compensadas[0],
        "mas_de_60": vencido,
        "sin_fecha": sum(1 for f in filas if f["dias"] is None),
        "nota": ("Pendiente = sin conciliar en el extracto bancario ni marcada como pagada en la ficha. "
                 "Los dias cuentan desde el vencimiento (fecha de factura + dias de pago). "
                 "Lo compensado contra la factura del grupo de la agencia ya no se debe."),
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
