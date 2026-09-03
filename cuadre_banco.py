# -*- coding: utf-8 -*-
"""cuadre_banco.py — cuadre del banco por pestañas (OLA B · bloque 2).

Lo que hace el controller a fin de mes con el extracto: cada movimiento va a
una pestaña —AR (cobros de clientes/OTAs), AP (pagos a proveedores), TARJETAS
(liquidaciones del TPV/adquirente), CAJA (ingresos de efectivo), VARIOS
(nominas, impuestos, comisiones bancarias, alquiler...)— y cada pestaña se
cuadra contra lo que la justifica. Lo que no se sabe clasificar queda en
SIN_CLASIFICAR, nunca se reparte a ojo.

Como se decide la pestaña, en orden:
  1. asignacion manual guardada (cuadre_banco_manual.json, clave = clave_movimiento)
  2. la conciliacion: `origen` AP/AR del informe (movimiento cruzado con factura)
  3. palabras clave del concepto (config_banco_pestanas.json amplia/sustituye)
  4. el nombre de un proveedor de proveedores.xlsx o de una OTA conocida
  5. si nada encaja: SIN_CLASIFICAR

Funciones puras sobre DataFrames; el unico fichero que se escribe es el de
asignaciones manuales (desde tab_cierre).
"""
import json
import os
import unicodedata
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt, _mes_a_rango

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PESTANAS = ("AR", "AP", "TARJETAS", "CAJA", "VARIOS", "SIN_CLASIFICAR")

PALABRAS_DEFECTO = {
    "TARJETAS": ["redsys", "tpv", "visa", "mastercard", "amex", "american express", "adquir",
                 "comercia", "liq tarjeta", "liq. tarjeta", "liquidacion tarjeta", "stripe",
                 "paypal", "adyen", "datafono", "universalpay", "servired", "4b"],
    "CAJA":     ["ingreso efectivo", "ingreso en efectivo", "ingreso caja", "remesa efectivo",
                 "deposito efectivo", "prosegur", "loomis", "ingreso metalico"],
    "VARIOS":   ["comision", "nomina", "seguridad social", "tgss", "aeat", "hacienda", "impuesto",
                 "tributo", "alquiler", "seguro", "intereses", "traspaso", "transferencia interna",
                 "prestamo", "cuota", "leasing", "renting", "devolucion recibo", "ayuntamiento"],
    "AR":       ["booking", "expedia", "hotelbeds", "hotusa", "agoda", "airbnb", "despegar",
                 "cobro", "liquidacion ota", "trivago"],
    "AP":       ["pago", "fra.", "fra ", "factura", "recibo"],
}
OTAS = ["booking", "expedia", "hotelbeds", "hotusa", "agoda", "airbnb", "despegar", "trivago", "hotels.com"]


def _norm(x):
    x = unicodedata.normalize("NFKD", _txt(x).lower())
    return "".join(c for c in x if not unicodedata.combining(c))


def palabras(datos_dir=None):
    p = {k: list(v) for k, v in PALABRAS_DEFECTO.items()}
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "config_banco_pestanas.json")
    try:
        with open(ruta, encoding="utf-8") as fh:
            extra = json.load(fh) or {}
        for k, v in extra.items():
            if k in PESTANAS and isinstance(v, list):
                p.setdefault(k, []).extend(_norm(x) for x in v)
    except Exception:
        pass
    return p


def manuales(datos_dir=None):
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "cuadre_banco_manual.json")
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def guardar_manual(clave, pestana, datos_dir=None):
    """La unica escritura del modulo: clave_movimiento -> pestaña (o '' para quitar)."""
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "cuadre_banco_manual.json")
    d = manuales(datos_dir)
    if pestana:
        d[clave] = pestana
    else:
        d.pop(clave, None)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)
    return d


def proveedores_conocidos(datos_dir=None):
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "proveedores.xlsx")
    out = []
    try:
        df = pd.read_excel(ruta)
        col = next((c for c in df.columns if "nombre" in c.lower()), None)
        if col:
            for v in df[col]:
                n = _norm(v)
                # la primera palabra "fuerte" del nombre (Makro, Endesa...)
                primera = next((w for w in n.split() if len(w) >= 4 and w not in ("hotel", "grupo")), n)
                if primera:
                    out.append(primera)
    except Exception:
        pass
    return out


def clasificar(mov, palabras_cfg, manual, proveedores):
    """(pestaña, via). `mov` es un dict del extracto ya cruzado (almacen_datos)."""
    from almacen_datos import clave_movimiento
    clave = clave_movimiento(mov)
    if clave in manual and manual[clave] in PESTANAS:
        return manual[clave], "manual"
    origen = _txt(mov.get("origen")).upper()
    estado = _txt(mov.get("estado")).upper()
    if estado == "CONCILIADO" and origen in ("AP", "AR"):
        return origen, "conciliacion"
    c = _norm(mov.get("concepto"))
    imp = _num(mov.get("importe"))
    for pest in ("TARJETAS", "CAJA", "VARIOS"):
        if any(w in c for w in palabras_cfg.get(pest, [])):
            return pest, "concepto"
    # AR: una OTA es AR con cualquier signo (una devolucion a Booking sigue
    # siendo cuenta de clientes); "cobro" y similares solo si entra dinero
    if any(o in c for o in OTAS):
        return "AR", "concepto"
    if imp > 0 and any(w in c for w in palabras_cfg.get("AR", [])):
        return "AR", "concepto"
    if any(p in c for p in proveedores) and imp < 0:
        return "AP", "proveedor"
    if any(w in c for w in palabras_cfg.get("AP", [])) and imp < 0:
        return "AP", "concepto"
    return "SIN_CLASIFICAR", ""


def cuadrar(mes, df_banco, ventas_fb=None, palabras_cfg=None, manual=None, proveedores=None):
    """Devuelve {mes, pestañas: {...}, movimientos: [...], saldo_final, resumen}."""
    ini, fin, mes = _mes_a_rango(mes)
    palabras_cfg = palabras_cfg or {k: list(v) for k, v in PALABRAS_DEFECTO.items()}
    manual = manual or {}
    proveedores = proveedores or []
    from almacen_datos import clave_movimiento
    movs = []
    saldo_final = None; fecha_saldo = None
    if df_banco is not None and not df_banco.empty:
        for _, r in df_banco.iterrows():
            f = _fecha(r.get("fecha"))
            if f is None or not (ini <= f <= fin):
                continue
            d = r.to_dict()
            pest, via = clasificar(d, palabras_cfg, manual, proveedores)
            imp = _num(r.get("importe"))
            movs.append({
                "clave": clave_movimiento(d), "fecha": f.isoformat(), "concepto": _txt(r.get("concepto")),
                "importe": imp, "pestana": pest, "via": via,
                "estado": _txt(r.get("estado")).upper() or "PENDIENTE",
                "factura_ref": _txt(r.get("factura_ref")), "hotel_id": _txt(r.get("hotel_id")),
            })
            s = _num(r.get("saldo"))
            if r.get("saldo") is not None and _txt(r.get("saldo")) and (fecha_saldo is None or f >= fecha_saldo):
                fecha_saldo, saldo_final = f, s
    movs.sort(key=lambda m: (m["fecha"], m["concepto"]))

    # TPV del mes (lo unico que justifica TARJETAS sin PMS: las ventas F&B)
    tpv = 0.0
    if ventas_fb is not None and not ventas_fb.empty and "total_venta" in ventas_fb.columns:
        for _, r in ventas_fb.iterrows():
            f = _fecha(r.get("fecha"))
            if f and ini <= f <= fin:
                tpv = round(tpv + _num(r.get("total_venta")), 2)

    pest = {}
    for p in PESTANAS:
        ms = [m for m in movs if m["pestana"] == p]
        total = round(sum(m["importe"] for m in ms), 2)
        con_ref = round(sum(m["importe"] for m in ms if m["factura_ref"]), 2)
        info = {"n": len(ms), "total": total, "cobros": round(sum(m["importe"] for m in ms if m["importe"] > 0), 2),
                "pagos": round(sum(m["importe"] for m in ms if m["importe"] < 0), 2)}
        if p in ("AP", "AR"):
            info.update({"justificado": con_ref, "diferencia": round(total - con_ref, 2),
                         "n_sin_factura": sum(1 for m in ms if not m["factura_ref"]),
                         "estado": "CUADRA" if all(m["factura_ref"] for m in ms) else ("PENDIENTE" if ms else "CUADRA"),
                         "nota": "Justificado = movimientos cruzados con una factura (conciliacion o asignacion manual)."})
        elif p == "TARJETAS":
            info.update({"justificado": tpv if tpv else None, "diferencia": round(total - tpv, 2) if tpv else None,
                         "estado": "INFO", "nota": ("Contra las ventas F&B del TPV del mes. Las tarjetas de habitaciones "
                                                   "las liquida el PMS: la diferencia es normal hasta tener conector.")})
        elif p == "CAJA":
            info.update({"justificado": None, "diferencia": None, "estado": "SIN_DATO",
                         "nota": "Los ingresos de efectivo se cuadran contra el arqueo de caja (Glory/CREPT), que Yve no tiene."})
        elif p == "VARIOS":
            info.update({"justificado": None, "diferencia": None, "estado": "INFO",
                         "nota": "Nominas, impuestos, comisiones bancarias, alquileres... Revisar que cada uno tenga su documento."})
        else:
            info.update({"justificado": None, "diferencia": None,
                         "estado": "CUADRA" if not ms else "PENDIENTE",
                         "nota": "Asigna cada movimiento a su pestaña; lo que quede aqui no esta cuadrado."})
        pest[p] = info

    total_mes = round(sum(m["importe"] for m in movs), 2)
    return {
        "mes": mes, "desde": ini.isoformat(), "hasta": fin.isoformat(),
        "movimientos": movs, "pestanas": pest, "n": len(movs),
        "total_mes": total_mes, "saldo_final": saldo_final, "fecha_saldo": fecha_saldo.isoformat() if fecha_saldo else "",
        "sin_clasificar": pest["SIN_CLASIFICAR"]["n"],
        "sin_conciliar": sum(1 for m in movs if m["estado"] != "CONCILIADO"),
        "ok": pest["SIN_CLASIFICAR"]["n"] == 0 and pest["AP"]["estado"] == "CUADRA" and pest["AR"]["estado"] == "CUADRA",
    }


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([{"pestana": k, **v} for k, v in res["pestanas"].items()]).to_excel(w, index=False, sheet_name="Resumen")
        for p in PESTANAS:
            ms = [m for m in res["movimientos"] if m["pestana"] == p]
            pd.DataFrame(ms or [{}]).to_excel(w, index=False, sheet_name=p[:31])
    buf.seek(0)
    return buf, f"cuadre_banco_{res['mes']}.xlsx"
