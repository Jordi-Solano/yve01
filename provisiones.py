# -*- coding: utf-8 -*-
"""provisiones.py — las dos provisiones del cierre que salen de lo que Yve ya tiene.

OLA A. Al cerrar el mes, el controller provisiona:

  1. **Gasto con albaran sin factura.** Mercancia que ha entrado (albaran) y
     que el proveedor todavia no ha facturado. Yve ya cruza albaranes con
     facturas (`matching_ap_albaran`): lo que sale `ALBARAN_SIN_FACTURAR` a
     fecha de corte es exactamente lo que hay que provisionar.
     Asiento: DEBE gasto (600 F&B / 629 otros) · HABER 4009 "Proveedores,
     facturas pendientes de recibir". Se revierte el dia 1 del mes siguiente.

  2. **Comisiones OTA del mes.** Las liquidaciones de las OTAs que Yve ha
     verificado y cuyo periodo cae en el mes. Se provisiona lo PACTADO cuando
     el contrato permite calcularlo (bruto x % pactado); si no, lo facturado,
     y se dice cual de los dos. Asiento: DEBE 628 · HABER cuenta de provision
     (en el plan del hotel que valido esto es la 20630; aqui es configurable).

Lo que NO hace, y lo dice: las comisiones DEVENGADAS de las que aun no ha
llegado liquidacion necesitan la produccion OTA del PMS. Eso es un conector
(Ola C), no un documento.

SOLO LEE. No escribe ningun fichero, no toca Oracle ni el clasificador. Las
cuentas salen de `hotel_config.json` si las define, si no de los defectos PGC.
"""
import calendar
import json
import os
import re
from datetime import date, datetime
from io import BytesIO

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NF = "NO_ENCONTRADO"

# Cuentas por defecto (PGC). El hotel puede sobreescribirlas en hotel_config.json:
#   {"cuenta_provision_albaranes": "4009", "cuenta_provision_comisiones": "20630"}
CUENTAS_DEFECTO = {
    "gasto_fb":               ("600", "Compras de mercaderias F&B"),
    "gasto_otros":            ("629", "Otros servicios"),
    "comision_ota":           ("628", "Comisiones de agencias y OTAs"),
    "provision_albaranes":    ("4009", "Proveedores, facturas pendientes de recibir"),
    "provision_comisiones":   ("4109", "Acreedores, comisiones pendientes de liquidar"),
}


# ── utilidades ───────────────────────────────────────────────────────────────

def _txt(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT", NF) else s


def _num(v):
    """Numero tolerante: '1.234,56', '450 EUR', NaN -> 0.0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if v != v else float(v)          # NaN != NaN
    s = _txt(v).replace("EUR", "").replace("€", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fecha(v):
    """date o None. Acepta Timestamp, 'dd/mm/aaaa', 'aaaa-mm-dd'."""
    s = _txt(v)
    if not s:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.date()
    # OJO: con dayfirst=True pandas lee '2026-08-01' como 8 de ENERO. El ISO
    # (aaaa-mm-dd) se reconoce primero y se lee tal cual; el resto, dia primero.
    iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", s))
    try:
        d = pd.to_datetime(s[:10], dayfirst=not iso, errors="coerce")
        if d is not pd.NaT and d == d:
            return d.date()
    except Exception:
        pass
    return None


def _mes_a_rango(mes):
    """'2026-08' -> (date(2026,8,1), date(2026,8,31)). Sin mes: el actual."""
    if not mes:
        hoy = date.today()
        mes = f"{hoy.year:04d}-{hoy.month:02d}"
    y, m = int(mes[:4]), int(mes[5:7])
    return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1]), mes


def _cuentas(datos_dir=None):
    c = dict(CUENTAS_DEFECTO)
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "hotel_config.json")
    try:
        with open(ruta, encoding="utf-8") as fh:
            cfg = json.load(fh)
        for clave, campo in (("provision_albaranes", "cuenta_provision_albaranes"),
                             ("provision_comisiones", "cuenta_provision_comisiones")):
            v = _txt(cfg.get(campo))
            if v:
                c[clave] = (v, c[clave][1])
    except Exception:
        pass
    return c


def _tipo_proveedor(nombre, proveedores):
    """FB u OTRAS segun proveedores.xlsx (misma tabla que usa el Libro Diario)."""
    n = _txt(nombre).lower()
    for k, v in proveedores.items():
        if k and (k in n or n in k):
            return v
    return "OTRAS"


def _proveedores(datos_dir=None):
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "proveedores.xlsx")
    out = {}
    try:
        df = pd.read_excel(ruta)
        col_n = next((c for c in df.columns if "nombre" in c.lower()), None)
        col_t = next((c for c in df.columns if "tipo" in c.lower()), None)
        if col_n and col_t:
            for _, r in df.iterrows():
                out[_txt(r[col_n]).lower()] = _txt(r[col_t]).upper() or "OTRAS"
    except Exception:
        pass
    return out


# ── 1 · albaranes sin factura ────────────────────────────────────────────────

def provision_albaranes(mes=None, hotel=None, procesadas_dir=None, reportes_dir=None,
                        datos_dir=None, hoy=None):
    """Lo entregado hasta fin de mes que a dia de hoy sigue sin factura."""
    import almacen_datos
    ini, fin, mes = _mes_a_rango(mes)
    hoy = hoy or date.today()
    ctas = _cuentas(datos_dir)
    provs = _proveedores(datos_dir)

    df = almacen_datos.albaranes(procesadas_dir, reportes_dir, hotel=hotel)
    filas, sin_cruzar = [], 0
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            estado = _txt(r.get("estado")).upper()
            f_ent = _fecha(r.get("fecha_entrega"))
            if f_ent and f_ent > fin:
                continue                        # entregado despues del corte
            if not estado:
                sin_cruzar += 1                  # nunca ha pasado por el cruce
                continue
            if estado != "ALBARAN_SIN_FACTURAR":
                continue
            base = _num(r.get("total_albaran"))
            tipo = _tipo_proveedor(r.get("nombre_proveedor"), provs)
            cta = ctas["gasto_fb"] if tipo == "FB" else ctas["gasto_otros"]
            filas.append({
                "numero_albaran":  _txt(r.get("numero_albaran")) or NF,
                "nombre_proveedor": _txt(r.get("nombre_proveedor")) or "Desconocido",
                "fecha_entrega":   f_ent.isoformat() if f_ent else "",
                "dias":            (hoy - f_ent).days if f_ent else None,
                "importe":         round(base, 2),
                "sin_importe":     base == 0,
                "cuenta_gasto":    cta[0],
                "hotel_id":        _txt(r.get("hotel_id")),
                "archivo":         _txt(r.get("archivo")),
            })

    por_prov = {}
    for f in filas:
        k = (f["nombre_proveedor"], f["cuenta_gasto"], f["hotel_id"])
        p = por_prov.setdefault(k, {"nombre_proveedor": k[0], "cuenta_gasto": k[1],
                                    "hotel_id": k[2], "n_albaranes": 0, "importe": 0.0})
        p["n_albaranes"] += 1
        p["importe"] = round(p["importe"] + f["importe"], 2)

    asientos = []
    cta_prov = ctas["provision_albaranes"]
    for p in por_prov.values():
        if p["importe"] <= 0:
            continue
        concepto = f"Provision {mes} albaranes sin factura — {p['nombre_proveedor']}"
        asientos.append({"fecha": fin.isoformat(), "cuenta": p["cuenta_gasto"],
                         "concepto": concepto, "debe": p["importe"], "haber": 0.0})
        asientos.append({"fecha": fin.isoformat(), "cuenta": cta_prov[0],
                         "concepto": concepto, "debe": 0.0, "haber": p["importe"]})
    total = round(sum(f["importe"] for f in filas), 2)
    return {
        "mes": mes, "corte": fin.isoformat(),
        "filas": filas, "por_proveedor": list(por_prov.values()),
        "total": total, "n": len(filas),
        "sin_importe": sum(1 for f in filas if f["sin_importe"]),
        "sin_cruzar": sin_cruzar,
        "cuenta_provision": {"codigo": cta_prov[0], "descripcion": cta_prov[1]},
        "asientos": asientos,
        "aviso": ("El estado 'sin facturar' es el de HOY: una factura que llegue "
                  "despues del corte ya no sale aqui aunque el mes quedara abierto."),
    }


# ── 2 · comisiones OTA del mes ───────────────────────────────────────────────

def provision_comisiones(mes=None, hotel=None, reportes_dir=None, datos_dir=None):
    """Las liquidaciones OTA verificadas cuyo periodo cae en el mes."""
    import almacen_datos
    ini, fin, mes = _mes_a_rango(mes)
    ctas = _cuentas(datos_dir)

    df = almacen_datos.reporte_verificacion(reportes_dir, hotel=hotel)
    filas = []
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            f_per = _fecha(r.get("periodo_inicio")) or _fecha(r.get("fecha"))
            if not f_per or not (ini <= f_per <= fin):
                continue
            bruto = _num(r.get("importe_bruto"))
            pct_pact = _num(r.get("porcentaje_pactado"))
            fact = _num(r.get("importe_comision_factura"))
            estado = _txt(r.get("estado")).upper()
            if bruto > 0 and pct_pact > 0 and estado not in ("OTA_DESCONOCIDA", "SIN_TARIFA_HOTEL", "SIN_TARIFA_PACTADA"):
                importe, base = round(bruto * pct_pact / 100, 2), "pactado"
            else:
                importe, base = round(fact, 2), "facturado"
            filas.append({
                "nombre_ota":       _txt(r.get("nombre_ota")) or "OTA",
                "numero_factura":   _txt(r.get("numero_factura")) or NF,
                "periodo_inicio":   f_per.isoformat(),
                "importe_bruto":    round(bruto, 2),
                "porcentaje_pactado": pct_pact,
                "importe_facturado": round(fact, 2),
                "importe_provision": importe,
                "base_provision":   base,
                "estado":           estado or NF,
                "discrepancia":     round(_num(r.get("discrepancia_euros")), 2),
                "hotel_id":         _txt(r.get("hotel_id")),
            })

    por_ota = {}
    for f in filas:
        k = (f["nombre_ota"], f["hotel_id"])
        p = por_ota.setdefault(k, {"nombre_ota": k[0], "hotel_id": k[1], "n_facturas": 0,
                                   "importe_provision": 0.0, "importe_facturado": 0.0})
        p["n_facturas"] += 1
        p["importe_provision"] = round(p["importe_provision"] + f["importe_provision"], 2)
        p["importe_facturado"] = round(p["importe_facturado"] + f["importe_facturado"], 2)

    asientos = []
    cta_g, cta_p = ctas["comision_ota"], ctas["provision_comisiones"]
    for p in por_ota.values():
        if p["importe_provision"] <= 0:
            continue
        concepto = f"Provision {mes} comisiones {p['nombre_ota']}"
        asientos.append({"fecha": fin.isoformat(), "cuenta": cta_g[0], "concepto": concepto,
                         "debe": p["importe_provision"], "haber": 0.0})
        asientos.append({"fecha": fin.isoformat(), "cuenta": cta_p[0], "concepto": concepto,
                         "debe": 0.0, "haber": p["importe_provision"]})
    return {
        "mes": mes, "corte": fin.isoformat(),
        "filas": filas, "por_ota": list(por_ota.values()),
        "total": round(sum(f["importe_provision"] for f in filas), 2),
        "total_facturado": round(sum(f["importe_facturado"] for f in filas), 2),
        "n": len(filas),
        "n_pactado": sum(1 for f in filas if f["base_provision"] == "pactado"),
        "cuenta_provision": {"codigo": cta_p[0], "descripcion": cta_p[1]},
        "asientos": asientos,
        "aviso": ("Solo comisiones con liquidacion recibida. Las devengadas sin "
                  "liquidacion necesitan la produccion OTA del PMS (conector)."),
    }


# ── 3 · el fichero del cierre ────────────────────────────────────────────────

def exportar_excel(mes=None, hotel=None, **dirs):
    """Un Excel con las dos provisiones y sus asientos. Devuelve (BytesIO, nombre)."""
    a = provision_albaranes(mes, hotel, dirs.get("procesadas_dir"), dirs.get("reportes_dir"),
                            dirs.get("datos_dir"))
    c = provision_comisiones(mes, hotel, dirs.get("reportes_dir"), dirs.get("datos_dir"))
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([
            {"provision": "Albaranes sin factura", "n": a["n"], "importe": a["total"],
             "cuenta": a["cuenta_provision"]["codigo"], "nota": a["aviso"]},
            {"provision": "Comisiones OTA", "n": c["n"], "importe": c["total"],
             "cuenta": c["cuenta_provision"]["codigo"], "nota": c["aviso"]},
        ]).to_excel(w, index=False, sheet_name="Resumen")
        pd.DataFrame(a["filas"] or [{}]).to_excel(w, index=False, sheet_name="Albaranes")
        pd.DataFrame(c["filas"] or [{}]).to_excel(w, index=False, sheet_name="Comisiones")
        pd.DataFrame((a["asientos"] + c["asientos"]) or [{}]).to_excel(w, index=False, sheet_name="Asientos")
    buf.seek(0)
    return buf, f"provisiones_{a['mes']}.xlsx"
