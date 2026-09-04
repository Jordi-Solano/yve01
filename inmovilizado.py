# -*- coding: utf-8 -*-
"""inmovilizado.py — FF&E / inmovilizado y amortizaciones (OLA B · bloque 4).

Registro de activos en `datos-referencia/inmovilizado.xlsx` (una fila por
activo) y, para cada mes, la cuota de amortizacion lineal, la acumulada, el
valor neto contable y el asiento 68x / 28x.

Reglas (PGC, amortizacion lineal por meses):
  cuota mensual = (coste - valor residual) / (vida util en años × 12)
  empieza el mes de alta (mes completo) y termina cuando la acumulada llega a
  (coste - residual) o el mes anterior a la baja. Nunca se amortiza mas del
  amortizable ni despues de la baja.

Vidas utiles por defecto (config_inmovilizado.json las cambia) y cuentas:
  CONSTRUCCIONES 33 · INSTALACIONES 10 · MAQUINARIA 10 · MOBILIARIO 10 ·
  INFORMATICA 4 · VEHICULOS 6 · LENCERIA_MENAJE 3 · OTRO 5

Ademas: facturas AP del mes cuya cuenta empieza por 2 (o cuyo importe supera el
umbral de activacion) se proponen como ALTAS PENDIENTES: no se dan de alta
solas, se listan para que alguien decida.

Funciones puras; las escrituras (alta/baja) estan en tab_cierre.
"""
import calendar
import json
import os
from datetime import date
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt, _mes_a_rango

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FICHERO = "inmovilizado.xlsx"
COLUMNAS = ["id", "descripcion", "categoria", "fecha_alta", "coste", "valor_residual", "vida_util_anios",
            "cuenta_activo", "cuenta_amortizacion", "cuenta_gasto", "fecha_baja", "documento", "hotel_id", "notas"]

CATEGORIAS = {
    #  categoria:        (vida años, cuenta activo, cuenta amort. acumulada, cuenta gasto, descripcion)
    "CONSTRUCCIONES":  (33, "211", "2811", "681", "Construcciones"),
    "INSTALACIONES":   (10, "212", "2812", "681", "Instalaciones tecnicas"),
    "MAQUINARIA":      (10, "213", "2813", "681", "Maquinaria"),
    "MOBILIARIO":      (10, "216", "2816", "681", "Mobiliario"),
    "INFORMATICA":     (4,  "217", "2817", "681", "Equipos para procesos de informacion"),
    "VEHICULOS":       (6,  "218", "2818", "681", "Elementos de transporte"),
    "LENCERIA_MENAJE": (3,  "219", "2819", "681", "Otro inmovilizado material (lenceria, menaje)"),
    "OTRO":            (5,  "219", "2819", "681", "Otro inmovilizado material"),
}
UMBRAL_ACTIVACION = 300.0     # por debajo, gasto del ejercicio (criterio habitual PYME)
UMBRAL_AVISO = 3000.0         # una factura AP de gasto por encima se pregunta: ¿es un activo?


def config(datos_dir=None):
    cfg = {"umbral_activacion": UMBRAL_ACTIVACION, "umbral_aviso": UMBRAL_AVISO, "vidas": {}}
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "config_inmovilizado.json")
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh) or {}
        if d.get("umbral_activacion"):
            cfg["umbral_activacion"] = float(d["umbral_activacion"])
        if d.get("umbral_aviso"):
            cfg["umbral_aviso"] = float(d["umbral_aviso"])
        for k, v in (d.get("vidas") or {}).items():
            if str(k).upper() in CATEGORIAS and _num(v) > 0:
                cfg["vidas"][str(k).upper()] = float(v)
    except Exception:
        pass
    return cfg


def vida_defecto(categoria, cfg=None):
    cat = str(categoria or "OTRO").upper()
    if cfg and cat in (cfg.get("vidas") or {}):
        return cfg["vidas"][cat]
    return CATEGORIAS.get(cat, CATEGORIAS["OTRO"])[0]


def normalizar_activo(d, cfg=None):
    """Un dict del usuario o del Excel -> fila completa y validada (o ValueError)."""
    cat = str(d.get("categoria") or "OTRO").upper().replace(" ", "_")
    if cat not in CATEGORIAS:
        cat = "OTRO"
    coste = _num(d.get("coste"))
    if coste <= 0:
        raise ValueError("El coste tiene que ser mayor que 0")
    alta = _fecha(d.get("fecha_alta"))
    if alta is None:
        raise ValueError("Falta la fecha de alta (dd/mm/aaaa)")
    vida = _num(d.get("vida_util_anios")) or vida_defecto(cat, cfg)
    resid = _num(d.get("valor_residual"))
    if resid < 0 or resid >= coste:
        raise ValueError("El valor residual tiene que estar entre 0 y el coste")
    baja = _fecha(d.get("fecha_baja"))
    vida_d, c_act, c_am, c_g, _ = CATEGORIAS[cat]
    return {
        "id": _txt(d.get("id")),
        "descripcion": _txt(d.get("descripcion")) or "Activo",
        "categoria": cat,
        "fecha_alta": alta.isoformat(),
        "coste": round(coste, 2),
        "valor_residual": round(resid, 2),
        "vida_util_anios": vida,
        "cuenta_activo": _txt(d.get("cuenta_activo")) or c_act,
        "cuenta_amortizacion": _txt(d.get("cuenta_amortizacion")) or c_am,
        "cuenta_gasto": _txt(d.get("cuenta_gasto")) or c_g,
        "fecha_baja": baja.isoformat() if baja else "",
        "documento": _txt(d.get("documento")),
        "hotel_id": _txt(d.get("hotel_id")),
        "notas": _txt(d.get("notas")),
    }


def _mi(d):
    """Indice de mes absoluto (año*12 + mes) para contar meses completos."""
    return d.year * 12 + d.month


def amortizacion_activo(act, mes):
    """Cuota del mes, acumulada a fin de mes, VNC y estado para UN activo."""
    ini, fin, mes = _mes_a_rango(mes)
    alta = _fecha(act.get("fecha_alta"))
    baja = _fecha(act.get("fecha_baja"))
    coste = _num(act.get("coste")); resid = _num(act.get("valor_residual"))
    vida = _num(act.get("vida_util_anios")) or 1
    amortizable = round(coste - resid, 2)
    n_meses = max(1, int(round(vida * 12)))
    cuota_base = round(amortizable / n_meses, 2)
    if alta is None or alta > fin:
        return {"cuota": 0.0, "acumulada": 0.0, "vnc": coste, "meses": 0, "estado": "NO_ALTA",
                "cuota_base": cuota_base, "amortizable": amortizable}
    # ultimo mes que se amortiza: el del cierre, o el anterior a la baja
    ultimo_mi = _mi(fin)
    if baja and _mi(baja) <= ultimo_mi:
        ultimo_mi = _mi(baja) - 1

    def acumulada_hasta(mi):
        n = max(0, min(n_meses, mi - _mi(alta) + 1))
        if n >= n_meses:
            return amortizable
        return min(amortizable, round(cuota_base * n, 2))

    acum_fin = acumulada_hasta(ultimo_mi)
    acum_ini = acumulada_hasta(min(ultimo_mi, _mi(ini) - 1))
    cuota = round(acum_fin - acum_ini, 2)
    vnc = round(coste - acum_fin, 2)
    if baja and _mi(baja) <= _mi(fin):
        estado = "BAJA"
    elif acum_fin >= amortizable - 0.005:
        estado = "AMORTIZADO"
    else:
        estado = "EN_CURSO"
    return {"cuota": cuota, "acumulada": acum_fin, "vnc": vnc, "meses": max(0, min(n_meses, ultimo_mi - _mi(alta) + 1)),
            "estado": estado, "cuota_base": cuota_base, "amortizable": amortizable}


def amortizar_mes(mes, df_activos, df_ap=None, cfg=None):
    """Devuelve {activos: [...], por_cuenta: [...], asientos: [...], altas_pendientes: [...], resumen}."""
    ini, fin, mes = _mes_a_rango(mes)
    cfg = cfg or {"umbral_activacion": UMBRAL_ACTIVACION, "umbral_aviso": UMBRAL_AVISO, "vidas": {}}
    activos = []
    if df_activos is not None and not df_activos.empty:
        for _, r in df_activos.iterrows():
            d = r.to_dict()
            try:
                act = normalizar_activo(d, cfg)
            except ValueError as e:
                activos.append({"id": _txt(d.get("id")), "descripcion": _txt(d.get("descripcion")), "categoria": _txt(d.get("categoria")),
                                "error": str(e), "cuota": 0.0, "acumulada": 0.0, "vnc": 0.0, "estado": "ERROR"})
                continue
            a = amortizacion_activo(act, mes)
            activos.append({**act, **a})
    por = {}
    for a in activos:
        if a.get("estado") in ("ERROR", "NO_ALTA") or not a.get("cuota"):
            continue
        k = (a["cuenta_gasto"], a["cuenta_amortizacion"], a["categoria"])
        p = por.setdefault(k, {"cuenta_gasto": k[0], "cuenta_amortizacion": k[1], "categoria": k[2], "cuota": 0.0, "n": 0})
        p["cuota"] = round(p["cuota"] + a["cuota"], 2); p["n"] += 1
    asientos = []
    for p in por.values():
        concepto = f"Amortizacion {mes} — {CATEGORIAS.get(p['categoria'], CATEGORIAS['OTRO'])[4]}"
        asientos.append({"fecha": fin.isoformat(), "cuenta": p["cuenta_gasto"], "desc_cuenta": "Amortizacion del inmovilizado material",
                         "concepto": concepto, "debe": p["cuota"], "haber": 0.0})
        asientos.append({"fecha": fin.isoformat(), "cuenta": p["cuenta_amortizacion"], "desc_cuenta": "Amortizacion acumulada",
                         "concepto": concepto, "debe": 0.0, "haber": p["cuota"]})
    # altas pendientes: facturas AP del mes con cuenta 2xx o por encima del umbral y no registradas
    docs = {_txt(a.get("documento")).upper() for a in activos if _txt(a.get("documento"))}
    pend = []
    if df_ap is not None and not df_ap.empty:
        for _, r in df_ap.iterrows():
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            f = _fecha(fecha)
            if not f or not (ini <= f <= fin):
                continue
            cta = _txt(r.get("cuenta_debe_gasto")) or _txt(r.get("cuenta_contable"))
            cta = cta[:-2] if cta.endswith(".0") else cta
            base = _num(r.get("base_imponible")) or _num(r.get("total_factura"))
            num = _txt(r.get("numero_factura"))
            es_2xx = cta.startswith("2") and len(cta) >= 3
            if (es_2xx or base >= float(cfg.get("umbral_aviso", UMBRAL_AVISO))) and num.upper() not in docs:
                pend.append({"numero_factura": num, "proveedor": _txt(r.get("nombre_proveedor")), "fecha": f.isoformat(),
                             "base": round(base, 2), "cuenta": cta,
                             "motivo": "cuenta de inmovilizado (2xx)" if es_2xx else "importe alto: ¿es un activo?"})
    total_cuota = round(sum(a.get("cuota", 0) for a in activos), 2)
    resumen = {"mes": mes, "n_activos": len(activos), "n_en_curso": sum(1 for a in activos if a.get("estado") == "EN_CURSO"),
               "n_amortizados": sum(1 for a in activos if a.get("estado") == "AMORTIZADO"),
               "n_error": sum(1 for a in activos if a.get("estado") == "ERROR"),
               "coste_total": round(sum(_num(a.get("coste")) for a in activos), 2),
               "acumulada_total": round(sum(a.get("acumulada", 0) for a in activos), 2),
               "vnc_total": round(sum(a.get("vnc", 0) for a in activos if a.get("estado") != "ERROR"), 2),
               "cuota_mes": total_cuota, "altas_pendientes": len(pend)}
    return {"mes": mes, "activos": activos, "por_cuenta": list(por.values()), "asientos": asientos,
            "altas_pendientes": pend, "resumen": resumen}


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([res["resumen"]]).to_excel(w, index=False, sheet_name="Resumen")
        pd.DataFrame(res["activos"] or [{}]).to_excel(w, index=False, sheet_name="Activos")
        pd.DataFrame(res["asientos"] or [{}]).to_excel(w, index=False, sheet_name="Asientos")
        pd.DataFrame(res["altas_pendientes"] or [{}]).to_excel(w, index=False, sheet_name="Altas pendientes")
    buf.seek(0)
    return buf, f"inmovilizado_{res['mes']}.xlsx"
