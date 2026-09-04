# -*- coding: utf-8 -*-
"""cierre_mes.py — asientos del mes y reconciliacion de cuentas (OLA B · bloque 1).

Monta el Libro Diario del mes con TODO lo que Yve ya tiene por documento y lo
reconcilia cuenta a cuenta contra lo que lo justifica. SOLO LEE. Las funciones
que calculan son puras (entran DataFrames, sale un dict) para poder probarlas
sin Flask; `recoger_fuentes` es la unica que toca disco.

Fuentes y asientos (PGC espanol, cuentas de `datos-referencia/plan_cuentas.xlsx`):
  AP     factura de proveedor      6xx gasto (D) · 472 IVA soportado (D) · 400 Proveedores (H)
  OTA    factura de comision       628 (D) · 410 Acreedores (H); el IVA depende del REGIMEN de
                                   la OTA (`regimen_ota`): "es" = OTA española, la comision lleva
                                   el IVA dentro; "ue" / "no_ue" = inversion del sujeto pasivo
                                   (472 D y 477 H por el 21 %). Defectos: Booking ue, Expedia
                                   no_ue, HotelBeds/Hotusa es; lo desconocido se trata como "es"
                                   (no se inventa una ISP). config_cierre.json "otas": {nombre: regimen}.
  F&B    ventas TPV del dia        570 Caja (D) · 700 Ventas F&B (H) · 477 IVA 10 % (H)
  AR     factura a credito emitida 430 Clientes (D) · 705 Alojamiento / 700 F&B (H) · 477 IVA 10 % (H)
  AR     cobro (fecha_cobro)       572 Banco (D) · 430 Clientes (H)
  BANCO  movimiento CONCILIADO     pago: 400 (D) · 572 (H)   cobro: 572 (D) · 430 (H)
  PROV   provisiones (provisiones.py)  albaran sin factura y comisiones OTA devengadas

Lo que Yve NO puede asentar porque el dato vive en otro sistema (PMS, caja,
tarjetas) NO se inventa: en la reconciliacion sale como "sin dato" o como
diferencia explicada (p. ej. ingresos de habitaciones segun DRR vs lo asentado).
"""
import calendar
import json
import os
from datetime import date
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt, _mes_a_rango

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

IVA_GENERAL = 21.0
IVA_REDUCIDO = 10.0          # alojamiento y restauracion en Espana

CUENTAS_BASE = {
    "400":  "Proveedores",
    "410":  "Acreedores por prestaciones de servicios",
    "4009": "Proveedores, facturas pendientes de recibir",
    "4109": "Acreedores, facturas pendientes de recibir",
    "430":  "Clientes",
    "472":  "H.P. IVA soportado",
    "477":  "H.P. IVA repercutido",
    "570":  "Caja",
    "572":  "Bancos c/c",
    "600":  "Compras de mercaderias F&B",
    "628":  "Comisiones de agencias y OTAs",
    "629":  "Otros servicios",
    "700":  "Ventas F&B",
    "705":  "Prestaciones de servicios — Alojamiento",
}
CONFIG_FILE = "config_cierre.json"
REGIMEN_OTA_DEFECTO = {"booking": "ue", "expedia": "no_ue", "hotels.com": "no_ue", "agoda": "no_ue",
                       "airbnb": "ue", "trivago": "ue", "hotelbeds": "es", "hotusa": "es", "despegar": "no_ue"}


def regimen_ota(nombre, cfg=None):
    """'es' (IVA incluido), 'ue' o 'no_ue' (inversion del sujeto pasivo)."""
    n = _txt(nombre).lower()
    tabla = dict(REGIMEN_OTA_DEFECTO)
    for k, v in ((cfg or {}).get("otas") or {}).items():
        if str(v).lower() in ("es", "ue", "no_ue"):
            tabla[str(k).lower()] = str(v).lower()
    if str((cfg or {}).get("ota_iva", "")).lower() == "incluido":
        return "es"
    for k, v in tabla.items():
        if k in n:
            return v
    return "es"


# ── configuracion y plan ─────────────────────────────────────────────────────
def config_cierre(datos_dir=None):
    cfg = {"ota_iva": "", "iva_fb": IVA_REDUCIDO, "iva_alojamiento": IVA_REDUCIDO, "otas": {}}
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), CONFIG_FILE)
    try:
        with open(ruta, encoding="utf-8") as fh:
            cfg.update({k: v for k, v in (json.load(fh) or {}).items() if v not in (None, "")})
    except Exception:
        pass
    return cfg


def plan_cuentas(datos_dir=None):
    """codigo -> descripcion. El plan del hotel manda; lo que falte, de CUENTAS_BASE."""
    plan = dict(CUENTAS_BASE)
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "plan_cuentas.xlsx")
    try:
        df = pd.read_excel(ruta)
        for _, r in df.iterrows():
            c = _cuenta_str(r.get("codigo_cuenta"))
            if c:
                plan[c] = _txt(r.get("descripcion")) or plan.get(c, "")
    except Exception:
        pass
    return plan


def _cuenta_str(v):
    s = _txt(v)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _en_mes(fecha, ini, fin):
    f = _fecha(fecha)
    return f is not None and ini <= f <= fin


def _r(x):
    return round(float(x or 0), 2)


# ── generador de asientos (puro) ─────────────────────────────────────────────
class _Diario:
    def __init__(self, plan):
        self.plan = plan
        self.asientos = []
        self.num = 0
        self.cuentas_fuera_plan = set()

    def nuevo(self, fecha, concepto, documento, origen, lineas, hotel=""):
        """lineas: [(cuenta, debe, haber)]. Un asiento que no cuadra NO se escribe."""
        lineas = [(str(c), _r(d), _r(h)) for c, d, h in lineas if _r(d) or _r(h)]
        if not lineas:
            return False
        if abs(sum(d for _, d, _ in lineas) - sum(h for _, _, h in lineas)) > 0.011:
            return False
        self.num += 1
        for c, d, h in lineas:
            if c not in self.plan:
                self.cuentas_fuera_plan.add(c)
            self.asientos.append({
                "num": self.num, "fecha": fecha.isoformat() if hasattr(fecha, "isoformat") else str(fecha),
                "cuenta": c, "desc_cuenta": self.plan.get(c, "(fuera del plan)"),
                "concepto": concepto, "debe": d, "haber": h,
                "documento": _txt(documento), "origen": origen, "hotel_id": _txt(hotel),
            })
        return True


def generar_asientos(mes, fuentes, plan=None, cfg=None):
    """fuentes: dict con DataFrames opcionales:
         ap, ar_ota, ventas_fb, reservas, banco  y listas: provisiones (asientos ya hechos)
    Devuelve {mes, asientos, resumen, fuentes, cuentas_fuera_plan, avisos}."""
    ini, fin, mes = _mes_a_rango(mes)
    plan = plan or dict(CUENTAS_BASE)
    cfg = cfg or {"ota_iva": "", "iva_fb": IVA_REDUCIDO, "iva_alojamiento": IVA_REDUCIDO, "otas": {}}
    D = _Diario(plan)
    cont = {"ap": 0, "ar_ota": 0, "ventas_fb": 0, "ar_facturas": 0, "ar_cobros": 0,
            "banco": 0, "provisiones": 0}
    saltados = {"ap_sin_total": 0, "ap_sin_cuadrar": 0, "ar_ota_sin_importe": 0}
    avisos = []

    # ── AP ───────────────────────────────────────────────────────────────
    ap = fuentes.get("ap")
    if ap is not None and not ap.empty:
        for _, r in ap.iterrows():
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            if not _en_mes(fecha, ini, fin):
                continue
            total = _num(r.get("total_factura"))
            if not total:
                saltados["ap_sin_total"] += 1
                continue
            base = _num(r.get("base_imponible"))
            iva = _num(r.get("cuota_iva"))
            if not base and not iva:
                pct = _num(r.get("porcentaje_iva")) or IVA_GENERAL
                base = _r(total / (1 + pct / 100)); iva = _r(total - base)
            elif not iva:
                iva = _r(total - base)
            elif not base:
                base = _r(total - iva)
            cta = _cuenta_str(r.get("cuenta_debe_gasto")) or _cuenta_str(r.get("cuenta_contable"))
            if not cta or cta.upper() == "REVISAR_MANUAL":
                cta = "600" if _txt(r.get("tipo_proveedor")).upper() == "FB" else "629"
            num = _txt(r.get("numero_factura")) or _txt(r.get("archivo"))
            prov = _txt(r.get("nombre_proveedor")) or "proveedor"
            ok = D.nuevo(_fecha(fecha), f"Fra. {num} — {prov}", num, "AP",
                         [(cta, base, 0), ("472", iva, 0), ("400", 0, total)], r.get("hotel_id"))
            if ok:
                cont["ap"] += 1
            else:
                saltados["ap_sin_cuadrar"] += 1

    # ── comisiones OTA ───────────────────────────────────────────────────
    otas_isp = set()
    ar = fuentes.get("ar_ota")
    if ar is not None and not ar.empty:
        for _, r in ar.iterrows():
            fecha = r.get("fecha") if _txt(r.get("fecha")) else r.get("periodo_fin")
            if not _en_mes(fecha, ini, fin):
                continue
            imp = _num(r.get("importe_comision")) or _num(r.get("importe_comision_factura"))
            if imp <= 0:
                saltados["ar_ota_sin_importe"] += 1
                continue
            num = _txt(r.get("numero_factura")) or "s/n"
            ota = _txt(r.get("nombre_ota")) or "OTA"
            reg = regimen_ota(ota, cfg)
            if reg == "es":     # OTA española: la comision lleva el IVA dentro
                base = _r(imp / (1 + IVA_GENERAL / 100)); iva = _r(imp - base)
                lineas = [("628", base, 0), ("472", iva, 0), ("410", 0, imp)]
            else:               # inversion del sujeto pasivo: la comision es la base; el IVA se autorrepercute
                iva = _r(imp * IVA_GENERAL / 100)
                lineas = [("628", imp, 0), ("472", iva, 0), ("477", 0, iva), ("410", 0, imp)]
                otas_isp.add(ota)
            if D.nuevo(_fecha(fecha), f"Comision {ota} — {num}", num, "OTA", lineas, r.get("hotel_id")):
                cont["ar_ota"] += 1

    # ── ventas F&B (TPV), un asiento por dia ─────────────────────────────
    vf = fuentes.get("ventas_fb")
    if vf is not None and not vf.empty and "total_venta" in vf.columns:
        por_dia = {}
        for _, r in vf.iterrows():
            f = _fecha(r.get("fecha"))
            if f is None or not (ini <= f <= fin):
                continue
            por_dia[f] = _r(por_dia.get(f, 0) + _num(r.get("total_venta")))
        pct = float(cfg.get("iva_fb", IVA_REDUCIDO))
        for f in sorted(por_dia):
            tot = por_dia[f]
            if tot <= 0:
                continue
            base = _r(tot / (1 + pct / 100)); iva = _r(tot - base)
            if D.nuevo(f, f"Ventas F&B TPV {f.isoformat()}", f"TPV-{f.isoformat()}", "FB",
                       [("570", tot, 0), ("700", 0, base), ("477", 0, iva)]):
                cont["ventas_fb"] += 1

    # ── AR direct bill: facturas emitidas y cobros ───────────────────────
    # El cobro registrado en AR Real (fecha_cobro) y el movimiento CONCILIADO
    # del banco con esa factura son el MISMO dinero: si el banco lo tiene, manda
    # el banco y el cobro de AR no se asienta (si no, 572 y 430 se duplican).
    cobrados_en_banco = set()
    bk0 = fuentes.get("banco")
    if bk0 is not None and not bk0.empty:
        for _, r in bk0.iterrows():
            if _txt(r.get("estado")).upper() == "CONCILIADO" and _num(r.get("importe")) > 0:
                ref = _txt(r.get("factura_ref")) or _txt(r.get("referencia"))
                if ref:
                    cobrados_en_banco.add(ref.upper())
    rv = fuentes.get("reservas")
    if rv is not None and not rv.empty:
        pct_h = float(cfg.get("iva_alojamiento", IVA_REDUCIDO)); pct_f = float(cfg.get("iva_fb", IVA_REDUCIDO))
        for _, r in rv.iterrows():
            estado = _txt(r.get("estado")).upper()
            if estado in ("PENDIENTE_FACTURA", ""):
                continue
            num = _txt(r.get("numero_reserva")) or _txt(r.get("numero")) or "s/n"
            cli = _txt(r.get("cliente")) or "cliente"
            total = _num(r.get("total")) or _num(r.get("importe"))
            f_em = r.get("fecha_emision") if _txt(r.get("fecha_emision")) else r.get("fecha_entrada")
            if total > 0 and _en_mes(f_em, ini, fin):
                hab = _num(r.get("importe_habitaciones")); fb = _num(r.get("importe_fb")); ext = _num(r.get("importe_extras"))
                if not (hab or fb or ext):
                    hab = total
                # reparto proporcional si el detalle no suma el total
                suma = _r(hab + fb + ext)
                if suma and abs(suma - total) > 0.011:
                    k = total / suma; hab, fb, ext = _r(hab * k), _r(fb * k), _r(ext * k)
                b_h = _r((hab + ext) / (1 + pct_h / 100)); i_h = _r(hab + ext - b_h)
                b_f = _r(fb / (1 + pct_f / 100)); i_f = _r(fb - b_f)
                iva = _r(i_h + i_f)
                # ajuste de redondeo al ultimo centimo sobre el IVA
                dif = _r(total - (b_h + b_f + iva)); iva = _r(iva + dif)
                if D.nuevo(_fecha(f_em), f"Fra. {num} — {cli}", num, "AR",
                           [("430", total, 0), ("705", 0, b_h), ("700", 0, b_f), ("477", 0, iva)], r.get("hotel_id")):
                    cont["ar_facturas"] += 1
            if estado in ("COBRADO", "COBRADA") and total > 0 and _en_mes(r.get("fecha_cobro"), ini, fin) \
                    and num.upper() not in cobrados_en_banco:
                if D.nuevo(_fecha(r.get("fecha_cobro")), f"Cobro fra. {num} — {cli}", num, "AR",
                           [("572", total, 0), ("430", 0, total)], r.get("hotel_id")):
                    cont["ar_cobros"] += 1

    # ── banco: solo lo conciliado (lo demas no se sabe que es) ───────────
    bk = fuentes.get("banco")
    if bk is not None and not bk.empty:
        for _, r in bk.iterrows():
            if _txt(r.get("estado")).upper() != "CONCILIADO":
                continue
            if not _en_mes(r.get("fecha"), ini, fin):
                continue
            imp = _num(r.get("importe"))
            if not imp:
                continue
            ref = _txt(r.get("factura_ref")) or _txt(r.get("referencia"))
            concepto = _txt(r.get("concepto"))[:60]
            if imp < 0:
                lineas = [("400", -imp, 0), ("572", 0, -imp)]; txt = f"Pago {ref or concepto}"
            else:
                lineas = [("572", imp, 0), ("430", 0, imp)]; txt = f"Cobro {ref or concepto}"
            if D.nuevo(_fecha(r.get("fecha")), txt, ref, "BANCO", lineas, r.get("hotel_id")):
                cont["banco"] += 1

    # ── provisiones (ya vienen como asientos de provisiones.py) ──────────
    for bloque in fuentes.get("provisiones") or []:
        filas = bloque.get("asientos") or []
        # provisiones.py da las lineas sueltas; se agrupan de dos en dos (gasto, provision)
        for i in range(0, len(filas) - 1, 2):
            a, b = filas[i], filas[i + 1]
            if D.nuevo(_fecha(a.get("fecha")) or fin, a.get("concepto", "Provision"), "", "PROVISION",
                       [(a["cuenta"], a.get("debe", 0), a.get("haber", 0)),
                        (b["cuenta"], b.get("debe", 0), b.get("haber", 0))]):
                cont["provisiones"] += 1

    debe = _r(sum(a["debe"] for a in D.asientos)); haber = _r(sum(a["haber"] for a in D.asientos))
    if saltados["ap_sin_total"]:
        avisos.append(f"{saltados['ap_sin_total']} factura(s) AP sin total: no se asientan")
    if saltados["ap_sin_cuadrar"]:
        avisos.append(f"{saltados['ap_sin_cuadrar']} factura(s) AP con base+IVA distinto del total: no se asientan")
    if D.cuentas_fuera_plan:
        avisos.append("Cuentas usadas que no estan en plan_cuentas.xlsx: " + ", ".join(sorted(D.cuentas_fuera_plan)))
    if otas_isp:
        avisos.append("Comisiones con inversion del sujeto pasivo (472/477 al 21 %): " + ", ".join(sorted(otas_isp))
                      + ". Si alguna factura con IVA espanol, ponla como \"es\" en config_cierre.json (\"otas\").")
    return {
        "mes": mes, "desde": ini.isoformat(), "hasta": fin.isoformat(),
        "asientos": D.asientos, "n_asientos": D.num, "n_lineas": len(D.asientos),
        "debe": debe, "haber": haber, "cuadra": abs(debe - haber) < 0.011,
        "fuentes": cont, "saltados": saltados,
        "cuentas_fuera_plan": sorted(D.cuentas_fuera_plan), "avisos": avisos,
    }


def mayor(asientos, plan=None):
    plan = plan or CUENTAS_BASE
    por = {}
    for a in asientos:
        p = por.setdefault(a["cuenta"], {"cuenta": a["cuenta"], "descripcion": plan.get(a["cuenta"], a.get("desc_cuenta", "")),
                                         "debe": 0.0, "haber": 0.0, "n": 0})
        p["debe"] = _r(p["debe"] + a["debe"]); p["haber"] = _r(p["haber"] + a["haber"]); p["n"] += 1
    out = []
    for p in sorted(por.values(), key=lambda x: x["cuenta"]):
        p["saldo"] = _r(p["debe"] - p["haber"])      # positivo = deudor
        out.append(p)
    return out


# ── reconciliacion (pura) ────────────────────────────────────────────────────
def _check(cuenta, concepto, libro, justificado, nota="", tolerancia=0.011):
    if justificado is None:
        return {"cuenta": cuenta, "concepto": concepto, "libro": _r(libro), "justificado": None,
                "diferencia": None, "estado": "SIN_DATO", "nota": nota}
    dif = _r(libro - justificado)
    return {"cuenta": cuenta, "concepto": concepto, "libro": _r(libro), "justificado": _r(justificado),
            "diferencia": dif, "estado": "CUADRA" if abs(dif) <= tolerancia else "DIFERENCIA", "nota": nota}


def reconciliar(mes, res, fuentes, drr=None):
    """Cuenta a cuenta: lo asentado (libro) contra lo que lo justifica."""
    ini, fin, mes = _mes_a_rango(mes)
    my = {m["cuenta"]: m for m in mayor(res["asientos"])}
    def s(c, lado=None):
        m = my.get(c, {"debe": 0.0, "haber": 0.0})
        return m["debe"] if lado == "D" else m["haber"] if lado == "H" else _r(m["debe"] - m["haber"])
    checks = []

    # 0 · el diario cuadra
    checks.append(_check("*", "Suma del Diario (debe = haber)", res["debe"], res["haber"],
                         "Cada asiento se escribe solo si cuadra; esto comprueba el total."))

    # 400 · proveedores: facturado en el mes - pagado (conciliado) en el mes
    ap = fuentes.get("ap")
    fact_ap = 0.0
    if ap is not None and not ap.empty:
        for _, r in ap.iterrows():
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            if _en_mes(fecha, ini, fin):
                fact_ap = _r(fact_ap + _num(r.get("total_factura")))
    checks.append(_check("400", "Proveedores: facturas AP del mes (haber)", s("400", "H"), fact_ap,
                         "Facturas AP con fecha en el mes; las sin total o sin cuadrar no entran (ver avisos)."))
    checks.append(_check("572", "Banco: pagos y cobros conciliados (movimiento neto)", s("572"),
                         _saldo_banco_conciliado(fuentes.get("banco"), ini, fin),
                         "Solo movimientos CONCILIADO del extracto entran en el Diario."))
    # movimientos del extracto SIN conciliar en el mes: dinero que se movio y no tiene asiento
    n_pend, imp_pend = _banco_pendiente(fuentes.get("banco"), ini, fin)
    checks.append({"cuenta": "572", "concepto": "Movimientos del extracto sin conciliar (sin asiento)",
                   "libro": 0.0, "justificado": _r(imp_pend), "diferencia": _r(-imp_pend),
                   "estado": "CUADRA" if n_pend == 0 else "PENDIENTE",
                   "nota": f"{n_pend} movimiento(s) por {imp_pend:,.2f} EUR esperan conciliacion en la pestaña Banco."})

    # 472 / 477 · IVA
    iva_ap = 0.0
    if ap is not None and not ap.empty:
        for _, r in ap.iterrows():
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            if _en_mes(fecha, ini, fin) and _num(r.get("total_factura")):
                iva = _num(r.get("cuota_iva"))
                if not iva:
                    base = _num(r.get("base_imponible")); tot = _num(r.get("total_factura"))
                    iva = _r(tot - base) if base else _r(tot - tot / (1 + (_num(r.get("porcentaje_iva")) or IVA_GENERAL) / 100))
                iva_ap = _r(iva_ap + iva)
    iva_lib_ap = _r(sum(a["debe"] for a in res["asientos"] if a["cuenta"] == "472" and a["origen"] == "AP"))
    checks.append(_check("472", "IVA soportado de facturas AP", iva_lib_ap, iva_ap,
                         "Suma de cuota_iva de las facturas AP del mes."))
    checks.append({"cuenta": "477", "concepto": "IVA repercutido del mes (ventas F&B, facturas AR, ISP OTA)",
                   "libro": s("477", "H"), "justificado": s("477", "H"), "diferencia": 0.0, "estado": "INFO",
                   "nota": "Base del modelo 303 (bloque fiscal de la Ola B)."})

    # 430 · clientes: facturado - cobrado en el mes
    checks.append(_check("430", "Clientes: facturas AR emitidas en el mes (debe)", s("430", "D"),
                         _facturado_ar(fuentes.get("reservas"), ini, fin),
                         "reservas_credito.xlsx con fecha_emision en el mes (FACTURADO/COBRADO/PENDIENTE de cobro)."))

    # 705 · alojamiento vs DRR (dato del PMS)
    if drr and drr.get("rooms_revenue_mtd") is not None:
        checks.append(_check("705", "Ingresos de alojamiento: asentado vs DRR (Rooms Revenue MTD)",
                             s("705", "H"), drr["rooms_revenue_mtd"],
                             "El DRR viene del PMS: la diferencia son ventas que no entran por documento (mostrador, tarjetas, OTAs cobradas).",
                             tolerancia=1.0))
    else:
        checks.append(_check("705", "Ingresos de alojamiento: asentado vs DRR", s("705", "H"), None,
                             "Sin DRR del mes no se puede contrastar. Sube el DRR en la pestaña DRR."))
    # 700 · F&B vs TPV
    vf = fuentes.get("ventas_fb"); tpv = 0.0
    if vf is not None and not vf.empty and "total_venta" in vf.columns:
        for _, r in vf.iterrows():
            if _en_mes(r.get("fecha"), ini, fin):
                tpv = _r(tpv + _num(r.get("total_venta")))
    tpv_lib = _r(sum(a["debe"] for a in res["asientos"] if a["cuenta"] == "570" and a["origen"] == "FB"))
    checks.append(_check("570/700", "Ventas F&B: asentado (caja) vs TPV", tpv_lib, tpv,
                         "Total de ventas_fb_diarias.xlsx en el mes (IVA incluido)."))

    # cuentas fuera de plan
    for c in res.get("cuentas_fuera_plan") or []:
        checks.append({"cuenta": c, "concepto": "Cuenta usada que no esta en plan_cuentas.xlsx",
                       "libro": s(c), "justificado": None, "diferencia": None, "estado": "REVISAR",
                       "nota": "Añadela al plan o corrige la asignacion de la factura."})

    n = {k: sum(1 for c in checks if c["estado"] == k) for k in ("CUADRA", "DIFERENCIA", "PENDIENTE", "SIN_DATO", "REVISAR", "INFO")}
    return {"mes": mes, "checks": checks, "resumen": n,
            "mayor": mayor(res["asientos"]),
            "ok": n["DIFERENCIA"] == 0 and n["REVISAR"] == 0}


def _saldo_banco_conciliado(bk, ini, fin):
    if bk is None or bk.empty:
        return 0.0
    tot = 0.0
    for _, r in bk.iterrows():
        if _txt(r.get("estado")).upper() == "CONCILIADO" and _en_mes(r.get("fecha"), ini, fin):
            tot = _r(tot + _num(r.get("importe")))
    return tot


def _banco_pendiente(bk, ini, fin):
    if bk is None or bk.empty:
        return 0, 0.0
    n, tot = 0, 0.0
    for _, r in bk.iterrows():
        if _txt(r.get("estado")).upper() != "CONCILIADO" and _en_mes(r.get("fecha"), ini, fin) and _num(r.get("importe")):
            n += 1; tot = _r(tot + _num(r.get("importe")))
    return n, tot


def _facturado_ar(rv, ini, fin):
    if rv is None or rv.empty:
        return 0.0
    tot = 0.0
    for _, r in rv.iterrows():
        estado = _txt(r.get("estado")).upper()
        if estado in ("PENDIENTE_FACTURA", ""):
            continue
        f_em = r.get("fecha_emision") if _txt(r.get("fecha_emision")) else r.get("fecha_entrada")
        if _en_mes(f_em, ini, fin):
            tot = _r(tot + (_num(r.get("total")) or _num(r.get("importe"))))
    return tot


# ── fuentes reales (lo unico que toca disco) ─────────────────────────────────
def recoger_fuentes(mes, hotel=None, procesadas_dir=None, reportes_dir=None, datos_dir=None):
    import almacen_datos as ALM
    f = {}
    try:
        f["ap"] = ALM.facturas_ap(procesadas_dir, reportes_dir, hotel=hotel or None)
    except Exception:
        f["ap"] = pd.DataFrame()
    try:
        f["ar_ota"] = ALM.facturas_ar(procesadas_dir, reportes_dir, hotel=hotel or None)
    except Exception:
        f["ar_ota"] = pd.DataFrame()
    dd = datos_dir or os.path.join(BASE_DIR, "datos-referencia")
    for clave, fichero in (("ventas_fb", "ventas_fb_diarias.xlsx"), ("reservas", "reservas_credito.xlsx")):
        try:
            df = pd.read_excel(os.path.join(dd, fichero))
            if hotel:
                df = ALM._filtrar_hotel(df, hotel)
            f[clave] = df
        except Exception:
            f[clave] = pd.DataFrame()
    try:
        bk, _ = ALM.movimientos_banco(datos_dir=dd, reportes_dir=reportes_dir)
        f["banco"] = ALM._filtrar_hotel(bk, hotel) if hotel else bk
    except Exception:
        f["banco"] = pd.DataFrame()
    try:
        import provisiones as PV
        f["provisiones"] = [PV.provision_albaranes(mes, hotel, procesadas_dir, reportes_dir, dd),
                            PV.provision_comisiones(mes, hotel, reportes_dir, dd)]
    except Exception:
        f["provisiones"] = []
    return f


def drr_del_mes(mes, hotel=None):
    """Rooms Revenue MTD del ultimo DRR del hotel, si es de ese mes. Solo lectura."""
    try:
        import dashboard as D
        ruta = D.drr_del_hotel(hotel=hotel)
        if not ruta:
            return None
        st = D._leer_drr_stats(ruta) or {}
        met = st.get("metricas") or {}
        rev = D.num_drr((met.get("Rooms Revenue") or {}).get("mtd"))
        if rev is None:
            rev = D.num_drr((met.get("Total Revenue") or {}).get("mtd"))
            concepto = "Total Revenue"
        else:
            concepto = "Rooms Revenue"
        # el mes del DRR: la fecha de cualquiera de sus dias
        fechas = [_fecha(d.get("fecha")) for d in (st.get("dias") or []) if d.get("fecha")]
        fechas = [f for f in fechas if f]
        if fechas and fechas[-1].strftime("%Y-%m") != mes:
            return None
        return {"rooms_revenue_mtd": rev, "concepto": concepto, "fichero": os.path.basename(ruta)} if rev is not None else None
    except Exception:
        return None


def cierre_completo(mes, hotel=None, **dirs):
    dd = dirs.get("datos_dir")
    plan = plan_cuentas(dd); cfg = config_cierre(dd)
    fuentes = recoger_fuentes(mes, hotel, **dirs)
    res = generar_asientos(mes, fuentes, plan, cfg)
    rec = reconciliar(mes, res, fuentes, drr_del_mes(res["mes"], hotel))
    return res, rec


def exportar_excel(res, rec):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(res["asientos"] or [{}]).to_excel(w, index=False, sheet_name="Libro Diario")
        pd.DataFrame(rec["mayor"] or [{}]).to_excel(w, index=False, sheet_name="Mayor")
        pd.DataFrame(rec["checks"] or [{}]).to_excel(w, index=False, sheet_name="Reconciliacion")
        pd.DataFrame([{"mes": res["mes"], "asientos": res["n_asientos"], "debe": res["debe"], "haber": res["haber"],
                       **{f"fuente_{k}": v for k, v in res["fuentes"].items()},
                       "avisos": " | ".join(res["avisos"])}]).to_excel(w, index=False, sheet_name="Resumen")
    buf.seek(0)
    return buf, f"cierre_{res['mes']}_asientos.xlsx"
