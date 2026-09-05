# -*- coding: utf-8 -*-
"""fiscal.py — bloque fiscal del cierre (OLA B · bloque 6): modelo 303, 349 y SII.

Sale de las MISMAS fuentes que los asientos del mes (cierre_mes.recoger_fuentes),
asi la cuota de IVA que declara el 303 es la que esta en el libro (472/477).
Nada de aqui se envia a la AEAT: el SII se PREPARA como libros registro
(expedidas / recibidas) listos para el envio, que exige certificado digital y
lo hace la gestoria o el conector, no Yve.

Criterio (revisable por la gestoria — cada cifra lleva su casilla):
  · Ventas alojamiento y F&B: IVA reducido (10 %) salvo config → casillas 04-06.
    Si la config pone 21 % en algo, va a 07-09.
  · Facturas AP: IVA soportado corriente → casillas 28-29 (base, cuota).
  · Comisiones OTA (cierre_mes.regimen_ota):
      'es'    → IVA incluido, soportado corriente 28-29.
      'ue'    → adquisicion intracomunitaria de servicios: devengado 10-11 y
                deducible 36-37; ademas fila del 349 con clave 'S' (Servicios).
      'no_ue' → inversion del sujeto pasivo (art. 84.Uno.2º): devengado 12-13,
                deducible 28-29. No va al 349.
  · Resultado (casilla 46 = 27 − 45): a ingresar si > 0, a compensar si < 0.
    Compensaciones de periodos anteriores (casilla 67) NO se conocen: aviso.
  · Bienes de inversion (30-31), importaciones (32-35) y prorrata (44): no se
    distinguen en los datos de Yve → 0 y aviso. Las compras de inmovilizado
    van hoy en 28-29 como corriente: la gestoria las mueve si toca.
  · NIF: las fuentes de Yve no traen NIF de proveedor/cliente/OTA. Se deja
    en blanco y se marca como pendiente; sin NIF el 349 y el SII no se pueden
    presentar. Se pueden completar en config_fiscal.json ("nif": {"booking": "NL..."}).
"""
import json
import os
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt, _mes_a_rango


def _es(v):
    """1234.5 -> '1.234,50' (formato español, como el resto del panel)."""
    return f"{float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
from cierre_mes import IVA_GENERAL, IVA_REDUCIDO, regimen_ota, config_cierre, _r

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = "config_fiscal.json"

# casillas del 303 que usa Yve (base, cuota)
CASILLAS = {
    "dev_4":   ("01", "03", "Regimen general 4 %"),
    "dev_10":  ("04", "06", "Regimen general 10 %"),
    "dev_21":  ("07", "09", "Regimen general 21 %"),
    "dev_aib": ("10", "11", "Adquisiciones intracomunitarias de bienes y servicios"),
    "dev_isp": ("12", "13", "Otras operaciones con inversion del sujeto pasivo"),
    "ded_int": ("28", "29", "IVA deducible: operaciones interiores corrientes"),
    "ded_inv": ("30", "31", "IVA deducible: bienes de inversion"),
    "ded_aib": ("36", "37", "IVA deducible: adquisiciones intracomunitarias corrientes"),
}


def config_fiscal(datos_dir=None):
    cfg = {"nif": {}, "nif_propio": "", "razon_social": "", "periodicidad": "mensual"}
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), CONFIG_FILE)
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh) or {}
        if isinstance(d.get("nif"), dict):
            cfg["nif"] = {str(k).lower(): _txt(v) for k, v in d["nif"].items()}
        for k in ("nif_propio", "razon_social", "periodicidad"):
            if _txt(d.get(k)):
                cfg[k] = _txt(d[k])
    except Exception:
        pass
    return cfg


def _nif(nombre, cfg):
    n = _txt(nombre).lower()
    for k, v in (cfg.get("nif") or {}).items():
        if k and k in n:
            return v
    return ""


def _pct_key(pct):
    p = round(float(pct or 0))
    return {4: "dev_4", 10: "dev_10", 21: "dev_21"}.get(p, "dev_21")


def _desglose_ap(r):
    """(base, cuota, pct) de una factura AP con los mismos criterios que cierre_mes."""
    total = _num(r.get("total_factura"))
    base = _num(r.get("base_imponible")); iva = _num(r.get("cuota_iva"))
    pct = _num(r.get("porcentaje_iva"))
    if not base and not iva:
        pct = pct or IVA_GENERAL
        base = _r(total / (1 + pct / 100)); iva = _r(total - base)
    elif not iva:
        iva = _r(total - base)
    elif not base:
        base = _r(total - iva)
    if not pct and base:
        pct = round(iva / base * 100) if base else IVA_GENERAL
    return base, iva, pct or IVA_GENERAL


def calcular(mes, fuentes, cfg=None, cfg_fiscal=None):
    """Devuelve {mes, m303, m349, sii, avisos, estado, cifra, detalle}. Puro."""
    ini, fin, mes = _mes_a_rango(mes)
    cfg = cfg or config_cierre()
    cf = cfg_fiscal or {"nif": {}, "nif_propio": "", "razon_social": "", "periodicidad": "mensual"}
    acc = {k: [0.0, 0.0] for k in CASILLAS}   # base, cuota
    avisos = []
    exp, rec = [], []          # libros SII
    m349 = {}                  # ota -> {nif, base, clave}
    nif_pend = set()

    def suma(k, base, cuota):
        acc[k][0] = _r(acc[k][0] + base); acc[k][1] = _r(acc[k][1] + cuota)

    # ── ventas F&B por TPV (factura simplificada resumen diario, F4) ───────
    vf = fuentes.get("ventas_fb")
    pct_f = float(cfg.get("iva_fb", IVA_REDUCIDO)); pct_h = float(cfg.get("iva_alojamiento", IVA_REDUCIDO))
    if vf is not None and not vf.empty and "total_venta" in vf.columns:
        por_dia = {}
        for _, r in vf.iterrows():
            f = _fecha(r.get("fecha"))
            if f and ini <= f <= fin:
                por_dia[f] = _r(por_dia.get(f, 0) + _num(r.get("total_venta")))
        for f in sorted(por_dia):
            tot = por_dia[f]
            if tot <= 0:
                continue
            base = _r(tot / (1 + pct_f / 100)); cuota = _r(tot - base)
            suma(_pct_key(pct_f), base, cuota)
            exp.append({"tipo_factura": "F4", "fecha": f.isoformat(), "numero": f"TPV-{f.isoformat()}",
                        "nif": "", "nombre": "Ventas TPV (resumen diario)", "clave_regimen": "01",
                        "base": base, "tipo": pct_f, "cuota": cuota, "total": tot, "origen": "FB"})

    # ── facturas emitidas direct bill ─────────────────────────────────────
    rv = fuentes.get("reservas")
    if rv is not None and not rv.empty:
        for _, r in rv.iterrows():
            estado = _txt(r.get("estado")).upper()
            if estado in ("PENDIENTE_FACTURA", ""):
                continue
            total = _num(r.get("total")) or _num(r.get("importe"))
            f_em = r.get("fecha_emision") if _txt(r.get("fecha_emision")) else r.get("fecha_entrada")
            f = _fecha(f_em)
            if total <= 0 or f is None or not (ini <= f <= fin):
                continue
            hab = _num(r.get("importe_habitaciones")); fb = _num(r.get("importe_fb")); ext = _num(r.get("importe_extras"))
            if not (hab or fb or ext):
                hab = total
            s = _r(hab + fb + ext)
            if s and abs(s - total) > 0.011:
                k = total / s; hab, fb, ext = _r(hab * k), _r(fb * k), _r(ext * k)
            b_h = _r((hab + ext) / (1 + pct_h / 100)); i_h = _r(hab + ext - b_h)
            b_f = _r(fb / (1 + pct_f / 100)); i_f = _r(fb - b_f)
            dif = _r(total - (b_h + b_f + i_h + i_f)); i_h = _r(i_h + dif)
            num = _txt(r.get("numero_reserva")) or _txt(r.get("numero")) or "s/n"
            cli = _txt(r.get("cliente")) or "cliente"
            nif = _txt(r.get("nif")) or _txt(r.get("cif")) or _nif(cli, cf)
            if not nif:
                nif_pend.add(cli)
            if b_h:
                suma(_pct_key(pct_h), b_h, i_h)
            if b_f:
                suma(_pct_key(pct_f), b_f, i_f)
            exp.append({"tipo_factura": "F1", "fecha": f.isoformat(), "numero": num, "nif": nif, "nombre": cli,
                        "clave_regimen": "01", "base": _r(b_h + b_f), "tipo": pct_h if not b_f else f"{pct_h}/{pct_f}",
                        "cuota": _r(i_h + i_f), "total": total, "origen": "AR"})

    # ── facturas recibidas AP ─────────────────────────────────────────────
    ap = fuentes.get("ap")
    if ap is not None and not ap.empty:
        for _, r in ap.iterrows():
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            f = _fecha(fecha)
            if f is None or not (ini <= f <= fin):
                continue
            total = _num(r.get("total_factura"))
            if not total:
                continue
            base, cuota, pct = _desglose_ap(r)
            suma("ded_int", base, cuota)
            prov = _txt(r.get("nombre_proveedor")) or "proveedor"
            nif = _txt(r.get("nif_proveedor")) or _txt(r.get("cif")) or _nif(prov, cf)
            if not nif:
                nif_pend.add(prov)
            rec.append({"tipo_factura": "F1", "fecha": f.isoformat(),
                        "numero": _txt(r.get("numero_factura")) or _txt(r.get("archivo")),
                        "nif": nif, "nombre": prov, "clave_regimen": "01", "base": base, "tipo": pct,
                        "cuota": cuota, "total": total, "inversion_sujeto_pasivo": "N", "origen": "AP"})

    # ── comisiones OTA ────────────────────────────────────────────────────
    ar = fuentes.get("ar_ota")
    if ar is not None and not ar.empty:
        for _, r in ar.iterrows():
            fecha = r.get("fecha") if _txt(r.get("fecha")) else r.get("periodo_fin")
            f = _fecha(fecha)
            if f is None or not (ini <= f <= fin):
                continue
            imp = _num(r.get("importe_comision")) or _num(r.get("importe_comision_factura"))
            if imp <= 0:
                continue
            ota = _txt(r.get("nombre_ota")) or "OTA"
            reg = regimen_ota(ota, cfg)
            nif = _nif(ota, cf)
            if not nif:
                nif_pend.add(ota)
            fila = {"tipo_factura": "F1", "fecha": f.isoformat(), "numero": _txt(r.get("numero_factura")) or "s/n",
                    "nif": nif, "nombre": ota, "origen": "OTA"}
            if reg == "es":
                base = _r(imp / (1 + IVA_GENERAL / 100)); cuota = _r(imp - base)
                suma("ded_int", base, cuota)
                fila.update({"clave_regimen": "01", "base": base, "tipo": IVA_GENERAL, "cuota": cuota,
                             "total": imp, "inversion_sujeto_pasivo": "N"})
            else:
                cuota = _r(imp * IVA_GENERAL / 100)
                if reg == "ue":
                    suma("dev_aib", imp, cuota); suma("ded_aib", imp, cuota)
                    m = m349.setdefault(ota.lower(), {"operador": ota, "nif": nif, "clave": "S", "base": 0.0})
                    m["base"] = _r(m["base"] + imp)
                    clave = "09"       # adquisiciones intracomunitarias de servicios
                else:
                    suma("dev_isp", imp, cuota); suma("ded_int", imp, cuota)
                    clave = "01"
                fila.update({"clave_regimen": clave, "base": imp, "tipo": IVA_GENERAL, "cuota": cuota,
                             "total": imp, "inversion_sujeto_pasivo": "S"})
            rec.append(fila)

    # ── modelo 303 ───────────────────────────────────────────────────────
    casillas = []
    for k, (cb, cc, nombre) in CASILLAS.items():
        casillas.append({"clave": k, "casilla_base": cb, "casilla_cuota": cc, "concepto": nombre,
                         "base": acc[k][0], "cuota": acc[k][1]})
    devengado = _r(sum(acc[k][1] for k in acc if k.startswith("dev_")))
    deducible = _r(sum(acc[k][1] for k in acc if k.startswith("ded_")))
    resultado = _r(devengado - deducible)
    m303 = {"casillas": casillas, "c27_devengado": devengado, "c45_deducible": deducible,
            "c46_resultado": resultado, "c67_compensar_anterior": None,
            "signo": "A INGRESAR" if resultado > 0 else ("A COMPENSAR" if resultado < 0 else "SIN ACTIVIDAD"),
            "periodicidad": cf.get("periodicidad", "mensual")}
    avisos.append("Casilla 67 (compensacion de periodos anteriores) no se conoce: el resultado es antes de compensar.")
    if acc["ded_inv"][0] == 0:
        avisos.append("Bienes de inversion (30-31), importaciones (32-35) y prorrata (44) no se distinguen: "
                      "las compras de inmovilizado van en 28-29 como corriente.")
    if cf.get("periodicidad") != "mensual":
        avisos.append("Periodicidad trimestral: sumar los 3 meses antes de presentar.")

    # ── modelo 349 ───────────────────────────────────────────────────────
    filas349 = sorted(m349.values(), key=lambda x: x["operador"])
    if filas349 and any(not x["nif"] for x in filas349):
        avisos.append("349: falta el NIF-IVA de " + ", ".join(x["operador"] for x in filas349 if not x["nif"])
                      + " (config_fiscal.json → nif).")

    # ── SII ──────────────────────────────────────────────────────────────
    sii = {"expedidas": exp, "recibidas": rec, "n_expedidas": len(exp), "n_recibidas": len(rec),
           "total_expedidas": _r(sum(x["total"] for x in exp)), "total_recibidas": _r(sum(x["total"] for x in rec)),
           "estado": "PREPARADO", "nota": "Libros listos; el envio a la AEAT exige certificado digital y no lo hace Yve."}
    if nif_pend:
        avisos.append(f"SII: {len(nif_pend)} contrapartes sin NIF (no se puede enviar sin el).")

    hay_datos = bool(exp or rec)
    estado = "PREPARADO" if hay_datos and not nif_pend else ("PENDIENTE" if hay_datos else "SIN_DATO")
    return {
        "mes": mes, "desde": ini.isoformat(), "hasta": fin.isoformat(),
        "m303": m303, "m349": {"filas": filas349, "total_base": _r(sum(x["base"] for x in filas349)),
                               "n": len(filas349), "nif_propio": cf.get("nif_propio", "")},
        "sii": sii, "avisos": avisos, "nif_pendientes": sorted(nif_pend),
        "estado": estado,
        "cifra": f"303 {m303['signo'].lower()}: {_es(resultado)} € · devengado {_es(devengado)} € · deducible {_es(deducible)} €",
        "detalle": (f"349: {len(filas349)} operadores UE, base {_es(_r(sum(x['base'] for x in filas349)))} € · "
                    f"SII: {len(exp)} expedidas / {len(rec)} recibidas"
                    + (f" · {len(nif_pend)} NIF pendientes" if nif_pend else "")),
    }


def hojas(res):
    """Hojas planas para el Excel / paquete: {nombre: [filas]}."""
    return {
        "Modelo 303": res["m303"]["casillas"] + [
            {"clave": "c27", "casilla_base": "", "casilla_cuota": "27", "concepto": "Total cuota devengada",
             "base": "", "cuota": res["m303"]["c27_devengado"]},
            {"clave": "c45", "casilla_base": "", "casilla_cuota": "45", "concepto": "Total a deducir",
             "base": "", "cuota": res["m303"]["c45_deducible"]},
            {"clave": "c46", "casilla_base": "", "casilla_cuota": "46", "concepto": f"Resultado ({res['m303']['signo']})",
             "base": "", "cuota": res["m303"]["c46_resultado"]}],
        "Modelo 349": res["m349"]["filas"],
        "SII expedidas": res["sii"]["expedidas"],
        "SII recibidas": res["sii"]["recibidas"],
        "Avisos": [{"aviso": a} for a in res["avisos"]],
    }


def fiscal_completo(mes, hotel=None, procesadas_dir=None, reportes_dir=None, datos_dir=None):
    import cierre_mes as CM
    fu = CM.recoger_fuentes(mes, hotel, procesadas_dir, reportes_dir, datos_dir)
    return calcular(mes, fu, CM.config_cierre(datos_dir), config_fiscal(datos_dir))


def resumen_para_paquete(mes, asientos=None, reconciliacion=None, hotel=None, **dirs):
    """Lo que consume paquete_cierre: {estado, cifra, detalle, hojas}. Ademas
    comprueba que la cuota del 303 coincide con el libro (477 − 472) si hay asientos."""
    res = fiscal_completo(mes, hotel, **dirs)
    out = {"estado": res["estado"], "cifra": res["cifra"], "detalle": res["detalle"], "hojas": hojas(res),
           "avisos": res["avisos"], "m303": res["m303"]}
    if asientos and asientos.get("asientos"):
        rep = _r(sum(a["haber"] - a["debe"] for a in asientos["asientos"] if a["cuenta"] == "477"))
        sop = _r(sum(a["debe"] - a["haber"] for a in asientos["asientos"] if a["cuenta"] == "472"))
        out["libro_477"] = rep; out["libro_472"] = sop
        out["cuadra_con_libro"] = abs(rep - res["m303"]["c27_devengado"]) <= 0.011 and \
            abs(sop - res["m303"]["c45_deducible"]) <= 0.011
        if not out["cuadra_con_libro"]:
            out["estado"] = "PENDIENTE"
            out["detalle"] += f" · NO cuadra con el libro (477 {_es(rep)} € / 472 {_es(sop)} €)"
    return out


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, filas in hojas(res).items():
            pd.DataFrame(filas or [{}]).to_excel(w, index=False, sheet_name=nombre[:31])
    buf.seek(0)
    return buf, f"fiscal_{res['mes']}.xlsx"
