# -*- coding: utf-8 -*-
"""tab_ficha_ap.py — Yve.01 · la ficha de una factura AP y lo que una persona decide sobre ella (b84).

Lo que pidio Jordi el 23 sep 2026 (con las notas de su contacto):
  - ver de que trata la factura (concepto, periodo facturado, lineas) y su asiento
    (cuenta de gasto, IVA, proveedor) SIN ir al Cierre;
  - poder CAMBIAR la cuenta (y que el asignador aprenda para ese proveedor);
  - vencimiento / dias de pago por factura, visible en el aging;
  - fecha contable = el mes en que se registra (ajustable), no la fecha impresa;
  - marcar PAGADA a mano (fecha, quien, desde que cuenta bancaria), y que se pueda
    extraer (columna, Excel, filtro);
  - descargar la ficha (PDF) y el asiento de una o varias facturas (Excel).

Los ajustes viven en datos-referencia/ajustes_ap.json (almacen_datos.guardar_ajuste_ap)
y se aplican al leer (almacen_datos.facturas_ap): el Excel del clasificador no se toca.

Rutas (todas con sesion; los POST pasan por el CSRF de /api/*):
  GET  /api/ap/ficha?clave=
  POST /api/ap/ajustar          {clave, cuenta_contable?, aplicar_proveedor?, vencimiento?, fecha_contable?, dias_pago?}
  POST /api/ap/pagar            {clave, fecha, cuenta, nota}
  POST /api/ap/despagar         {clave}
  GET  /api/ap/ficha.pdf?clave=
  GET  /api/exportar/ap_asientos?claves=a,b,c   (vacio = todas las del hotel)
  GET  /api/cuentas_bancarias · POST /api/cuentas_bancarias {nombre, iban}
"""
import io
import json
import os
from datetime import date, datetime

import pandas as pd
from flask import Blueprint, jsonify, request, send_file
from flask_login import login_required, current_user

ficha_ap_bp = Blueprint("ficha_ap", __name__)
CUENTAS_BANCARIAS_FILE = "cuentas_bancarias.json"


def _usuario():
    try:
        return getattr(current_user, "username", None) or getattr(current_user, "id", "") or ""
    except Exception:
        return ""


def _ddir():
    from tenant_dirs import datos_dir
    return str(datos_dir())


def _s(v):
    """Texto limpio: '' para NaN/None/NO_ENCONTRADO; sin '.0' en los codigos."""
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat", "no_encontrado", "null"):
        return ""
    return s


def _cta(v):
    s = _s(v)
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s


def _num(v):
    try:
        f = float(v)
        return None if f != f else round(f, 2)
    except (TypeError, ValueError):
        return None


def _iso(v):
    from almacen_datos import _iso as iso
    return iso(v)


def _facturas():
    from dashboard import cargar_datos_ap
    return cargar_datos_ap()


def _fila(clave):
    """La factura con esa clave (numero, o fichero si no hay numero) del hotel activo."""
    from almacen_datos import clave_ap
    df = _facturas()
    if df is None or df.empty:
        return None, df
    for fila in df.to_dict("records"):
        if clave_ap(fila) == clave:
            return fila, df
    return None, df


# ── asiento de una factura (mismo criterio que cierre_mes) ─────────────────
def asiento_de(fila, plan=None):
    from cierre_mes import plan_cuentas, IVA_GENERAL
    plan = plan or plan_cuentas(_ddir())
    total = _num(fila.get("total_factura")) or 0.0
    base = _num(fila.get("base_imponible"))
    iva = _num(fila.get("cuota_iva"))
    pct = _num(fila.get("porcentaje_iva"))
    if total and base is None and iva is None:
        p = pct or IVA_GENERAL
        base = round(total / (1 + p / 100), 2); iva = round(total - base, 2)
    elif total and iva is None:
        iva = round(total - (base or 0), 2)
    elif total and base is None:
        base = round(total - (iva or 0), 2)
    base = base or 0.0; iva = iva or 0.0
    from cierre_mes import cuentas_ap
    cta, acreedor = cuentas_ap(fila)        # b88: comision de agencia → 628 / 410
    lineas = [{"cuenta": cta, "nombre": plan.get(cta, ""), "debe": base, "haber": 0.0}]
    if iva:
        lineas.append({"cuenta": "472", "nombre": plan.get("472", "H.P. IVA soportado"), "debe": iva, "haber": 0.0})
    lineas.append({"cuenta": acreedor, "nombre": plan.get(acreedor, "Proveedores" if acreedor == "400" else "Acreedores por prestaciones de servicios"), "debe": 0.0, "haber": total})
    return {"fecha": _iso(fila.get("fecha_contable")) or _iso(fila.get("fecha_factura") or fila.get("fecha")),
            "concepto": f"Fra. {_s(fila.get('numero_factura')) or _s(fila.get('archivo'))} — {_s(fila.get('nombre_proveedor'))}",
            "lineas": lineas, "cuadra": round(sum(l["debe"] for l in lineas) - sum(l["haber"] for l in lineas), 2) == 0.0,
            "base": base, "iva": iva, "total": total, "cuenta_gasto": cta}


def _lineas_de(fila):
    from almacen_datos import lineas_factura, _txt
    try:
        df = lineas_factura()
    except Exception:
        return []
    if df is None or df.empty:
        return []
    arch = _txt(fila.get("archivo")); num = _txt(fila.get("numero_factura"))
    out = []
    for r in df.to_dict("records"):
        if (arch and _txt(r.get("archivo")) == arch) or (num and _txt(r.get("numero_factura")) == num and _txt(r.get("nombre_proveedor")) == _txt(fila.get("nombre_proveedor"))):
            out.append({"n": int(_num(r.get("n_linea")) or 0), "descripcion": _s(r.get("descripcion")), "cantidad": _num(r.get("cantidad")),
                        "unidad": _s(r.get("unidad")), "precio_unitario": _num(r.get("precio_unitario")), "importe": _num(r.get("importe"))})
    return sorted(out, key=lambda x: x["n"])


def _es_com(fila):
    from cierre_mes import es_comision_agencia
    return es_comision_agencia(fila)


def _comision_de(fila):
    """b88: si es la factura de comision de un contrato de grupo, lo que hay que ver."""
    if not _es_com(fila):
        return None
    tot = _num(fila.get("total_factura")) or 0.0
    comp = _num(fila.get("compensado")) or 0.0
    try:
        from compensaciones import puede_pagar_comision
        se_puede, motivo = puede_pagar_comision(fila)
    except Exception:
        se_puede, motivo = True, ""
    return {"contrato_id": _s(fila.get("comision_contrato")), "evento": _s(fila.get("comision_evento")),
            "esperada": _num(fila.get("comision_esperada")), "facturada": _num(fila.get("comision_facturada")),
            "diferencia": _num(fila.get("comision_diferencia")), "estado": _s(fila.get("comision_estado")),
            "pagador": _s(fila.get("comision_pagador")), "factura_grupo": _s(fila.get("grupo_numero")) or _s(fila.get("factura_grupo")),
            "vinculo": _s(fila.get("comision_vinculo")),
            # b89: compensada contra la factura del grupo (paga la agencia) o se paga despues de cobrarla (paga el cliente)
            "grupo_estado": _s(fila.get("grupo_estado")).upper(), "compensado": round(comp, 2),
            "saldo": round(max(0.0, tot - comp), 2), "se_puede_pagar": bool(se_puede), "motivo_no_pagar": motivo}


def ficha(clave):
    from almacen_datos import ajustes_ap, clave_ap
    from cierre_mes import plan_cuentas, criterio_fecha_ap, config_cierre
    fila, _ = _fila(clave)
    if fila is None:
        return None
    plan = plan_cuentas(_ddir())
    aj = ajustes_ap().get(clave_ap(fila)) or {}
    pag = aj.get("pagada") or {}
    return {
        "clave": clave_ap(fila),
        "numero_factura": _s(fila.get("numero_factura")), "archivo": _s(fila.get("archivo")),
        "proveedor": _s(fila.get("nombre_proveedor")), "nif": _s(fila.get("NIF_proveedor")),
        "concepto": _s(fila.get("descripcion_concepto")),
        "periodo_inicio": _iso(fila.get("periodo_inicio")), "periodo_fin": _iso(fila.get("periodo_fin")),
        "fecha_factura": _iso(fila.get("fecha_factura") or fila.get("fecha")),
        "fecha_registro": _iso(fila.get("fecha_registro")),
        "fecha_contable": _iso(fila.get("fecha_contable")),
        "criterio_fecha": "comision" if _es_com(fila) else criterio_fecha_ap(config_cierre(_ddir())),
        "comision": _comision_de(fila),
        "dias_pago": int(_num(fila.get("dias_pago")) or 30), "vencimiento": _iso(fila.get("vencimiento")),
        "base": _num(fila.get("base_imponible")), "porcentaje_iva": _num(fila.get("porcentaje_iva")),
        "cuota_iva": _num(fila.get("cuota_iva")), "total": _num(fila.get("total_factura")),
        "tipo": _s(fila.get("tipo_proveedor")).upper() or "OTRAS",
        "cuenta_contable": _cta(fila.get("cuenta_contable")), "cuenta_nombre": plan.get(_cta(fila.get("cuenta_contable")), ""),
        "cuenta_ajustada": bool(fila.get("cuenta_ajustada")),
        "estado": (_s(fila.get("estado_matching")) or _s(fila.get("estado"))).upper() or "PENDIENTE",
        "accion": _s(fila.get("accion")).upper(), "importes_cuadran": _s(fila.get("importes_cuadran")),
        "aviso_importes": _s(fila.get("aviso_importes")), "hotel_id": _s(fila.get("hotel_id")),
        "pagada": bool(pag), "pagada_fecha": pag.get("fecha", ""), "pagada_cuenta": pag.get("cuenta", ""),
        "pagada_por": pag.get("por", ""), "pagada_nota": pag.get("nota", ""),
        "compensado": _num(aj.get("compensado")) or 0.0, "compensada": bool(fila.get("compensada")) and str(fila.get("compensada")) != "nan",
        "asiento": asiento_de(fila, plan), "lineas": _lineas_de(fila),
        "historial": (aj.get("historial") or [])[-20:],
        "plan_cuentas": [{"codigo": k, "nombre": v} for k, v in sorted(plan.items()) if str(k)[:1] in ("2", "6")],
        "cuentas_bancarias": cuentas_bancarias(),
    }


# ── cuentas bancarias (para "marcar pagada") ───────────────────────────────
def cuentas_bancarias():
    try:
        with open(os.path.join(_ddir(), CUENTAS_BANCARIAS_FILE), encoding="utf-8") as fh:
            d = json.load(fh)
        return [c for c in d if isinstance(c, dict) and c.get("nombre")] if isinstance(d, list) else []
    except Exception:
        return []


def guardar_cuenta_bancaria(nombre, iban=""):
    nombre = str(nombre or "").strip()[:80]; iban = str(iban or "").replace(" ", "").upper()[:40]
    if not nombre:
        raise ValueError("falta el nombre de la cuenta")
    cs = cuentas_bancarias()
    for c in cs:
        if c["nombre"].lower() == nombre.lower():
            if iban:
                c["iban"] = iban
            break
    else:
        cs.append({"nombre": nombre, "iban": iban})
    ruta = os.path.join(_ddir(), CUENTAS_BANCARIAS_FILE)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cs, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)
    return cs


# ── que el asignador aprenda la cuenta corregida ───────────────────────────
def aprender_cuenta_proveedor(nombre_proveedor, cuenta, usuario=""):
    """La proxima factura de este proveedor sale con esta cuenta. Si esta en el maestro
    proveedores.xlsx se corrige ahi (el maestro manda); si no, en proveedores_aprendidos.json."""
    import cuentas_proveedor as CP
    tipo = "FB" if str(cuenta).startswith("60") else "OTRAS"
    ruta_m = os.path.join(_ddir(), "proveedores.xlsx")
    donde = "aprendidos"
    try:
        dfm = pd.read_excel(ruta_m)
        if "nombre_proveedor" in dfm.columns:
            k = CP.clave_proveedor(nombre_proveedor)
            mask = dfm["nombre_proveedor"].map(lambda v: CP.clave_proveedor(v) == k)
            if mask.any():
                dfm = dfm.copy()
                dfm["cuenta_contable"] = dfm["cuenta_contable"].astype(object)
                dfm.loc[mask, "cuenta_contable"] = str(cuenta)
                if "tipo" in dfm.columns:
                    dfm["tipo"] = dfm["tipo"].astype(object); dfm.loc[mask, "tipo"] = tipo
                tmp = ruta_m + ".tmp"; dfm.to_excel(tmp, index=False); os.replace(tmp, ruta_m)
                donde = "proveedores.xlsx"
    except Exception:
        pass
    if donde == "aprendidos":
        apr = CP.aprendidos(_ddir())
        k = CP.clave_proveedor(nombre_proveedor)
        apr[k] = {"proveedor": nombre_proveedor, "tipo": tipo, "cuenta": str(cuenta),
                  "origen": f"corregida a mano por {usuario or 'usuario'} el {date.today().isoformat()}",
                  "primera_factura": (apr.get(k) or {}).get("primera_factura", "")}
        CP.guardar_aprendidos(apr, _ddir())
    return donde


# ── exportaciones ──────────────────────────────────────────────────────────
def excel_asientos(claves=None):
    """Un Excel con el asiento de cada factura (una fila por linea) + hoja Facturas."""
    from almacen_datos import clave_ap
    from cierre_mes import plan_cuentas
    df = _facturas()
    plan = plan_cuentas(_ddir())
    filas, cab = [], []
    if df is not None and not df.empty:
        quiero = set(claves or [])
        for fila in df.to_dict("records"):
            k = clave_ap(fila)
            if quiero and k not in quiero:
                continue
            a = asiento_de(fila, plan)
            for n, l in enumerate(a["lineas"], 1):
                filas.append({"fecha": a["fecha"], "asiento": k, "linea": n, "cuenta": l["cuenta"], "nombre_cuenta": l["nombre"],
                              "debe": l["debe"], "haber": l["haber"], "concepto": a["concepto"], "documento": _s(fila.get("numero_factura")),
                              "proveedor": _s(fila.get("nombre_proveedor")), "nif": _s(fila.get("NIF_proveedor"))})
            cab.append(_fila_excel(fila))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(filas or [{"fecha": "", "asiento": "sin facturas", "linea": "", "cuenta": "", "nombre_cuenta": "", "debe": "", "haber": "", "concepto": "", "documento": "", "proveedor": "", "nif": ""}]).to_excel(w, index=False, sheet_name="Asientos")
        pd.DataFrame(cab or [{"numero_factura": "sin facturas"}]).to_excel(w, index=False, sheet_name="Facturas")
    buf.seek(0)
    return buf, f"asientos_ap_{date.today().strftime('%Y%m%d')}.xlsx"


def _fila_excel(fila):
    return {
        "numero_factura": _s(fila.get("numero_factura")), "proveedor": _s(fila.get("nombre_proveedor")), "nif": _s(fila.get("NIF_proveedor")),
        "fecha_factura": _iso(fila.get("fecha_factura") or fila.get("fecha")), "fecha_contable": _iso(fila.get("fecha_contable")),
        "periodo_inicio": _iso(fila.get("periodo_inicio")), "periodo_fin": _iso(fila.get("periodo_fin")),
        "concepto": _s(fila.get("descripcion_concepto")), "base": _num(fila.get("base_imponible")), "iva_pct": _num(fila.get("porcentaje_iva")),
        "cuota_iva": _num(fila.get("cuota_iva")), "total": _num(fila.get("total_factura")), "cuenta": _cta(fila.get("cuenta_contable")),
        "tipo": _s(fila.get("tipo_proveedor")).upper(), "estado": (_s(fila.get("estado_matching")) or _s(fila.get("estado"))).upper(),
        "comision_de_grupo": _s(fila.get("comision_evento")) if _es_com(fila) else "", "comision_estado": _s(fila.get("comision_estado")),
        "aprobacion": _s(fila.get("accion")).upper(), "vencimiento": _iso(fila.get("vencimiento")), "dias_pago": _num(fila.get("dias_pago")),
        "compensado": _num(fila.get("compensado")) or 0.0,
        "pagada": "SI" if fila.get("pagada") else "NO", "fecha_pago": _s(fila.get("pagada_fecha")), "cuenta_bancaria": _s(fila.get("pagada_cuenta")),
        "pagada_por": _s(fila.get("pagada_por")), "hotel_id": _s(fila.get("hotel_id")), "archivo": _s(fila.get("archivo")),
    }


def excel_facturas_ap():
    """El Excel AP de verdad (antes /api/exportar/ap devolvia cuatro filas inventadas)."""
    df = _facturas()
    filas = [_fila_excel(f) for f in df.to_dict("records")] if df is not None and not df.empty else []
    buf = io.BytesIO()
    pd.DataFrame(filas or [{"numero_factura": "Sin facturas de proveedor: sube facturas con Procesar archivos"}]).to_excel(buf, index=False, sheet_name="AP - Proveedores")
    buf.seek(0)
    return buf, f"AP_facturas_{date.today().strftime('%Y%m%d')}.xlsx"


def pdf_ficha(f):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4); W, H = A4
    y = H - 20 * mm

    def eur(x):
        return "—" if x is None else f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

    def linea(txt, dy=6, bold=False, size=10):
        nonlocal y
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size); c.drawString(20 * mm, y, txt[:110]); y -= dy * mm

    linea(f"Yve.01 · Ficha de factura {f['numero_factura'] or f['archivo']}", 9, True, 14)
    linea(f"{f['proveedor']}  ·  NIF {f['nif'] or '—'}", 7, False, 11)
    linea(f"Fecha factura {f['fecha_factura'] or '—'} · Fecha contable {f['fecha_contable'] or '—'} · Vence {f['vencimiento'] or '—'} ({f['dias_pago']} días)")
    if f["periodo_inicio"] or f["periodo_fin"]:
        linea(f"Periodo facturado: {f['periodo_inicio'] or '?'} → {f['periodo_fin'] or '?'}")
    linea(f"Concepto: {f['concepto'] or '—'}")
    linea(f"Base {eur(f['base'])} · IVA {f['porcentaje_iva'] or 0:g} % {eur(f['cuota_iva'])} · Total {eur(f['total'])}", 8, True)
    linea(f"Cuenta {f['cuenta_contable']} {f['cuenta_nombre']} · Tipo {f['tipo']} · Estado {f['estado']} · Aprobación {f['accion'] or 'sin decisión'}")
    linea(("PAGADA el " + f['pagada_fecha'] + (" desde " + f['pagada_cuenta'] if f['pagada_cuenta'] else "") + (" (" + f['pagada_por'] + ")" if f['pagada_por'] else "")) if f["pagada"] else "Pendiente de pago", 9, True)
    linea("Asiento", 6, True, 11)
    for l in f["asiento"]["lineas"]:
        linea(f"   {l['cuenta']:<6} {l['nombre'][:38]:<38}  debe {eur(l['debe']) if l['debe'] else '':>14}  haber {eur(l['haber']) if l['haber'] else '':>14}", 5)
    if f["lineas"]:
        y -= 3 * mm; linea("Líneas de la factura", 6, True, 11)
        for l in f["lineas"][:40]:
            linea(f"   {l['n']:>2}. {l['descripcion'][:45]:<45} {l['cantidad'] if l['cantidad'] is not None else '':>8} {l['unidad']:<5} {eur(l['precio_unitario']) if l['precio_unitario'] is not None else '':>12} {eur(l['importe']) if l['importe'] is not None else '':>12}", 5, False, 9)
            if y < 25 * mm:
                c.showPage(); y = H - 20 * mm
    c.setFont("Helvetica", 8); c.drawString(20 * mm, 12 * mm, f"Generado por Yve.01 el {datetime.now().strftime('%d/%m/%Y %H:%M')} · documento {f['archivo']}")
    c.save(); buf.seek(0)
    return buf, f"factura_{(f['numero_factura'] or 'sin_numero').replace('/', '-')}.pdf"


# ── rutas ──────────────────────────────────────────────────────────────────
@ficha_ap_bp.route("/api/ap/ficha")
@login_required
def api_ficha():
    clave = str(request.args.get("clave") or "").strip()
    f = ficha(clave) if clave else None
    if not f:
        return jsonify({"ok": False, "error": "factura no encontrada"}), 404
    return jsonify({"ok": True, **f})


def _audit(accion, detalle):
    try:
        from dashboard import _audit as _a
        _a(accion, detalle)
    except Exception:
        pass


@ficha_ap_bp.route("/api/ap/ajustar", methods=["POST"])
@login_required
def api_ajustar():
    from almacen_datos import guardar_ajuste_ap
    d = request.get_json(silent=True) or {}
    clave = str(d.get("clave") or "").strip()
    fila, _ = _fila(clave)
    if fila is None:
        return jsonify({"ok": False, "error": "factura no encontrada"}), 404
    cambios = {}
    for k in ("cuenta_contable", "vencimiento", "fecha_contable", "dias_pago"):
        if k in d:
            cambios[k] = d.get(k)
    if cambios.get("fecha_contable") and _es_com(fila):
        return jsonify({"ok": False, "error": "la factura de comisión de una agencia se imputa por su fecha de factura (regla de finanzas)"}), 400
    if "cuenta_contable" in cambios:
        cta = _cta(cambios["cuenta_contable"])
        if cta and not cta.isdigit():
            return jsonify({"ok": False, "error": "la cuenta tiene que ser numerica (p. ej. 629)"}), 400
        cambios["cuenta_contable"] = cta
    for k in ("vencimiento", "fecha_contable"):
        if cambios.get(k) and not _iso(cambios[k]):
            return jsonify({"ok": False, "error": f"{k}: fecha no valida"}), 400
        if cambios.get(k):
            cambios[k] = _iso(cambios[k])
    if not cambios:
        return jsonify({"ok": False, "error": "nada que cambiar"}), 400
    try:
        guardar_ajuste_ap(clave, cambios, _usuario())
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    donde = ""
    if cambios.get("cuenta_contable") and d.get("aplicar_proveedor"):
        donde = aprender_cuenta_proveedor(_s(fila.get("nombre_proveedor")), cambios["cuenta_contable"], _usuario())
    _audit("AJUSTE_AP", f"{clave}: {cambios}" + (f" · proveedor→{donde}" if donde else ""))
    return jsonify({"ok": True, "aprendido_en": donde, **(ficha(clave) or {})})


@ficha_ap_bp.route("/api/ap/pagar", methods=["POST"])
@login_required
def api_pagar():
    from almacen_datos import guardar_ajuste_ap
    d = request.get_json(silent=True) or {}
    clave = str(d.get("clave") or "").strip()
    fila, _ = _fila(clave)
    if fila is None:
        return jsonify({"ok": False, "error": "factura no encontrada"}), 404
    # b89, regla de finanzas: si la factura del grupo la paga el cliente final, primero se
    # cobra la factura del grupo y despues se paga la de comision de la agencia
    if _es_com(fila):
        from compensaciones import puede_pagar_comision
        se_puede, motivo = puede_pagar_comision(fila)
        if not se_puede:
            return jsonify({"ok": False, "error": motivo, "regla": True}), 409
    fecha = _iso(d.get("fecha")) or date.today().isoformat()
    cuenta = str(d.get("cuenta") or "").strip()
    if not cuenta:
        return jsonify({"ok": False, "error": "di desde que cuenta bancaria se ha pagado"}), 400
    if d.get("iban") or not any(c["nombre"].lower() == cuenta.lower() for c in cuentas_bancarias()):
        guardar_cuenta_bancaria(cuenta, d.get("iban", ""))
    guardar_ajuste_ap(clave, {"pagada": {"fecha": fecha, "cuenta": cuenta, "nota": d.get("nota", "")}}, _usuario())
    _audit("AP_PAGADA", f"{clave} el {fecha} desde {cuenta}")
    return jsonify({"ok": True, **(ficha(clave) or {})})


@ficha_ap_bp.route("/api/ap/despagar", methods=["POST"])
@login_required
def api_despagar():
    from almacen_datos import guardar_ajuste_ap
    d = request.get_json(silent=True) or {}
    clave = str(d.get("clave") or "").strip()
    fila, _ = _fila(clave)
    if fila is None:
        return jsonify({"ok": False, "error": "factura no encontrada"}), 404
    guardar_ajuste_ap(clave, {"pagada": None}, _usuario())
    _audit("AP_DESPAGADA", clave)
    return jsonify({"ok": True, **(ficha(clave) or {})})


@ficha_ap_bp.route("/api/ap/ficha.pdf")
@login_required
def api_ficha_pdf():
    clave = str(request.args.get("clave") or "").strip()
    f = ficha(clave) if clave else None
    if not f:
        return jsonify({"ok": False, "error": "factura no encontrada"}), 404
    buf, nombre = pdf_ficha(f)
    return send_file(buf, as_attachment=True, download_name=nombre, mimetype="application/pdf")


@ficha_ap_bp.route("/api/exportar/ap_asientos")
@login_required
def api_exportar_asientos():
    claves = [c for c in str(request.args.get("claves") or "").split(",") if c.strip()]
    buf, nombre = excel_asientos(claves or None)
    return send_file(buf, as_attachment=True, download_name=nombre, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@ficha_ap_bp.route("/api/cuentas_bancarias", methods=["GET", "POST"])
@login_required
def api_cuentas_bancarias():
    if request.method == "POST":
        d = request.get_json(silent=True) or {}
        try:
            return jsonify({"ok": True, "cuentas": guardar_cuenta_bancaria(d.get("nombre"), d.get("iban"))})
        except ValueError as e:
            return jsonify({"ok": False, "error": str(e)}), 400
    return jsonify({"ok": True, "cuentas": cuentas_bancarias()})
