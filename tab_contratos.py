# -*- coding: utf-8 -*-
"""tab_contratos.py — Yve.01 · AR > Contratos (b87).

Los contratos de grupo que ha leido Yve (fotos o PDF), con lo que finanzas dijo
que manda (24 sep 2026): la comision de la agencia segun el CAMPO DE COMISIONES
del contrato (tarifa neta = sin factura de comision; % = la agencia factura su
comision aparte) y QUIEN PAGA la factura del grupo. Lo que el contrato no dice
se pregunta aqui; no se supone.

Rutas (con sesion; el POST pasa por el CSRF de /api/*):
  GET  /api/ar/contratos
  POST /api/ar/contratos/decidir      {id, modo?: 'neta'|'porcentaje', pct?: {alojamiento, fb, salas}, pagador?: 'agencia'|'cliente'}
  POST /api/ar/contratos/vincular     {id, clave}   b88: una persona une la factura de comision de la agencia
  POST /api/ar/contratos/desvincular  {id, clave}   b88: "esta factura no es de este contrato"
"""
import os

import pandas as pd
from flask import Blueprint, jsonify, request
from flask_login import login_required, current_user

import contratos_grupo as CG

contratos_bp = Blueprint("contratos", __name__)


def _usuario():
    try:
        return getattr(current_user, "username", None) or ""
    except Exception:
        return ""


def _hotel():
    try:
        import censo_hoteles
        return censo_hoteles.activo()
    except Exception:
        return ""


def _audit(accion, detalle):
    try:
        from dashboard import _audit as _a
        _a(accion, detalle)
    except Exception:
        pass


def facturas_grupo(datos_dir=None):
    """{numero: {estado, total, fecha_emision, fecha_cobro, cliente, hotel_id}} de reservas_credito.xlsx."""
    ruta = os.path.join(CG._dd(datos_dir), "reservas_credito.xlsx")
    out = {}
    if not os.path.exists(ruta):
        return out
    try:
        df = pd.read_excel(ruta)
    except Exception:
        return out
    for f in df.to_dict("records"):
        num = CG._txt(f.get("numero_reserva")) or CG._txt(f.get("numero"))
        if not num:
            continue
        clave = num + "|" + CG._txt(f.get("hotel_id"))
        out[clave] = {"numero": num, "estado": CG._txt(f.get("estado")).upper(),
                      "total": CG._r(CG._f(f.get("total")) or CG._f(f.get("importe"))),
                      "fecha_emision": CG._txt(f.get("fecha_emision"))[:10], "fecha_cobro": CG._txt(f.get("fecha_cobro"))[:10],
                      "cliente": CG._txt(f.get("cliente")), "hotel_id": CG._txt(f.get("hotel_id"))}
    return out


def _facturas_ap(hotel):
    try:
        import almacen_datos
        df = almacen_datos.facturas_ap(hotel=hotel or None)
        return df.to_dict("records") if df is not None and not df.empty else []
    except Exception:
        return []


def _mini(f):
    return {"clave": CG._clave_ap(f), "numero": CG._txt(f.get("numero_factura")),
            "fecha": CG._iso(f.get("fecha_factura") if CG._txt(f.get("fecha_factura")) else f.get("fecha")),
            "total": CG._r(CG._f(f.get("total_factura"))), "base": CG.base_factura(f),
            "proveedor": CG._txt(f.get("nombre_proveedor"))}


def vista_contratos(hotel=None, datos_dir=None):
    fac = facturas_grupo(datos_dir)
    contratos = CG.del_hotel(CG.leer(datos_dir), hotel)
    filas_ap = _facturas_ap(hotel)
    enl = CG.enlazar(contratos, filas_ap) if contratos else {}
    por_clave = {CG._clave_ap(f): f for f in filas_ap}
    out = []
    for c in contratos:
        v = CG.vista(c)
        num = CG._txt(c.get("factura_grupo"))
        v["factura"] = fac.get(num + "|" + CG._txt(c.get("hotel_id"))) or fac.get(num + "|") or ({"numero": num} if num else {})
        e = enl.get(c["id"]) or {}
        if e.get("clave") and e["clave"] in por_clave:
            fr = por_clave[e["clave"]]
            v["factura_comision_vista"] = dict(_mini(fr), origen=e.get("origen"), **CG.estado_comision(c, fr))
        else:
            v["factura_comision_vista"] = {}
        v["candidatas"] = [_mini(por_clave[k]) for k in (e.get("candidatas") or []) if k in por_clave]
        out.append(v)
    # primero lo que falta decidir; dentro de cada grupo, lo mas reciente arriba
    pend = [c for c in out if c["pendientes"]]
    resto = [c for c in out if not c["pendientes"]]
    pend.sort(key=lambda c: c.get("fecha_entrada") or "", reverse=True)
    resto.sort(key=lambda c: c.get("fecha_entrada") or "", reverse=True)
    return pend + resto


@contratos_bp.route("/api/ar/contratos")
@login_required
def api_contratos():
    lista = vista_contratos(_hotel())
    return jsonify({"ok": True, "contratos": lista, "n_pendientes": sum(1 for c in lista if c["pendientes"])})


@contratos_bp.route("/api/ar/contratos/decidir", methods=["POST"])
@login_required
def api_decidir():
    d = request.get_json(silent=True) or {}
    cid = str(d.get("id") or "").strip()
    visibles = {c.get("id") for c in CG.del_hotel(CG.leer(), _hotel())}
    if not cid or cid not in visibles:
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    if d.get("modo") is None and d.get("pagador") is None:
        return jsonify({"ok": False, "error": "nada que decidir"}), 400
    try:
        c = CG.decidir(cid, modo=d.get("modo"), pct=d.get("pct"), pagador=d.get("pagador"), usuario=_usuario())
    except KeyError:
        return jsonify({"ok": False, "error": "contrato no encontrado"}), 404
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    CG.sincronizar_factura(c)
    _audit("CONTRATO_DECIDIDO", f"{c.get('contrato') or cid}: modo={d.get('modo')} pct={d.get('pct')} pagador={d.get('pagador')}")
    v = next((x for x in vista_contratos(_hotel()) if x.get("id") == cid), CG.vista(c))
    return jsonify({"ok": True, "contrato": v})


def _visible(cid):
    return cid and cid in {c.get("id") for c in CG.del_hotel(CG.leer(), _hotel())}


@contratos_bp.route("/api/ar/contratos/vincular", methods=["POST"])
@login_required
def api_vincular():
    d = request.get_json(silent=True) or {}
    cid = str(d.get("id") or "").strip()
    clave = str(d.get("clave") or "").strip()
    if not _visible(cid):
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    if clave not in {CG._clave_ap(f) for f in _facturas_ap(_hotel())}:
        return jsonify({"ok": False, "error": "esa factura no esta en AP de este hotel"}), 404
    try:
        CG.vincular(cid, clave, usuario=_usuario())
    except (KeyError, ValueError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    _audit("COMISION_VINCULADA", f"{cid}: {clave}")
    v = next((x for x in vista_contratos(_hotel()) if x.get("id") == cid), None)
    return jsonify({"ok": True, "contrato": v})


@contratos_bp.route("/api/ar/contratos/desvincular", methods=["POST"])
@login_required
def api_desvincular():
    d = request.get_json(silent=True) or {}
    cid = str(d.get("id") or "").strip()
    if not _visible(cid):
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    try:
        CG.desvincular(cid, str(d.get("clave") or ""), usuario=_usuario())
    except KeyError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    _audit("COMISION_SEPARADA", f"{cid}: {d.get('clave')}")
    v = next((x for x in vista_contratos(_hotel()) if x.get("id") == cid), None)
    return jsonify({"ok": True, "contrato": v})
