# -*- coding: utf-8 -*-
"""tab_contratos.py — Yve.01 · AR > Contratos (b87).

Los contratos de grupo que ha leido Yve (fotos o PDF), con lo que finanzas dijo
que manda (24 sep 2026): la comision de la agencia segun el CAMPO DE COMISIONES
del contrato (tarifa neta = sin factura de comision; % = la agencia factura su
comision aparte) y QUIEN PAGA la factura del grupo. Lo que el contrato no dice
se pregunta aqui; no se supone.

Rutas (con sesion; el POST pasa por el CSRF de /api/*):
  GET  /api/ar/contratos
  POST /api/ar/contratos/decidir   {id, modo?: 'neta'|'porcentaje', pct?: {alojamiento, fb, salas}, pagador?: 'agencia'|'cliente'}
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


def vista_contratos(hotel=None, datos_dir=None):
    fac = facturas_grupo(datos_dir)
    out = []
    for c in CG.del_hotel(CG.leer(datos_dir), hotel):
        v = CG.vista(c)
        num = CG._txt(c.get("factura_grupo"))
        v["factura"] = fac.get(num + "|" + CG._txt(c.get("hotel_id"))) or fac.get(num + "|") or ({"numero": num} if num else {})
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
