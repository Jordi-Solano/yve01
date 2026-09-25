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
  POST /api/ar/compensar              {id, importe, fecha?, nota?}   b89: 410/430, solo si paga la agencia
  POST /api/ar/compensar/anular       {id, comp_id}                  b89
  GET  /api/ar/contratos/<id>/beos                                   b91: las BEO del contrato (una por dia)
  GET  /api/ar/contratos/<id>/beo.pdf                                b91: en PDF, con el formato de la BEO de ejemplo
"""
from flask import Blueprint, jsonify, request, send_file
from flask_login import login_required, current_user

import beo_contrato as BEO
import compensaciones as CMP
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


def facturas_grupo(datos_dir=None, completas=None):
    """{"numero|hotel": {numero, estado, total, compensado, saldo, fecha_emision, fecha_cobro, cliente, hotel_id}}
    de reservas_credito.xlsx (lo que ve la pantalla; `completas` = compensaciones.leer_facturas_grupo)."""
    completas = CMP.leer_facturas_grupo(datos_dir) if completas is None else completas
    out = {}
    for clave, f in completas.items():
        out[clave] = {"numero": f["numero"], "numero_factura": f.get("numero_factura") or f["numero"], "estado": f["estado"], "total": f["total"],
                      "compensado": f["compensado"], "saldo": f["saldo"],
                      "fecha_emision": CG._txt(f.get("fecha_emision"))[:10], "fecha_cobro": CG._txt(f.get("fecha_cobro"))[:10],
                      "cliente": CG._txt(f.get("cliente")), "hotel_id": f["hotel_id"]}
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


def _contexto(hotel=None, datos_dir=None):
    """Lo que hace falta para ver o tocar los contratos de un hotel: los contratos, sus
    facturas de grupo, las facturas AP con su enlace y las compensaciones."""
    completas = CMP.leer_facturas_grupo(datos_dir)
    contratos = CG.del_hotel(CG.leer(datos_dir), hotel)
    filas_ap = _facturas_ap(hotel)
    return {"completas": completas, "fac": facturas_grupo(datos_dir, completas), "contratos": contratos,
            "enl": CG.enlazar(contratos, filas_ap) if contratos else {},
            "por_clave": {CG._clave_ap(f): f for f in filas_ap}, "comps": CMP.leer(datos_dir)}


def _grupo_y_comision(c, ctx):
    """(fila de la factura del grupo, fila AP de la comision unida) del contrato `c`."""
    fg = CMP.fila_grupo(c.get("factura_grupo"), c.get("hotel_id"), ctx["completas"])
    e = ctx["enl"].get(c["id"]) or {}
    fr = ctx["por_clave"].get(e.get("clave")) if e.get("clave") else None
    return fg, fr


def vista_contratos(hotel=None, datos_dir=None):
    ctx = _contexto(hotel, datos_dir)
    fac, enl, por_clave = ctx["fac"], ctx["enl"], ctx["por_clave"]
    out = []
    for c in ctx["contratos"]:
        v = CG.vista(c)
        num = CG._txt(c.get("factura_grupo"))
        v["factura"] = fac.get(num + "|" + CG._txt(c.get("hotel_id"))) or fac.get(num + "|") or ({"numero": num} if num else {})
        e = enl.get(c["id"]) or {}
        if e.get("clave") and e["clave"] in por_clave:
            fr = por_clave[e["clave"]]
            v["factura_comision_vista"] = dict(_mini(fr), origen=e.get("origen"), **CG.estado_comision(c, fr))
            v["factura_comision_vista"]["pagada"] = bool(fr.get("pagada")) and str(fr.get("pagada")) != "nan"
        else:
            v["factura_comision_vista"] = {}
        v["candidatas"] = [_mini(por_clave[k]) for k in (e.get("candidatas") or []) if k in por_clave]
        # b89: compensar la comision contra la factura del grupo (solo si paga la agencia)
        fg, fr = _grupo_y_comision(c, ctx)
        v["compensacion"] = CMP.vista(c, fg, fr, ctx["comps"])
        # b91: la BEO (orden de servicio) del contrato, con el formato de la de ejemplo
        try:
            v["beo"] = BEO.cotejo(c)
        except Exception:
            v["beo"] = {"n": 0}
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
    # b89: con compensaciones hechas (paga la agencia) no se cambia a "paga el cliente" ni
    # a "tarifa neta" sin anularlas antes: dejarian un 410/430 que la regla no permite
    if (d.get("pagador") == "cliente" or d.get("modo") == "neta") and any(x.get("contrato_id") == cid for x in CMP.leer()):
        return jsonify({"ok": False, "error": "Este contrato tiene compensaciones registradas: anúlalas antes de cambiar esto.", "regla": True}), 409
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


# ── b89: compensar la comision de la agencia contra la factura del grupo ─────
@contratos_bp.route("/api/ar/compensar", methods=["POST"])
@login_required
def api_compensar():
    """{id (contrato), importe, fecha?, nota?}. 409 si la regla de finanzas no lo deja
    (paga el cliente final, factura sin emitir o ya cobrada, sin factura de comision)."""
    d = request.get_json(silent=True) or {}
    cid = str(d.get("id") or "").strip()
    if not _visible(cid):
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    ctx = _contexto(_hotel())
    c = next((x for x in ctx["contratos"] if x.get("id") == cid), None)
    if c is None:
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    fg, fr = _grupo_y_comision(c, ctx)
    try:
        reg = CMP.compensar(c, fg, fr, d.get("importe"), d.get("fecha"), usuario=_usuario(), nota=d.get("nota", ""))
    except CMP.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    _audit("COMPENSACION", f"{c.get('contrato') or cid}: {reg['importe']:.2f} EUR de {reg['factura_comision']} contra {reg['factura_grupo']} ({reg['fecha']})")
    v = next((x for x in vista_contratos(_hotel()) if x.get("id") == cid), None)
    return jsonify({"ok": True, "compensacion": reg, "contrato": v})


@contratos_bp.route("/api/ar/compensar/anular", methods=["POST"])
@login_required
def api_anular_compensacion():
    """{id (contrato), comp_id}. 409 si la factura del grupo ya esta cobrada o la de comision pagada."""
    d = request.get_json(silent=True) or {}
    cid = str(d.get("id") or "").strip()
    comp_id = str(d.get("comp_id") or "").strip()
    if not _visible(cid):
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    if not any(x.get("id") == comp_id and x.get("contrato_id") == cid for x in CMP.leer()):
        return jsonify({"ok": False, "error": "compensacion no encontrada en este contrato"}), 404
    try:
        c = CMP.anular(comp_id, usuario=_usuario())
    except CMP.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except KeyError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    _audit("COMPENSACION_ANULADA", f"{cid}: {comp_id} ({c.get('importe')} EUR)")
    v = next((x for x in vista_contratos(_hotel()) if x.get("id") == cid), None)
    return jsonify({"ok": True, "contrato": v})


# ── b91: la BEO del contrato ────────────────────────────────────────────────
def _contrato_visible(cid):
    return next((c for c in CG.del_hotel(CG.leer(), _hotel()) if c.get("id") == cid), None)


@contratos_bp.route("/api/ar/contratos/<cid>/beos")
@login_required
def api_beos_contrato(cid):
    c = _contrato_visible(cid)
    if c is None:
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    lista = BEO.beos(c)
    return jsonify({"ok": True, "beos": lista, "cotejo": BEO.cotejo(c, lista)})


@contratos_bp.route("/api/ar/contratos/<cid>/beo.pdf")
@login_required
def api_beo_pdf(cid):
    c = _contrato_visible(cid)
    if c is None:
        return jsonify({"ok": False, "error": "contrato no encontrado en este hotel"}), 404
    lista = BEO.beos(c)
    if not lista:
        return jsonify({"ok": False, "error": "el contrato no trae servicios de evento (F&B, salas, programa) para una BEO"}), 404
    nombre = "BEO_" + "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in (c.get("contrato") or cid))[:40] + ".pdf"
    return send_file(BEO.pdf(lista), mimetype="application/pdf", as_attachment=False, download_name=nombre)
