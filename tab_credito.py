# -*- coding: utf-8 -*-
"""tab_credito.py — Yve.01 · AR > Peticion de credito (b90).

El proceso y las reglas estan en peticiones_credito.py. Aqui solo las rutas.

Rutas (con sesion; los POST JSON pasan por el CSRF de /api/*; la subida del
informe es multipart y comprueba el token ella misma):
  GET  /api/credito/peticiones                 las del hotel activo (vista de grupo: todas)
  POST /api/credito/peticion                   {cliente}  crea (o devuelve la abierta)
  GET  /api/credito/peticion/<pid>
  POST /api/credito/peticion/<pid>/guardar     {referencias?, fiscal?, informa?, comercial?}
  POST /api/credito/peticion/<pid>/informa     multipart: fichero (PDF o imagen)
  GET  /api/credito/peticion/<pid>/informa     descarga el informe adjunto
  POST /api/credito/peticion/<pid>/firmar      firma 1: quien pide (b92: con la firma de Direccion apagada
                                               en config_aprobaciones.json, esta firma APRUEBA)
  POST /api/credito/peticion/<pid>/direccion   {decision: aprobar|rechazar, nota}  firma 2: solo rol Direccion, nunca quien pidio
  POST /api/credito/peticion/<pid>/reabrir     una rechazada vuelve a borrador
"""
import hmac

from flask import Blueprint, jsonify, request, send_file, session
from flask_login import login_required, current_user

import peticiones_credito as PC

credito_bp = Blueprint("credito", __name__)


def _usuario():
    try:
        return getattr(current_user, "username", None) or ""
    except Exception:
        return ""


def _nombre():
    try:
        return getattr(current_user, "nombre", None) or _usuario()
    except Exception:
        return _usuario()


def _rol():
    try:
        return getattr(current_user, "rol", "") or ""
    except Exception:
        return ""


def _hotel():
    try:
        import censo_hoteles
        return censo_hoteles.activo()
    except Exception:
        return ""


def _hotel_guardar():
    """El hotel con el que se etiqueta una peticion nueva (la regla de siempre: el
    activo; con un solo hotel, ese; con varios y ninguno elegido, sin hotel)."""
    try:
        import censo_hoteles
        return censo_hoteles.para_guardar()
    except Exception:
        return ""


def _audit(accion, detalle):
    try:
        from dashboard import _audit as _a
        _a(accion, detalle)
    except Exception:
        pass


def _visibles():
    """Las del hotel activo (igualdad estricta); en la vista de grupo, todas."""
    h = _hotel()
    return [p for p in PC.leer() if not h or PC._txt(p.get("hotel_id")) == h]


def _la(pid):
    return next((p for p in _visibles() if p.get("id") == pid), None)


def _vista(p):
    return PC.vista(p, _usuario(), _rol())


def _fila(p):
    """Lo que ensena la lista (sin el historial entero)."""
    co = p.get("comercial") or {}
    fs = (p.get("firmas") or {}).get("solicitante") or {}
    fd = (p.get("firmas") or {}).get("direccion") or {}
    return {"id": p.get("id"), "cliente": p.get("cliente"), "estado": p.get("estado"), "hotel_id": p.get("hotel_id"),
            "hotel": PC._nombre_hotel(p.get("hotel_id")), "limite": co.get("limite"), "revision": co.get("revision"),
            "potencial": co.get("potencial"), "creada": p.get("creada"), "creada_por": p.get("creada_por"),
            "solicitante": fs.get("nombre") or "", "firmada": fs.get("cuando") or "",
            "direccion": fd.get("nombre") or "", "decision": fd.get("decision") or "", "nota_direccion": fd.get("nota") or "",
            "firmada_direccion": fd.get("cuando") or "", "firma_unica": bool(p.get("firma_unica")),
            "puede_firmar_direccion": PC.puede_firmar_direccion(p, _usuario(), _rol())[0]}


@credito_bp.route("/api/credito/peticiones")
@login_required
def api_peticiones():
    # primero lo pendiente de Direccion, luego los borradores, luego el resto; lo mas nuevo arriba
    lista = sorted(_visibles(), key=lambda p: p.get("creada") or "", reverse=True)
    lista.sort(key=lambda p: {"PENDIENTE_DIRECCION": 0, "BORRADOR": 1}.get(p.get("estado"), 2))
    filas = [_fila(p) for p in lista]
    return jsonify({"ok": True, "peticiones": filas, "rol": _rol(), "es_direccion": _rol() == PC.ROL_DIRECCION,
                    "firma_direccion": PC.firma_direccion_activa(),
                    "n_pendientes_direccion": sum(1 for f in filas if f["estado"] == "PENDIENTE_DIRECCION"),
                    "n_para_mi": sum(1 for f in filas if f["puede_firmar_direccion"])})


@credito_bp.route("/api/credito/peticion", methods=["POST"])
@login_required
def api_nueva():
    d = request.get_json(silent=True) or {}
    try:
        p, creada = PC.nueva(d.get("cliente"), _hotel_guardar(), _usuario(), _nombre())
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    h = _hotel()
    if h and PC._txt(p.get("hotel_id")) != h:
        return jsonify({"ok": False, "error": f"El cliente ya tiene una petición abierta en otro hotel del grupo ({p.get('id')})."}), 409
    if creada:
        _audit("CREDITO_PETICION", f"{p['id']} {p['cliente']}")
    return jsonify({"ok": True, "creada": creada, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>")
@login_required
def api_una(pid):
    p = _la(pid)
    if p is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    return jsonify({"ok": True, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>/guardar", methods=["POST"])
@login_required
def api_guardar(pid):
    if _la(pid) is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    try:
        p = PC.guardar(pid, request.get_json(silent=True) or {}, _usuario())
    except PC.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except (ValueError, KeyError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    return jsonify({"ok": True, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>/informa", methods=["GET", "POST"])
@login_required
def api_informa(pid):
    p = _la(pid)
    if p is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    if request.method == "GET":
        ruta = PC.ruta_informa(p)
        if not ruta:
            return jsonify({"ok": False, "error": "no hay informe adjunto"}), 404
        return send_file(ruta, as_attachment=True, download_name=(p.get("informa") or {}).get("nombre_fichero") or "informe.pdf")
    # multipart: el guardia de CSRF de /api/* no mira los multipart; se mira aqui
    tok = request.headers.get("X-CSRF-Token") or ""
    if not tok or not hmac.compare_digest(tok, session.get("csrf_token", "")):
        return jsonify({"ok": False, "error": "CSRF inválido", "csrf_error": True}), 403
    f = request.files.get("fichero")
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "falta el fichero"}), 400
    try:
        p = PC.adjuntar_informa(pid, f.filename, f.read(PC.MAX_INFORMA + 1), _usuario())
    except PC.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except (ValueError, KeyError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    _audit("CREDITO_INFORMA", f"{pid}: {f.filename}")
    return jsonify({"ok": True, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>/firmar", methods=["POST"])
@login_required
def api_firmar(pid):
    if _la(pid) is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    try:
        p = PC.firmar_solicitante(pid, _usuario(), _nombre())
    except PC.PermisoError as e:
        return jsonify({"ok": False, "error": str(e)}), 403
    except PC.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except KeyError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    _audit("CREDITO_FIRMA_SOLICITANTE", f"{pid} {p.get('cliente')} limite {(p.get('comercial') or {}).get('limite')}")
    if p.get("estado") == "PENDIENTE_DIRECCION":
        try:    # aviso a Direccion (push, si el servidor lo tiene configurado); no bloquea
            from notificaciones import enviar_push
            enviar_push("Petición de crédito para firmar", f"{p.get('cliente')}: límite {(p.get('comercial') or {}).get('limite')} € · pide {_nombre()}",
                        url="/app?tab=ar_real", tipo="credito_pendiente_firma", roles=[PC.ROL_DIRECCION])
        except Exception:
            pass
    else:       # b92: firma de Direccion apagada → aprobada con una firma
        _audit("CREDITO_APROBADO_UNA_FIRMA", f"{pid} {p.get('cliente')} limite {p.get('limite_aprobado')}")
    return jsonify({"ok": True, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>/direccion", methods=["POST"])
@login_required
def api_direccion(pid):
    if _la(pid) is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    d = request.get_json(silent=True) or {}
    try:
        p = PC.firmar_direccion(pid, _usuario(), _rol(), d.get("decision"), d.get("nota", ""), _nombre())
    except PC.PermisoError as e:
        return jsonify({"ok": False, "error": str(e)}), 403
    except PC.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    _audit("CREDITO_FIRMA_DIRECCION", f"{pid} {p.get('cliente')}: {d.get('decision')} {p.get('limite_aprobado') or ''}")
    return jsonify({"ok": True, "peticion": _vista(p)})


@credito_bp.route("/api/credito/peticion/<pid>/reabrir", methods=["POST"])
@login_required
def api_reabrir(pid):
    if _la(pid) is None:
        return jsonify({"ok": False, "error": "petición no encontrada en este hotel"}), 404
    try:
        p = PC.reabrir(pid, _usuario())
    except PC.ReglaError as e:
        return jsonify({"ok": False, "error": str(e), "regla": True}), 409
    except KeyError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    _audit("CREDITO_REABIERTA", pid)
    return jsonify({"ok": True, "peticion": _vista(p)})
