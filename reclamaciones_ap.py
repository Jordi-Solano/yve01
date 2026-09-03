"""
reclamaciones_ap.py — Yve.01 (OLA A)
Reclamar al PROVEEDOR una factura rectificativa o un abono.

Mismo patron que reclamaciones_ota: dato que ya tenemos (matching AP y las
decisiones del panel de aprobacion) -> borrador -> gate humano ("Aprobar y
enviar") -> Brevo -> registro. Nada sale sin que una persona pulse el boton.

Que se reclama:
- Facturas con incidencia de matching (`_ESTADOS_INCIDENCIA` del panel:
  DIFERENCIA_IMPORTE, FACTURA_SIN_ALBARAN, ...) que NO estan aprobadas.
- Facturas RECHAZADAS en "Facturas por aprobar" (se pide el abono).

El borrador se escribe SIN IA: los datos de la factura ya son estructurados
(numero, proveedor, importe, que no cuadra), asi que una plantilla con las
cifras es mas fiable que pedirle a un modelo que las repita. Se puede editar.

Estado por tenant en datos_dir()/reclamaciones_ap.json (escritura segura) y
ultimo email usado por proveedor en reclamaciones_ap_contactos.json.
NO toca aprobaciones_ap.xlsx ni nada de Oracle.
"""
import html
import json
import os
import unicodedata
from datetime import datetime

import pandas as pd
from flask import Blueprint, jsonify, request

from tenant_dirs import datos_dir as _t_ddir

recl_ap_bp = Blueprint("reclamaciones_ap", __name__)

TIPO_ABONO = "ABONO"            # factura rechazada: que la anulen / abonen
TIPO_CORRECCION = "CORRECCION"  # no cuadra: factura rectificativa


# ── ficheros de estado ──────────────────────────────────────────────────────
def _ddir():
    return _t_ddir()


def _estado_path():
    return os.path.join(_ddir(), "reclamaciones_ap.json")


def _contactos_path():
    return os.path.join(_ddir(), "reclamaciones_ap_contactos.json")


def _safe_write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _estado():
    return _load_json(_estado_path(), {})


def _guardar_estado(st):
    _safe_write_json(_estado_path(), st)


def _contactos():
    return _load_json(_contactos_path(), {})


def _guardar_contactos(c):
    _safe_write_json(_contactos_path(), c)


# ── utilidades ──────────────────────────────────────────────────────────────
def _txt(v):
    if v is None:
        return ""
    try:
        if isinstance(v, float) and v != v:
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _num(v):
    try:
        if isinstance(v, str):
            v = v.replace("€", "").strip()
            if "," in v:                       # "1.234,56"
                v = v.replace(".", "").replace(",", ".")
        f = float(v)
        return 0.0 if f != f else round(f, 2)
    except Exception:
        return 0.0


def _norm(x):
    x = unicodedata.normalize("NFKD", _txt(x).lower())
    return " ".join("".join(c for c in x if c.isalnum() or c.isspace()).split())


def _fmt(n):
    s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"


def _emails_proveedores():
    """nombre normalizado -> email_contacto de proveedores.xlsx."""
    ruta = os.path.join(_ddir(), "proveedores.xlsx")
    out = {}
    if not os.path.exists(ruta):
        return out
    try:
        df = pd.read_excel(ruta)
    except Exception:
        return out
    if "nombre_proveedor" not in df.columns or "email_contacto" not in df.columns:
        return out
    for _, r in df.iterrows():
        n, e = _norm(r.get("nombre_proveedor")), _txt(r.get("email_contacto"))
        if n and e and "@" in e:
            out[n] = e
    return out


def _email_de(proveedor, mapa):
    n = _norm(proveedor)
    if not n:
        return ""
    if n in mapa:
        return mapa[n]
    for k, v in mapa.items():          # "Makro" ~ "Makro Cash & Carry SL"
        if k and (k in n or n in k):
            return v
    return ""


def _entidad():
    """Quien firma: el hotel activo o el nombre configurado."""
    try:
        import censo_hoteles
        h = censo_hoteles.activo()
        if h:
            return censo_hoteles.nombre_de(h) or h
    except Exception:
        pass
    try:
        with open(os.path.join(_ddir(), "hotel_config.json"), encoding="utf-8") as f:
            return _txt((json.load(f) or {}).get("hotel_nombre"))
    except Exception:
        return ""


def _usuario():
    try:
        from flask_login import current_user
        return getattr(current_user, "username", None) or "sistema"
    except Exception:
        return "sistema"


# ── candidatas ──────────────────────────────────────────────────────────────
def candidatas():
    """Facturas AP que hay que reclamar al proveedor (hotel activo)."""
    try:
        import app_aprobacion_ap as P
        from dashboard import cargar_datos_ap
        df = cargar_datos_ap()
        acciones = P._acciones_por_clave()
        incidencias = P._ESTADOS_INCIDENCIA
        clave_de = P.clave_factura
        nombre_hotel = P._nombre_hotel
    except Exception:
        return []
    if df is None or df.empty:
        return []
    mapa = _emails_proveedores()
    out = []
    for _, r in df.iterrows():
        clave = clave_de(r)
        accion = _txt(acciones.get(clave, "")).upper()
        estado = _txt(r.get("estado_matching")).upper()
        if accion == "APROBADA":
            continue
        if accion == "RECHAZADA":
            tipo = TIPO_ABONO
        elif estado in incidencias:
            tipo = TIPO_CORRECCION
        else:
            continue
        prov = _txt(r.get("nombre_proveedor"))
        out.append({
            "id":              clave,
            "tipo":            tipo,
            "numero_factura":  _txt(r.get("numero_factura")),
            "proveedor":       prov,
            "email_proveedor": _email_de(prov, mapa),
            "fecha_factura":   _txt(r.get("fecha_factura")),
            "total_factura":   _num(r.get("total_factura")),
            "estado_matching": estado,
            "detalle":         _txt(r.get("detalle_matching")) or _txt(r.get("alerta_detalle")),
            "comentario":      _txt(r.get("comentario")) if accion == "RECHAZADA" else "",
            "hotel_id":        _txt(r.get("hotel_id")),
            "hotel":           nombre_hotel(r.get("hotel_id")),
        })
    return out


# ── borrador (plantilla, sin IA) ────────────────────────────────────────────
_MOTIVOS = {
    "DIFERENCIA_IMPORTE":    "el importe facturado no coincide con el del albaran/pedido",
    "DIFERENCIA_PO_IMPORTE": "el importe facturado no coincide con el del pedido",
    "DIFERENCIA_LINEA":      "hay lineas facturadas que no coinciden con lo recibido",
    "DISCREPANCIA":          "los datos facturados no coinciden con nuestros registros",
    "DISCREPANCIA_PO":       "la factura no coincide con el pedido autorizado",
    "FACTURA_SIN_ALBARAN":   "no consta ningun albaran de entrega asociado a esta factura",
    "ALERTA_CONSUMO":        "el consumo facturado se sale del rango habitual",
}


def redactar(c, idioma="es", firma_nombre="", firma_entidad=""):
    num = c.get("numero_factura") or "(sin numero)"
    total = _fmt(_num(c.get("total_factura")))
    prov = c.get("proveedor") or "proveedor"
    hotel = c.get("hotel") or firma_entidad or ""
    fecha = c.get("fecha_factura") or ""
    motivo = _MOTIVOS.get(c.get("estado_matching", ""), "los datos facturados no coinciden con nuestros registros")
    detalle = c.get("detalle") or c.get("comentario") or ""
    # Sin hotel conocido no se inventa ("nuestro hotel"): se omite la frase.
    emitida = f", emitida a {hotel}" if hotel else ""
    issued = f", issued to {hotel}" if hotel else ""
    firma = "\n".join(x for x in (firma_nombre or "[Nombre]", firma_entidad or hotel) if x)

    if idioma == "en":
        if c.get("tipo") == TIPO_ABONO:
            asunto = f"Credit note request — invoice {num} ({total})"
            cuerpo = (f"Dear {prov},\n\n"
                      f"Invoice {num}{' dated ' + fecha if fecha else ''} for {total}{issued} "
                      f"has been rejected in our approval process"
                      f"{' for the following reason: ' + detalle if detalle else ''}.\n\n"
                      f"Please issue a credit note cancelling invoice {num} for {total}.\n\n"
                      f"Kind regards,\n{firma}")
        else:
            asunto = f"Corrected invoice request — invoice {num} ({total})"
            cuerpo = (f"Dear {prov},\n\n"
                      f"While reviewing invoice {num}{' dated ' + fecha if fecha else ''} for {total}{issued}, "
                      f"we found that {motivo}"
                      f"{': ' + detalle if detalle else ''}.\n\n"
                      f"Please issue a corrected invoice (or a credit note for the difference) so we can process the payment.\n\n"
                      f"Kind regards,\n{firma}")
        return asunto, cuerpo

    if c.get("tipo") == TIPO_ABONO:
        asunto = f"Solicitud de abono — factura {num} ({total})"
        cuerpo = (f"Estimados {prov},\n\n"
                  f"La factura {num}{' de fecha ' + fecha if fecha else ''} por importe de {total}{emitida} "
                  f"ha sido rechazada en nuestro proceso de aprobacion"
                  f"{' por el siguiente motivo: ' + detalle if detalle else ''}.\n\n"
                  f"Les rogamos emitan una factura rectificativa (abono) que anule la factura {num} "
                  f"por {total}.\n\n"
                  f"Quedamos a su disposicion para cualquier aclaracion.\n\n"
                  f"Atentamente,\n{firma}")
    else:
        asunto = f"Solicitud de factura rectificativa — factura {num} ({total})"
        cuerpo = (f"Estimados {prov},\n\n"
                  f"Al revisar la factura {num}{' de fecha ' + fecha if fecha else ''} por importe de {total}{emitida}, "
                  f"hemos detectado que {motivo}"
                  f"{': ' + detalle if detalle else ''}.\n\n"
                  f"Les rogamos emitan una factura rectificativa (o un abono por la diferencia) "
                  f"para poder tramitar el pago.\n\n"
                  f"Quedamos a su disposicion para cualquier aclaracion.\n\n"
                  f"Atentamente,\n{firma}")
    return asunto, cuerpo


def _cifras_que_faltan(cuerpo, c):
    """Un email de reclamacion tiene que citar la factura y el importe."""
    faltan = []
    num = _txt(c.get("numero_factura"))
    if num and num not in cuerpo:
        faltan.append("el numero de factura " + num)
    tot = _num(c.get("total_factura"))
    if tot:
        variantes = {_fmt(tot), f"{tot:.2f}", f"{tot:,.2f}", f"{tot:.2f}".replace(".", ","),
                     f"{tot:.0f}" if tot == int(tot) else ""}
        if not any(v and v in cuerpo for v in variantes):
            faltan.append("el importe " + _fmt(tot))
    return faltan


# ── endpoints ───────────────────────────────────────────────────────────────
@recl_ap_bp.route("/api/reclamaciones_ap/list")
def list_reclamaciones():
    st = _estado()
    cont = _contactos()
    items = []
    for c in candidatas():
        s = st.get(c["id"], {})
        it = dict(c)
        it["estado"] = s.get("estado", "PENDIENTE")
        it["asunto"] = s.get("asunto", "")
        it["cuerpo"] = s.get("cuerpo", "")
        it["destinatario"] = (s.get("destinatario") or cont.get(_norm(c["proveedor"]), "")
                              or c["email_proveedor"])
        it["fecha_generado"] = s.get("fecha_generado", "")
        it["fecha_enviada"] = s.get("fecha_enviada", "")
        it["tiene_borrador"] = bool(s.get("cuerpo"))
        items.append(it)
    n_pend = sum(1 for i in items if i["estado"] == "PENDIENTE")
    total = round(sum(i["total_factura"] for i in items if i["estado"] != "DESCARTADA"), 2)
    return jsonify({"ok": True, "items": items, "n_pendientes": n_pend, "total_en_disputa": total})


@recl_ap_bp.route("/api/reclamaciones_ap/generar", methods=["POST"])
def generar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    idioma = data.get("idioma", "es")
    cand = {c["id"]: c for c in candidatas()}
    if rid not in cand:
        return jsonify({"ok": False, "error": "factura no encontrada entre las reclamables"}), 404
    try:
        from flask_login import current_user
        nombre = _txt(getattr(current_user, "nombre", ""))
    except Exception:
        nombre = ""
    asunto, cuerpo = redactar(cand[rid], idioma, firma_nombre=nombre, firma_entidad=_entidad())
    st = _estado()
    s = st.get(rid, {})
    s.update({"asunto": asunto, "cuerpo": cuerpo, "estado": s.get("estado", "PENDIENTE"),
              "fecha_generado": datetime.now().strftime("%Y-%m-%d %H:%M"), "idioma": idioma})
    st[rid] = s
    _guardar_estado(st)
    return jsonify({"ok": True, "asunto": asunto, "cuerpo": cuerpo})


@recl_ap_bp.route("/api/reclamaciones_ap/editar", methods=["POST"])
def editar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    if not rid:
        return jsonify({"ok": False, "error": "falta id"}), 400
    st = _estado()
    s = st.get(rid, {})
    for k in ("asunto", "cuerpo", "destinatario"):
        if k in data:
            s[k] = str(data[k]).strip() if k == "destinatario" else str(data[k])
    s.setdefault("estado", "PENDIENTE")
    st[rid] = s
    _guardar_estado(st)
    return jsonify({"ok": True})


@recl_ap_bp.route("/api/reclamaciones_ap/descartar", methods=["POST"])
def descartar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    if not rid:
        return jsonify({"ok": False, "error": "falta id"}), 400
    st = _estado()
    s = st.get(rid, {})
    s["estado"] = "DESCARTADA"
    st[rid] = s
    _guardar_estado(st)
    return jsonify({"ok": True})


@recl_ap_bp.route("/api/reclamaciones_ap/aprobar_enviar", methods=["POST"])
def aprobar_enviar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    destinatario = _txt(data.get("destinatario"))
    asunto = _txt(data.get("asunto"))
    cuerpo = _txt(data.get("cuerpo"))
    if not rid:
        return jsonify({"ok": False, "error": "falta id"}), 400
    if not destinatario or "@" not in destinatario:
        return jsonify({"ok": False, "error": "Introduce un email de destino válido"}), 400
    if not asunto or not cuerpo:
        return jsonify({"ok": False, "error": "Falta asunto o cuerpo del email"}), 400

    st = _estado()
    ya = st.get(rid, {})
    if ya.get("estado") == "ENVIADA":         # idempotencia: nunca dos veces
        return jsonify({"ok": False, "ya_enviada": True,
                        "error": ("Esta reclamación ya se envió"
                                  + (" el " + ya["fecha_enviada"] if ya.get("fecha_enviada") else "")
                                  + (" a " + ya["destinatario"] if ya.get("destinatario") else "")
                                  + ". No se reenvía.")}), 409

    cand = {c["id"]: c for c in candidatas()}
    c = cand.get(rid)
    if not c:
        return jsonify({"ok": False, "error": "La factura ya no está entre las reclamables (¿aprobada?)"}), 404
    faltan = _cifras_que_faltan(cuerpo, c)
    if faltan:
        return jsonify({"ok": False, "sin_cifras": True,
                        "error": ("El email no se envía: el texto no menciona " + ", ".join(faltan)
                                  + ". Una reclamación tiene que llevar las cifras.")}), 422

    cuerpo_html = ('<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f172a;'
                   'white-space:pre-wrap;line-height:1.5">' + html.escape(cuerpo) + "</div>")
    try:
        from notificaciones import enviar_email
        ok = enviar_email(destinatario, asunto, cuerpo_html, "reclamacion_ap")
    except Exception as e:
        return jsonify({"ok": False, "error": "Error enviando: " + str(e)[:140]}), 500
    if not ok:
        return jsonify({"ok": False, "error": "El proveedor de email no pudo enviar (revisa la configuración de Brevo)."}), 500

    s = st.get(rid, {})
    s.update({"estado": "ENVIADA", "asunto": asunto, "cuerpo": cuerpo, "destinatario": destinatario,
              "fecha_enviada": datetime.now().strftime("%Y-%m-%d %H:%M"), "enviada_por": _usuario(),
              "tipo": c["tipo"], "numero_factura": c["numero_factura"], "proveedor": c["proveedor"]})
    st[rid] = s
    _guardar_estado(st)
    if c["proveedor"]:
        cont = _contactos()
        cont[_norm(c["proveedor"])] = destinatario
        _guardar_contactos(cont)
    try:
        from dashboard import _audit
        _audit("RECLAMACION_AP_ENVIADA",
               f"{c['tipo']} · {c['proveedor']} · factura {c['numero_factura'] or rid} · {destinatario}",
               _usuario())
    except Exception:
        pass
    return jsonify({"ok": True})
