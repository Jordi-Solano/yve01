"""
reclamaciones_ota.py — Yve.01
Loop de reclamación automática de comisiones OTA.

Patrón: dato que ya tenemos (verificador_comisiones) -> Claude redacta ->
gate de aprobación humana -> Brevo envía -> registro/historial.

- Lee las discrepancias del último verificacion_*.xlsx del tenant.
- El borrador se genera BAJO DEMANDA (no en cada carga) y se cachea.
- Nada se envía hasta que un humano pulsa "Aprobar y enviar".
- Estado por tenant en datos_dir()/reclamaciones_ota.json (escritura segura).
- Recuerda el último email usado por OTA en reclamaciones_ota_contactos.json.

NOTA (Render free tier): el disco es efímero; este estado se pierde en cada
deploy. Asumido para el demo; depende de la persistencia de pago.
"""
import os, json, glob, html
from datetime import datetime
from flask import Blueprint, request, jsonify, session
import pandas as pd
from tenant_dirs import datos_dir as _t_ddir, reportes_dir as _t_rdir

recl_ota_bp = Blueprint("reclamaciones_ota", __name__)

# Contactos por defecto por OTA (copia local para no depender de la API key al arrancar)
_OTA_DEFAULT_CONTACTS = {
    "booking.com": "finance@booking.com",
    "booking": "finance@booking.com",
    "expedia": "hotelbilling@expedia.com",
    "hotels.com": "hotelbilling@hotels.com",
    "despegar": "contratos@despegar.com",
}
NF = "NO_ENCONTRADO"


def _ddir():
    return _t_ddir()


def _rdir():
    return _t_rdir()


def _estado_path():
    return os.path.join(_ddir(), "reclamaciones_ota.json")


def _contactos_path():
    return os.path.join(_ddir(), "reclamaciones_ota_contactos.json")


def _safe_write_json(path, obj):
    """Escritura segura: tmp + assert de tamaño + os.replace (nunca open('w') directo sobre el destino)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = json.dumps(obj, ensure_ascii=False, indent=2)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(data)
    assert os.path.getsize(tmp) >= len(data.encode("utf-8")) - 4
    os.replace(tmp, path)


def _load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default


def _estado():
    return _load_json(_estado_path(), {})


def _guardar_estado(st):
    _safe_write_json(_estado_path(), st)


def _contactos():
    return _load_json(_contactos_path(), {})


def _guardar_contactos(c):
    _safe_write_json(_contactos_path(), c)


def _num(v):
    try:
        s = str(v).replace("EUR", "").replace("€", "").replace("%", "").strip()
        if not s or s == NF:
            return None
        return round(float(s.replace(",", ".")), 2)
    except Exception:
        return None


def _default_contacto(ota):
    o = (ota or "").lower().strip()
    for k, v in _OTA_DEFAULT_CONTACTS.items():
        if k in o:
            return v
    return ""


def _ultimo_reporte():
    files = sorted(glob.glob(os.path.join(_rdir(), "verificacion_*.xlsx")), reverse=True)
    return files[0] if files else None


def _discrepancias():
    """Lista de discrepancias (estado DISCREPANCIA) del último reporte de verificación."""
    rep = _ultimo_reporte()
    if not rep:
        return []
    df = None
    for sheet in ("Detalle", 0):
        try:
            df = pd.read_excel(rep, sheet_name=sheet)
            break
        except Exception:
            df = None
    if df is None or "estado" not in df.columns:
        return []
    dd = df[df["estado"].astype(str).str.upper() == "DISCREPANCIA"]
    out = []
    for _, r in dd.iterrows():
        num = str(r.get("numero_factura", "") or "")
        ota = str(r.get("nombre_ota", "") or "")
        per = str(r.get("periodo_inicio", "") or "")
        rid = num + "|" + ota + "|" + per
        out.append({
            "id": rid,
            "numero_factura": num,
            "numero_reserva": str(r.get("numero_reserva", "") or ""),
            "ota": ota,
            "hotel": str(r.get("nombre_hotel", "") or ""),
            "periodo_inicio": per,
            "periodo_fin": str(r.get("periodo_fin", "") or ""),
            "importe_bruto": _num(r.get("importe_bruto")),
            "comision_contrato": _num(r.get("porcentaje_pactado")),
            "comision_cobrada": _num(r.get("porcentaje_factura")),
            "diferencia_pp": _num(r.get("diferencia_pp")),
            "importe_reclamable": _num(r.get("discrepancia_euros")),
        })
    return out


@recl_ota_bp.route("/api/reclamaciones_ota/list")
def list_reclamaciones():
    st = _estado()
    cont = _contactos()
    disc = _discrepancias()
    items = []
    for d in disc:
        s = st.get(d["id"], {})
        it = dict(d)
        it["estado"] = s.get("estado", "PENDIENTE")
        it["asunto"] = s.get("asunto", "")
        it["cuerpo"] = s.get("cuerpo", "")
        it["destinatario"] = s.get("destinatario") or cont.get(d["ota"].lower(), "") or _default_contacto(d["ota"])
        it["fecha_generado"] = s.get("fecha_generado", "")
        it["fecha_enviada"] = s.get("fecha_enviada", "")
        it["tiene_borrador"] = bool(s.get("cuerpo"))
        items.append(it)
    n_pend = sum(1 for i in items if i["estado"] == "PENDIENTE")
    total_reclamable = round(sum((i["importe_reclamable"] or 0) for i in items if i["estado"] != "DESCARTADA"), 2)
    return jsonify({"ok": True, "items": items, "n_pendientes": n_pend, "total_reclamable": total_reclamable})


@recl_ota_bp.route("/api/reclamaciones_ota/generar", methods=["POST"])
def generar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    idioma = data.get("idioma", "es")
    disc = {d["id"]: d for d in _discrepancias()}
    if rid not in disc:
        return jsonify({"ok": False, "error": "discrepancia no encontrada"}), 404
    try:
        from generador_emails import generar_reclamacion_ota
        em = generar_reclamacion_ota(disc[rid], idioma)
    except Exception as e:
        return jsonify({"ok": False, "error": "No se pudo redactar: " + str(e)[:140]}), 500
    st = _estado()
    s = st.get(rid, {})
    s.update({
        "asunto": em["asunto"], "cuerpo": em["cuerpo"],
        "estado": s.get("estado", "PENDIENTE"),
        "fecha_generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "idioma": idioma,
    })
    st[rid] = s
    _guardar_estado(st)
    return jsonify({"ok": True, "asunto": em["asunto"], "cuerpo": em["cuerpo"]})


@recl_ota_bp.route("/api/reclamaciones_ota/editar", methods=["POST"])
def editar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    if not rid:
        return jsonify({"ok": False, "error": "falta id"}), 400
    st = _estado()
    s = st.get(rid, {})
    if "asunto" in data:
        s["asunto"] = str(data["asunto"])
    if "cuerpo" in data:
        s["cuerpo"] = str(data["cuerpo"])
    if "destinatario" in data:
        s["destinatario"] = str(data["destinatario"]).strip()
    s.setdefault("estado", "PENDIENTE")
    st[rid] = s
    _guardar_estado(st)
    return jsonify({"ok": True})


@recl_ota_bp.route("/api/reclamaciones_ota/descartar", methods=["POST"])
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


@recl_ota_bp.route("/api/reclamaciones_ota/aprobar_enviar", methods=["POST"])
def aprobar_enviar():
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get("id")
    destinatario = (data.get("destinatario") or "").strip()
    asunto = (data.get("asunto") or "").strip()
    cuerpo = (data.get("cuerpo") or "").strip()
    if not rid:
        return jsonify({"ok": False, "error": "falta id"}), 400
    if not destinatario or "@" not in destinatario:
        return jsonify({"ok": False, "error": "Introduce un email de destino válido"}), 400
    if not asunto or not cuerpo:
        return jsonify({"ok": False, "error": "Falta asunto o cuerpo del email"}), 400
    disc = {d["id"]: d for d in _discrepancias()}
    ota = disc.get(rid, {}).get("ota", "")
    num = disc.get(rid, {}).get("numero_factura", "")
    cuerpo_html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0f172a;white-space:pre-wrap;line-height:1.5">' + html.escape(cuerpo) + "</div>"
    try:
        from notificaciones import enviar_email
        ok = enviar_email(destinatario, asunto, cuerpo_html, "reclamacion_ota")
    except Exception as e:
        return jsonify({"ok": False, "error": "Error enviando: " + str(e)[:140]}), 500
    if not ok:
        return jsonify({"ok": False, "error": "El proveedor de email no pudo enviar (revisa la configuración de Brevo)."}), 500
    st = _estado()
    s = st.get(rid, {})
    s.update({
        "estado": "ENVIADA", "asunto": asunto, "cuerpo": cuerpo, "destinatario": destinatario,
        "fecha_enviada": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })
    st[rid] = s
    _guardar_estado(st)
    # Recordar el email de esta OTA para la próxima
    if ota:
        c = _contactos()
        c[ota.lower()] = destinatario
        _guardar_contactos(c)
    # Auditoría (import perezoso; dashboard ya está cargado en runtime)
    try:
        from dashboard import _audit
        _audit("RECLAMACION_OTA_ENVIADA", ota + " · factura " + str(num) + " · " + destinatario, session.get("username", "sistema"))
    except Exception:
        pass
    return jsonify({"ok": True})
