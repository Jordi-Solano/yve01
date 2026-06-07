"""
notificaciones.py — Yve.01
Sistema de notificaciones por email para alertas financieras.
Uso:
  python notificaciones.py              → escanea reportes y envía alertas pendientes
  python notificaciones.py --check      → solo muestra alertas sin enviar
  Importable: from notificaciones import escanear_alertas, enviar_pendientes
"""

import os, json, glob, smtplib, sys
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

import pandas as pd

BASE_DIR       = Path(__file__).parent
REPORTES_DIR   = BASE_DIR / "reportes"
PROCESADAS_DIR = BASE_DIR / "facturas-procesadas"
REFERENCIA_DIR = BASE_DIR / "datos-referencia"
CONFIG_PATH    = REFERENCIA_DIR / "hotel_config.json"
HISTORIAL_PATH = REFERENCIA_DIR / "notificaciones_historial.json"

# ── Config helpers ────────────────────────────────────────────────────────

def _load_env():
    """Carga variables del .env."""
    env_path = BASE_DIR / ".env"
    env = {}
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")
    return env

def _load_config():
    """Carga hotel_config.json."""
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def _get_destinatario(config):
    """Devuelve email del Financial Controller."""
    alertas = config.get("alertas", {})
    email = alertas.get("email", "")
    if not email:
        # Fallback: email del FC en usuarios
        email = config.get("usuarios", {}).get("fc_email", "")
    return email

def _alertas_activas(config):
    """Devuelve dict con qué alertas están activas."""
    alertas = config.get("alertas", {})
    return {
        "ar_discrepancia":  alertas.get("ar_discrepancia", True),
        "ar_falta_di":      alertas.get("ar_falta_di", True),
        "drr_oob":          alertas.get("drr_oob", True),
        "ap_discrepancia":  alertas.get("ap_discrepancia", True),
    }

# ── Historial ─────────────────────────────────────────────────────────────

def _load_historial():
    if HISTORIAL_PATH.exists():
        try:
            return json.loads(HISTORIAL_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []

def _save_historial(historial):
    REFERENCIA_DIR.mkdir(exist_ok=True)
    HISTORIAL_PATH.write_text(json.dumps(historial, ensure_ascii=False, indent=2), encoding="utf-8")

def _registrar(tipo, asunto, destinatario, estado, detalle=""):
    historial = _load_historial()
    historial.append({
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tipo": tipo,
        "asunto": asunto,
        "destinatario": destinatario,
        "estado": estado,
        "detalle": detalle,
    })
    # Mantener últimos 200
    _save_historial(historial[-200:])

# ── Email ─────────────────────────────────────────────────────────────────

def enviar_email(destinatario, asunto, cuerpo_html, tipo="general"):
    """Envía un email vía SMTP. Devuelve True/False."""
    env = _load_env()
    smtp_server = env.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port   = int(env.get("SMTP_PORT", "587"))
    smtp_user   = env.get("SMTP_USER", "")
    smtp_pass   = env.get("SMTP_PASSWORD", "")

    if not smtp_user or not smtp_pass:
        _registrar(tipo, asunto, destinatario, "error", "SMTP no configurado")
        return False

    msg = MIMEMultipart("alternative")
    msg["From"]    = smtp_user
    msg["To"]      = destinatario
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo_html, "html", "utf-8"))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
        _registrar(tipo, asunto, destinatario, "enviado")
        return True
    except Exception as e:
        _registrar(tipo, asunto, destinatario, "error", str(e)[:200])
        return False

# ── Plantilla email ───────────────────────────────────────────────────────



def enviar_slack(webhook_url, mensaje, asunto="", tipo="general"):
    import urllib.request as _req, json as _j
    if not webhook_url or "hooks.slack.com" not in webhook_url:
        _registrar(tipo, asunto or "slack", "slack", "error", "Webhook no configurado")
        return False
    try:
        text = ("*" + (asunto or "Yve") + "*" + chr(10) + mensaje)
        payload = _j.dumps({"text": text}).encode("utf-8")
        req = _req.Request(webhook_url, data=payload,
                           headers={"Content-Type": "application/json"}, method="POST")
        with _req.urlopen(req, timeout=10) as r:
            ok = r.read() == b"ok"
        _registrar(tipo, asunto, "slack", "enviado" if ok else "warning")
        return ok
    except Exception as e:
        _registrar(tipo, asunto, "slack", "error", str(e)[:200])
        return False



def enviar_whatsapp(numero_destino, mensaje, asunto="", tipo="general"):
    """Envia WhatsApp via Twilio. Vars: TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_FROM"""
    import os
    sid   = os.environ.get("TWILIO_ACCOUNT_SID", "")
    token = os.environ.get("TWILIO_AUTH_TOKEN", "")
    desde = os.environ.get("TWILIO_WHATSAPP_FROM", "")
    if not all([sid, token, desde, numero_destino]):
        _registrar(tipo, asunto, "whatsapp", "error", "Twilio no configurado")
        return False
    try:
        from twilio.rest import Client
        client = Client(sid, token)
        fr = ("whatsapp:" + desde) if not desde.startswith("whatsapp:") else desde
        to = ("whatsapp:" + numero_destino) if not numero_destino.startswith("whatsapp:") else numero_destino
        body = (("*" + asunto + "*\n") if asunto else "") + mensaje
        msg = client.messages.create(from_=fr, to=to, body=body[:1600])
        ok = msg.status not in ("failed", "undelivered")
        _registrar(tipo, asunto, "whatsapp:" + numero_destino, "enviado" if ok else "error", msg.status)
        return ok
    except Exception as e:
        _registrar(tipo, asunto, "whatsapp", "error", str(e)[:200])
        return False

def enviar_por_canales(asunto, html, texto, tipo="general"):
    """Envia por todos los canales activos segun notif_config.json."""
    cfg = _load_config()
    ch  = cfg.get("canales", {})
    res = {}
    if ch.get("email") and cfg.get("email"):
        res["email"] = enviar_email(cfg["email"], asunto, html, tipo)
    if ch.get("slack") and cfg.get("slack_webhook"):
        res["slack"] = enviar_slack(cfg["slack_webhook"], texto, asunto, tipo)
    if ch.get("whatsapp") and cfg.get("whatsapp"):
        res["whatsapp"] = enviar_whatsapp(cfg["whatsapp"], texto, asunto, tipo)
    return res

def _email_html(titulo, items, color="#3b82f6"):
    """Genera HTML para un email de alerta."""
    rows = "".join(
        f'<tr><td style="padding:8px 12px;border-bottom:1px solid #eee">{it}</td></tr>'
        for it in items
    )
    return f"""
    <div style="font-family:-apple-system,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:{color};color:#fff;padding:16px 20px;border-radius:12px 12px 0 0">
        <h2 style="margin:0;font-size:18px">Yve.01 — {titulo}</h2>
        <p style="margin:4px 0 0;font-size:12px;opacity:.8">{datetime.now().strftime("%d/%m/%Y %H:%M")}</p>
      </div>
      <div style="background:#fff;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 12px 12px;padding:4px 0">
        <table style="width:100%;font-size:14px;color:#1f2937">{rows}</table>
      </div>
      <p style="font-size:11px;color:#9ca3af;margin-top:12px;text-align:center">
        Notificación automática de Yve.01 — Dashboard Financiero
      </p>
    </div>"""

# ── Escaneo de alertas ────────────────────────────────────────────────────

def _safe_float(v):
    try:
        if v is None or str(v).strip() in ("", "nan", "None", "NO_ENCONTRADO"):
            return 0.0
        return float(str(v).replace(",", "").replace("€", "").replace("EUR", "").strip())
    except Exception:
        return 0.0

def _ultimo_excel(patron, directorio):
    hits = glob.glob(str(directorio / patron))
    if not hits:
        return None
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return hits[0]


def escanear_alertas():
    """Escanea reportes y devuelve lista de alertas pendientes."""
    alertas = []

    # ── AR: discrepancias y falta DI ──
    for patron in ["doble_imposicion_*.xlsx", "verificacion_*.xlsx"]:
        ruta = _ultimo_excel(patron, REPORTES_DIR)
        if ruta:
            try:
                df = pd.read_excel(ruta)
                if "estado" in df.columns:
                    disc = df[df["estado"].astype(str).str.upper() == "DISCREPANCIA"]
                    for _, r in disc.iterrows():
                        imp = _safe_float(r.get("discrepancia_euros", 0))
                        alertas.append({
                            "tipo": "ar_discrepancia",
                            "msg": f"Discrepancia AR: factura {r.get('numero_factura','?')} — "
                                   f"OTA {r.get('nombre_ota','?')} — {imp:,.2f} EUR",
                        })
                if "estado_di" in df.columns:
                    falta = df[df["estado_di"].astype(str).str.upper() == "FALTA_CERTIFICADO_DI"]
                    for _, r in falta.iterrows():
                        alertas.append({
                            "tipo": "ar_falta_di",
                            "msg": f"Falta certificado DI: factura {r.get('numero_factura','?')} — "
                                   f"OTA {r.get('nombre_ota','?')}",
                        })
            except Exception:
                pass
            break  # Solo el más reciente

    # ── AP: discrepancias PO ──
    for patron in ["facturas_contabilizadas_*.xlsx", "facturas_ap_*.xlsx"]:
        ruta = _ultimo_excel(patron, PROCESADAS_DIR)
        if ruta:
            try:
                df = pd.read_excel(ruta)
                est_col = None
                for c in ["estado_matching", "estado"]:
                    if c in df.columns:
                        est_col = c
                        break
                if est_col:
                    disc = df[df[est_col].astype(str).str.upper() == "DISCREPANCIA_PO"]
                    for _, r in disc.iterrows():
                        alertas.append({
                            "tipo": "ap_discrepancia",
                            "msg": f"Discrepancia PO: {r.get('nombre_proveedor','?')} — "
                                   f"factura {r.get('numero_factura','?')}",
                        })
            except Exception:
                pass
            break

    # ── DRR: Out of Balance ──
    ruta = _ultimo_excel("drr_procesado_*.xlsx", REPORTES_DIR)
    if ruta:
        try:
            df = pd.read_excel(ruta, sheet_name="Alertas", header=None)
            for _, row in df.iterrows():
                dia_val = row.iloc[0]
                if isinstance(dia_val, (int, float)) and 1 <= dia_val <= 31:
                    estado_txt = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""
                    if "OUT" in estado_txt.upper():
                        diff = _safe_float(row.iloc[4])
                        alertas.append({
                            "tipo": "drr_oob",
                            "msg": f"DRR día {int(dia_val)}: Out of Balance — diferencia {diff:,.2f} EUR",
                        })
        except Exception:
            pass

    return alertas


def enviar_pendientes(solo_check=False):
    """Escanea alertas y envía por email las que estén activas."""
    config = _load_config()
    destinatario = _get_destinatario(config)
    activas = _alertas_activas(config)
    alertas = escanear_alertas()

    # Filtrar por tipo activo
    filtradas = [a for a in alertas if activas.get(a["tipo"], True)]

    if solo_check:
        return filtradas

    if not filtradas:
        return []

    if not destinatario:
        for a in filtradas:
            _registrar(a["tipo"], a["msg"][:60], "sin destinatario", "error", "No hay email configurado")
        return filtradas

    # Agrupar por tipo
    grupos = {}
    for a in filtradas:
        grupos.setdefault(a["tipo"], []).append(a["msg"])

    TITULOS = {
        "ar_discrepancia": ("Discrepancias AR Detectadas", "#ef4444"),
        "ar_falta_di": ("Certificados DI Pendientes", "#f97316"),
        "drr_oob": ("DRR — Días Out of Balance", "#ef4444"),
        "ap_discrepancia": ("Discrepancias AP (Proveedores)", "#f97316"),
    }

    enviados = 0
    for tipo, msgs in grupos.items():
        titulo, color = TITULOS.get(tipo, (tipo, "#3b82f6"))
        hotel = config.get("hotel", {}).get("nombre", "Hotel")
        asunto = f"[Yve.01] {hotel} — {titulo}"
        html = _email_html(titulo, msgs, color)
        if enviar_email(destinatario, asunto, html, tipo):
            enviados += 1

    return filtradas


# ── CLI ───────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Yve.01 — Sistema de Notificaciones")
    print("=" * 60)

    solo_check = "--check" in sys.argv

    config = _load_config()
    dest = _get_destinatario(config)
    activas = _alertas_activas(config)

    print(f"  Destinatario: {dest or '⚠ No configurado'}")
    print(f"  Alertas activas: {sum(1 for v in activas.values() if v)}/4")

    alertas = escanear_alertas()
    filtradas = [a for a in alertas if activas.get(a["tipo"], True)]

    print(f"\n  Alertas encontradas: {len(alertas)} total, {len(filtradas)} activas")
    for a in filtradas:
        print(f"    [{a['tipo']}] {a['msg']}")

    if solo_check:
        print("\n  (modo --check, no se envían emails)")
        return

    if not filtradas:
        print("\n  ✅ Sin alertas pendientes")
        return

    if not dest:
        print("\n  ⚠ No se pueden enviar: falta email del Financial Controller")
        return

    env = _load_env()
    if not env.get("SMTP_USER"):
        print("\n  ⚠ SMTP no configurado — añade SMTP_USER y SMTP_PASSWORD al .env")
        return

    print(f"\n  Enviando {len(filtradas)} alerta(s) a {dest}...")
    enviar_pendientes()
    print("  ✅ Proceso completado. Revisa el historial en datos-referencia/notificaciones_historial.json")

    print("=" * 60)


if __name__ == "__main__":
    main()
