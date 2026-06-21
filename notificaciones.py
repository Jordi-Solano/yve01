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
CONFIG_PATH       = REFERENCIA_DIR / "hotel_config.json"
NOTIF_CONFIG_PATH = REFERENCIA_DIR / "notif_config.json"
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

def _load_notif_config():
    """Carga notif_config.json (canales, email destino, alertas)."""
    defaults = {
        "canales": {"email": True, "whatsapp": False, "slack": False, "push": True},
        "email": "", "whatsapp": "", "slack_webhook": "",
        "alertas": {"ar_discrepancia": True, "ar_falta_di": True,
                    "ap_discrepancia": True, "drr_oob": True,
                    "banco_sin_conciliar": True, "factura_pendiente_firma": False},
        "frecuencia": "inmediata",
    }
    if NOTIF_CONFIG_PATH.exists():
        try:
            saved = json.loads(NOTIF_CONFIG_PATH.read_text(encoding="utf-8"))
            defaults.update(saved)
        except Exception:
            pass
    return defaults

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

def _enviar_via_resend(destinatario, asunto, cuerpo_html):
    """Envía email via Resend HTTP API. Funciona en Render free tier (no bloquea SMTP)."""
    import urllib.request, urllib.error, json as _json
    api_key = os.environ.get("RESEND_API_KEY", "")
    if not api_key:
        return False, "RESEND_API_KEY no configurado"
    env = _load_env()
    nombre_hotel = env.get("HOTEL_NOMBRE", "Hotel")
    payload = _json.dumps({
        "from": "Yve.01 <onboarding@resend.dev>",
        "to": [destinatario],
        "subject": asunto,
        "html": cuerpo_html,
    }).encode()
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=payload,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            body = resp.read().decode()
            return True, f"OK id={body[:60]}"
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"[RESEND ERROR] {e.code}: {body}")  # visible en logs de Render
        _registrar("resend_error", "Resend HTTP error", destinatario, "error", f"{e.code}: {body[:300]}")
        return False, f"Resend {e.code}: {body[:200]}"
    except Exception as e:
        _registrar("resend_error", "Resend connection error", destinatario, "error", str(e)[:120])
        return False, str(e)[:120]


def enviar_email(destinatario, asunto, cuerpo_html, tipo="general"):
    """Envía un email. Usa Resend si hay RESEND_API_KEY, sino SMTP."""
    env = _load_env()

    # Prioridad 1: Resend (HTTP API — funciona en Render free tier)
    if os.environ.get("RESEND_API_KEY"):
        ok, msg = _enviar_via_resend(destinatario, asunto, cuerpo_html)
        _registrar(tipo, asunto, destinatario, "enviado" if ok else "error", msg)
        return ok

    # Prioridad 2: SMTP directo (no funciona en Render free tier)
    smtp_server = env.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port   = int(env.get("SMTP_PORT", "587"))
    smtp_user   = env.get("SMTP_USER", "")
    smtp_pass   = env.get("SMTP_PASSWORD", "")

    if not smtp_user or not smtp_pass:
        _registrar(tipo, asunto, destinatario, "error", "SMTP no configurado")
        return False

    msg = MIMEMultipart("alternative")
    nombre_hotel = env.get("HOTEL_NOMBRE", "Hotel")
    msg["From"]    = f"Yve.01 · {nombre_hotel} <{smtp_user}>"
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

def test_smtp():
    """Verifica la configuración SMTP enviando un email de prueba."""
    env = _load_env()
    smtp_server = env.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port   = int(env.get("SMTP_PORT", "587"))
    smtp_user   = env.get("SMTP_USER", "")
    smtp_pass   = env.get("SMTP_PASSWORD", "")
    dest        = env.get("NOTIF_EMAIL", smtp_user)

    if not smtp_user or not smtp_pass:
        return {"ok": False, "error": "SMTP_USER o SMTP_PASSWORD no configurados en variables de entorno"}

    try:
        import smtplib as _s
        with _s.SMTP(smtp_server, smtp_port, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_user, smtp_pass)
        return {"ok": True, "message": f"Conexión SMTP correcta con {smtp_user} en {smtp_server}:{smtp_port}"}
    except Exception as e:
        err = str(e)
        hint = ""
        if "534" in err or "Username and Password not accepted" in err:
            hint = " → Usa una Contraseña de Aplicación de Gmail (no tu contraseña normal)"
        elif "535" in err:
            hint = " → Credenciales incorrectas. Verifica SMTP_USER y SMTP_PASSWORD"
        elif "Connection refused" in err or "timed out" in err:
            hint = f" → No se puede conectar a {smtp_server}:{smtp_port}"
        elif "STARTTLS" in err:
            hint = " → El servidor no soporta STARTTLS. Prueba puerto 465 con SSL"
        return {"ok": False, "error": err[:200] + hint}

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

def enviar_telegram(mensaje: str) -> bool:
    """Envía un mensaje via Telegram Bot API."""
    env = _load_env()
    token  = env.get("TELEGRAM_BOT_TOKEN", "")
    chat_id = env.get("TELEGRAM_CHAT_ID", "")
    if not token or not chat_id:
        return False
    try:
        import urllib.request, json as _json
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        data = _json.dumps({"chat_id": chat_id, "text": f"🔔 Yve.01\n{mensaje}", "parse_mode": "HTML"}).encode()
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=10)
        return True
    except Exception as e:
        print(f"[Telegram] Error: {e}")
        return False

def enviar_por_canales(asunto, html, texto, tipo="general"):
    """Envia por todos los canales activos segun notif_config.json."""
    cfg = _load_notif_config()
    ch  = cfg.get("canales", {})
    res = {}
    if ch.get("email") and cfg.get("email"):
        res["email"] = enviar_email(cfg["email"], asunto, html, tipo)
    if ch.get("slack") and cfg.get("slack_webhook"):
        res["slack"] = enviar_slack(cfg["slack_webhook"], texto, asunto, tipo)
    if ch.get("whatsapp") and cfg.get("whatsapp"):
        res["whatsapp"] = enviar_whatsapp(cfg["whatsapp"], texto, asunto, tipo)
    return res

def _email_html(titulo, items, color="#3b82f6", footer_note=None):
    """Genera HTML profesional para un email de alerta Yve.01."""
    from datetime import datetime
    date_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    color_light = "rgba(59,130,246,.08)" if color == "#3b82f6" else "rgba(239,68,68,.06)"
    rows = "".join(
        f'<tr><td style="padding:10px 16px;border-bottom:1px solid #1e293b;font-size:14px;color:#cbd5e1;line-height:1.5">'
        f'<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:{color};margin-right:10px;vertical-align:middle"></span>'
        f'{it}</td></tr>'
        for it in items
    )
    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — {titulo}</title></head>
<body style="margin:0;padding:0;background:#0f172a;font-family:-apple-system,BlinkMacSystemFont,'Inter',sans-serif">
  <div style="max-width:600px;margin:0 auto;padding:24px 16px">
    <!-- Header -->
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:24px;padding-bottom:20px;border-bottom:1px solid #1e293b">
      <div style="width:10px;height:10px;border-radius:50%;background:#3b82f6;box-shadow:0 0 8px #3b82f6;flex-shrink:0"></div>
      <span style="font-size:20px;font-weight:800;color:#f1f5f9;letter-spacing:-.5px">Yve<span style="color:#60a5fa">.01</span></span>
      <span style="font-size:11px;background:rgba(59,130,246,.15);color:#60a5fa;border:1px solid rgba(59,130,246,.25);border-radius:12px;padding:2px 8px;margin-left:4px">Dashboard</span>
    </div>
    <!-- Alert Card -->
    <div style="background:#1e293b;border-radius:16px;overflow:hidden;border:1px solid #334155;margin-bottom:20px">
      <!-- Alert header -->
      <div style="background:{color};padding:18px 20px">
        <div style="font-size:18px;font-weight:700;color:#fff;margin-bottom:4px">{titulo}</div>
        <div style="font-size:12px;color:rgba(255,255,255,.7)">{date_str} &nbsp;·&nbsp; Yve.01 Finance Dashboard</div>
      </div>
      <!-- Alert items -->
      <table style="width:100%;border-collapse:collapse">
        {rows}
      </table>
    </div>
    <!-- CTA -->
    <div style="text-align:center;margin-bottom:24px">
      <a href="https://yve01.onrender.com" style="display:inline-block;background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;padding:13px 32px;border-radius:12px;font-size:14px;font-weight:700;text-decoration:none;box-shadow:0 4px 20px rgba(59,130,246,.4)">
        Ver en el dashboard →
      </a>
    </div>
    <!-- Footer -->
    <div style="text-align:center;font-size:11px;color:#475569;padding-top:16px;border-top:1px solid #1e293b">
      <strong style="color:#64748b">Yve.01</strong> &nbsp;·&nbsp; Sistema financiero AI para hoteles &nbsp;·&nbsp; Barcelona, España<br>
      <span style="margin-top:6px;display:block">Este email fue generado automáticamente. No respondas a este mensaje.</span>
    </div>
  </div>
</body>
</html>"""


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
