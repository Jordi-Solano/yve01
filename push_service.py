"""
push_service.py — Yve.01
Notificaciones push web reales (Web Push API + VAPID), segmentadas por rol y tenant.

- La clave PÚBLICA VAPID no es secreta (el navegador la recibe). Va en el código
  y puede sobreescribirse con la variable de entorno VAPID_PUBLIC_KEY.
- La clave PRIVADA VAPID es SECRETA. Se lee SOLO de la variable de entorno
  VAPID_PRIVATE_KEY (nunca se guarda en git). Sin ella, el push se desactiva
  de forma elegante (el resto de la app sigue funcionando igual).

Cada suscripción guarda: endpoint, keys, rol, tenant, username. El envío puede
filtrarse por tenant (nunca cruza datos entre hoteles/cadenas) y por rol
(cada alerta va solo a quien le corresponde). El rol 'admin' recibe todo.

Suscripciones en datos-referencia/push_subscriptions.json.
"""
import os, json, time
from pathlib import Path

BASE_DIR       = Path(__file__).parent
REFERENCIA_DIR = BASE_DIR / "datos-referencia"
SUBS_PATH      = REFERENCIA_DIR / "push_subscriptions.json"

# Clave pública VAPID por defecto (generada para Yve.01). Sobreescribible por env.
VAPID_PUBLIC_KEY = os.environ.get("VAPID_PUBLIC_KEY", "BA4APGeKyarjFuTQX95jO03et4SuK3LSyG7JaCQnlBFlkaHWVISoR8kiRgu3m2E8_Xj3Al6Zqgp4tkrllC94ZUU")


def _private_key() -> str:
    return (os.environ.get("VAPID_PRIVATE_KEY") or "").strip()


def push_enabled() -> bool:
    """True solo si hay clave privada configurada (en Render)."""
    return bool(_private_key())


def _contact() -> str:
    return os.environ.get("VAPID_CONTACT_EMAIL", "barnar749@gmail.com")


# ── Almacén de suscripciones ────────────────────────────────────────────────

def _load_subs() -> list:
    if SUBS_PATH.exists():
        try:
            return json.loads(SUBS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _save_subs(subs: list) -> None:
    REFERENCIA_DIR.mkdir(exist_ok=True)
    SUBS_PATH.write_text(json.dumps(subs, ensure_ascii=False, indent=2), encoding="utf-8")


def add_subscription(sub: dict, rol=None, tenant=None, username=None) -> bool:
    """Guarda (o refresca) una suscripción push con su rol y tenant. Dedup por endpoint."""
    if not sub or not sub.get("endpoint"):
        return False
    subs = _load_subs()
    ep = sub["endpoint"]
    subs = [s for s in subs if s.get("endpoint") != ep]
    sub = {"endpoint": sub.get("endpoint"), "keys": sub.get("keys", {}),
           "expirationTime": sub.get("expirationTime")}
    sub["rol"]      = rol or "income_auditor"
    sub["tenant"]   = tenant or "default"
    sub["username"] = username or ""
    sub["_ts"]      = int(time.time())
    subs.append(sub)
    _save_subs(subs)
    return True


def remove_subscription(endpoint: str) -> bool:
    if not endpoint:
        return False
    subs = _load_subs()
    nuevas = [s for s in subs if s.get("endpoint") != endpoint]
    _save_subs(nuevas)
    return len(nuevas) != len(subs)


def count_subscriptions() -> int:
    return len(_load_subs())


# ── Envío ───────────────────────────────────────────────────────────────────

def send_push(title, body, url="/app", tag="yve-alert", icon=None,
              require_interaction=True, renotify=False,
              roles=None, tenant=None) -> dict:
    """
    Envía una notificación push a las suscripciones que coincidan con:
      · tenant (si se indica): solo dispositivos de ese hotel/cadena.
      · roles (si se indica): solo esos roles (el rol 'admin' siempre recibe).
    Limpia suscripciones caducadas (404/410). Nunca lanza excepción.
    """
    if not push_enabled():
        return {"ok": False, "reason": "no_vapid_private_key", "sent": 0, "total": 0}
    try:
        from pywebpush import webpush, WebPushException
    except Exception as e:
        return {"ok": False, "reason": "pywebpush_no_disponible: " + str(e)[:80],
                "sent": 0, "total": 0}

    subs = _load_subs()
    allowed = (set(roles) | {"admin"}) if roles else None
    objetivo = []
    for s in subs:
        if tenant is not None and (s.get("tenant") or "default") != tenant:
            continue
        if allowed is not None and (s.get("rol") or "") not in allowed:
            continue
        objetivo.append(s)

    if not objetivo:
        return {"ok": True, "sent": 0, "total": 0, "reason": "sin_destinatarios"}

    payload = json.dumps({
        "title": title, "body": body, "url": url, "tag": tag,
        "icon": icon or "/static/icons/yve-logo-192.png",
        "requireInteraction": bool(require_interaction), "renotify": bool(renotify),
    })
    priv = _private_key()
    sub_claim = "mailto:" + _contact()

    sent, stale = 0, []
    for s in objetivo:
        try:
            webpush(subscription_info={"endpoint": s["endpoint"], "keys": s.get("keys", {})},
                    data=payload, vapid_private_key=priv,
                    vapid_claims={"sub": sub_claim}, ttl=86400,
                    headers={"Urgency": "high"})
            sent += 1
        except WebPushException as e:
            code = getattr(getattr(e, "response", None), "status_code", None)
            if code in (404, 410):
                stale.append(s.get("endpoint"))
        except Exception:
            pass

    if stale:
        restantes = [s for s in subs if s.get("endpoint") not in stale]
        _save_subs(restantes)

    return {"ok": True, "sent": sent, "total": len(objetivo), "removed": len(stale)}


if __name__ == "__main__":
    print("VAPID public key:", VAPID_PUBLIC_KEY[:24], "...")
    print("Push habilitado (hay clave privada):", push_enabled())
    print("Suscripciones registradas:", count_subscriptions())
