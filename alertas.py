# -*- coding: utf-8 -*-
"""alertas.py — Yve.01 · aviso por email cuando la app devuelve errores (b82).

Dos cosas distintas, y solo una la puede hacer la propia app:

  1. ERRORES (esto): cualquier respuesta 5xx o excepcion sin capturar se apunta
     y se manda un email a YVE_ALERTAS_EMAIL con la ruta, el usuario, el hotel y
     el error. Para no inundar el buzon hay un margen (YVE_ALERTAS_MIN, 15 min
     por defecto): el primer error avisa al momento; los siguientes dentro del
     margen se acumulan y salen en UN email resumen cuando pasa el margen.
     Ademas, al arrancar manda un "Yve.01 ha arrancado" (deploy o reinicio):
     si te llega sin haber desplegado, la app se cayo y Render la levanto.
     YVE_ALERTAS_ARRANQUE=off lo quita.

  2. CAIDAS: si la app no responde, no puede avisar de nada. Eso lo hace un
     vigilante EXTERNO (UptimeRobot, Better Stack…) que pide /health cada pocos
     minutos y avisa por email si no contesta 200. La configuracion esta en la
     documentacion para Jordi (paso a paso), no en el codigo.

El envio va por Brevo (notificaciones.enviar_email) en un hilo aparte para no
retrasar la respuesta al usuario. Sin YVE_ALERTAS_EMAIL no se manda nada (se
sigue apuntando en memoria: /admin/api/alertas lo enseña).
"""
import os
import threading
import traceback
from datetime import datetime, timezone

_lock = threading.Lock()
_estado = {"pendientes": [], "ultimo_envio": None, "enviados": 0, "total": 0, "ultimos": []}
_MAX_ULTIMOS = 20
_SINCRONO = False          # los tests lo ponen a True para no depender de hilos


def destinatario(env=None):
    env = env if env is not None else os.environ
    return (env.get("YVE_ALERTAS_EMAIL") or "").strip()


def margen_min(env=None):
    env = env if env is not None else os.environ
    try:
        return max(0, int(env.get("YVE_ALERTAS_MIN", "15")))
    except ValueError:
        return 15


def _ahora():
    return datetime.now(timezone.utc)


def _enviar(asunto, cuerpo_html):
    dest = destinatario()
    if not dest:
        return False

    def _go():
        try:
            from notificaciones import enviar_email
            enviar_email(dest, asunto, cuerpo_html, tipo="alerta")
        except Exception as e:
            print(f"[alertas] no se pudo enviar: {e}")
    if _SINCRONO:
        _go()
        return True
    threading.Thread(target=_go, daemon=True, name="alerta_email").start()
    return True


def _html_error(e):
    return (f"<li><b>{e['fecha'][11:19]} UTC</b> · {e['metodo']} <code>{e['ruta']}</code> → {e['status']}"
            f" · usuario <i>{e.get('usuario') or '-'}</i> · hotel <i>{e.get('hotel') or '-'}</i>"
            f"<br><code style='white-space:pre-wrap'>{(e.get('error') or '')[:600]}</code></li>")


def registrar(ruta, metodo="GET", status=500, error="", usuario="", hotel="", traza=""):
    """Apunta un error y decide si toca email ahora. Devuelve 'enviado' | 'acumulado' | 'sin_email'."""
    ev = {"fecha": _ahora().isoformat(timespec="seconds"), "ruta": str(ruta)[:200], "metodo": metodo,
          "status": int(status), "error": str(error)[:1500], "usuario": str(usuario or "")[:60],
          "hotel": str(hotel or "")[:60], "traza": str(traza or "")[-2500:]}
    with _lock:
        _estado["total"] += 1
        _estado["ultimos"] = ([ev] + _estado["ultimos"])[:_MAX_ULTIMOS]
        if not destinatario():
            return "sin_email"
        ult = _estado["ultimo_envio"]
        margen = margen_min() * 60
        if ult is None or (_ahora() - ult).total_seconds() >= margen:
            _estado["ultimo_envio"] = _ahora()
            _estado["enviados"] += 1
            cuerpo = (f"<p>Yve.01 ha devuelto un error.</p><ul>{_html_error(ev)}</ul>"
                      + (f"<pre style='font-size:12px'>{ev['traza'][-1800:]}</pre>" if ev["traza"] else "")
                      + f"<p>Si hay mas errores en los proximos {margen_min()} minutos, llegaran juntos en un solo email.</p>")
            _enviar(f"⚠ Yve.01: error {ev['status']} en {ev['metodo']} {ev['ruta'][:60]}", cuerpo)
            return "enviado"
        _estado["pendientes"].append(ev)
        return "acumulado"


def vaciar_pendientes(forzar=False):
    """Manda el resumen de lo acumulado si ya paso el margen (lo llama cada peticion, es barato)."""
    with _lock:
        if not _estado["pendientes"]:
            return False
        ult = _estado["ultimo_envio"]
        if not forzar and ult is not None and (_ahora() - ult).total_seconds() < margen_min() * 60:
            return False
        pend = _estado["pendientes"]; _estado["pendientes"] = []
        _estado["ultimo_envio"] = _ahora(); _estado["enviados"] += 1
    por_ruta = {}
    for e in pend:
        por_ruta[e["ruta"]] = por_ruta.get(e["ruta"], 0) + 1
    cuerpo = (f"<p>Yve.01 ha devuelto <b>{len(pend)} errores</b> mas en los ultimos {margen_min()} minutos.</p>"
              f"<p>Por ruta: " + " · ".join(f"<code>{r}</code> ×{n}" for r, n in sorted(por_ruta.items(), key=lambda x: -x[1])) + "</p>"
              f"<ul>{''.join(_html_error(e) for e in pend[-15:])}</ul>")
    _enviar(f"⚠ Yve.01: {len(pend)} errores mas ({', '.join(list(por_ruta)[:2])})", cuerpo)
    return True


def avisar_arranque(env=None):
    env = env if env is not None else os.environ
    if (env.get("YVE_ALERTAS_ARRANQUE") or "on").strip().lower() == "off" or not destinatario(env):
        return False
    try:
        import subprocess
        rev = subprocess.run(["git", "rev-parse", "--short", "HEAD"], capture_output=True, text=True, timeout=5).stdout.strip()
    except Exception:
        rev = ""
    rev = rev or (env.get("RENDER_GIT_COMMIT") or "")[:7]
    cuerpo = (f"<p>Yve.01 ha arrancado a las {_ahora().strftime('%Y-%m-%d %H:%M:%S')} UTC"
              + (f" (commit <code>{rev}</code>)" if rev else "") + ".</p>"
              "<p>Si no acabas de desplegar, es que la app se habia caido y Render la ha vuelto a levantar: mira los logs de Render.</p>")
    return _enviar("ℹ Yve.01 ha arrancado", cuerpo)


def resumen():
    with _lock:
        return {"configurado": bool(destinatario()), "email": destinatario(), "margen_min": margen_min(),
                "total": _estado["total"], "enviados": _estado["enviados"], "pendientes": len(_estado["pendientes"]),
                "ultimo_envio": _estado["ultimo_envio"].isoformat(timespec="seconds") if _estado["ultimo_envio"] else None,
                "ultimos": [{k: v for k, v in e.items() if k != "traza"} for e in _estado["ultimos"]]}


def _reset():
    """Solo para tests."""
    with _lock:
        _estado.update({"pendientes": [], "ultimo_envio": None, "enviados": 0, "total": 0, "ultimos": []})


def instalar(app):
    """Engancha los avisos a la app Flask: excepciones sin capturar y respuestas 5xx."""
    from flask import request, g, got_request_exception

    def _quien():
        try:
            from flask_login import current_user
            u = getattr(current_user, "id", "") if getattr(current_user, "is_authenticated", False) else ""
        except Exception:
            u = ""
        try:
            import censo_hoteles
            h = censo_hoteles.activo() or ""
        except Exception:
            h = ""
        return u, h

    def _al_excepcion(sender, exception, **extra):
        try:
            g._alerta_apuntada = True
            u, h = _quien()
            registrar(request.path, request.method, 500, f"{type(exception).__name__}: {exception}", u, h,
                      "".join(traceback.format_exception(type(exception), exception, exception.__traceback__)))
        except Exception:
            pass

    got_request_exception.connect(_al_excepcion, app, weak=False)

    @app.after_request
    def _tras_respuesta(resp):
        try:
            if resp.status_code >= 500 and not getattr(g, "_alerta_apuntada", False):
                u, h = _quien()
                cuerpo = ""
                try:
                    if resp.is_json:
                        cuerpo = str(resp.get_json(silent=True) or "")[:500]
                    elif not resp.direct_passthrough and (resp.mimetype or "").startswith("text/"):
                        cuerpo = resp.get_data(as_text=True)[:300]
                except Exception:
                    cuerpo = ""
                registrar(request.path, request.method, resp.status_code, cuerpo, u, h)
            vaciar_pendientes()
        except Exception:
            pass
        return resp
    return True
