"""panel_admin.py - Yve.01 Admin Panel"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Blueprint, request, jsonify
from flask_login import login_required, current_user
from auth import (init_login, inicializar_usuarios, listar_usuarios,
                  crear_usuario, cambiar_password, toggle_activo, ROLES_VALIDOS)
from pathlib import Path
from version_estaticos import SELLO as SELLO_ESTATICOS

bp = Blueprint("admin", __name__, url_prefix="/admin")
BASE_DIR = Path(__file__).parent


def _admin_required(f):
    from functools import wraps
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if current_user.rol != "admin":
            return jsonify({"error": "Acceso denegado"}), 403
        return f(*args, **kwargs)
    return decorated


@bp.route("/")
@login_required
def index():
    if current_user.rol != "admin":
        return "<h3>Acceso denegado</h3>", 403
    return HTML.replace("__ASSETS_V__", SELLO_ESTATICOS)


@bp.route("/api/usuarios")
@_admin_required
def api_usuarios():
    return jsonify(listar_usuarios())


@bp.route("/api/crear_usuario", methods=["POST"])
@_admin_required
def api_crear():
    d = request.get_json(force=True) or {}
    result = crear_usuario(
        d.get("username", "").strip(),
        d.get("password", ""),
        d.get("nombre", ""),
        d.get("email", ""),
        d.get("rol", "financial_controller"),
        tenant=d.get("tenant", "default").strip() or "default",
    )
    return jsonify({"ok": result is True, "error": result if result is not True else None})


@bp.route("/api/cambiar_password", methods=["POST"])
@_admin_required
def api_cambiar_pass():
    d = request.get_json(force=True) or {}
    ok = cambiar_password(d.get("username", ""), d.get("nueva_password", ""))
    return jsonify({"ok": ok})


@bp.route("/api/toggle_usuario", methods=["POST"])
@_admin_required
def api_toggle():
    d = request.get_json(force=True) or {}
    ok = toggle_activo(d.get("username", ""))
    return jsonify({"ok": ok})


@bp.route("/api/stats")
@_admin_required
def api_admin_stats():
    import glob, time
    from datetime import datetime

    try:
        users = json.loads((BASE_DIR / "datos-referencia" / "usuarios.json").read_text())
        n_users  = len(users)
        n_active = sum(1 for u in users if u.get("activo", True))
    except Exception:
        n_users = n_active = 0

    try:
        hotels = json.loads((__import__("pathlib").Path(__import__("tenant_dirs").datos_dir()) / "hoteles.json").read_text())
        n_hotels = len(hotels)
    except Exception:
        n_hotels = 0

    n_reports = len(glob.glob(str(BASE_DIR / "reportes" / "*.xlsx")))

    uptime = "-"
    try:
        with open("/proc/uptime") as f:
            secs = float(f.read().split()[0])
        h = int(secs // 3600); m = int((secs % 3600) // 60)
        uptime = str(h) + "h " + str(m) + "m"
    except Exception:
        pass

    return jsonify({
        "usuarios": n_users,
        "usuarios_activos": n_active,
        "hoteles": n_hotels,
        "reportes_generados": n_reports,
        "uptime": uptime,
        "version": "Yve.01 Beta",
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
    })


@bp.route("/api/audit")
@_admin_required
def api_audit():
    """Returns recent audit log entries."""
    ruta = BASE_DIR / "datos-referencia" / "audit_log.json"
    if not ruta.exists():
        return jsonify({"entries": []})
    try:
        entries = json.loads(ruta.read_text())
        return jsonify({"entries": entries[-50:]})  # last 50
    except Exception:
        return jsonify({"entries": []})

@bp.route("/api/hoteles")
@_admin_required
def api_hoteles():
    try:
        hotels = json.loads((__import__("pathlib").Path(__import__("tenant_dirs").datos_dir()) / "hoteles.json").read_text())
        return jsonify({"ok": True, "hoteles": hotels})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Copias de seguridad (b81): fuera de Render, diarias, con restauracion ──
@bp.route("/api/copias")
@_admin_required
def api_copias():
    import copia_seguridad as CS
    return jsonify({"ok": True, **CS.resumen(), "copias": CS.listar()})


@bp.route("/api/copias/ahora", methods=["POST"])
@_admin_required
def api_copias_ahora():
    import copia_seguridad as CS
    if not CS.configurado():
        return jsonify({"ok": False, "error": "Sin destino configurado: pon YVE_BACKUP_S3_BUCKET (+ endpoint, key, secret) o YVE_BACKUP_DIR en Render."}), 400
    r = CS.hacer_copia(motivo="manual")
    return jsonify(r), (200 if r.get("ok") else 500)


@bp.route("/api/copias/restaurar", methods=["POST"])
@_admin_required
def api_copias_restaurar():
    """Deja los datos como en la copia elegida. Pide escribir RESTAURAR: no es un boton mas."""
    import copia_seguridad as CS
    d = request.get_json(silent=True) or {}
    nombre = str(d.get("nombre") or "").strip()
    if not nombre or "/" in nombre or ".." in nombre or not nombre.endswith(".zip"):
        return jsonify({"ok": False, "error": "nombre de copia no valido"}), 400
    if str(d.get("confirmar") or "").strip().upper() != "RESTAURAR":
        return jsonify({"ok": False, "error": "Escribe RESTAURAR para confirmar: se sustituyen TODOS los datos por los de la copia (antes se guarda una copia de lo actual)."}), 400
    r = CS.restaurar(nombre, tipo=(d.get("destino") or None))
    try:
        from dashboard import _audit
        _audit("RESTAURAR_COPIA", f"{nombre} ok={r.get('ok')} ficheros={r.get('ficheros')}")
    except Exception:
        pass
    return jsonify(r), (200 if r.get("ok") else 500)


@bp.route("/api/copias/descargar")
@_admin_required
def api_copias_descargar():
    """Baja una copia al ordenador del admin (para guardarla aparte o mirarla)."""
    import copia_seguridad as CS
    import tempfile
    from flask import send_file, after_this_request
    nombre = str(request.args.get("nombre") or "").strip()
    if not nombre or "/" in nombre or ".." in nombre or not nombre.endswith(".zip"):
        return jsonify({"ok": False, "error": "nombre de copia no valido"}), 400
    tmpdir = tempfile.mkdtemp(prefix="yve_desc_")
    ruta = os.path.join(tmpdir, nombre)
    try:
        CS.descargar_copia(nombre, ruta, tipo=(request.args.get("destino") or None))
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 404

    @after_this_request
    def _limpiar(resp):
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)
        return resp
    return send_file(ruta, as_attachment=True, download_name=nombre, mimetype="application/zip")


@bp.route("/api/alertas")
@_admin_required
def api_alertas():
    import alertas as AL
    return jsonify({"ok": True, **AL.resumen()})


@bp.route("/api/alertas/probar", methods=["POST"])
@_admin_required
def api_alertas_probar():
    """Manda un email de prueba al YVE_ALERTAS_EMAIL para comprobar que llega."""
    import alertas as AL
    if not AL.destinatario():
        return jsonify({"ok": False, "error": "Sin YVE_ALERTAS_EMAIL en Render: no hay a quien avisar."}), 400
    try:
        from notificaciones import enviar_email
        ok = bool(enviar_email(AL.destinatario(), "✅ Yve.01: prueba de alertas",
                               "<p>Si lees esto, los avisos de error de Yve.01 llegan a este buzon.</p>", tipo="alerta"))
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500
    return jsonify({"ok": ok, "email": AL.destinatario()}), (200 if ok else 500)


@bp.route("/api/hoteles/eliminar", methods=["POST"])
@_admin_required
def api_eliminar_hotel():
    data = request.get_json(silent=True) or {}
    hotel_id = data.get("id")
    if not hotel_id:
        return jsonify({"ok": False, "error": "ID requerido"}), 400
    try:
        path = __import__("pathlib").Path(__import__("tenant_dirs").datos_dir()) / "hoteles.json"
        hotels = json.loads(path.read_text())
        hotels = [h for h in hotels if h.get("id") != hotel_id]
        path.write_text(json.dumps(hotels, indent=2, ensure_ascii=False))
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500




HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<title>Admin - Yve.01</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--red:#ef4444}
*{box-sizing:border-box;margin:0;padding:0}
html{scrollbar-gutter:stable}
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;padding:calc(28px + env(safe-area-inset-top)) 28px 40px;-webkit-font-smoothing:antialiased;position:relative;min-height:100vh}
body::before{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;background:radial-gradient(900px 500px at 90% -5%,rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10),transparent 60%),radial-gradient(700px 400px at -5% 105%,rgba(139,92,246,.08),transparent 55%)}
body>*{position:relative;z-index:1}
h1{font-size:20px;font-weight:800;margin-bottom:6px}
.sub{font-size:13px;color:var(--mut);margin-bottom:24px}
.sg{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;margin-bottom:24px}
/* Misma burbuja que el dashboard (.sc): fondo plano, mismo borde y radio.
   Antes tenia degradado y una barra azul a la izquierda y se notaba que era
   otra pantalla. */
.kc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px;position:relative;overflow:hidden;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s}
.kc:hover{border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4);transform:translateY(-1px)}

.kl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px}
.kv{font-size:28px;font-weight:800;color:var(--acc2);letter-spacing:-1px;line-height:1}
.card{background:linear-gradient(170deg,rgba(30,41,59,.92),rgba(20,28,45,.82));border:1px solid var(--s2);border-radius:16px;padding:22px;margin-bottom:18px;box-shadow:0 4px 20px rgba(0,0,0,.18)}
.ct{font-size:15px;font-weight:700;margin-bottom:16px;display:flex;align-items:center;gap:8px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:768px){.g2{grid-template-columns:1fr}body{padding:calc(16px + env(safe-area-inset-top)) 14px 32px}.card{padding:16px;overflow-x:auto;min-width:0}.g2>div{min-width:0}.sg{grid-template-columns:repeat(2,1fr)}[style*="grid-template-columns:1fr 1fr"]{grid-template-columns:1fr!important}.ut{font-size:12px}h1{font-size:18px}input,select,textarea{font-size:16px!important}}
.admin-top{display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap}
.admin-back{display:inline-flex;align-items:center;gap:6px;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.12);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.35);color:var(--acc2);padding:9px 15px;border-radius:10px;font-size:13px;font-weight:600;text-decoration:none;-webkit-tap-highlight-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2)}
.admin-back:active,.admin-back:hover{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.22);border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.6);color:var(--acc2)}
.brand-dot{width:11px;height:11px;border-radius:50%;background:var(--acc);box-shadow:0 0 14px var(--acc);flex-shrink:0}
label{display:block;font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px;margin-top:12px;font-weight:600}
label:first-child{margin-top:0}
input,select,textarea{width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:10px;padding:11px 14px;font-size:16px;font-family:inherit;outline:none;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s}
input:focus,select:focus,textarea:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)}
.btn{padding:9px 18px;border:none;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit}
.bp{background:linear-gradient(135deg,var(--acc),var(--acc-dark,#2563eb));color:#fff;box-shadow:0 4px 14px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.3)}
.bp:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.45)}
.bd{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);color:#ef4444}
.bsm{padding:5px 12px;font-size:11px;border-radius:7px}
.ut{width:100%;border-collapse:collapse;font-size:13px}
.ut th{padding:9px 12px;font-size:10px;color:#94a3b8;text-transform:uppercase;border-bottom:1px solid #334155;text-align:left}
.ut td{padding:9px 12px;border-bottom:1px solid rgba(51,65,85,.4);vertical-align:middle}
.msg{font-size:12px;text-align:center;margin-top:10px;min-height:16px}
.ok{background:rgba(34,197,94,.1);color:#22c55e;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;display:inline-block}
.off{background:rgba(100,116,139,.1);color:#64748b;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;display:inline-block}
</style>
<script src="/static/yve-tema.js?v=__ASSETS_V__"></script>
</head>
<body>
<div class="admin-top">
  <a class="admin-back" href="/app">← Volver a Yve</a>
  <span class="brand-dot"></span>
  <h1 style="margin:0">Admin · <span style="color:var(--acc2)">Yve.01</span></h1>
</div>
<div class="sg">
  <div class="kc"><div class="kl">Usuarios</div><div class="kv" id="su">-</div></div>
  <div class="kc"><div class="kl">Activos</div><div class="kv" id="sa">-</div></div>
  <div class="kc"><div class="kl">Hoteles</div><div class="kv" id="sh">-</div></div>
  <div class="kc"><div class="kl">Reportes</div><div class="kv" id="sr">-</div></div>
  <div class="kc"><div class="kl">Uptime</div><div class="kv" style="font-size:15px" id="sup">-</div></div>
  <div class="kc"><div class="kl">Version</div><div class="kv" style="font-size:13px" id="sv">-</div></div>
</div>
<div class="g2">
  <div>
    <div class="card">
      <div class="ct">Usuarios</div>
      <table class="ut"><thead><tr><th>User</th><th>Nombre</th><th>Rol</th><th>Estado</th><th></th></tr></thead>
      <tbody id="utb"></tbody></table>
    </div>
    <div class="card">
      <div class="ct">Crear usuario</div>
      <label>Username</label><input id="nu">
      <label>Password</label><input id="np" type="password">
      <label>Nombre</label><input id="nn">
      <label>Email</label><input id="ne" type="email">
      <label>Rol</label>
      <select id="nr">
        <option value="financial_controller">Financial Controller</option>
        <option value="income_auditor">Income Auditor</option>
        <option value="fb_manager">F&amp;B Manager</option>
        <option value="jefe_otras">Jefe Servicios</option>
        <option value="admin">Admin</option>
      </select>
      <button class="btn bp" style="margin-top:16px;width:100%" onclick="crearU()">+ Crear</button>
      <div class="msg" id="mu"></div>
    </div>
  </div>
  <div>
    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
        <div class="ct" style="margin:0">Hoteles</div>
        <button class="btn bsm" onclick="document.getElementById('nuevo-hotel-form').style.display=document.getElementById('nuevo-hotel-form').style.display==='none'?'block':'none'" style="font-size:11px;padding:5px 12px">+ Añadir hotel</button>
      </div>
      <div id="nuevo-hotel-form" style="display:none;background:#0f172a;border-radius:10px;padding:14px;margin-bottom:14px">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px">
          <input id="hn-nombre" placeholder="Nombre del hotel" style="background:#1e293b;border:1px solid #334155;color:#f1f5f9;padding:8px;border-radius:7px;font-size:12px">
          <input id="hn-ciudad" placeholder="Ciudad" style="background:#1e293b;border:1px solid #334155;color:#f1f5f9;padding:8px;border-radius:7px;font-size:12px">
          <input id="hn-hab" placeholder="Habitaciones" type="number" value="100" style="background:#1e293b;border:1px solid #334155;color:#f1f5f9;padding:8px;border-radius:7px;font-size:12px">
          <input id="hn-grupo" list="hn-grupos" placeholder="Grupo / cadena (p. ej. Cadena Llevant)" style="background:#1e293b;border:1px solid #334155;color:#f1f5f9;padding:8px;border-radius:7px;font-size:12px">
          <datalist id="hn-grupos"></datalist>
        </div>
        <button onclick="crearHotel()" class="btn bsm" style="font-size:12px;padding:6px 14px">Registrar hotel</button>
        <span id="hotel-msg" style="font-size:12px;margin-left:10px"></span>
      </div>
      <table class="ut"><thead><tr><th>Hotel</th><th>Ciudad</th><th>Hab</th><th>Grupo</th><th></th></tr></thead>
      <tbody id="htb"></tbody></table>
    </div>
    <div class="card" style="border-color:rgba(34,197,94,.2)" id="copias-card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
        <div class="ct" style="margin:0">💾 Copias de seguridad (fuera de Render)</div>
        <button class="btn bsm" onclick="loadCopias()" style="font-size:11px">↺ Actualizar</button>
      </div>
      <div id="copias-resumen" style="font-size:12px;color:#94a3b8;margin-bottom:10px">Cargando…</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">
        <button class="btn bsm" id="btn-copia-ahora" onclick="copiaAhora()" style="font-size:11px">💾 Copiar ahora</button>
        <span id="copias-msg" style="font-size:11px;color:#64748b;align-self:center"></span>
      </div>
      <div id="copias-lista" style="max-height:220px;overflow-y:auto;font-size:11px;font-family:monospace;color:#cbd5e1"></div>
    </div>
    <div class="card" style="border-color:rgba(239,68,68,.2)" id="alertas-card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
        <div class="ct" style="margin:0">🚨 Avisos de error por email</div>
        <button class="btn bsm" onclick="loadAlertas()" style="font-size:11px">↺ Actualizar</button>
      </div>
      <div id="alertas-resumen" style="font-size:12px;color:#94a3b8;margin-bottom:10px">Cargando…</div>
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">
        <button class="btn bsm" onclick="probarAlertas()" style="font-size:11px">📧 Enviar prueba</button>
        <span id="alertas-msg" style="font-size:11px;color:#64748b"></span>
      </div>
      <div id="alertas-lista" style="max-height:160px;overflow-y:auto;font-size:11px;font-family:monospace;color:#cbd5e1"></div>
    </div>
    <div class="card" style="border-color:rgba(239,68,68,.2)">
      <div class="ct" style="color:#ef4444">Herramientas</div>
      <button class="btn bd bsm" onclick="cleanCache()">Limpiar cache</button>
      <div class="msg" id="md"></div>
    </div>
    <div class="card" style="border-color:rgba(148,163,184,.2)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
        <div class="ct" style="margin:0">Registro de actividad</div>
        <button class="btn bsm" onclick="loadAudit()" style="font-size:11px">↺ Actualizar</button>
      </div>
      <div id="audit-log" style="max-height:200px;overflow-y:auto;font-size:11px;color:#94a3b8;font-family:monospace"></div>
    </div>

    <!-- LEADS / Solicitudes self-service -->
    <div class="card" style="border-color:rgba(34,197,94,.2)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
        <div class="ct" style="margin:0">🎯 Solicitudes de hoteles (leads)</div>
        <button class="btn bsm" onclick="loadLeads()" style="font-size:11px">↺ Actualizar</button>
      </div>
      <div id="leads-list" style="max-height:300px;overflow-y:auto"><div style="color:#64748b;font-size:13px;padding:12px">Cargando...</div></div>
    </div>
    <div class="card" style="border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2)">
      <div class="ct" style="color:var(--acc2)">Conexiones</div>
      <div style="display:flex;flex-direction:column;gap:8px">
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('smtp')" style="min-width:120px;font-size:11px">📧 Test SMTP</button>
          <span id="smtp-status" style="font-size:11px;color:#64748b"></span>
        </div>
        <div style="display:flex;align-items:center;gap:8px;margin-top:8px">
          <button class="btn bsm" onclick="checkSystemHealth()" style="min-width:120px;font-size:11px;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1);border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.3);color:var(--acc2)">🔍 Estado sistema</button>
          <div id="health-summary" style="font-size:11px;color:#64748b"></div>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('oracle')" style="min-width:120px;font-size:11px">🔴 Test Oracle</button>
          <a href="/api/oracle/export_excel" class="btn bsm" style="text-decoration:none;font-size:11px">⬇ GL Excel</a>
          <a href="/api/oracle/dryrun" target="_blank" class="btn bsm" style="text-decoration:none;font-size:11px">👁 Preview</a>
          <span id="oracle-status" style="font-size:11px;color:#64748b"></span>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('telegram')" style="min-width:120px;font-size:11px">✈️ Test Telegram</button>
          <span id="telegram-status" style="font-size:11px;color:#64748b"></span>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('stripe')" style="min-width:120px;font-size:11px">💳 Test Stripe</button>
          <span id="stripe-status" style="font-size:11px;color:#64748b"></span>
        </div>
      </div>
    </div>
  </div>
</div>
<script>
async function ls(){const r=await fetch('/admin/api/stats'),d=await r.json();['u','a','h','r','up','v'].forEach((k,i)=>{const el=document.getElementById('s'+k);if(el)el.textContent=[d.usuarios,d.usuarios_activos,d.hoteles,d.reportes_generados,d.uptime,d.version][i];});}
async function lu(){const r=await fetch('/admin/api/usuarios'),us=await r.json();document.getElementById('utb').innerHTML=us.map(u=>'<tr><td><b>'+u.username+'</b></td><td>'+(u.nombre||'-')+'</td><td style="color:var(--acc2)">'+u.rol+'</td><td><span class="'+(u.activo!==false?'ok':'off')+'">'+(u.activo!==false?'Activo':'Inactivo')+'</span></td><td><button class="btn bd bsm" onclick="tU(\''+u.username+'\')">Toggle</button></td></tr>').join('');}
async function lh(){const r=await fetch('/admin/api/hoteles'),d=await r.json();if(!d.ok)return;document.getElementById('htb').innerHTML=d.hoteles.map(h=>'<tr><td><b>'+h.nombre+'</b></td><td style="color:#94a3b8">'+(h.ciudad||'-')+'</td><td>'+(h.habitaciones||'-')+'</td><td style="color:#94a3b8">'+(h.grupo||'-')+'</td><td><button class="btn bd bsm" onclick="dH(\''+h.id+'\')">×</button></td></tr>').join('');
  const dl=document.getElementById('hn-grupos');if(dl)dl.innerHTML=[...new Set(d.hoteles.map(h=>h.grupo).filter(g=>g&&g!=='Principal'))].map(g=>'<option value="'+g.replace(/"/g,'&quot;')+'">').join('');}
async function loadLeads() {
  const el = document.getElementById('leads-list');
  try {
    const r = await fetch('/api/admin/leads');
    const d = await r.json();
    if (!d.ok || !d.leads || !d.leads.length) {
      el.innerHTML = '<div style="color:#64748b;font-size:13px;padding:12px">Sin solicitudes todavía. Aparecerán aquí cuando un hotel use /unirse.</div>';
      return;
    }
    el.innerHTML = d.leads.reverse().map(function(l){
      var planColor = {starter:'#60a5fa', pro:'#a78bfa', multi:'#22c55e'}[l.plan_sugerido] || '#64748b';
      var fecha = l.fecha ? new Date(l.fecha).toLocaleDateString('es-ES',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}) : '';
      return '<div style="border:1px solid #334155;border-radius:8px;padding:12px;margin-bottom:8px;font-size:12px">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">' +
          '<b style="color:#f1f5f9;font-size:13px">'+l.hotel+'</b>' +
          '<span style="background:'+planColor+'22;color:'+planColor+';padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;text-transform:uppercase">'+l.plan_sugerido+'</span>' +
        '</div>' +
        '<div style="color:#94a3b8;line-height:1.6">' +
          '👤 '+l.nombre+' · <a href="mailto:'+l.email+'" style="color:var(--acc2)">'+l.email+'</a>'+(l.telefono?' · 📞 '+l.telefono:'')+'<br>' +
          '📍 '+(l.ciudad||'—')+' · 🏨 '+(l.habitaciones||0)+' hab'+(l.grupo?' · Grupo: '+l.grupo:'')+'<br>' +
          (l.notas?'<span style="color:#64748b">📝 '+l.notas+'</span><br>':'') +
          '<span style="color:#475569;font-size:11px">'+fecha+'</span>' +
        '</div>' +
      '</div>';
    }).join('');
  } catch(e) {
    el.innerHTML = '<div style="color:#ef4444;font-size:13px;padding:12px">Error cargando leads</div>';
  }
}

async function loadAudit() {
  const el = document.getElementById('audit-log');
  if (!el) return;
  try {
    const r = await fetch('/admin/api/audit');
    const d = await r.json();
    const entries = d.entries || [];
    if (!entries.length) { el.textContent = 'Sin registros aún.'; return; }
    el.innerHTML = entries.slice().reverse().map(e =>
      '<div style="padding:3px 0;border-bottom:1px solid #1e293b">' +
      '<span style="color:#475569">' + e.ts + '</span> ' +
      '<span style="color:var(--acc2)">' + e.accion + '</span> ' +
      '<span>' + e.detalle + '</span>' +
      ' <span style="color:#334155">— ' + e.usuario + '</span>' +
      '</div>'
    ).join('');
  } catch(e) { el.textContent = 'Error cargando audit log.'; }
}
// Load audit log on page load
setTimeout(loadAudit, 1000);
async function resetPw(u) {
  const pw = prompt('Nueva contraseña para ' + u + ' (mínimo 6 caracteres):');
  if (!pw || pw.length < 6) { alert('Contraseña muy corta o cancelado'); return; }
  const r = await fetch('/admin/api/cambiar_password', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,nueva_password:pw})});
  const d = await r.json();
  showMsg(d.ok ? '✓ Contraseña cambiada para ' + u : '✗ ' + (d.error||'Error'), d.ok);
}
async function tU(u){await fetch('/admin/api/toggle_usuario',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u})});lu();ls();loadLeads();}
async function dH(id){if(!confirm('Eliminar '+id+'?'))return;await fetch('/admin/api/hoteles/eliminar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});lh();ls();}
async function crearU(){const d={username:document.getElementById('nu').value.trim(),password:document.getElementById('np').value,nombre:document.getElementById('nn').value.trim(),email:document.getElementById('ne').value.trim(),rol:document.getElementById('nr').value};const m=document.getElementById('mu');if(!d.username||!d.password){m.style.color='#ef4444';m.textContent='Faltan campos.';return;}const r=await fetch('/admin/api/crear_usuario',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});const res=await r.json();m.style.color=res.ok?'#22c55e':'#ef4444';m.textContent=res.ok?'Creado: '+d.username:(res.error||'Error');if(res.ok){lu();ls();}}
async function crearHotel(){
  const d={nombre:document.getElementById('hn-nombre').value.trim(),ciudad:document.getElementById('hn-ciudad').value.trim(),
    habitaciones:document.getElementById('hn-hab').value,grupo:document.getElementById('hn-grupo').value.trim()};
  const m=document.getElementById('hotel-msg');
  if(!d.nombre){m.style.color='#ef4444';m.textContent='El nombre es obligatorio';return;}
  const r=await fetch('/admin/api/hoteles/crear',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});
  const res=await r.json();m.style.color=res.ok?'#22c55e':'#ef4444';
  m.textContent=res.ok?'✓ Hotel '+res.nombre+' registrado':(res.error||'Error');
  if(res.ok){lh();ls();document.getElementById('nuevo-hotel-form').style.display='none';
    document.getElementById('hn-nombre').value='';document.getElementById('hn-ciudad').value='';}
}
async function testConn(type) {
  const el = document.getElementById(type + '-status');
  el.style.color = '#64748b'; el.textContent = '⏳ Probando...';
  try {
    const r = await fetch('/api/test_' + type, {method:'POST'});
    const d = await r.json();
    el.style.color = d.ok ? '#22c55e' : '#ef4444';
    el.textContent = (d.ok ? '✓ ' : '✗ ') + (d.message || d.error || 'Error');
  } catch(e) { el.style.color = '#ef4444'; el.textContent = '✗ Error de red'; }
}
async function loadCopias(){
  const res=document.getElementById('copias-resumen'), lst=document.getElementById('copias-lista');
  try{
    const r=await fetch('/admin/api/copias',{cache:'no-store'}); const d=await r.json();
    if(!d.configurado){res.innerHTML='<span style="color:#f59e0b">Sin destino configurado.</span> En Render → Environment pon <code>YVE_BACKUP_S3_BUCKET</code>, <code>YVE_BACKUP_S3_ENDPOINT</code>, <code>YVE_BACKUP_S3_KEY</code> y <code>YVE_BACKUP_S3_SECRET</code> (Cloudflare R2, Backblaze B2 o AWS S3). Hasta entonces no hay copia diaria.';lst.innerHTML='';return;}
    const u=d.ultima_ok||{}; const ult=d.ultima||{};
    res.innerHTML='Destino: <b>'+(d.destinos||[]).join(', ')+'</b> · cada día a las <b>'+d.hora+' UTC</b> · se guardan <b>'+d.retencion_dias+' días</b><br>'+
      (u.fecha?('Última copia buena: <b>'+u.fecha.replace('T',' ').slice(0,16)+' UTC</b> ('+u.ficheros+' ficheros, '+Math.round((u.bytes||0)/1024)+' KB)'+(d.al_dia?' <span class="ok">al día</span>':' <span style="color:#f59e0b">hace más de un día</span>')):'<span style="color:#f59e0b">Todavía no hay ninguna copia buena.</span>')+
      (ult.fecha&&!ult.ok?('<br><span style="color:#ef4444">Último intento FALLÓ ('+ult.fecha.replace('T',' ').slice(0,16)+'): '+((ult.errores||[]).join(' · ')||'error')+'</span>'):'')+
      (d.ultima_restauracion&&d.ultima_restauracion.fecha?('<br>Última restauración: '+d.ultima_restauracion.fecha.replace('T',' ').slice(0,16)+' UTC · '+d.ultima_restauracion.nombre+(d.ultima_restauracion.ok?' (ok)':' (FALLÓ: '+(d.ultima_restauracion.error||'')+')')):'');
    const cs=(d.copias||[]).filter(c=>c.nombre);
    lst.innerHTML=cs.length?cs.map(c=>'<div style="display:flex;gap:8px;align-items:center;padding:3px 0;border-bottom:1px solid #1e293b"><span style="flex:1">'+c.nombre+'</span><span style="color:#64748b">'+Math.round((c.bytes||0)/1024)+' KB · '+c.destino+'</span><a class="btn bsm" style="font-size:10px;padding:2px 8px;text-decoration:none" href="/admin/api/copias/descargar?nombre='+encodeURIComponent(c.nombre)+'&destino='+c.destino+'">⬇</a><button class="btn bd bsm" style="font-size:10px;padding:2px 8px" onclick="restaurarCopia(\''+c.nombre+'\',\''+c.destino+'\')">Restaurar</button></div>').join(''):'<div style="color:#64748b">Sin copias todavía.</div>';
  }catch(e){res.textContent='No se pudo leer el estado de las copias.';}
}
async function copiaAhora(){
  const b=document.getElementById('btn-copia-ahora'), m=document.getElementById('copias-msg');
  b.disabled=true; m.style.color='#64748b'; m.textContent='Copiando…';
  try{const r=await fetch('/admin/api/copias/ahora',{method:'POST'}); const d=await r.json();
    m.style.color=d.ok?'#22c55e':'#ef4444'; m.textContent=d.ok?('Copia hecha: '+d.nombre+' ('+d.ficheros+' ficheros)'):('Falló: '+((d.errores||[]).join(' · ')||d.error||'error'));
  }catch(e){m.style.color='#ef4444';m.textContent='Error de red';}
  b.disabled=false; loadCopias();
}
async function restaurarCopia(nombre,destino){
  const p=prompt('Vas a SUSTITUIR todos los datos actuales por los de la copia\n'+nombre+'\n\nAntes se guarda una copia de lo que hay ahora ("antes_de_restaurar_…").\n\nEscribe RESTAURAR para confirmar:');
  if(p===null)return; const m=document.getElementById('copias-msg'); m.style.color='#64748b'; m.textContent='Restaurando…';
  try{const r=await fetch('/admin/api/copias/restaurar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nombre,destino,confirmar:p})}); const d=await r.json();
    m.style.color=d.ok?'#22c55e':'#ef4444'; m.textContent=d.ok?('Restaurada: '+d.ficheros+' ficheros de '+nombre):('No se restauró: '+(d.error||'error'));
  }catch(e){m.style.color='#ef4444';m.textContent='Error de red';}
  loadCopias();
}
async function loadAlertas(){
  const res=document.getElementById('alertas-resumen'), lst=document.getElementById('alertas-lista');
  try{const r=await fetch('/admin/api/alertas',{cache:'no-store'}); const d=await r.json();
    res.innerHTML=(d.configurado?('Avisa a <b>'+d.email+'</b> (margen '+d.margen_min+' min entre emails)'):'<span style="color:#f59e0b">Sin <code>YVE_ALERTAS_EMAIL</code> en Render: los errores se apuntan aquí pero no se avisa a nadie.</span>')+
      '<br>Errores desde el último arranque: <b>'+d.total+'</b> · emails enviados: '+d.enviados+(d.pendientes?(' · '+d.pendientes+' esperando el resumen'):'')+(d.ultimo_envio?(' · último email '+d.ultimo_envio.replace('T',' ').slice(0,16)+' UTC'):'');
    lst.innerHTML=(d.ultimos||[]).length?d.ultimos.map(e=>'<div style="padding:2px 0;border-bottom:1px solid #1e293b"><span style="color:#64748b">'+e.fecha.replace('T',' ').slice(5,16)+'</span> <b>'+e.status+'</b> '+e.metodo+' '+e.ruta+' <span style="color:#94a3b8">'+(e.usuario||'')+'</span><br><span style="color:#f87171">'+(e.error||'').slice(0,160)+'</span></div>').join(''):'<div style="color:#64748b">Sin errores desde el arranque.</div>';
  }catch(e){res.textContent='No se pudo leer el estado de las alertas.';}
}
async function probarAlertas(){
  const m=document.getElementById('alertas-msg'); m.style.color='#64748b'; m.textContent='Enviando…';
  try{const r=await fetch('/admin/api/alertas/probar',{method:'POST'}); const d=await r.json();
    m.style.color=d.ok?'#22c55e':'#ef4444'; m.textContent=d.ok?('Enviado a '+d.email+': mira el buzón'):(d.error||'No se pudo enviar');
  }catch(e){m.style.color='#ef4444';m.textContent='Error de red';}
}
async function checkSystemHealth(){
  const el=document.getElementById('health-summary'); el.textContent='Comprobando…';
  try{const r=await fetch('/api/health/detalle',{cache:'no-store'}); const d=await r.json(); const c=d.components||{};
    el.innerHTML=Object.keys(c).map(k=>'<span style="color:'+(c[k].ok?'#22c55e':'#f59e0b')+'">'+(c[k].ok?'●':'○')+' '+k+'</span>: '+(c[k].msg||'')).join(' · ')+' · estado <b>'+d.status+'</b>';
  }catch(e){el.textContent='No responde';}
}
async function cleanCache(){const d=document.getElementById('md');d.textContent='Cache limpiado';d.style.color='#22c55e';}
ls();lu();lh();loadCopias();loadAlertas();
</script>
</body>
</html>"""

@bp.route("/api/hoteles/crear", methods=["POST"])
@_admin_required
def api_crear_hotel():
    """Registra un nuevo hotel en el grupo."""
    data = request.get_json(force=True, silent=True) or {}
    nombre    = data.get("nombre", "").strip()
    ciudad    = data.get("ciudad", "").strip()
    categoria = data.get("categoria", "4★")
    hab       = int(data.get("habitaciones", 100))
    grupo     = (data.get("grupo") or "").strip() or "Principal"

    if not nombre:
        return jsonify({"ok": False, "error": "El nombre es obligatorio"}), 400

    ruta = __import__("pathlib").Path(__import__("tenant_dirs").datos_dir()) / "hoteles.json"
    hotels = json.loads(ruta.read_text()) if ruta.exists() else []

    import hashlib, time
    hid = "H" + hashlib.md5((nombre + str(time.time())).encode()).hexdigest()[:6].upper()
    hotels.append({
        "id": hid, "nombre": nombre, "ciudad": ciudad,
        "categoria": categoria, "habitaciones": hab,
        "grupo": grupo, "activo": True,
        "modulos": ["ar", "ap", "drr", "banco", "fb"],
        "creado": __import__("datetime").datetime.now().strftime("%Y-%m-%d"),
    })
    ruta.write_text(json.dumps(hotels, ensure_ascii=False, indent=2))
    return jsonify({"ok": True, "id": hid, "nombre": nombre})


