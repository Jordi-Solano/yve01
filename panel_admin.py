"""panel_admin.py - Yve.01 Admin Panel"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Blueprint, request, jsonify
from flask_login import login_required, current_user
from auth import (init_login, inicializar_usuarios, listar_usuarios,
                  crear_usuario, cambiar_password, toggle_activo, ROLES_VALIDOS)
from pathlib import Path

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
    return HTML


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
        hotels = json.loads((BASE_DIR / "datos-referencia" / "hoteles.json").read_text())
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
        hotels = json.loads((BASE_DIR / "datos-referencia" / "hoteles.json").read_text())
        return jsonify({"ok": True, "hoteles": hotels})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/api/hoteles/eliminar", methods=["POST"])
@_admin_required
def api_eliminar_hotel():
    data = request.get_json(silent=True) or {}
    hotel_id = data.get("id")
    if not hotel_id:
        return jsonify({"ok": False, "error": "ID requerido"}), 400
    try:
        path = BASE_DIR / "datos-referencia" / "hoteles.json"
        hotels = json.loads(path.read_text())
        hotels = [h for h in hotels if h.get("id") != hotel_id]
        path.write_text(json.dumps(hotels, indent=2, ensure_ascii=False))
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500




HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin - Yve.01</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--red:#ef4444}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;padding:28px;-webkit-font-smoothing:antialiased}
h1{font-size:20px;font-weight:800;margin-bottom:6px}
.sub{font-size:13px;color:var(--mut);margin-bottom:24px}
.sg{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;margin-bottom:24px}
.kc{background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px}
.kl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;font-weight:600}
.kv{font-size:26px;font-weight:800;color:var(--acc2);letter-spacing:-1px;line-height:1}
.card{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:22px;margin-bottom:18px}
.ct{font-size:14px;font-weight:700;margin-bottom:16px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:768px){.g2{grid-template-columns:1fr}}
label{display:block;font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px;margin-top:12px;font-weight:600}
label:first-child{margin-top:0}
input,select{width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:9px;padding:10px 13px;font-size:14px;font-family:inherit;outline:none}
.btn{padding:9px 18px;border:none;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit}
.bp{background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff}
.bd{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);color:#ef4444}
.bsm{padding:5px 12px;font-size:11px;border-radius:7px}
.ut{width:100%;border-collapse:collapse;font-size:13px}
.ut th{padding:9px 12px;font-size:10px;color:#94a3b8;text-transform:uppercase;border-bottom:1px solid #334155;text-align:left}
.ut td{padding:9px 12px;border-bottom:1px solid rgba(51,65,85,.4);vertical-align:middle}
.msg{font-size:12px;text-align:center;margin-top:10px;min-height:16px}
.ok{background:rgba(34,197,94,.1);color:#22c55e;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;display:inline-block}
.off{background:rgba(100,116,139,.1);color:#64748b;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;display:inline-block}
</style>
</head>
<body>
<h1>Admin - Yve.01</h1>
<div class="sub">Panel de administracion - Solo accesible para administradores</div>
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
          <select id="hn-cat" style="background:#1e293b;border:1px solid #334155;color:#f1f5f9;padding:8px;border-radius:7px;font-size:12px">
            <option>3★</option><option selected>4★</option><option>5★</option>
          </select>
        </div>
        <button onclick="crearHotel()" class="btn bsm" style="font-size:12px;padding:6px 14px">Registrar hotel</button>
        <span id="hotel-msg" style="font-size:12px;margin-left:10px"></span>
      </div>
      <table class="ut"><thead><tr><th>Hotel</th><th>Ciudad</th><th>Hab</th><th>Cat.</th><th></th></tr></thead>
      <tbody id="htb"></tbody></table>
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
    <div class="card" style="border-color:rgba(59,130,246,.2)">
      <div class="ct" style="color:#60a5fa">Conexiones</div>
      <div style="display:flex;flex-direction:column;gap:8px">
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('smtp')" style="min-width:120px;font-size:11px">📧 Test SMTP</button>
          <span id="smtp-status" style="font-size:11px;color:#64748b"></span>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <button class="btn bsm" onclick="testConn('oracle')" style="min-width:120px;font-size:11px">🔴 Test Oracle</button>
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
async function lu(){const r=await fetch('/admin/api/usuarios'),us=await r.json();document.getElementById('utb').innerHTML=us.map(u=>'<tr><td><b>'+u.username+'</b></td><td>'+(u.nombre||'-')+'</td><td style="color:#60a5fa">'+u.rol+'</td><td><span class="'+(u.activo!==false?'ok':'off')+'">'+(u.activo!==false?'Activo':'Inactivo')+'</span></td><td><button class="btn bd bsm" onclick="tU(\''+u.username+'\')">Toggle</button></td></tr>').join('');}
async function lh(){const r=await fetch('/admin/api/hoteles'),d=await r.json();if(!d.ok)return;document.getElementById('htb').innerHTML=d.hoteles.map(h=>'<tr><td><b>'+h.nombre+'</b></td><td style="color:#94a3b8">'+(h.ciudad||'-')+'</td><td>'+(h.habitaciones||'-')+'</td><td style="color:#94a3b8">'+(h.categoria||'-')+'</td><td><button class="btn bd bsm" onclick="dH(\''+h.id+'\')">×</button></td></tr>').join('');}
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
      '<span style="color:#60a5fa">' + e.accion + '</span> ' +
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
async function tU(u){await fetch('/admin/api/toggle_usuario',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u})});lu();ls();}
async function dH(id){if(!confirm('Eliminar '+id+'?'))return;await fetch('/admin/api/hoteles/eliminar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});lh();ls();}
async function crearU(){const d={username:document.getElementById('nu').value.trim(),password:document.getElementById('np').value,nombre:document.getElementById('nn').value.trim(),email:document.getElementById('ne').value.trim(),rol:document.getElementById('nr').value};const m=document.getElementById('mu');if(!d.username||!d.password){m.style.color='#ef4444';m.textContent='Faltan campos.';return;}const r=await fetch('/admin/api/crear_usuario',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});const res=await r.json();m.style.color=res.ok?'#22c55e':'#ef4444';m.textContent=res.ok?'Creado: '+d.username:(res.error||'Error');if(res.ok){lu();ls();}}
async function crearHotel(){
  const d={nombre:document.getElementById('hn-nombre').value.trim(),ciudad:document.getElementById('hn-ciudad').value.trim(),
    habitaciones:document.getElementById('hn-hab').value,categoria:document.getElementById('hn-cat').value};
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
async function cleanCache(){const d=document.getElementById('md');d.textContent='Cache limpiado';d.style.color='#22c55e';}
ls();lu();lh();
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
    grupo     = data.get("grupo", "Principal")

    if not nombre:
        return jsonify({"ok": False, "error": "El nombre es obligatorio"}), 400

    ruta = BASE_DIR / "datos-referencia" / "hoteles.json"
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


