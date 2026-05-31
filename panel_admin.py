"""
panel_admin.py — Yve.01
Panel de administración de usuarios. Solo accesible para rol admin.
Se integra como blueprint o se ejecuta standalone en puerto 5006.
"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, redirect
from flask_login import login_required, current_user
from auth import (init_login, inicializar_usuarios, listar_usuarios,
                  crear_usuario, cambiar_password, toggle_activo, ROLES_VALIDOS)

app = Flask(__name__)
app.secret_key = "yve01-secret-key-change-in-production"
init_login(app)
inicializar_usuarios()


def _admin_required(f):
    from functools import wraps
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if current_user.rol != "admin":
            return jsonify({"error": "Acceso denegado — solo admin"}), 403
        return f(*args, **kwargs)
    return decorated


@app.route("/")
@login_required
def index():
    if current_user.rol != "admin":
        return "<h3>Acceso denegado</h3><p>Solo administradores.</p>", 403
    return HTML


@app.route("/api/usuarios")
@_admin_required
def api_usuarios():
    return jsonify(listar_usuarios())


@app.route("/api/crear_usuario", methods=["POST"])
@_admin_required
def api_crear():
    d = request.get_json(force=True) or {}
    result = crear_usuario(
        d.get("username", "").strip(),
        d.get("password", ""),
        d.get("nombre", ""),
        d.get("email", ""),
        d.get("rol", "income_auditor"),
    )
    if result is True:
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": result}), 400


@app.route("/api/cambiar_password", methods=["POST"])
@_admin_required
def api_cambiar_pass():
    d = request.get_json(force=True) or {}
    result = cambiar_password(d.get("username", ""), d.get("password", ""))
    if result is True:
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": result}), 400


@app.route("/api/toggle_usuario", methods=["POST"])
@_admin_required
def api_toggle():
    d = request.get_json(force=True) or {}
    result = toggle_activo(d.get("username", ""))
    if result is not None:
        return jsonify({"ok": True, "activo": result})
    return jsonify({"ok": False, "error": "Usuario no encontrado"}), 404


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — Admin</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--red:#ef4444;--ora:#f97316}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh}
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc);display:inline-block;margin-right:8px}
.logo-name{font-size:20px;font-weight:800;color:var(--acc2)}
.logo-tag{font-size:11px;color:var(--mut);margin-left:8px}
.nav-right{margin-left:auto}
.nav-right a{color:var(--acc2);text-decoration:none;font-size:.85rem;padding:6px 14px;border:1px solid var(--s2);border-radius:8px;transition:.15s}
.nav-right a:hover{border-color:var(--acc)}
.main{max-width:900px;margin:32px auto;padding:0 24px}
.card{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px;margin-bottom:24px}
.card h2{font-size:1.1rem;margin-bottom:16px}
table{width:100%;border-collapse:collapse;font-size:.85rem}
th{background:rgba(51,65,85,.6);color:var(--mut);font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;padding:10px 12px;text-align:left;border-bottom:1px solid var(--s2)}
td{padding:10px 12px;border-bottom:1px solid rgba(51,65,85,.4)}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.7rem;font-weight:700}
.b-on{background:rgba(34,197,94,.15);color:#86efac}.b-off{background:rgba(239,68,68,.15);color:#fca5a5}
.b-role{background:rgba(59,130,246,.15);color:#93c5fd}
.btn{padding:6px 12px;border-radius:8px;font-size:.8rem;font-weight:600;cursor:pointer;border:none;transition:.15s}
.btn-sm{padding:4px 10px;font-size:.75rem}
.btn-blue{background:var(--acc);color:#fff}.btn-blue:hover{background:var(--acc2)}
.btn-red{background:rgba(239,68,68,.2);color:#fca5a5}.btn-red:hover{background:rgba(239,68,68,.3)}
.btn-grn{background:rgba(34,197,94,.2);color:#86efac}.btn-grn:hover{background:rgba(34,197,94,.3)}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
@media(max-width:600px){.form-row{grid-template-columns:1fr}}
input,select{width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:8px;padding:9px 12px;font-size:.85rem;outline:none}
input:focus,select:focus{border-color:var(--acc)}
label{font-size:.75rem;color:var(--mut);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;display:block;margin-top:10px}
#msg{display:none;padding:10px;border-radius:8px;font-size:.85rem;margin-bottom:14px}
#msg.ok{display:block;background:rgba(34,197,94,.1);color:#86efac;border:1px solid rgba(34,197,94,.2)}
#msg.err{display:block;background:rgba(239,68,68,.1);color:#fca5a5;border:1px solid rgba(239,68,68,.2)}
</style>
</head>
<body>
<nav class="nav">
  <span class="logo-dot"></span><span class="logo-name">Yve.01</span>
  <span class="logo-tag">Administración</span>
  <div class="nav-right"><a href="http://localhost:5001">← Dashboard</a></div>
</nav>

<div class="main">
  <!-- Crear usuario -->
  <div class="card">
    <h2>Crear Usuario</h2>
    <div id="msg"></div>
    <div class="form-row">
      <div><label>Username</label><input id="nu-user" placeholder="nuevo_usuario"></div>
      <div><label>Contraseña</label><input id="nu-pass" type="password" placeholder="••••••"></div>
    </div>
    <div class="form-row">
      <div><label>Nombre</label><input id="nu-nombre" placeholder="Nombre completo"></div>
      <div><label>Email</label><input id="nu-email" type="email" placeholder="email@hotel.com"></div>
    </div>
    <div class="form-row">
      <div><label>Rol</label>
        <select id="nu-rol">
          <option value="financial_controller">Financial Controller</option>
          <option value="income_auditor">Income Auditor</option>
          <option value="fb_manager">F&B Manager</option>
          <option value="jefe_otras">Jefe de OTRAS</option>
          <option value="admin">Admin</option>
        </select>
      </div>
      <div style="display:flex;align-items:flex-end"><button class="btn btn-blue" onclick="crearUsuario()">+ Crear</button></div>
    </div>
  </div>

  <!-- Lista usuarios -->
  <div class="card">
    <h2>Usuarios</h2>
    <table>
      <thead><tr><th>Username</th><th>Nombre</th><th>Email</th><th>Rol</th><th>Estado</th><th>Acciones</th></tr></thead>
      <tbody id="usr-tbody"></tbody>
    </table>
  </div>
</div>

<script>
async function loadUsers() {
  const r = await fetch('/api/usuarios');
  const users = await r.json();
  const tbody = document.getElementById('usr-tbody');
  tbody.innerHTML = users.map(function(u) {
    const est = u.activo !== false
      ? '<span class="badge b-on">Activo</span>'
      : '<span class="badge b-off">Inactivo</span>';
    const toggleBtn = u.activo !== false
      ? '<button class="btn btn-sm btn-red" onclick="toggleUser(\'' + u.username + '\')">Desactivar</button>'
      : '<button class="btn btn-sm btn-grn" onclick="toggleUser(\'' + u.username + '\')">Activar</button>';
    return '<tr>'
      + '<td style="font-weight:700">' + u.username + '</td>'
      + '<td>' + (u.nombre || '') + '</td>'
      + '<td style="color:var(--dim)">' + (u.email || '') + '</td>'
      + '<td><span class="badge b-role">' + u.rol + '</span></td>'
      + '<td>' + est + '</td>'
      + '<td style="display:flex;gap:6px">'
      + toggleBtn
      + '<button class="btn btn-sm btn-blue" onclick="resetPass(\'' + u.username + '\')">Reset pass</button>'
      + '</td></tr>';
  }).join('');
}

function showMsg(text, ok) {
  const el = document.getElementById('msg');
  el.textContent = text;
  el.className = ok ? 'ok' : 'err';
  setTimeout(function() { el.className = ''; }, 4000);
}

async function crearUsuario() {
  const d = {
    username: document.getElementById('nu-user').value.trim(),
    password: document.getElementById('nu-pass').value,
    nombre:   document.getElementById('nu-nombre').value.trim(),
    email:    document.getElementById('nu-email').value.trim(),
    rol:      document.getElementById('nu-rol').value,
  };
  if (!d.username || !d.password) { showMsg('Username y contraseña requeridos', false); return; }
  const r = await fetch('/api/crear_usuario', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(d)});
  const res = await r.json();
  if (res.ok) { showMsg('Usuario creado: ' + d.username, true); loadUsers(); }
  else { showMsg(res.error || 'Error', false); }
}

async function toggleUser(username) {
  const r = await fetch('/api/toggle_usuario', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({username})});
  const res = await r.json();
  if (res.ok) loadUsers();
}

async function resetPass(username) {
  const pw = prompt('Nueva contraseña para ' + username + ':');
  if (!pw) return;
  const r = await fetch('/api/cambiar_password', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({username, password:pw})});
  const res = await r.json();
  if (res.ok) showMsg('Contraseña cambiada para ' + username, true);
  else showMsg(res.error || 'Error', false);
}

loadUsers();
</script>
</body>
</html>"""

if __name__ == "__main__":
    print("=" * 60)
    print("  Yve.01 — Panel Admin")
    print("  http://localhost:5006")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5006, debug=False)
