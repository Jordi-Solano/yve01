"""
login.py — Yve.01
Página de login. Ejecutar: python login.py
Abre en: http://localhost:5005
"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, redirect, jsonify
from flask_login import login_user, logout_user, current_user
from auth import init_login, login as auth_login, inicializar_usuarios

app = Flask(__name__)
app.secret_key = "yve01-secret-key-change-in-production"
init_login(app)
inicializar_usuarios()


@app.route("/")
def index():
    if current_user.is_authenticated:
        return redirect("http://localhost:5001")
    return HTML


@app.route("/login", methods=["POST"])
def do_login():
    data = request.get_json(force=True, silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    user = auth_login(username, password)
    if user:
        login_user(user, remember=True)
        return jsonify({"ok": True, "nombre": user.nombre, "rol": user.rol})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401


@app.route("/logout")
def do_logout():
    logout_user()
    return redirect("/")


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — Login</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--red:#ef4444}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  min-height:100vh;display:flex;align-items:center;justify-content:center}
.login-card{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:40px 36px;
  width:100%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,.4);animation:fadeIn .3s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
.logo{text-align:center;margin-bottom:28px}
.logo-dot{display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--acc);
  box-shadow:0 0 12px var(--acc);vertical-align:middle;margin-right:8px}
.logo-name{font-size:26px;font-weight:800;color:var(--acc2);letter-spacing:-.5px}
.logo-sub{font-size:12px;color:var(--mut);margin-top:4px}
label{display:block;font-size:.8rem;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;margin-top:18px}
label:first-of-type{margin-top:0}
input{width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:10px;
  padding:12px 14px;font-size:.95rem;outline:none;transition:.15s;font-family:inherit}
input:focus{border-color:var(--acc);box-shadow:0 0 0 2px rgba(59,130,246,.15)}
.btn-login{width:100%;margin-top:24px;padding:13px;background:linear-gradient(135deg,var(--acc),#1d4ed8);
  color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer;
  transition:.15s;box-shadow:0 0 20px rgba(59,130,246,.3)}
.btn-login:hover{box-shadow:0 0 30px rgba(59,130,246,.5);transform:translateY(-1px)}
.btn-login:disabled{opacity:.5;cursor:not-allowed;transform:none}
.error{display:none;margin-top:14px;padding:10px 14px;background:rgba(239,68,68,.1);
  border:1px solid rgba(239,68,68,.25);border-radius:8px;color:#fca5a5;font-size:.85rem;text-align:center}
.error.on{display:block}
.users-hint{margin-top:20px;padding:14px;background:var(--bg);border:1px solid var(--s2);border-radius:10px;
  font-size:.75rem;color:var(--mut);line-height:1.6}
.users-hint b{color:var(--acc2)}
</style>
</head>
<body>

<div class="login-card">
  <div class="logo">
    <span class="logo-dot"></span><span class="logo-name">Yve.01</span>
    <div class="logo-sub">Dashboard Financiero</div>
  </div>

  <label>Usuario</label>
  <input id="username" placeholder="username" autofocus>

  <label>Contraseña</label>
  <input id="password" type="password" placeholder="••••••••">

  <button class="btn-login" id="btn-login" onclick="doLogin()">Iniciar sesión</button>

  <div class="error" id="error"></div>

  <div class="users-hint">
    Usuarios por defecto:<br>
    <b>admin</b> / admin123 — Administrador<br>
    <b>fc_user</b> / hotel2024 — Financial Controller<br>
    <b>auditor</b> / hotel2024 — Income Auditor<br>
    <b>fbmanager</b> / hotel2024 — F&B Manager
  </div>
</div>

<script>
async function doLogin() {
  const btn = document.getElementById('btn-login');
  const err = document.getElementById('error');
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value;
  if (!username || !password) { err.textContent = 'Introduce usuario y contraseña'; err.classList.add('on'); return; }

  btn.disabled = true;
  btn.textContent = 'Verificando...';
  err.classList.remove('on');

  try {
    const r = await fetch('/login', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({username, password})
    });
    const data = await r.json();
    if (data.ok) {
      btn.textContent = 'Bienvenido, ' + data.nombre;
      setTimeout(function() { window.location.href = 'http://localhost:5001'; }, 800);
    } else {
      err.textContent = data.error || 'Credenciales incorrectas';
      err.classList.add('on');
      btn.disabled = false;
      btn.textContent = 'Iniciar sesión';
    }
  } catch(e) {
    err.textContent = 'Error de conexión';
    err.classList.add('on');
    btn.disabled = false;
    btn.textContent = 'Iniciar sesión';
  }
}

document.getElementById('password').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doLogin();
});
document.getElementById('username').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') document.getElementById('password').focus();
});
</script>
</body>
</html>"""

if __name__ == "__main__":
    import socket
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = "127.0.0.1"
    print("=" * 60)
    print("  Yve.01 — Login")
    print("=" * 60)
    print("  Escritorio:  http://localhost:5005")
    print(f"  Movil:       http://{ip}:5005")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5005, debug=False)
