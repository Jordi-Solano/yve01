"""
login.py — Yve.01
Blueprint de autenticación (login / logout). Se registra en dashboard.py.
Página de login servida en /login. Sin proceso ni puerto propio.
"""

from flask import Blueprint, request, redirect, jsonify
from flask_login import login_user, logout_user, current_user
from auth import login as auth_login

bp = Blueprint("auth", __name__)


@bp.route("/login")
def login_page():
    if current_user.is_authenticated:
        return redirect("/")
    return HTML


@bp.route("/api/login", methods=["POST"])
def do_login():
    data = request.get_json(force=True, silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    user = auth_login(username, password)
    if user:
        login_user(user, remember=True)
        return jsonify({"ok": True, "nombre": user.nombre, "rol": user.rol})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401


@bp.route("/logout")
def do_logout():
    logout_user()
    return redirect("/login")


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Iniciar sesión</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--red:#ef4444;--grn:#22c55e}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;position:relative;overflow:hidden}
body::before{content:"";position:fixed;inset:0;z-index:0;pointer-events:none;
  background:
    radial-gradient(1000px 500px at var(--bx1,82%) var(--by1,-12%), rgba(59,130,246,.18), transparent 60%),
    radial-gradient(700px 400px at var(--bx2,-10%) var(--by2,112%), rgba(139,92,246,.12), transparent 55%);
  animation:bgFloat 12s ease-in-out infinite alternate}
@keyframes bgFloat{
  0%{--bx1:82%;--by1:-12%;--bx2:-10%;--by2:112%}
  50%{--bx1:88%;--by1:5%;--bx2:5%;--by2:105%}
  100%{--bx1:75%;--by1:-8%;--bx2:-5%;--by2:118%}
}
/* CSS custom props not animatable in all browsers — use transform instead */
body::after{content:"";position:fixed;width:600px;height:600px;
  border-radius:50%;
  background:radial-gradient(circle,rgba(59,130,246,.08),transparent 70%);
  top:-100px;right:-100px;z-index:0;pointer-events:none;
  animation:blobA 15s ease-in-out infinite alternate}
@keyframes blobA{
  0%{transform:translate(0,0) scale(1)}
  50%{transform:translate(-40px,60px) scale(1.1)}
  100%{transform:translate(30px,-30px) scale(0.95)}
}
.wrap{position:relative;z-index:1;width:100%;max-width:410px}
.login-card{background:linear-gradient(160deg,rgba(30,41,59,.94),rgba(15,23,42,.96));
  border:1px solid var(--s2);border-radius:20px;padding:38px 34px;
  box-shadow:0 24px 70px rgba(0,0,0,.55),0 0 0 1px rgba(59,130,246,.07) inset;
  backdrop-filter:blur(12px);animation:rise .4s cubic-bezier(.2,.8,.2,1)}
@keyframes rise{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.brand{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.brand-dot{width:11px;height:11px;border-radius:50%;background:var(--acc);box-shadow:0 0 14px var(--acc)}
.brand-name{font-size:24px;font-weight:800;letter-spacing:-.6px;color:#fff}
.brand-name span{color:var(--acc2)}
.brand-sub{font-size:12.5px;color:var(--mut);margin-bottom:26px;line-height:1.5}
.heading{font-size:15px;font-weight:700;color:var(--tx);margin-bottom:18px}
label{display:block;font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px;margin-top:16px;font-weight:600}
label:first-of-type{margin-top:0}
input{width:100%;background:rgba(15,23,42,.7);border:1px solid var(--s2);color:var(--tx);border-radius:11px;
  padding:13px 14px;font-size:14.5px;outline:none;font-family:inherit}
input::placeholder{color:var(--dim)}
input:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(59,130,246,.18)}
.btn-login{width:100%;margin-top:24px;padding:14px;background:linear-gradient(135deg,var(--acc),#1d4ed8);
  color:#fff;border:none;border-radius:11px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;
  box-shadow:0 8px 24px rgba(59,130,246,.32)}
.btn-login:hover{box-shadow:0 12px 32px rgba(59,130,246,.5),0 0 0 1px rgba(59,130,246,.3);transform:translateY(-1px)}
.btn-login:focus-visible{outline:2px solid var(--acc2);outline-offset:3px}
.btn-login:active{transform:translateY(0)}
.btn-login:disabled{opacity:.6;cursor:not-allowed;transform:none}
.error{display:none;margin-top:14px;padding:11px 14px;background:rgba(239,68,68,.1);
  border:1px solid rgba(239,68,68,.28);border-radius:10px;color:#fca5a5;font-size:13px;text-align:center}
.error.on{display:block;animation:rise .2s ease}
.demo{margin-top:26px;padding-top:22px;border-top:1px solid var(--s2)}
.demo-h{font-size:10.5px;color:var(--dim);text-transform:uppercase;letter-spacing:.7px;margin-bottom:11px;font-weight:600}
.chips{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.chip{background:rgba(51,65,85,.45);border:1px solid var(--s2);border-radius:10px;padding:9px 11px;
  cursor:pointer;text-align:left;font-family:inherit;transition:.15s}
.chip:hover{border-color:var(--acc);background:rgba(59,130,246,.1);box-shadow:0 0 12px rgba(59,130,246,.15)}
.chip-role{font-size:12px;font-weight:700;color:var(--acc2);display:block}
.chip-user{font-size:10.5px;color:var(--mut);margin-top:1px}
.foot{display:flex;align-items:center;justify-content:center;gap:7px;margin-top:22px;font-size:11.5px;color:var(--dim)}
.foot svg{width:13px;height:13px;opacity:.8}
</style>
</head>
<body>
<div class="wrap">
  <div class="login-card">
    <div class="brand"><span class="brand-dot"></span><span class="brand-name">Yve<span>.01</span></span></div>
    <div class="brand-sub">Automatización financiera para hoteles</div>

    <div class="heading">Inicia sesión</div>

    <label>Usuario</label>
    <input id="username" placeholder="tu usuario" autocomplete="username" autofocus>

    <label>Contraseña</label>
    <input id="password" type="password" placeholder="••••••••" autocomplete="current-password">

    <button class="btn-login" id="btn-login" onclick="doLogin()">Entrar al panel</button>
    <div class="error" id="error"></div>

    <div class="demo">
      <div class="demo-h">Accesos de demostración — pulsa para rellenar</div>
      <div class="chips">
        <button class="chip" onclick="fill('fc_user','hotel2024')"><span class="chip-role">Financial Controller</span><span class="chip-user">fc_user</span></button>
        <button class="chip" onclick="fill('auditor','hotel2024')"><span class="chip-role">Income Auditor</span><span class="chip-user">auditor</span></button>
        <button class="chip" onclick="fill('fbmanager','hotel2024')"><span class="chip-role">F&B Manager</span><span class="chip-user">fbmanager</span></button>
        <button class="chip" onclick="fill('admin','admin123')"><span class="chip-role">Administrador</span><span class="chip-user">admin</span></button>
      </div>
    </div>
  </div>
  <div class="foot">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l8 4v6c0 5-3.5 8-8 10-4.5-2-8-5-8-10V6l8-4z"/></svg>
    Sesión cifrada · Acceso por roles
  </div>
</div>

<script>
function fill(u,p){
  document.getElementById('username').value=u;
  document.getElementById('password').value=p;
  document.getElementById('btn-login').focus();
}
async function doLogin() {
  const btn = document.getElementById('btn-login');
  const err = document.getElementById('error');
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value;
  if (!username || !password) { err.textContent = 'Introduce usuario y contraseña'; err.classList.add('on'); return; }
  btn.disabled = true; btn.textContent = 'Verificando...'; err.classList.remove('on');
  try {
    const r = await fetch('/api/login', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({username, password})
    });
    const data = await r.json();
    if (data.ok) {
      btn.textContent = 'Bienvenido, ' + data.nombre;
      const next = new URLSearchParams(window.location.search).get('next') || '/';
      setTimeout(function() { window.location.href = next; }, 600);
    } else {
      err.textContent = data.error || 'Credenciales incorrectas';
      err.classList.add('on'); btn.disabled = false; btn.textContent = 'Entrar al panel';
    }
  } catch(e) {
    err.textContent = 'Error de conexión'; err.classList.add('on');
    btn.disabled = false; btn.textContent = 'Entrar al panel';
  }
}
document.getElementById('password').addEventListener('keydown', function(e){ if(e.key==='Enter') doLogin(); });
document.getElementById('username').addEventListener('keydown', function(e){ if(e.key==='Enter') document.getElementById('password').focus(); });
</script>
</body>
</html>"""
