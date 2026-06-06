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
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect x='9' y='9' width='14' height='14' rx='3' fill='%2300c8a0' transform='rotate(45 16 16)'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Iniciar sesión</title>
<style>
:root{
  --bg:#07090e;--s1:#0b1018;--s2:#111b28;--s3:#182234;--s4:#1e2c40;
  --bd:#19202f;--bd2:#1f2d42;
  --acc:#2d7ff9;--acc2:#5a9bff;--acc3:#99c0ff;
  --teal:#00c8a0;--teal2:#26dab4;
  --tx:#eef1f7;--mut:#7a8fa8;--dim:#3e5068;
  --red:#ff3b55;--grn:#00e08a;
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  background:var(--bg);color:var(--tx);
  font-family:'Plus Jakarta Sans',system-ui,sans-serif;
  min-height:100vh;display:flex;align-items:center;justify-content:center;
  padding:24px;position:relative;overflow:hidden
}
body::before{content:"";position:fixed;inset:0;z-index:0;pointer-events:none;
  background:
    radial-gradient(900px 500px at 80% -15%,rgba(45,127,249,.12),transparent 65%),
    radial-gradient(700px 400px at -8% 110%,rgba(0,200,160,.08),transparent 60%)}
.wrap{position:relative;z-index:1;width:100%;max-width:410px}
.login-card{
  background:linear-gradient(160deg,rgba(11,16,24,.96),rgba(7,9,14,.98));
  border:1px solid var(--bd2);border-radius:22px;padding:38px 34px;
  box-shadow:0 32px 80px rgba(0,0,0,.65),0 0 0 1px rgba(255,255,255,.03);
  backdrop-filter:blur(12px);animation:rise .4s cubic-bezier(.2,.8,.2,1)
}
@keyframes rise{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:none}}
.brand{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.brand-mark{
  width:34px;height:34px;
  background:linear-gradient(135deg,var(--teal),var(--acc));
  border-radius:10px;flex-shrink:0;
  display:flex;align-items:center;justify-content:center;
  box-shadow:0 0 20px rgba(0,200,160,.32)
}
.brand-mark::after{
  content:'Y';color:#fff;
  font-family:'Syne',sans-serif;
  font-weight:800;font-size:17px;line-height:1
}
.brand-dot{width:11px;height:11px;border-radius:50%;background:var(--teal);box-shadow:0 0 14px rgba(0,200,160,.5)}
.brand-name{font-family:'Syne',sans-serif;font-size:26px;font-weight:800;letter-spacing:-.5px;color:var(--tx)}
.brand-name span{color:var(--teal)}
.brand-sub{font-size:12.5px;color:var(--mut);margin-bottom:28px;line-height:1.6;margin-top:2px}
.heading{font-size:15px;font-weight:700;color:var(--tx);margin-bottom:18px}
label{display:block;font-size:10.5px;color:var(--mut);text-transform:uppercase;letter-spacing:.7px;margin-bottom:7px;margin-top:16px;font-weight:600}
label:first-of-type{margin-top:0}
input{
  width:100%;background:rgba(7,9,14,.8);
  border:1px solid var(--bd2);color:var(--tx);border-radius:12px;
  padding:13px 16px;font-size:14.5px;outline:none;font-family:inherit
}
input::placeholder{color:var(--dim)}
input:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(45,127,249,.18)}
.btn-login{
  width:100%;margin-top:24px;padding:14px;
  background:linear-gradient(135deg,var(--acc),#1a5ee8);
  color:#fff;border:none;border-radius:12px;
  font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;
  box-shadow:0 8px 28px rgba(45,127,249,.4),0 0 0 1px rgba(45,127,249,.2)
}
.btn-login:hover{box-shadow:0 12px 36px rgba(45,127,249,.6);transform:translateY(-1px)}
.btn-login:active{transform:translateY(0)}
.btn-login:disabled{opacity:.5;cursor:not-allowed;transform:none}
.error{display:none;margin-top:14px;padding:11px 14px;background:rgba(255,59,85,.08);
  border:1px solid rgba(255,59,85,.2);border-radius:10px;color:#ff8a9e;font-size:13px;text-align:center}
.error.on{display:block;animation:rise .2s ease}
.demo{margin-top:26px;padding-top:22px;border-top:1px solid var(--bd)}
.demo-h{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.8px;margin-bottom:11px;font-weight:700}
.chips{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.chip{
  background:rgba(255,255,255,.03);border:1px solid var(--bd);border-radius:10px;
  padding:9px 12px;cursor:pointer;text-align:left;font-family:inherit;transition:all .15s
}
.chip:hover{border-color:var(--acc);background:rgba(45,127,249,.08)}
.chip-role{font-size:12px;font-weight:700;color:var(--acc2);display:block}
.chip-user{font-size:10.5px;color:var(--mut);margin-top:2px;font-family:'DM Mono',monospace}
.foot{display:flex;align-items:center;justify-content:center;gap:7px;margin-top:22px;font-size:11.5px;color:var(--dim)}
.foot svg{width:13px;height:13px;opacity:.7}
</style>
</head>
<body>
<div class="wrap">
  <div class="login-card">
    <div class="brand"><div class="brand-mark"></div><span class="brand-name">Yve<span>.01</span></span></div>
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
