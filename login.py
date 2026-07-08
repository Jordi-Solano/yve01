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
        from flask import session
        session["tenant_id"] = getattr(user, "tenant", "default") or "default"
        session.pop("hotel_activo", None)
        # Cuentas de ejemplo: si su tenant está vacío, generar datos demo con sus nombres
        _DEMOS = {
            "solmar":   [{"nombre": "Cadena Sol", "hoteles": ["Hotel Sol Mar", "Hotel Sol Playa"]}],
            "gestoria": [{"nombre": "Gestoría Nord", "hoteles": ["Hotel Pirineus", "Hotel Vall d'Aran"]}],
        }
        if username in _DEMOS:
            try:
                import os as _os, pandas as _pd
                from tenant_dirs import datos_dir as _t_ddir
                _kpis = _os.path.join(_t_ddir(), "kpis_hoteles.xlsx")
                if not _os.path.exists(_kpis) or _pd.read_excel(_kpis).empty:
                    from demo_generator import generar_demo
                    generar_demo(_DEMOS[username])
            except Exception as _e:
                print(f"[login] demo seed warning: {_e}")
        return jsonify({"ok": True, "nombre": user.nombre, "rol": user.rol, "tenant": session["tenant_id"]})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401


@bp.route("/logout")
def do_logout():
    logout_user()
    from flask import session
    session.pop("tenant_id", None)
    session.pop("hotel_activo", None)
    return redirect("/login")


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5,viewport-fit=cover">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%230f172a'/%3E%3Crect width='32' height='32' rx='8' fill='url(%23g)'/%3E%3Cdefs%3E%3ClinearGradient id='g' x1='0' y1='0' x2='32' y2='32' gradientUnits='userSpaceOnUse'%3E%3Cstop offset='0' stop-color='%233b82f6' stop-opacity='.15'/%3E%3Cstop offset='1' stop-color='%23a78bfa' stop-opacity='.08'/%3E%3C/linearGradient%3E%3C/defs%3E%3Ccircle cx='16' cy='10' r='3' fill='%233b82f6'/%3E%3Cpath d='M10 6 L16 16 L22 6' stroke='%233b82f6' stroke-width='2.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3Cline x1='16' y1='16' x2='16' y2='26' stroke='%2360a5fa' stroke-width='2.5' stroke-linecap='round'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Iniciar sesión</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--red:#ef4444;--grn:#22c55e}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px 16px;position:relative;overflow-x:hidden;overflow-y:auto}
@media(max-width:480px){
  body{padding:16px;align-items:flex-start;padding-top:40px}
  .login-card{padding:28px 22px;border-radius:16px}
  .brand-name{font-size:22px}
  .chips{grid-template-columns:1fr}
}
/* ── Fondo premium: degradado limpio, sin formas ── */
body{background:linear-gradient(180deg, #101a2e 0%, #0c1424 55%, #090e1a 100%)}
.wrap{position:relative;z-index:1;width:100%;max-width:410px}
.login-card{position:relative;background:linear-gradient(170deg,rgba(23,32,50,.92),rgba(13,20,35,.96));
  border:1px solid rgba(148,163,184,.28);border-radius:20px;padding:38px 34px;
  backdrop-filter:blur(14px);animation:rise .45s cubic-bezier(.2,.8,.2,1)}
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
.error{display:none;margin-top:14px;padding:12px 16px;background:rgba(239,68,68,.12);
  border:1px solid rgba(239,68,68,.35);border-radius:10px;color:#fca5a5;font-size:13px;text-align:center;font-weight:500}
.error.on{display:block;animation:rise .2s ease}
.demo{margin-top:26px;padding-top:22px;border-top:1px solid var(--s2)}
.demo-h{font-size:10.5px;color:var(--dim);text-transform:uppercase;letter-spacing:.7px;margin-bottom:11px;font-weight:600}
.chips{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.chip{background:rgba(51,65,85,.45);border:1px solid var(--s2);border-radius:10px;padding:12px 14px;
  cursor:pointer;text-align:left;font-family:inherit;transition:.15s;-webkit-tap-highlight-color:rgba(59,130,246,.2)}
.chip:hover{border-color:var(--acc);background:rgba(59,130,246,.1);box-shadow:0 0 12px rgba(59,130,246,.15)}
.chip-role{font-size:12px;font-weight:700;color:var(--acc2);display:block}
.chip-user{font-size:10.5px;color:var(--mut);margin-top:1px}
.foot{display:flex;align-items:center;justify-content:center;gap:7px;margin-top:22px;font-size:11.5px;color:var(--dim)}
.foot svg{width:13px;height:13px;opacity:.8}
</style>
</head>
<body>
<!-- ── Pantalla de inicio (splash) ── -->
<style>
#yve-splash{position:fixed;inset:0;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;
  background:linear-gradient(180deg,#101a2e 0%,#0c1424 55%,#090e1a 100%);padding:24px;
  transition:opacity .55s ease,visibility .55s ease}
#yve-splash.hide{opacity:0;visibility:hidden;pointer-events:none}
#yve-splash .sp-logo{width:110px;height:110px;border-radius:27px;box-shadow:0 22px 60px rgba(0,0,0,.55);
  animation:spPop .6s cubic-bezier(.2,.8,.2,1)}
#yve-splash .sp-brand{margin-top:24px;font-size:31px;font-weight:800;letter-spacing:-.8px;color:#fff;
  animation:spFade .6s ease .12s both}
#yve-splash .sp-brand span{color:#60a5fa}
#yve-splash .sp-sub{margin-top:9px;font-size:13px;color:#94a3b8;animation:spFade .6s ease .22s both}
#yve-splash .sp-loader{margin-top:30px;width:32px;height:32px;border-radius:50%;
  border:3px solid rgba(148,163,184,.22);border-top-color:#3b82f6;animation:spSpin .8s linear infinite}
#yve-splash .sp-skip{position:absolute;bottom:calc(26px + env(safe-area-inset-bottom));left:50%;transform:translateX(-50%);
  background:none;border:none;color:#64748b;font-size:12.5px;cursor:pointer;font-family:inherit;padding:10px 14px;
  text-decoration:underline;-webkit-tap-highlight-color:rgba(59,130,246,.2)}
#yve-splash .sp-skip:hover{color:#94a3b8}
@keyframes spPop{from{opacity:0;transform:scale(.82)}to{opacity:1;transform:scale(1)}}
@keyframes spFade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
@keyframes spSpin{to{transform:rotate(360deg)}}
</style>
<div id="yve-splash" role="status" aria-label="Cargando Yve.01">
  <img class="sp-logo" src="/static/icons/yve-logo-192.png" alt="Yve.01">
  <div class="sp-brand">Yve<span>.01</span></div>
  <div class="sp-sub">Automatización financiera para hoteles</div>
  <div class="sp-loader"></div>
  <button class="sp-skip" onclick="yveSkipSplash()">No volver a mostrar</button>
</div>
<script>
(function(){
  var sp = document.getElementById('yve-splash');
  if (!sp) return;
  var skip = false;
  try { skip = localStorage.getItem('yve_skip_splash') === '1'; } catch(e){}
  function quitar(){ if (sp && sp.parentNode) sp.parentNode.removeChild(sp);
    var u = document.getElementById('username'); if (u) { try { u.focus(); } catch(e){} } }
  function ocultar(){ sp.classList.add('hide'); setTimeout(quitar, 600); }
  window.yveSkipSplash = function(){ try { localStorage.setItem('yve_skip_splash','1'); } catch(e){} clearTimeout(_spT); ocultar(); };
  if (skip) { quitar(); return; }
  var _spT = setTimeout(ocultar, 2600);
})();
</script>
<div class="wrap">
  <div class="login-card">
    <div class="brand"><span class="brand-dot"></span><span class="brand-name">Yve<span>.01</span></span></div>
    <div class="brand-sub">Automatización financiera para hoteles</div>

    

    <div class="heading" data-i18n="login.titulo">Inicia sesión</div>

    <label data-i18n="login.usuario">Usuario</label>
    <input id="username" placeholder="tu usuario" autocomplete="username" autofocus>

    <label data-i18n="login.password">Contraseña</label>
    <input id="password" type="password" placeholder="••••••••" autocomplete="current-password">

    <button class="btn-login" id="btn-login" onclick="doLogin()" data-i18n="login.boton">Entrar al panel</button>

    
    <div class="error" id="error"></div>

    <div style="text-align:center;margin-top:12px;font-size:11px;color:#475569">
      <div style="display:flex;justify-content:center;gap:12px;margin-bottom:8px">
        <button onclick="setLoginLang('es')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Español" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇪🇸</button>
        <button onclick="setLoginLang('en')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="English" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇬🇧</button>
        <button onclick="setLoginLang('ca')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Català" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🏴󠁥󠁳󠁣󠁴󠁿</button>
        <button onclick="setLoginLang('fr')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Français" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇫🇷</button>
        <button onclick="setLoginLang('de')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Deutsch" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇩🇪</button>
        <button onclick="setLoginLang('it')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Italiano" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇮🇹</button>
        <button onclick="setLoginLang('pt')" style="background:none;border:none;cursor:pointer;font-size:20px;opacity:.6;transition:.15s" title="Português" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=.6">🇵🇹</button>
      </div>
    </div>
    <div class="demo">
      <div class="demo-h">🏨 Clientes de ejemplo — cada cuenta solo ve SUS datos (1 clic para entrar)</div>
      <div class="chips">
        <button class="chip" style="border-color:rgba(245,158,11,.35)" onclick="quick('solmar','demo123')"><span class="chip-role" style="color:#fbbf24">☀️ Cadena Sol</span><span class="chip-user">2 hoteles · solmar</span></button>
        <button class="chip" style="border-color:rgba(168,85,247,.35)" onclick="quick('gestoria','demo123')"><span class="chip-role" style="color:#c084fc">🧾 Gestoría Nord</span><span class="chip-user">2 hoteles · gestoria</span></button>
      </div>
      <div class="demo-h" style="margin-top:16px">👥 Equipo del hotel — roles (tenant principal)</div>
      <div class="chips">
        <button class="chip" onclick="quick('admin','admin123')"><span class="chip-role">🔑 Administrador</span><span class="chip-user">admin</span></button>
        <button class="chip" onclick="quick('fc_user','hotel2024')"><span class="chip-role">💰 Financial Controller</span><span class="chip-user">fc_user</span></button>
        <button class="chip" onclick="quick('auditor','hotel2024')"><span class="chip-role">🔍 Income Auditor</span><span class="chip-user">auditor</span></button>
        <button class="chip" onclick="quick('fbmanager','hotel2024')"><span class="chip-role">🍽️ F&B Manager</span><span class="chip-user">fbmanager</span></button>
      </div>
    </div>
  </div>
  <div class="foot">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l8 4v6c0 5-3.5 8-8 10-4.5-2-8-5-8-10V6l8-4z"/></svg>
    Sesión cifrada · Acceso por roles
  </div>
</div>

<script>
function setLoginLang(lang) {
  localStorage.setItem('yve_lang', lang);
  // Translate login labels
  var labels = {
    es: {t:'Iniciar sesión', u:'USUARIO', p:'CONTRASEÑA', b:'Entrar al panel'},
    en: {t:'Log in', u:'USERNAME', p:'PASSWORD', b:'Enter panel'},
    ca: {t:'Iniciar sessió', u:'USUARI', p:'CONTRASENYA', b:'Entrar al tauler'},
    fr: {t:'Se connecter', u:'UTILISATEUR', p:'MOT DE PASSE', b:'Accéder au tableau'},
    de: {t:'Anmelden', u:'BENUTZER', p:'PASSWORT', b:'Dashboard öffnen'},
    it: {t:'Accedi', u:'UTENTE', p:'PASSWORD', b:'Accedi al pannello'},
    pt: {t:'Entrar', u:'USUÁRIO', p:'SENHA', b:'Acessar painel'},
  };
  var l = labels[lang] || labels.es;
  var title = document.querySelector('h1');
  if (title) title.textContent = l.t;
  var pu = document.querySelector('label[for="username"], label:first-of-type');
  var pp = document.querySelector('label[for="password"]');
  var btn = document.getElementById('btn-login');
  if (pu) pu.textContent = l.u;
  if (pp) pp.textContent = l.p;
  if (btn) btn.textContent = l.b;
}
function fill(u,p){
  var un = document.getElementById('username');
  var pw = document.getElementById('password');
  var btn = document.getElementById('btn-login');
  un.value=u; pw.value=p;
  // Flash green to confirm fill
  un.style.transition='border-color .2s';
  pw.style.transition='border-color .2s';
  un.style.borderColor='var(--grn)';
  pw.style.borderColor='var(--grn)';
  setTimeout(function(){ un.style.borderColor=''; pw.style.borderColor=''; }, 800);
  btn.focus();
}
function quick(u, p) {
  fill(u, p);
  doLogin();
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
      err.textContent = data.error || 'Usuario o contraseña incorrectos';
      err.classList.add('on'); btn.disabled = false; btn.textContent = 'Entrar al panel';
    btn.style.animation='shake .3s ease'; setTimeout(function(){ btn.style.animation=''; }, 400);
    }
  } catch(e) {
    err.textContent = 'Error de conexión'; err.classList.add('on');
    btn.disabled = false; btn.textContent = 'Entrar al panel';
  }
}
document.getElementById('password').addEventListener('keydown', function(e){ if(e.key==='Enter') doLogin(); });
document.getElementById('username').addEventListener('keydown', function(e){ if(e.key==='Enter') document.getElementById('password').focus(); });

// I18N Login — con caché
const _i18nCache = {};
const _i18nOriginal = {}; // textos ES originales — para restaurar al volver a español
let _i18nData = {};
let _i18nLang = localStorage.getItem('yve_lang') || 'es';

function _saveOriginals() {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const k = el.getAttribute('data-i18n');
    if (!_i18nOriginal[k]) _i18nOriginal[k] = el.textContent;
  });
}

async function loadI18n(lang) {
  _saveOriginals();
  if (lang === 'es') {
    _i18nData = {}; _i18nLang = 'es';
    applyI18n(_i18nOriginal); // restaura textos originales
    localStorage.setItem('yve_lang', 'es'); return;
  }
  if (_i18nCache[lang]) {
    _i18nData = _i18nCache[lang]; _i18nLang = lang;
    applyI18n(_i18nData); localStorage.setItem('yve_lang', lang); return;
  }
  try {
    const r = await fetch('/static/i18n/' + lang + '.json');
    const data = await r.json();
    _i18nCache[lang] = data; _i18nData = data; _i18nLang = lang;
    applyI18n(data); localStorage.setItem('yve_lang', lang);
  } catch(e) { console.warn('i18n error:', e); }
}

function t(key) { return _i18nData[key] || _i18nOriginal[key] || key; }

function applyI18n(data) {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const k = el.getAttribute('data-i18n');
    if (data[k] !== undefined) el.textContent = data[k];
  });
}

async function cambiarIdioma(lang) {
  fetch('/api/set_lang/' + lang);
  await loadI18n(lang);
  document.querySelectorAll('.lang-btn').forEach(b => {
    b.style.background = b.dataset.lang === lang ? 'rgba(59,130,246,.15)' : 'transparent';
    b.style.color = b.dataset.lang === lang ? 'var(--acc2)' : 'var(--mut)';
    b.style.fontWeight = b.dataset.lang === lang ? '700' : '400';
  });
}

loadI18n(_i18nLang);
setTimeout(() => {
  ['en','ca','fr','de','it','pt'].forEach(l => {
    if (!_i18nCache[l]) fetch('/static/i18n/' + l + '.json').then(r=>r.json()).then(d=>{_i18nCache[l]=d;}).catch(()=>{});
  });
}, 1500);

</script>
</body>
</html>"""
