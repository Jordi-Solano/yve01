"""
signup.py — Yve Self-Service Registration
Ruta: /signup (pública)
Crea cuenta financial_controller + entrada de hotel
"""
import os, json
from pathlib import Path
from flask import Blueprint, request, redirect, Response, jsonify
from auth import crear_usuario, _find
from flask_login import login_user

signup_bp = Blueprint('signup', __name__)
BASE = Path(__file__).parent
HOTELES_PATH = BASE / "datos-referencia" / "hoteles.json"


@signup_bp.route('/signup')
def signup_page():
    return Response(SIGNUP_HTML, mimetype='text/html')


@signup_bp.route('/api/signup', methods=['POST'])
def api_signup():
    d = request.get_json(silent=True) or {}
    hotel    = (d.get('hotel') or '').strip()
    nombre   = (d.get('nombre') or '').strip()
    email    = (d.get('email') or '').strip().lower()
    password = d.get('password') or ''
    rooms    = d.get('habitaciones') or 0
    grupo    = (d.get('grupo') or '').strip()

    if not hotel or not nombre or not email or not password:
        return jsonify({'ok': False, 'error': 'Faltan campos obligatorios'}), 400
    if len(password) < 6:
        return jsonify({'ok': False, 'error': 'La contraseña debe tener al menos 6 caracteres'}), 400

    # Username = email local part, evita colisiones
    username = email.split('@')[0]
    base_username = username
    i = 1
    while _find(username):
        username = f"{base_username}{i}"
        i += 1

    # Crear usuario como financial_controller
    res = crear_usuario(username, password, nombre, email, 'financial_controller')
    if res is not True:
        return jsonify({'ok': False, 'error': res}), 400

    # Registrar hotel en hoteles.json (best-effort)
    try:
        hoteles = json.loads(HOTELES_PATH.read_text()) if HOTELES_PATH.exists() else []
        new_id = f"H{len(hoteles)+1:03d}"
        hoteles.append({
            'id': new_id,
            'nombre': hotel,
            'categoria': '4★',
            'habitaciones': int(rooms) if str(rooms).isdigit() else 0,
            'ciudad': '',
            'grupo': grupo or hotel,
            'contacto': email,
            'modulos': ['AP', 'AR', 'DRR', 'Banco'],
            'owner_username': username,
        })
        HOTELES_PATH.write_text(json.dumps(hoteles, indent=2, ensure_ascii=False))
    except Exception as e:
        print(f"[SIGNUP] hotel registration warning: {e}")

    # Auto-select plan based on room count
    rooms_int = int(rooms) if str(rooms).isdigit() else 0
    plan = 'multi' if (grupo and grupo.strip().lower() not in ('', hotel.lower()))            else ('pro' if rooms_int > 150 else 'starter')
    return jsonify({'ok': True, 'username': username,
                    'plan': plan, 'checkout_url': f'/checkout/{plan}'})


SIGNUP_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Crea tu cuenta — Yve.01</title>
<meta property="og:type" content="website">
<meta property="og:url" content="https://yve01.onrender.com/signup">
<meta property="og:title" content="Crear cuenta gratis | Yve.01">
<meta property="og:description" content="Empieza tu prueba gratuita de 14 días. Setup en 15 minutos. Sin tarjeta de crédito.">
<link rel="canonical" href="https://yve01.onrender.com/signup">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--red:#ef4444;--grn:#22c55e}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;position:relative;overflow-x:hidden;-webkit-font-smoothing:antialiased}
body::before{content:"";position:fixed;inset:0;z-index:0;pointer-events:none;background:radial-gradient(900px 500px at 80% -15%,rgba(59,130,246,.14),transparent 60%),radial-gradient(700px 400px at -8% 110%,rgba(139,92,246,.1),transparent 55%)}
.wrap{position:relative;z-index:1;width:100%;max-width:480px}
.card{background:linear-gradient(160deg,rgba(30,41,59,.94),rgba(15,23,42,.96));border:1px solid var(--s2);border-radius:20px;padding:38px 34px;box-shadow:0 24px 70px rgba(0,0,0,.5);animation:rise .4s cubic-bezier(.2,.8,.2,1)}
@keyframes rise{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:none}}
.brand{display:flex;align-items:center;gap:9px;margin-bottom:4px}
.brand .dot{width:10px;height:10px;border-radius:50%;background:var(--acc);box-shadow:0 0 12px var(--acc)}
.brand .name{font-size:24px;font-weight:800;letter-spacing:-.5px}
.brand .name span{color:var(--acc2)}
.sub{font-size:13px;color:var(--mut);margin-bottom:26px}
label{display:block;font-size:10.5px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px;margin-top:16px;font-weight:600}
label:first-of-type{margin-top:0}
input{width:100%;background:rgba(15,23,42,.7);border:1px solid var(--s2);color:var(--tx);border-radius:11px;padding:12px 14px;font-size:14.5px;outline:none;font-family:inherit;transition:.15s}
input:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(59,130,246,.15)}
input.err{border-color:var(--red)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
@media(max-width:420px){.row2{grid-template-columns:1fr}}
.btn{width:100%;margin-top:24px;padding:14px;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;border:none;border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;box-shadow:0 8px 28px rgba(59,130,246,.4);transition:.15s}
.btn:hover{box-shadow:0 12px 36px rgba(59,130,246,.6);transform:translateY(-1px)}
.btn:disabled{opacity:.6;cursor:not-allowed;transform:none}
.err-msg{display:none;margin-top:14px;padding:11px 14px;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.28);border-radius:10px;color:#fca5a5;font-size:13px;text-align:center}
.err-msg.on{display:block}
.foot{text-align:center;margin-top:22px;font-size:13px;color:var(--dim)}
.foot a{color:var(--acc2);text-decoration:none}
.trial{background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);border-radius:10px;padding:11px 14px;font-size:12.5px;color:#86efac;margin-bottom:22px;text-align:center}
.success{text-align:center;padding:30px 10px}
.success .ic{width:60px;height:60px;background:rgba(34,197,94,.15);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 18px}
.success h2{font-size:22px;font-weight:800;margin-bottom:10px}
.success p{font-size:14px;color:var(--mut);line-height:1.7;margin-bottom:24px}
</style>
</head>
<body>
<div class="wrap">
  <div class="card" id="card">
    <div class="brand"><div class="dot"></div><div class="name">Yve<span>.01</span></div></div>
    <div class="sub">Crea tu cuenta y empieza a automatizar</div>
    <div class="trial">✓ 14 días gratis · Sin tarjeta de crédito · Setup en 15 min</div>

    <label>Nombre del hotel *</label>
    <input id="f-hotel" placeholder="Hotel Example Barcelona" autocomplete="organization">
    <div class="row2">
      <div>
        <label>Habitaciones</label>
        <input id="f-rooms" type="number" placeholder="120" min="1">
      </div>
      <div>
        <label>Grupo (opcional)</label>
        <input id="f-grupo" placeholder="Si es cadena">
      </div>
    </div>

    <label>Tu nombre *</label>
    <input id="f-nombre" placeholder="Nombre y apellidos" autocomplete="name">

    <label>Email *</label>
    <input id="f-email" type="email" placeholder="tu@hotel.com" autocomplete="email">

    <label>Contraseña *</label>
    <input id="f-pass" type="password" placeholder="Mínimo 6 caracteres" autocomplete="new-password">

    <button class="btn" id="btn-signup" onclick="doSignup()">Crear cuenta gratis →</button>
    <div class="err-msg" id="err"></div>
    <div class="foot">¿Ya tienes cuenta? <a href="/login">Inicia sesión</a></div>
  </div>
</div>
<script>
function showErr(m){const e=document.getElementById('err');e.textContent='⚠️ '+m;e.classList.add('on');}
async function doSignup(){
  const hotel=document.getElementById('f-hotel').value.trim();
  const nombre=document.getElementById('f-nombre').value.trim();
  const email=document.getElementById('f-email').value.trim();
  const pass=document.getElementById('f-pass').value;
  const rooms=document.getElementById('f-rooms').value;
  const grupo=document.getElementById('f-grupo').value.trim();
  document.querySelectorAll('.err').forEach(e=>e.classList.remove('err'));
  document.getElementById('err').classList.remove('on');
  let bad=false;
  [['f-hotel',hotel],['f-nombre',nombre],['f-email',email],['f-pass',pass]].forEach(([id,val])=>{
    if(!val){document.getElementById(id).classList.add('err');bad=true;}
  });
  if(bad){showErr('Completa todos los campos obligatorios.');return;}
  if(pass.length<6){document.getElementById('f-pass').classList.add('err');showErr('La contraseña debe tener al menos 6 caracteres.');return;}
  const btn=document.getElementById('btn-signup');
  btn.disabled=true;btn.textContent='Creando cuenta...';
  try{
    const r=await fetch('/api/signup',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({hotel,nombre,email,password:pass,habitaciones:rooms,grupo})});
    const res=await r.json();
    if(res.ok){
      document.getElementById('card').innerHTML=
        '<div class="success"><div class="ic">✓</div>'
        +'<h2>¡Cuenta creada!</h2>'
        +'<p>Tu usuario es <b style="color:var(--acc2)">'+res.username+'</b>.<br>'
        +'Redirigiendo al pago...</p>'
        +'<p style="font-size:12px;color:#64748b;margin-top:8px">Si no quieres pagar ahora, '
        +'<a href="/login" style="color:var(--acc2)">inicia sesión</a> con 14 días de prueba.</p></div>';
      setTimeout(function(){ window.location.href = res.checkout_url || '/login'; }, 2000);
    }else{
      btn.disabled=false;btn.textContent='Crear cuenta gratis →';
      showErr(res.error||'Error al crear la cuenta.');
    }
  }catch(e){
    btn.disabled=false;btn.textContent='Crear cuenta gratis →';
    showErr('Error de conexión. Inténtalo de nuevo.');
  }
}
document.getElementById('f-pass').addEventListener('keydown',e=>{if(e.key==='Enter')doSignup();});
</script>
</body>
</html>"""
