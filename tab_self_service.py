"""
tab_self_service.py — Yve.01 Multi-Hotel Self-Service Registration
Allows hotels to register themselves without admin intervention.
Route: /unirse  (también /join)
"""
import os, json, smtplib
from pathlib import Path
from email.mime.text import MIMEText
from flask import Blueprint, request, jsonify, Response

self_service_bp = Blueprint('self_service', __name__)
BASE_DIR  = Path(__file__).parent
DATOS     = BASE_DIR / 'datos-referencia'
LEADS_FILE = DATOS / 'leads_hoteles.json'

def _load_leads():
    if LEADS_FILE.exists():
        return json.loads(LEADS_FILE.read_text())
    return []

def _save_leads(leads):
    DATOS.mkdir(exist_ok=True)
    LEADS_FILE.write_text(json.dumps(leads, indent=2, ensure_ascii=False))

def _notify_admin(lead):
    """Send email notification to Yve admin about new hotel registration."""
    try:
        smtp_server = os.environ.get('SMTP_SERVER','')
        smtp_user   = os.environ.get('SMTP_USER','')
        smtp_pass   = os.environ.get('SMTP_PASSWORD','')
        if not all([smtp_server, smtp_user, smtp_pass]): return
        msg = MIMEText(f"""Nuevo hotel registrado en Yve.01\n
Hotel:        {lead['hotel']}
Ciudad:       {lead['ciudad']}
Habitaciones: {lead['habitaciones']}
Contacto:     {lead['nombre']} — {lead['email']}
Teléfono:     {lead.get('telefono','')}
Plan sugerido:{lead['plan_sugerido']}
Grupo:        {lead.get('grupo','')}
Notas:        {lead.get('notas','')}

Panel: https://yve01.onrender.com/admin
""")
        msg['Subject'] = f"[Yve.01] Nuevo lead: {lead['hotel']}"
        msg['From'] = smtp_user
        msg['To']   = smtp_user  # admin = same email
        with smtplib.SMTP(smtp_server, int(os.environ.get('SMTP_PORT',587))) as s:
            s.starttls()
            s.login(smtp_user, smtp_pass)
            s.sendmail(smtp_user, smtp_user, msg.as_string())
    except Exception as e:
        print(f"[SELF_SERVICE] Email warning: {e}")


@self_service_bp.route('/unirse')
@self_service_bp.route('/join')
def join_page():
    return Response(JOIN_HTML, mimetype='text/html')


@self_service_bp.route('/api/solicitar_demo', methods=['POST'])
def api_solicitar_demo():
    d        = request.get_json(silent=True) or {}
    hotel    = (d.get('hotel') or '').strip()
    nombre   = (d.get('nombre') or '').strip()
    email    = (d.get('email') or '').strip().lower()
    ciudad   = (d.get('ciudad') or '').strip()
    habitaciones = d.get('habitaciones') or 0
    telefono = (d.get('telefono') or '').strip()
    grupo    = (d.get('grupo') or '').strip()
    notas    = (d.get('notas') or '').strip()

    if not hotel or not nombre or not email:
        return jsonify({'ok': False, 'error': 'Faltan campos obligatorios'}), 400

    # Auto-select plan
    hab_int = int(habitaciones) if str(habitaciones).isdigit() else 0
    plan = 'multi' if grupo else ('pro' if hab_int > 150 else 'starter')

    lead = {
        'hotel': hotel, 'nombre': nombre, 'email': email,
        'ciudad': ciudad, 'habitaciones': hab_int,
        'telefono': telefono, 'grupo': grupo, 'notas': notas,
        'plan_sugerido': plan, 'estado': 'nuevo',
        'fecha': __import__('datetime').datetime.now().isoformat(),
    }

    leads = _load_leads()
    # Avoid duplicate by email
    if any(l['email'] == email for l in leads):
        return jsonify({'ok': False, 'error': 'Ya tienes una solicitud registrada con este email.'}), 400

    leads.append(lead)
    _save_leads(leads)
    _notify_admin(lead)

    # Redirect to checkout
    return jsonify({'ok': True, 'plan': plan, 'checkout_url': f'/checkout/{plan}'})


@self_service_bp.route('/api/admin/leads')
def api_admin_leads():
    from flask_login import current_user
    if not current_user.is_authenticated or current_user.rol != 'admin':
        return jsonify({'error': 'Unauthorized'}), 403
    return jsonify({'ok': True, 'leads': _load_leads()})


JOIN_HTML = """<!DOCTYPE html>
<html lang="es"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Únete a Yve.01 — Software de finanzas para hoteles</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700;800;900&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:#0f172a;color:#f1f5f9;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px}
.card{background:#1e293b;border:1px solid #334155;border-radius:20px;padding:40px 36px;width:100%;max-width:480px}
.logo{display:flex;align-items:center;gap:10px;margin-bottom:28px}
.logo-dot{width:10px;height:10px;border-radius:50%;background:#3b82f6}
.logo-name{font-size:20px;font-weight:800;color:#f1f5f9}
.logo-name span{color:#3b82f6}
h1{font-size:22px;font-weight:800;margin-bottom:8px}
p.sub{font-size:14px;color:#94a3b8;margin-bottom:28px;line-height:1.5}
label{display:block;font-size:12px;font-weight:600;color:#94a3b8;margin-bottom:5px;margin-top:16px;letter-spacing:.04em;text-transform:uppercase}
input,textarea,select{width:100%;background:#0f172a;border:1px solid #334155;color:#f1f5f9;border-radius:10px;padding:11px 14px;font-size:14px;font-family:inherit;outline:none;transition:.15s}
input:focus,textarea:focus,select:focus{border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.12)}
textarea{resize:vertical;min-height:70px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.btn{width:100%;background:linear-gradient(135deg,#3b82f6,#7c3aed);color:#fff;border:none;border-radius:12px;padding:14px;font-size:15px;font-weight:700;cursor:pointer;margin-top:24px;transition:.15s}
.btn:hover{opacity:.9;transform:translateY(-1px)}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
.err-msg{display:none;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);border-radius:8px;padding:10px 14px;font-size:13px;color:#fca5a5;margin-top:12px}
.err-msg.on{display:block}
.success{text-align:center;padding:20px 0}
.success .ic{font-size:48px;margin-bottom:16px}
.success h2{font-size:22px;font-weight:800;margin-bottom:8px}
.success p{color:#94a3b8;font-size:14px;line-height:1.5;margin-bottom:20px}
.success a{display:inline-block;background:linear-gradient(135deg,#3b82f6,#7c3aed);color:#fff;text-decoration:none;border-radius:10px;padding:12px 28px;font-weight:700}
.trust{display:flex;gap:16px;margin-top:20px;flex-wrap:wrap}
.trust-item{display:flex;align-items:center;gap:6px;font-size:12px;color:#64748b}
</style>
</head><body>
<div class="card">
  <div class="logo"><div class="logo-dot"></div><div class="logo-name">Yve<span>.01</span></div></div>
  <div id="form-wrap">
    <h1>Empieza en 15 minutos</h1>
    <p class="sub">Sin instalaciones. Sin consultores. Sin contrato de permanencia.<br>14 días de prueba gratuita, cancela cuando quieras.</p>
    <label>Nombre del hotel *</label>
    <input id="f-hotel" type="text" placeholder="Hotel Costa Dorada Sitges">
    <div class="row2">
      <div>
        <label>Tu nombre *</label>
        <input id="f-nombre" type="text" placeholder="Ana García">
      </div>
      <div>
        <label>Ciudad</label>
        <input id="f-ciudad" type="text" placeholder="Barcelona">
      </div>
    </div>
    <label>Email de trabajo *</label>
    <input id="f-email" type="email" placeholder="ana@hotelcostadorada.com">
    <div class="row2">
      <div>
        <label>Teléfono</label>
        <input id="f-tel" type="tel" placeholder="+34 600 000 000">
      </div>
      <div>
        <label>Habitaciones</label>
        <input id="f-rooms" type="number" placeholder="120" min="1" max="2000">
      </div>
    </div>
    <label>Grupo hotelero (si aplica)</label>
    <input id="f-grupo" type="text" placeholder="Grupo Calipolis, NH, etc.">
    <label>¿En qué área necesitas más ayuda?</label>
    <textarea id="f-notas" placeholder="AP (facturas proveedores), AR (comisiones OTAs), DRR, conciliación bancaria..."></textarea>
    <button class="btn" id="btn-submit" onclick="doJoin()">Solicitar acceso gratuito →</button>
    <div class="err-msg" id="err"></div>
    <div class="trust">
      <div class="trust-item">✓ Sin tarjeta hasta confirmar</div>
      <div class="trust-item">✓ 14 días gratis</div>
      <div class="trust-item">✓ Cancela cuando quieras</div>
    </div>
  </div>
</div>
<script>
async function doJoin(){
  const hotel  = document.getElementById('f-hotel').value.trim();
  const nombre = document.getElementById('f-nombre').value.trim();
  const email  = document.getElementById('f-email').value.trim();
  if(!hotel||!nombre||!email){
    const e=document.getElementById('err');e.textContent='⚠️ Completa los campos obligatorios.';e.classList.add('on');return;
  }
  const btn=document.getElementById('btn-submit');
  btn.disabled=true;btn.textContent='Enviando...';
  try{
    const r=await fetch('/api/solicitar_demo',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        hotel,nombre,email,
        ciudad:document.getElementById('f-ciudad').value.trim(),
        habitaciones:document.getElementById('f-rooms').value,
        telefono:document.getElementById('f-tel').value.trim(),
        grupo:document.getElementById('f-grupo').value.trim(),
        notas:document.getElementById('f-notas').value.trim()
      })});
    const d=await r.json();
    if(d.ok){
      document.getElementById('form-wrap').innerHTML=
        '<div class="success"><div class="ic">🎉</div>'
        +'<h2>¡Solicitud recibida!</h2>'
        +'<p>Nos pondremos en contacto en menos de 24h para configurar tu cuenta.<br>'
        +'Si quieres empezar ahora mismo, puedes activar tu plan directamente:</p>'
        +'<a href="'+d.checkout_url+'">Activar plan ahora →</a>'
        +'<p style="font-size:12px;color:#475569;margin-top:16px">O espera a que te contactemos sin coste.</p>'
        +'</div>';
    }else{
      const e=document.getElementById('err');e.textContent='⚠️ '+(d.error||'Error');e.classList.add('on');
      btn.disabled=false;btn.textContent='Solicitar acceso gratuito →';
    }
  }catch(ex){
    const e=document.getElementById('err');e.textContent='⚠️ Error de conexión.';e.classList.add('on');
    btn.disabled=false;btn.textContent='Solicitar acceso gratuito →';
  }
}
</script>
</body></html>
"""
