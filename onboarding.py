"""
onboarding.py — Yve.01 Onboarding Wizard
5-step guided setup for new hotels
"""
from flask import Blueprint, Response, request, redirect, session, jsonify
import json, os

onboarding_bp = Blueprint('onboarding', __name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'datos-referencia', 'hotel_config.json')

def _save_config(data: dict):
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    existing = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH) as f: existing = json.load(f)
        except: pass
    existing.update(data)
    with open(CONFIG_PATH, 'w') as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)

@onboarding_bp.route('/onboarding/save', methods=['POST'])
def save_step():
    data = request.get_json(force=True, silent=True) or {}
    _save_config(data)
    return jsonify({'ok': True})

@onboarding_bp.route('/onboarding/config')
def get_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f: return jsonify(json.load(f))
    return jsonify({})

@onboarding_bp.route('/onboarding/complete', methods=['POST'])
def complete():
    _save_config({'configurado': True})
    return jsonify({'ok': True, 'redirect': '/'})

ONBOARDING_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5">
<title>Configuración inicial | Yve.01</title>
<style>
:root{--bg:#0f1117;--s1:#0f172a;--s2:#1e293b;--s3:#334155;--acc:#3b82f6;--acc2:#60a5fa;--grn:#22c55e;--ora:#f59e0b;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:Inter,system-ui,sans-serif;min-height:100vh;display:flex;flex-direction:column}
.wrap{max-width:680px;margin:0 auto;padding:40px 20px;flex:1;width:100%}
.logo{font-size:22px;font-weight:800;color:#fff;margin-bottom:40px;text-align:center}
.logo span{color:var(--acc2)}.logo small{display:block;font-size:11px;color:var(--dim);font-weight:400;margin-top:2px}

/* Progress steps */
.steps{display:flex;justify-content:center;align-items:center;gap:0;margin-bottom:44px}
.step{display:flex;flex-direction:column;align-items:center;gap:4px;flex:1;max-width:80px}
.step-num{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;border:2px solid var(--s3);color:var(--dim);background:var(--s1);transition:.3s;z-index:1}
.step.active .step-num{border-color:var(--acc);color:var(--acc2);background:rgba(59,130,246,.1)}
.step.done .step-num{border-color:var(--grn);color:var(--grn);background:rgba(34,197,94,.1)}
.step-lbl{font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.4px;text-align:center}
.step.active .step-lbl{color:var(--acc2)}
.step-line{flex:1;height:2px;background:var(--s3);margin:0 -2px;margin-bottom:16px;position:relative;top:-8px}
.step-line.done{background:var(--grn)}

/* Card */
.card{background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:32px}
.card h2{font-size:20px;font-weight:700;margin-bottom:6px}
.card .sub{font-size:13px;color:var(--mut);margin-bottom:28px;line-height:1.5}

/* Form */
.fg{margin-bottom:18px}
.fg label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);display:block;margin-bottom:6px}
.fg input,.fg select,.fg textarea{width:100%;background:var(--bg);border:1px solid var(--s3);color:var(--tx);padding:11px 14px;border-radius:10px;font-size:13px;outline:none;transition:.2s;font-family:inherit}
.fg input:focus,.fg select:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(59,130,246,.1)}
.fg .hint{font-size:11px;color:var(--dim);margin-top:5px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}

/* Checkboxes */
.check-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-top:8px}
.check-item{background:var(--bg);border:1px solid var(--s3);border-radius:10px;padding:10px 12px;cursor:pointer;transition:.2s;display:flex;align-items:center;gap:8px;font-size:12px}
.check-item:hover{border-color:var(--acc)}
.check-item.selected{border-color:var(--acc);background:rgba(59,130,246,.08);color:var(--acc2)}
.check-item input{display:none}

/* Buttons */
.btn-next{background:var(--acc);border:none;color:#fff;padding:13px 28px;border-radius:12px;font-size:14px;font-weight:700;cursor:pointer;width:100%;margin-top:24px;transition:.2s}
.btn-next:hover{background:#2563eb}
.btn-prev{background:none;border:1px solid var(--s3);color:var(--mut);padding:13px 20px;border-radius:12px;font-size:13px;cursor:pointer;margin-top:12px;margin-right:10px;transition:.2s}
.btn-prev:hover{border-color:var(--acc2);color:var(--acc2)}
.btns{display:flex;flex-direction:column}

/* Step 5 — complete */
.success{text-align:center;padding:16px 0}
.success .icon{font-size:56px;margin-bottom:16px}
.success h2{font-size:24px;margin-bottom:10px}
.success p{color:var(--mut);margin-bottom:24px;line-height:1.6}
.feature-list{text-align:left;background:var(--bg);border-radius:12px;padding:16px 20px;margin-bottom:24px}
.feature-list li{list-style:none;padding:5px 0;font-size:13px;color:var(--mut)}
.feature-list li::before{content:"✓ ";color:var(--grn);font-weight:700}

@media(max-width:600px){.row2,.row3{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrap">
  <div class="logo">Yve<span>.01</span><small>Configuración inicial del hotel</small></div>

  <!-- Step progress -->
  <div class="steps" id="steps-bar">
    <div class="step active" id="s1"><div class="step-num">1</div><div class="step-lbl">Hotel</div></div>
    <div class="step-line" id="sl1"></div>
    <div class="step" id="s2"><div class="step-num">2</div><div class="step-lbl">OTAs</div></div>
    <div class="step-line" id="sl2"></div>
    <div class="step" id="s3"><div class="step-num">3</div><div class="step-lbl">Finanzas</div></div>
    <div class="step-line" id="sl3"></div>
    <div class="step" id="s4"><div class="step-num">4</div><div class="step-lbl">Equipo</div></div>
    <div class="step-line" id="sl4"></div>
    <div class="step" id="s5"><div class="step-num">5</div><div class="step-lbl">Listo</div></div>
  </div>

  <!-- Step 1: Hotel info -->
  <div class="card" id="step-1">
    <h2>🏨 Datos del hotel</h2>
    <p class="sub">Información básica de tu propiedad. Puedes cambiarla después desde Ajustes.</p>
    <div class="fg"><label>Nombre del hotel *</label><input id="f-nombre" placeholder="Ej: Hotel Calipolis Sitges" autocomplete="organization"></div>
    <div class="row2">
      <div class="fg"><label>Ciudad *</label><input id="f-ciudad" placeholder="Barcelona" autocomplete="address-level2"></div>
      <div class="fg"><label>País</label><select id="f-pais"><option value="ES">España</option><option value="PT">Portugal</option><option value="FR">Francia</option><option value="IT">Italia</option><option value="DE">Alemania</option></select></div>
    </div>
    <div class="row3">
      <div class="fg"><label>Habitaciones *</label><input id="f-hab" type="number" min="1" placeholder="142"></div>
      <div class="fg"><label>Estrellas</label><select id="f-estrellas"><option>3</option><option selected>4</option><option>5</option></select></div>
      <div class="fg"><label>NIF/CIF</label><input id="f-nif" placeholder="A12345678"></div>
    </div>
    <div class="fg"><label>Email de contacto finanzas *</label><input id="f-email" type="email" placeholder="finanzas@tuhotel.com" autocomplete="email"></div>
    <div class="btns">
      <button class="btn-next" onclick="nextStep(1)">Siguiente → OTAs</button>
    </div>
  </div>

  <!-- Step 2: OTAs -->
  <div class="card" style="display:none" id="step-2">
    <h2>💳 OTAs activas</h2>
    <p class="sub">Selecciona las plataformas de distribución que usas. Yve verificará automáticamente las comisiones de cada factura.</p>
    <div class="fg">
      <label>Canales de distribución</label>
      <div class="check-grid" id="otas-grid">
        <div class="check-item selected" onclick="toggleOTA(this,'Booking.com')">🌐 Booking.com</div>
        <div class="check-item selected" onclick="toggleOTA(this,'Expedia')">✈️ Expedia</div>
        <div class="check-item" onclick="toggleOTA(this,'Hotusa')">🏨 Hotusa</div>
        <div class="check-item" onclick="toggleOTA(this,'HotelBeds')">🛏 HotelBeds</div>
        <div class="check-item" onclick="toggleOTA(this,'Airbnb')">🏠 Airbnb</div>
        <div class="check-item" onclick="toggleOTA(this,'Agoda')">🌏 Agoda</div>
        <div class="check-item" onclick="toggleOTA(this,'TripAdvisor')">⭐ TripAdvisor</div>
        <div class="check-item" onclick="toggleOTA(this,'Directo')">🏷 Directo</div>
      </div>
    </div>
    <div class="fg">
      <label>Comisión Booking.com (%)</label>
      <input id="f-comm-booking" type="number" step="0.1" value="15" min="0" max="40">
      <div class="hint">La comisión pactada con Booking. Yve alertará si la factura difiere.</div>
    </div>
    <div class="fg">
      <label>Comisión Expedia (%)</label>
      <input id="f-comm-expedia" type="number" step="0.1" value="18" min="0" max="40">
    </div>
    <div class="btns">
      <button class="btn-next" onclick="nextStep(2)">Siguiente → Finanzas</button>
      <button class="btn-prev" onclick="prevStep(2)">← Atrás</button>
    </div>
  </div>

  <!-- Step 3: Finance config -->
  <div class="card" style="display:none" id="step-3">
    <h2>💰 Configuración financiera</h2>
    <p class="sub">Parámetros contables del hotel. Se usan para validar facturas y generar asientos Oracle.</p>
    <div class="row2">
      <div class="fg"><label>IVA estándar (%)</label><select id="f-iva"><option value="10" selected>10% (hotelero)</option><option value="21">21% (general)</option><option value="0">0% (exento)</option></select></div>
      <div class="fg"><label>Días de pago proveedores</label><input id="f-dias-pago" type="number" value="30" min="0" max="120"></div>
    </div>
    <div class="fg">
      <label>Divisa principal</label>
      <select id="f-moneda"><option value="EUR" selected>EUR — Euro</option><option value="USD">USD — Dólar</option><option value="GBP">GBP — Libra</option></select>
    </div>
    <div class="fg">
      <label>Sistema ERP contable</label>
      <select id="f-erp">
        <option value="oracle">Oracle Fusion</option>
        <option value="sap">SAP</option>
        <option value="sage">Sage</option>
        <option value="a3">A3 ERP</option>
        <option value="manual">Sin ERP / Manual</option>
      </select>
      <div class="hint">Yve puede exportar asientos en el formato de tu ERP.</div>
    </div>
    <div class="fg">
      <label>¿El hotel tiene operaciones internacionales?</label>
      <div class="check-grid">
        <div class="check-item selected" id="intl-yes" onclick="toggleIntl(true)">✅ Sí (necesito DI)</div>
        <div class="check-item" id="intl-no" onclick="toggleIntl(false)">❌ No</div>
      </div>
      <div class="hint">Afecta al módulo de certificados de Doble Imposición (AR).</div>
    </div>
    <div class="btns">
      <button class="btn-next" onclick="nextStep(3)">Siguiente → Equipo</button>
      <button class="btn-prev" onclick="prevStep(3)">← Atrás</button>
    </div>
  </div>

  <!-- Step 4: Team -->
  <div class="card" style="display:none" id="step-4">
    <h2>👥 Equipo y notificaciones</h2>
    <p class="sub">Define quién recibe alertas. Puedes añadir más usuarios desde el panel de administración.</p>
    <div class="fg">
      <label>Financial Controller (email)</label>
      <input id="f-fc-email" type="email" placeholder="fc@tuhotel.com" autocomplete="email">
      <div class="hint">Recibirá alertas de discrepancias AR, facturas pendientes y Out of Balance.</div>
    </div>
    <div class="fg">
      <label>Jefe de departamento AP (email)</label>
      <input id="f-ap-email" type="email" placeholder="compras@tuhotel.com">
      <div class="hint">Recibirá solicitudes de aprobación de facturas AP.</div>
    </div>
    <div class="fg">
      <label>¿Cómo quieres recibir las alertas?</label>
      <div class="check-grid">
        <div class="check-item selected" onclick="this.classList.toggle('selected')" data-ch="email">📧 Email</div>
        <div class="check-item" onclick="this.classList.toggle('selected')" data-ch="telegram">✈️ Telegram</div>
        <div class="check-item" onclick="this.classList.toggle('selected')" data-ch="push">🔔 Push (app)</div>
      </div>
    </div>
    <div class="fg">
      <label>¿Cuándo enviar alertas?</label>
      <select id="f-alert-time">
        <option value="realtime">En tiempo real (cada cambio)</option>
        <option value="daily">Resumen diario (09:00)</option>
        <option value="manual">Solo manual</option>
      </select>
    </div>
    <div class="btns">
      <button class="btn-next" onclick="nextStep(4)">Finalizar configuración →</button>
      <button class="btn-prev" onclick="prevStep(4)">← Atrás</button>
    </div>
  </div>

  <!-- Step 5: Complete -->
  <div class="card" style="display:none" id="step-5">
    <div class="success">
      <div class="icon">🎉</div>
      <h2>¡Todo listo, <span id="hotel-name-finish">tu hotel</span>!</h2>
      <p>Yve.01 está configurado y listo para automatizar tus finanzas hoteleras.</p>
      <ul class="feature-list">
        <li>Verificación automática de comisiones OTA</li>
        <li>3-way matching para facturas AP</li>
        <li>DRR con detección de Out of Balance</li>
        <li>Ciclo AR Real para grupos corporativos</li>
        <li>Dashboard multi-hotel con Calipolis</li>
      </ul>
      <button class="btn-next" onclick="goToDashboard()">Ir al dashboard →</button>
    </div>
  </div>
</div>

<script>
var step = 1;
var config = {};
var selectedOTAs = ['Booking.com', 'Expedia'];
var intl = true;

function toggleOTA(el, name) {
  el.classList.toggle('selected');
  if (el.classList.contains('selected')) {
    if (!selectedOTAs.includes(name)) selectedOTAs.push(name);
  } else {
    selectedOTAs = selectedOTAs.filter(o => o !== name);
  }
}

function toggleIntl(yes) {
  intl = yes;
  document.getElementById('intl-yes').classList.toggle('selected', yes);
  document.getElementById('intl-no').classList.toggle('selected', !yes);
}

function validateStep(n) {
  if (n === 1) {
    var nombre = document.getElementById('f-nombre').value.trim();
    var email  = document.getElementById('f-email').value.trim();
    var hab    = document.getElementById('f-hab').value;
    if (!nombre) { alert('El nombre del hotel es obligatorio'); return false; }
    if (!email || !email.includes('@')) { alert('Email de contacto inválido'); return false; }
    if (!hab || parseInt(hab) < 1) { alert('Introduce el número de habitaciones'); return false; }
  }
  return true;
}

function collectStep(n) {
  if (n === 1) {
    config.hotel_nombre    = document.getElementById('f-nombre').value.trim();
    config.hotel_ciudad    = document.getElementById('f-ciudad').value.trim();
    config.hotel_pais      = document.getElementById('f-pais').value;
    config.hotel_habitaciones = parseInt(document.getElementById('f-hab').value) || 0;
    config.hotel_estrellas = parseInt(document.getElementById('f-estrellas').value);
    config.hotel_nif       = document.getElementById('f-nif').value.trim();
    config.hotel_email     = document.getElementById('f-email').value.trim();
    config.hotel_tag       = config.hotel_nombre.split(' ').slice(-1)[0];
  } else if (n === 2) {
    config.otas_activas    = selectedOTAs;
    config.comision_booking = parseFloat(document.getElementById('f-comm-booking').value) || 15;
    config.comision_expedia = parseFloat(document.getElementById('f-comm-expedia').value) || 18;
  } else if (n === 3) {
    config.iva_tipo       = parseInt(document.getElementById('f-iva').value);
    config.dias_pago_proveedor = parseInt(document.getElementById('f-dias-pago').value) || 30;
    config.moneda         = document.getElementById('f-moneda').value;
    config.erp_sistema    = document.getElementById('f-erp').value;
    config.opera_internacionalmente = intl;
  } else if (n === 4) {
    config.fc_email       = document.getElementById('f-fc-email').value.trim();
    config.ap_email       = document.getElementById('f-ap-email').value.trim();
    config.alert_time     = document.getElementById('f-alert-time').value;
    config.canales_alerta = [...document.querySelectorAll('.check-item.selected[data-ch]')]
                            .map(el => el.dataset.ch);
  }
}

async function saveToServer(data) {
  try {
    await fetch('/onboarding/save', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(data)
    });
  } catch(e) { console.warn('Save error:', e); }
}

async function nextStep(current) {
  if (!validateStep(current)) return;
  collectStep(current);
  await saveToServer(config);
  
  // Update progress
  document.getElementById('s' + current).classList.remove('active');
  document.getElementById('s' + current).classList.add('done');
  document.getElementById('s' + current).querySelector('.step-num').textContent = '✓';
  if (current < 5) document.getElementById('sl' + current).classList.add('done');
  
  step = current + 1;
  document.getElementById('step-' + current).style.display = 'none';
  document.getElementById('step-' + step).style.display = 'block';
  document.getElementById('s' + step).classList.add('active');
  window.scrollTo(0, 0);
  
  if (step === 5) {
    var nombre = config.hotel_nombre || 'tu hotel';
    document.getElementById('hotel-name-finish').textContent = nombre.split(' ').slice(-1)[0];
    await saveToServer({configurado: true});
  }
}

function prevStep(current) {
  document.getElementById('step-' + current).style.display = 'none';
  step = current - 1;
  document.getElementById('step-' + step).style.display = 'block';
  document.getElementById('s' + current).classList.remove('active');
  document.getElementById('s' + step).classList.remove('done');
  document.getElementById('s' + step).querySelector('.step-num').textContent = step;
  document.getElementById('s' + step).classList.add('active');
  if (current <= 5) {
    var sl = document.getElementById('sl' + (current-1));
    if (sl) sl.classList.remove('done');
  }
  window.scrollTo(0, 0);
}

async function goToDashboard() {
  await fetch('/onboarding/complete', {method:'POST'});
  window.location.href = '/';
}
</script>
</body>
</html>"""

@onboarding_bp.route('/onboarding')
@onboarding_bp.route('/onboarding/')
def onboarding_page():
    return Response(ONBOARDING_HTML, mimetype='text/html')
