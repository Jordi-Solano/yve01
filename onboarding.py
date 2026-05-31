"""
onboarding.py — Yve.01
Formulario multi-paso para configurar un hotel nuevo.
Ejecutar: python onboarding.py
Abre en: http://localhost:5003
"""

import os, json
from flask import Flask, jsonify, request
import pandas as pd

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
REFERENCIA_DIR = os.path.join(BASE_DIR, "datos-referencia")
CONFIG_PATH    = os.path.join(REFERENCIA_DIR, "hotel_config.json")

app = Flask(__name__)

# ── API ───────────────────────────────────────────────────────────────────

@app.route("/api/config", methods=["GET"])
def get_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return jsonify(json.load(f))
    return jsonify(None)


@app.route("/api/save", methods=["POST"])
def save_config():
    data = request.get_json(force=True)
    if not data:
        return jsonify({"ok": False, "error": "No data"}), 400

    os.makedirs(REFERENCIA_DIR, exist_ok=True)

    # 1. Guardar hotel_config.json
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 2. Actualizar comisiones_pactadas.xlsx
    otas = data.get("otas", [])
    if otas:
        rows = []
        for o in otas:
            rows.append({
                "OTA": o.get("nombre", ""),
                "Porcentaje_Comision": float(o.get("porcentaje", 0)),
                "Mercado": o.get("mercado", "Internacional"),
            })
        df = pd.DataFrame(rows)
        df.to_excel(os.path.join(REFERENCIA_DIR, "comisiones_pactadas.xlsx"), index=False)

    # 3. Actualizar proveedores.xlsx
    proveedores = data.get("proveedores", [])
    if proveedores:
        rows = []
        for p in proveedores:
            rows.append({
                "nombre_proveedor": p.get("nombre", ""),
                "tipo": p.get("tipo", "OTRAS"),
                "cuenta_contable": p.get("cuenta", ""),
                "email_contacto": p.get("email", ""),
                "porcentaje_iva_habitual": float(p.get("iva", 21)),
            })
        df = pd.DataFrame(rows)
        df.to_excel(os.path.join(REFERENCIA_DIR, "proveedores.xlsx"), index=False)

    return jsonify({"ok": True})


@app.route("/")
def index():
    return HTML


# ── HTML ──────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — Configuración del Hotel</title>
<style>
:root{
  --bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;
  --acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;
  --tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh;line-height:1.5}

/* Nav */
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px}
.logo{display:flex;align-items:baseline;gap:10px}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc)}
.logo-name{font-size:20px;font-weight:800;color:var(--acc2);letter-spacing:-.5px}
.logo-tag{font-size:11px;color:var(--mut);font-weight:400}

/* Progress */
.progress-wrap{max-width:700px;margin:32px auto 0;padding:0 24px}
.steps-bar{display:flex;align-items:center;gap:0;margin-bottom:32px}
.step-dot{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;border:2px solid var(--s2);color:var(--dim);background:var(--s1);flex-shrink:0;transition:.3s;cursor:default}
.step-dot.active{border-color:var(--acc);color:#fff;background:var(--acc);box-shadow:0 0 12px rgba(59,130,246,.4)}
.step-dot.done{border-color:var(--grn);color:#fff;background:var(--grn)}
.step-line{flex:1;height:2px;background:var(--s2);transition:.3s}
.step-line.done{background:var(--grn)}
.step-labels{display:flex;justify-content:space-between;margin-top:8px}
.step-label{font-size:10px;color:var(--dim);text-align:center;width:72px;transition:.2s}
.step-label.active{color:var(--acc2)}

/* Card */
.card{max-width:700px;margin:0 auto 40px;background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:32px;animation:fadeIn .25s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
.card h2{font-size:1.3rem;font-weight:700;margin-bottom:6px}
.card p.sub{font-size:.85rem;color:var(--mut);margin-bottom:24px}
@media(max-width:640px){.card{margin:0 12px 40px;padding:20px}}

/* Form */
label{display:block;font-size:.8rem;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;margin-top:18px}
label:first-of-type{margin-top:0}
input,select,textarea{width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:10px;padding:11px 14px;font-size:.9rem;outline:none;transition:.15s;font-family:inherit}
input:focus,select:focus,textarea:focus{border-color:var(--acc);box-shadow:0 0 0 2px rgba(59,130,246,.15)}
input.err,select.err{border-color:var(--red)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
@media(max-width:480px){.row2{grid-template-columns:1fr}}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
@media(max-width:600px){.row3{grid-template-columns:1fr}}

/* Buttons */
.btn-row{display:flex;justify-content:space-between;margin-top:28px;gap:12px}
.btn{padding:11px 24px;border-radius:10px;font-size:.9rem;font-weight:700;cursor:pointer;border:none;transition:.15s}
.btn-prev{background:var(--s2);color:var(--tx)}
.btn-prev:hover{background:var(--s3)}
.btn-next{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;box-shadow:0 0 16px rgba(59,130,246,.3)}
.btn-next:hover{box-shadow:0 0 24px rgba(59,130,246,.5);transform:translateY(-1px)}
.btn-next:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
.btn-add{background:rgba(59,130,246,.1);color:var(--acc2);border:1px dashed var(--acc);border-radius:10px;padding:10px;font-size:.85rem;cursor:pointer;width:100%;margin-top:14px;font-weight:600;transition:.15s}
.btn-add:hover{background:rgba(59,130,246,.2)}
.btn-del{background:none;border:none;color:var(--red);cursor:pointer;font-size:1.1rem;padding:4px 8px;border-radius:6px;transition:.15s;flex-shrink:0}
.btn-del:hover{background:rgba(239,68,68,.15)}

/* Dynamic rows */
.dyn-row{background:var(--bg);border:1px solid var(--s2);border-radius:12px;padding:16px;margin-top:12px;position:relative;animation:fadeIn .2s ease}
.dyn-row .row-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.dyn-row .row-head span{font-size:.85rem;font-weight:700;color:var(--acc3)}

/* Summary */
.sum-section{margin-bottom:20px}
.sum-section h3{font-size:.9rem;color:var(--acc2);margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px}
.sum-grid{display:grid;grid-template-columns:140px 1fr;gap:4px 12px;font-size:.85rem}
.sum-grid .k{color:var(--mut)}.sum-grid .v{color:var(--tx);font-weight:600}
.sum-table{width:100%;font-size:.8rem;border-collapse:collapse;margin-top:6px}
.sum-table th{text-align:left;color:var(--mut);font-weight:600;padding:6px 10px;border-bottom:1px solid var(--s2)}
.sum-table td{padding:6px 10px;border-bottom:1px solid rgba(51,65,85,.4)}

/* Success */
.success{text-align:center;padding:40px 20px}
.success .icon{font-size:64px;margin-bottom:16px}
.success h2{font-size:1.5rem;margin-bottom:8px;color:var(--grn)}
.success p{color:var(--mut);font-size:.95rem;line-height:1.6}
.success .btn-dash{margin-top:24px;display:inline-block;padding:12px 28px;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;border-radius:10px;font-weight:700;text-decoration:none;font-size:.95rem}
</style>
</head>
<body>

<nav class="nav">
  <div class="logo">
    <div class="logo-dot"></div>
    <span class="logo-name">Yve.01</span>
    <span class="logo-tag">Configuración</span>
  </div>
  <div style="flex:1"></div>
  <a href="http://localhost:5001" style="background:var(--acc);color:#fff;text-decoration:none;padding:6px 12px;border-radius:8px;font-size:12px;font-weight:600;transition:.15s;white-space:nowrap" onmouseover="this.style.background='var(--acc2)'" onmouseout="this.style.background='var(--acc)'">← Dashboard</a>
</nav>

<div class="progress-wrap">
  <div class="steps-bar" id="steps-bar"></div>
</div>

<div class="card" id="card"></div>

<script>
// ── State ──
const STEPS = ['Hotel','OTAs','Proveedores','Usuarios','Resumen'];
let step = 0;
let D = {hotel:{},otas:[],proveedores:[],usuarios:{},alertas:{}};

// ── Progress bar ──
function renderBar() {
  const bar = document.getElementById('steps-bar');
  bar.innerHTML = '';
  STEPS.forEach((s, i) => {
    if (i > 0) {
      const line = document.createElement('div');
      line.className = 'step-line' + (i <= step ? ' done' : '');
      bar.appendChild(line);
    }
    const dot = document.createElement('div');
    dot.className = 'step-dot' + (i === step ? ' active' : i < step ? ' done' : '');
    dot.textContent = i < step ? '✓' : (i + 1);
    bar.appendChild(dot);
  });
}

// ── Navigation ──
function prev() { if (step > 0) { step--; render(); } }
function next() {
  if (!validate()) return;
  save();
  if (step < 4) { step++; render(); }
}

function save() {
  if (step === 0) {
    D.hotel = {
      nombre: v('h-nombre'), direccion: v('h-dir'), nif: v('h-nif'),
      pms: v('h-pms'), contable: v('h-contable')
    };
  } else if (step === 1) {
    D.otas = [];
    document.querySelectorAll('.ota-row').forEach(r => {
      const nombre = r.querySelector('.ota-nombre').value.trim();
      if (nombre) D.otas.push({
        nombre,
        porcentaje: r.querySelector('.ota-pct').value || '15',
        mercado: r.querySelector('.ota-mercado').value
      });
    });
  } else if (step === 2) {
    D.proveedores = [];
    document.querySelectorAll('.prov-row').forEach(r => {
      const nombre = r.querySelector('.prov-nombre').value.trim();
      if (nombre) D.proveedores.push({
        nombre,
        tipo: r.querySelector('.prov-tipo').value,
        cuenta: r.querySelector('.prov-cuenta').value || '600',
        email: r.querySelector('.prov-email').value,
        iva: r.querySelector('.prov-iva').value || '21'
      });
    });
  } else if (step === 3) {
    D.usuarios = {
      fc_nombre: v('u-fc-n'), fc_email: v('u-fc-e'),
      ia_nombre: v('u-ia-n'), ia_email: v('u-ia-e'),
      fb_nombre: v('u-fb-n'), fb_email: v('u-fb-e'),
      otras_nombre: v('u-ot-n'), otras_email: v('u-ot-e')
    };
    function chk(id) { const el = document.getElementById(id); return el ? el.checked : true; }
    D.alertas = {
      email: v('al-email'),
      ar_discrepancia: chk('al-ar-disc'),
      ar_falta_di: chk('al-ar-di'),
      drr_oob: chk('al-drr-oob'),
      ap_discrepancia: chk('al-ap-disc'),
    };
  }
}

function v(id) { const el = document.getElementById(id); return el ? el.value.trim() : ''; }

function validate() {
  let ok = true;
  document.querySelectorAll('.err').forEach(e => e.classList.remove('err'));
  if (step === 0) {
    ['h-nombre','h-nif'].forEach(id => {
      if (!v(id)) { document.getElementById(id).classList.add('err'); ok = false; }
    });
  } else if (step === 1) {
    if (!document.querySelectorAll('.ota-row').length) ok = false;
  }
  return ok;
}

// ── Render ──
function render() {
  renderBar();
  const c = document.getElementById('card');
  if (step === 0) c.innerHTML = stepHotel();
  else if (step === 1) c.innerHTML = stepOTAs();
  else if (step === 2) c.innerHTML = stepProv();
  else if (step === 3) c.innerHTML = stepUsers();
  else if (step === 4) c.innerHTML = stepSummary();
  // Restore values
  if (step === 0) {
    setV('h-nombre', D.hotel.nombre); setV('h-dir', D.hotel.direccion);
    setV('h-nif', D.hotel.nif); setV('h-pms', D.hotel.pms); setV('h-contable', D.hotel.contable);
  } else if (step === 3) {
    setV('u-fc-n', D.usuarios.fc_nombre); setV('u-fc-e', D.usuarios.fc_email);
    setV('u-ia-n', D.usuarios.ia_nombre); setV('u-ia-e', D.usuarios.ia_email);
    setV('u-fb-n', D.usuarios.fb_nombre); setV('u-fb-e', D.usuarios.fb_email);
    setV('u-ot-n', D.usuarios.otras_nombre); setV('u-ot-e', D.usuarios.otras_email);
    if (D.alertas && D.alertas.email) setV('al-email', D.alertas.email);
  }
}

function setV(id, val) { const el = document.getElementById(id); if (el && val) el.value = val; }

// ── Step 1: Hotel ──
function stepHotel() {
  return '<h2>Datos del Hotel</h2><p class="sub">Información básica de tu establecimiento</p>'
    + '<label>Nombre del hotel *</label><input id="h-nombre" placeholder="Hotel Example Barcelona">'
    + '<label>Dirección</label><input id="h-dir" placeholder="C/ Example 123, Barcelona">'
    + '<label>NIF / CIF *</label><input id="h-nif" placeholder="B12345678">'
    + '<div class="row2">'
    + '<div><label>PMS (sistema de gestión)</label><select id="h-pms">'
    + '<option value="Opera">Oracle Opera</option><option value="Mews">Mews</option>'
    + '<option value="Protel">Protel</option><option value="Sihot">Sihot</option>'
    + '<option value="Otro">Otro</option></select></div>'
    + '<div><label>Sistema contable</label><select id="h-contable">'
    + '<option value="SAP">SAP</option><option value="Oracle">Oracle Financials</option>'
    + '<option value="A3">A3</option><option value="Sage">Sage</option>'
    + '<option value="Otro">Otro</option></select></div>'
    + '</div>'
    + btnRow(false, true);
}

// ── Step 2: OTAs ──
function stepOTAs() {
  let rows = '';
  const otas = D.otas.length ? D.otas : [{nombre:'Booking.com',porcentaje:'15',mercado:'Internacional'}];
  otas.forEach((o, i) => { rows += otaRow(i, o); });
  return '<h2>OTAs y Comisiones Pactadas</h2><p class="sub">Añade las OTAs con las que trabaja el hotel y su comisión</p>'
    + '<div id="ota-list">' + rows + '</div>'
    + '<button class="btn-add" onclick="addOTA()">+ Añadir OTA</button>'
    + btnRow(true, true);
}

function otaRow(i, o) {
  return '<div class="dyn-row ota-row"><div class="row-head"><span>OTA #' + (i+1) + '</span>'
    + '<button class="btn-del" onclick="this.closest(\'.ota-row\').remove()" title="Eliminar">&times;</button></div>'
    + '<div class="row3">'
    + '<div><label>Nombre OTA</label><input class="ota-nombre" value="' + (o.nombre||'') + '" placeholder="Booking.com"></div>'
    + '<div><label>Comisión %</label><input class="ota-pct" type="number" step="0.5" value="' + (o.porcentaje||15) + '"></div>'
    + '<div><label>Mercado</label><select class="ota-mercado">'
    + '<option value="Internacional"' + (o.mercado==='Internacional'?' selected':'') + '>Internacional</option>'
    + '<option value="Nacional"' + (o.mercado==='Nacional'?' selected':'') + '>Nacional</option></select></div>'
    + '</div></div>';
}

function addOTA() {
  const list = document.getElementById('ota-list');
  const n = list.querySelectorAll('.ota-row').length;
  list.insertAdjacentHTML('beforeend', otaRow(n, {}));
}

// ── Step 3: Proveedores ──
function stepProv() {
  let rows = '';
  D.proveedores.forEach((p, i) => { rows += provRow(i, p); });
  return '<h2>Proveedores Habituales</h2><p class="sub">Proveedores de F&B y otros servicios del hotel</p>'
    + '<div id="prov-list">' + rows + '</div>'
    + '<button class="btn-add" onclick="addProv()">+ Añadir Proveedor</button>'
    + btnRow(true, true);
}

function provRow(i, p) {
  return '<div class="dyn-row prov-row"><div class="row-head"><span>Proveedor #' + (i+1) + '</span>'
    + '<button class="btn-del" onclick="this.closest(\'.prov-row\').remove()" title="Eliminar">&times;</button></div>'
    + '<div class="row2">'
    + '<div><label>Nombre</label><input class="prov-nombre" value="' + (p.nombre||'') + '" placeholder="Makro Cash & Carry"></div>'
    + '<div><label>Tipo</label><select class="prov-tipo">'
    + '<option value="FB"' + (p.tipo==='FB'?' selected':'') + '>F&B</option>'
    + '<option value="OTRAS"' + (p.tipo==='OTRAS'?' selected':'') + '>OTRAS</option></select></div>'
    + '</div><div class="row3">'
    + '<div><label>Cuenta contable</label><input class="prov-cuenta" value="' + (p.cuenta||'600') + '" placeholder="600"></div>'
    + '<div><label>Email</label><input class="prov-email" type="email" value="' + (p.email||'') + '" placeholder="facturas@proveedor.es"></div>'
    + '<div><label>IVA habitual %</label><input class="prov-iva" type="number" value="' + (p.iva||'21') + '"></div>'
    + '</div></div>';
}

function addProv() {
  const list = document.getElementById('prov-list');
  const n = list.querySelectorAll('.prov-row').length;
  list.insertAdjacentHTML('beforeend', provRow(n, {}));
}

// ── Step 4: Usuarios ──
function stepUsers() {
  const al = D.alertas || {};
  return '<h2>Usuarios y Roles</h2><p class="sub">Personas clave del equipo financiero del hotel</p>'
    + userBlock('Financial Controller','u-fc')
    + userBlock('Income Auditor','u-ia')
    + userBlock('F&B Manager','u-fb')
    + userBlock('Jefe de OTRAS','u-ot')
    + '<div style="margin-top:24px;padding:18px;background:var(--bg);border:1px solid var(--s2);border-radius:10px">'
    + '<div style="font-size:.85rem;font-weight:700;color:var(--ora);margin-bottom:12px">🔔 Configuración de Alertas</div>'
    + '<label>Email para notificaciones</label>'
    + '<input id="al-email" type="email" placeholder="controller@hotel.com" value="' + (al.email||'') + '">'
    + '<div style="margin-top:14px;display:grid;gap:10px">'
    + alertToggle('al-ar-disc', 'Discrepancias AR (comisiones OTA)', al.ar_discrepancia !== false)
    + alertToggle('al-ar-di', 'Falta certificado DI', al.ar_falta_di !== false)
    + alertToggle('al-drr-oob', 'DRR: días Out of Balance', al.drr_oob !== false)
    + alertToggle('al-ap-disc', 'Discrepancias AP (proveedores)', al.ap_discrepancia !== false)
    + '</div></div>'
    + btnRow(true, true);
}

function alertToggle(id, label, checked) {
  return '<label style="display:flex;align-items:center;gap:10px;cursor:pointer;margin:0;text-transform:none;letter-spacing:0;font-size:.85rem;color:var(--tx)">'
    + '<input type="checkbox" id="' + id + '"' + (checked ? ' checked' : '') + ' style="width:18px;height:18px;accent-color:var(--acc)">'
    + label + '</label>';
}

function userBlock(role, prefix) {
  return '<div style="margin-top:18px;padding:14px;background:var(--bg);border:1px solid var(--s2);border-radius:10px">'
    + '<div style="font-size:.85rem;font-weight:700;color:var(--acc3);margin-bottom:8px">' + role + '</div>'
    + '<div class="row2">'
    + '<div><label>Nombre</label><input id="' + prefix + '-n" placeholder="Nombre completo"></div>'
    + '<div><label>Email</label><input id="' + prefix + '-e" type="email" placeholder="email@hotel.com"></div>'
    + '</div></div>';
}

// ── Step 5: Summary ──
function stepSummary() {
  let otaTbl = '<table class="sum-table"><thead><tr><th>OTA</th><th>Comisión</th><th>Mercado</th></tr></thead><tbody>';
  D.otas.forEach(o => { otaTbl += '<tr><td>' + o.nombre + '</td><td>' + o.porcentaje + '%</td><td>' + o.mercado + '</td></tr>'; });
  otaTbl += '</tbody></table>';

  let provTbl = '';
  if (D.proveedores.length) {
    provTbl = '<table class="sum-table"><thead><tr><th>Proveedor</th><th>Tipo</th><th>Cuenta</th><th>IVA</th></tr></thead><tbody>';
    D.proveedores.forEach(p => { provTbl += '<tr><td>' + p.nombre + '</td><td>' + p.tipo + '</td><td>' + p.cuenta + '</td><td>' + p.iva + '%</td></tr>'; });
    provTbl += '</tbody></table>';
  } else {
    provTbl = '<p style="color:var(--dim);font-size:.85rem">No se añadieron proveedores</p>';
  }

  const u = D.usuarios;
  let userHtml = '';
  [['Financial Controller', u.fc_nombre, u.fc_email],
   ['Income Auditor', u.ia_nombre, u.ia_email],
   ['F&B Manager', u.fb_nombre, u.fb_email],
   ['Jefe de OTRAS', u.otras_nombre, u.otras_email]].forEach(([role, n, e]) => {
    if (n) userHtml += '<div class="sum-grid"><span class="k">' + role + '</span><span class="v">' + n + ' &lt;' + (e||'') + '&gt;</span></div>';
  });
  if (!userHtml) userHtml = '<p style="color:var(--dim);font-size:.85rem">No se configuraron usuarios</p>';

  return '<h2>Resumen de Configuración</h2><p class="sub">Revisa los datos antes de confirmar</p>'
    + '<div class="sum-section"><h3>Hotel</h3><div class="sum-grid">'
    + '<span class="k">Nombre</span><span class="v">' + (D.hotel.nombre||'—') + '</span>'
    + '<span class="k">NIF</span><span class="v">' + (D.hotel.nif||'—') + '</span>'
    + '<span class="k">Dirección</span><span class="v">' + (D.hotel.direccion||'—') + '</span>'
    + '<span class="k">PMS</span><span class="v">' + (D.hotel.pms||'—') + '</span>'
    + '<span class="k">Contable</span><span class="v">' + (D.hotel.contable||'—') + '</span>'
    + '</div></div>'
    + '<div class="sum-section"><h3>OTAs (' + D.otas.length + ')</h3>' + otaTbl + '</div>'
    + '<div class="sum-section"><h3>Proveedores (' + D.proveedores.length + ')</h3>' + provTbl + '</div>'
    + '<div class="sum-section"><h3>Usuarios</h3>' + userHtml + '</div>'
    + '<div class="btn-row"><button class="btn btn-prev" onclick="prev()">← Anterior</button>'
    + '<button class="btn btn-next" id="btn-confirm" onclick="confirm()">✓ Confirmar y Guardar</button></div>';
}

// ── Confirm ──
async function confirm() {
  const btn = document.getElementById('btn-confirm');
  btn.disabled = true;
  btn.textContent = 'Guardando...';
  try {
    const resp = await fetch('/api/save', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(D)
    });
    const result = await resp.json();
    if (result.ok) {
      document.getElementById('card').innerHTML =
        '<div class="success">'
        + '<div class="icon">🏨</div>'
        + '<h2>¡' + D.hotel.nombre + ' configurado!</h2>'
        + '<p>Yve.01 está listo para procesar facturas.<br>'
        + 'Se han guardado ' + D.otas.length + ' OTAs y ' + D.proveedores.length + ' proveedores.</p>'
        + '<a class="btn-dash" href="http://localhost:5001">Ir al Dashboard →</a>'
        + '</div>';
      document.querySelector('.progress-wrap').style.display = 'none';
    } else {
      btn.disabled = false;
      btn.textContent = '✓ Confirmar y Guardar';
      alert('Error: ' + (result.error || 'desconocido'));
    }
  } catch(e) {
    btn.disabled = false;
    btn.textContent = '✓ Confirmar y Guardar';
    alert('Error de conexión');
  }
}

// ── Helpers ──
function btnRow(showPrev, showNext) {
  return '<div class="btn-row">'
    + (showPrev ? '<button class="btn btn-prev" onclick="prev()">← Anterior</button>' : '<div></div>')
    + (showNext ? '<button class="btn btn-next" onclick="next()">Siguiente →</button>' : '')
    + '</div>';
}

// ── Init ──
// Load existing config if any
(async () => {
  try {
    const r = await fetch('/api/config');
    const cfg = await r.json();
    if (cfg && cfg.hotel) {
      D = cfg;
      step = 4; // Go to summary
    }
  } catch(e) {}
  render();
})();
</script>
</body>
</html>"""

if __name__ == '__main__':
    import socket
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = "127.0.0.1"
    print("=" * 60)
    print("  Yve.01 — Onboarding")
    print("=" * 60)
    print("  Escritorio:  http://localhost:5003")
    print(f"  Movil:       http://{ip}:5003")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5003, debug=False)
