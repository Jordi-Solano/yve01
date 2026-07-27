"""
app_aprobacion_ap.py — Yve.01 Módulo AP
Dashboard de aprobaciones AP para jefes de departamento. Puerto 5002.
Ejecutar: python app_aprobacion_ap.py
"""

import os, glob, json
from datetime import date, datetime
import pandas as pd
from flask import Blueprint, jsonify, request, Response
from flask_login import login_required

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
PROCESADAS_DIR   = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_DIR = os.path.join(BASE_DIR, "aprobaciones")
os.makedirs(APROBACIONES_DIR, exist_ok=True)

APRO_FILE = os.path.join(APROBACIONES_DIR, "aprobaciones_ap.xlsx")
NF        = "NO_ENCONTRADO"

bp = Blueprint("aprob_ap", __name__, url_prefix="/aprobaciones-ap")

@bp.before_request
@login_required
def _require_login():
    """Protege todas las rutas del blueprint: exige sesión iniciada."""
    pass


# ── Datos ─────────────────────────────────────────────────────────────────

def cargar_facturas_ap():
    excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_contabilizadas_*.xlsx")), reverse=True)
    if not excels:
        excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_ap_*.xlsx")), reverse=True)
    if not excels:
        return pd.DataFrame()
    return pd.read_excel(excels[0])

def cargar_aprobaciones():
    if not os.path.exists(APRO_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(APRO_FILE)
    except Exception:
        return pd.DataFrame()

def safe_str(v):
    if v is None or str(v).strip() in (NF, "nan", "None", ""):
        return ""
    return str(v).strip()

def facturas_a_lista(df):
    if df.empty:
        return []
    df_apro = cargar_aprobaciones()
    apro_map = {}
    if not df_apro.empty and "numero_factura" in df_apro.columns:
        ultimas = df_apro.sort_values("fecha_hora").groupby("numero_factura").last()
        apro_map = ultimas["accion"].to_dict() if "accion" in ultimas.columns else {}

    filas = []
    for _, r in df.iterrows():
        num = safe_str(r.get("numero_factura", ""))
        filas.append({
            "archivo":           safe_str(r.get("archivo")),
            "numero_factura":    num,
            "nombre_proveedor":  safe_str(r.get("nombre_proveedor")),
            "NIF_proveedor":     safe_str(r.get("NIF_proveedor")),
            "tipo_proveedor":    safe_str(r.get("tipo_proveedor","OTRAS")),
            "descripcion":       safe_str(r.get("descripcion_concepto")),
            "base_imponible":    safe_str(r.get("base_imponible")),
            "cuota_iva":         safe_str(r.get("cuota_iva")),
            "total_factura":     safe_str(r.get("total_factura")),
            "cuenta_debe":       safe_str(r.get("cuenta_debe_gasto")),
            "estado_asignacion": safe_str(r.get("estado_asignacion")),
            "estado_matching":   safe_str(r.get("estado_matching","")),
            "alerta_detalle":    safe_str(r.get("detalle_matching",r.get("alerta_detalle",""))),
            "departamento_po":   safe_str(r.get("departamento_po","General")),
            "accion":            apro_map.get(num, ""),
        })
    return filas

@login_required
@bp.route("/api/facturas")
def api_facturas():
    df   = cargar_facturas_ap()
    dept = request.args.get("departamento","").strip().lower()
    rows = facturas_a_lista(df)
    if dept:
        rows = [r for r in rows if dept in r.get("departamento_po","").lower()
                or dept == "todos"]
    return jsonify(rows)

@bp.route("/api/stats")
def api_stats():
    df   = cargar_facturas_ap()
    rows = facturas_a_lista(df)
    total  = len(rows)
    def cnt(key, val): return sum(1 for r in rows if r.get(key,"") == val)
    return jsonify({
        "total":         total,
        "match_ok":      cnt("estado_matching","MATCH_3WAY_OK") + cnt("estado_matching","MATCH_CORRECTO"),
        "discrepancias": cnt("estado_matching","DISCREPANCIA") + cnt("estado_matching","DISCREPANCIA_PO"),
        "alertas":       cnt("estado_matching","ALERTA_CONSUMO"),
        "sin_po":        cnt("estado_matching","SIN_PO"),
        "manuales":      cnt("estado_asignacion","SIN_REGLA"),
        "aprobadas":     cnt("accion","APROBADA"),
        "rechazadas":    cnt("accion","RECHAZADA"),
        "pendientes":    sum(1 for r in rows if not r.get("accion","")),
    })

@login_required
@bp.route("/api/accion", methods=["POST"])
def api_accion():
    data = request.get_json(force=True)
    num_fac     = data.get("numero_factura","")
    accion      = data.get("accion","")
    comentario  = data.get("comentario","")
    departamento= data.get("departamento","")
    aprobador   = data.get("aprobador","Jefe de Departamento")

    if not num_fac or not accion or not comentario or not departamento:
        return jsonify({"ok": False, "error": "Faltan campos obligatorios"}), 400

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nueva = pd.DataFrame([{
        "fecha_hora":    now_str,
        "numero_factura":num_fac,
        "accion":        accion,
        "comentario":    comentario,
        "departamento":  departamento,
        "aprobador":     aprobador,
    }])

    if os.path.exists(APRO_FILE):
        df_ex = pd.read_excel(APRO_FILE)
        df_ex = pd.concat([df_ex, nueva], ignore_index=True)
    else:
        df_ex = nueva

    with pd.ExcelWriter(APRO_FILE, engine="openpyxl") as w:
        df_ex.to_excel(w, index=False, sheet_name="Aprobaciones_AP")
    return jsonify({"ok": True})

@bp.route("/")
def index():
    return HTML

HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — AP Aprobaciones</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;
  --acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;--yel:#eab308;--pur:#8b5cf6}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;min-height:100vh}
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 18px;height:56px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100}
.logo{font-size:17px;font-weight:800;color:var(--acc2)}.logo em{color:var(--mut);font-style:normal;font-size:11px;margin-left:6px}
.pill{font-size:11px;color:var(--mut);background:var(--s2);padding:3px 10px;border-radius:20px}
.sel-dept{background:var(--s2);color:var(--tx);border:1px solid var(--s3);padding:6px 10px;border-radius:8px;font-size:12px;cursor:pointer}
.btn-ref{background:none;border:1px solid var(--s2);color:var(--mut);padding:6px 12px;border-radius:7px;font-size:12px;cursor:pointer}
.btn-ref:hover{border-color:var(--acc);color:var(--acc2)}
.nm{flex:1}
.main{padding:16px;max-width:680px;margin:0 auto}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px}
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:12px;padding:14px}
.sc-lbl{font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.sc-v{font-size:22px;font-weight:800}
.c-b{color:var(--acc2)}.c-g{color:var(--grn)}.c-r{color:var(--red)}.c-o{color:var(--ora)}.c-p{color:var(--pur)}
.card{background:#fff;border-radius:14px;padding:16px;margin-bottom:10px;color:#1e293b;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.card-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}
.prov-name{font-size:15px;font-weight:700;color:#0f172a}
.prov-num{font-size:11px;color:#64748b;margin-top:2px}
.badge{display:inline-block;font-size:9px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.3px}
.b-fb{background:#fef3c7;color:#92400e;border:1px solid #fcd34d}
.b-otras{background:#ede9fe;color:#5b21b6;border:1px solid #c4b5fd}
.b-ok{background:#dcfce7;color:#166534;border:1px solid #86efac}
.b-disc{background:#fee2e2;color:#991b1b;border:1px solid #fca5a5}
.b-alert{background:#dbeafe;color:#1e40af;border:1px solid #93c5fd}
.b-nopo{background:#fef9c3;color:#713f12;border:1px solid #fde047}
.b-man{background:#f3e8ff;color:#6b21a8;border:1px solid #d8b4fe}
.b-apr{background:#dcfce7;color:#166534;border:1px solid #86efac}
.b-rec{background:#fee2e2;color:#991b1b;border:1px solid #fca5a5}
.b-pen{background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:12px}
.ii .lbl{font-size:9px;color:#64748b;text-transform:uppercase;letter-spacing:.4px}
.ii .val{font-size:13px;font-weight:600;color:#1e293b}
.alerta-box{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:8px 10px;font-size:11px;color:#1d4ed8;margin-bottom:12px;line-height:1.5}
.alerta-box.warn{background:#fff7ed;border-color:#fed7aa;color:#c2410c}
.alerta-box.err{background:#fef2f2;border-color:#fecaca;color:#dc2626}
.dept-row{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.dept-label{font-size:11px;color:#475569;font-weight:600}
.dept-select{flex:1;background:#f8fafc;border:1.5px solid #e2e8f0;color:#334155;padding:7px 10px;border-radius:8px;font-size:12px}
.dept-select:focus{border-color:#3b82f6;outline:none}
textarea{width:100%;border:1.5px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:12px;color:#334155;resize:none;height:52px;background:#f8fafc}
textarea:focus{border-color:#3b82f6;outline:none;background:#fff}
.btn-row{display:flex;gap:8px;margin-top:10px}
.btn{flex:1;padding:11px;border-radius:10px;border:none;font-size:13px;font-weight:700;cursor:pointer}
.btn.ok{background:#22c55e;color:#fff}.btn.ko{background:#ef4444;color:#fff}
.btn.disabled{opacity:.4;cursor:not-allowed}
.cuenta-tag{display:inline-flex;align-items:center;gap:5px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:6px;padding:3px 8px;font-size:10px;color:#0369a1;font-weight:700;margin-bottom:12px}
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);padding:10px 20px;border-radius:24px;font-size:12px;font-weight:700;color:#fff;white-space:nowrap;opacity:0;transition:opacity .3s;pointer-events:none;z-index:300}
.toast.on{opacity:1}
.empty{text-align:center;padding:40px;color:var(--mut);font-size:13px}
.tabs{display:flex;background:var(--s1);border-bottom:1px solid var(--s2);padding:0 18px}
.tab{flex:1;text-align:center;padding:10px 0;font-size:12px;font-weight:600;color:var(--dim);border-bottom:2px solid transparent;cursor:pointer}
.tab.on{color:var(--acc2);border-bottom-color:var(--acc2)}
.hist-item{background:var(--s1);border:1px solid var(--s2);border-radius:12px;padding:12px 14px;margin-bottom:8px;display:flex;align-items:center;gap:10px}
.hist-icon{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.hist-icon.a{background:rgba(34,197,94,.15)}.hist-icon.r{background:rgba(239,68,68,.15)}
.hist-info .n{font-size:13px;font-weight:600;color:var(--tx)}.hist-info .d{font-size:11px;color:var(--dim);margin-top:1px}
.hist-accion{margin-left:auto;font-size:9px;font-weight:700;padding:3px 8px;border-radius:6px}
.hist-accion.a{background:rgba(34,197,94,.15);color:#4ade80}
.hist-accion.r{background:rgba(239,68,68,.15);color:#f87171}
</style>
</head>
<body>
<nav class="nav">
  <div class="logo">Yve.01 <em>AP — Aprobaciones</em></div>
  <div class="nm"></div>
  <select class="sel-dept" id="dept-filter" onchange="loadData()">
    <option value="">Todos los departamentos</option>
    <option value="f&b">F&B</option>
    <option value="administracion">Administración</option>
    <option value="mantenimiento">Mantenimiento</option>
    <option value="housekeeping">Housekeeping</option>
    <option value="seguridad">Seguridad</option>
  </select>
  <button class="btn-ref" onclick="loadData()">↻</button>
</nav>

<div class="tabs">
  <div class="tab on" id="tab-fact" onclick="showTab('facturas',this)">Facturas AP</div>
  <div class="tab" id="tab-hist" onclick="showTab('historial',this)">Historial</div>
</div>

<div class="main">
  <div class="stats" id="stats-bar">
    <div class="sc"><div class="sc-lbl">Total</div><div class="sc-v c-b" id="s-tot">—</div></div>
    <div class="sc"><div class="sc-lbl">Match OK</div><div class="sc-v c-g" id="s-ok">—</div></div>
    <div class="sc"><div class="sc-lbl">Pendientes</div><div class="sc-v c-o" id="s-pend">—</div></div>
  </div>

  <div id="tab-facturas">
    <div id="lista"></div>
  </div>

  <div id="tab-historial" style="display:none">
    <div id="hist-empty" class="empty">Sin aprobaciones todavía.</div>
    <div id="hist-list"></div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
const historial = [];

function showTab(tab, el) {
  document.getElementById('tab-facturas').style.display = tab==='facturas'?'':'none';
  document.getElementById('tab-historial').style.display = tab==='historial'?'':'none';
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  el.classList.add('on');
}

function bMatch(m) {
  const mp = {
    MATCH_3WAY_OK:['b-ok','✓ 3-Way OK'], MATCH_CORRECTO:['b-ok','✓ Match OK'],
    DISCREPANCIA_PO:['b-disc','✗ Discrepancia PO'], DISCREPANCIA:['b-disc','✗ Discrepancia'],
    ALERTA_CONSUMO:['b-alert','⚠ Alerta Consumo'], SIN_PO:['b-nopo','? Sin PO'],
    SIN_DATOS_POS:['b-nopo','~ Sin POS'], SIN_IMPORTE:['b-man','~ Sin Importe'],
  };
  const[c,l]=mp[m]||['b-pen',m||'—'];
  return '<span class="badge '+c+'">'+l+'</span>';
}

function alertClass(m) {
  if(m==='ALERTA_CONSUMO')return 'warn';
  if(m==='DISCREPANCIA_PO'||m==='DISCREPANCIA')return 'err';
  return '';
}

async function loadData() {
  const dept = document.getElementById('dept-filter').value;
  const url  = '/aprobaciones-ap/api/facturas' + (dept ? '?departamento='+encodeURIComponent(dept) : '');
  const [fr, sr] = await Promise.all([fetch(url), fetch('/aprobaciones-ap/api/stats')]);
  const rows = await fr.json();
  const stats= await sr.json();

  document.getElementById('s-tot').textContent  = stats.total??'—';
  document.getElementById('s-ok').textContent   = stats.match_ok??'—';
  document.getElementById('s-pend').textContent = stats.pendientes??'—';

  const lista = document.getElementById('lista');
  if(!rows.length){lista.innerHTML='<div class="empty">No hay facturas AP para este departamento.</div>';return;}

  lista.innerHTML = rows.map((r,i) => {
    const needsAlert = r.estado_matching && r.estado_matching !== 'MATCH_3WAY_OK' && r.estado_matching !== 'MATCH_CORRECTO' && r.alerta_detalle;
    const alertHtml  = needsAlert ? '<div class="alerta-box '+alertClass(r.estado_matching)+'">'+r.alerta_detalle+'</div>' : '';
    return '<div class="card" id="card-'+i+'">' +
      '<div class="card-top">' +
        '<div><div class="prov-name">'+r.nombre_proveedor+'</div><div class="prov-num">'+r.numero_factura+'</div></div>' +
        '<span class="badge '+(r.tipo_proveedor==='FB'?'b-fb':'b-otras')+'">'+r.tipo_proveedor+'</span>' +
      '</div>' +
      '<div class="info-grid">' +
        '<div class="ii"><div class="lbl">Total</div><div class="val">'+(r.total_factura||'—')+' €</div></div>' +
        '<div class="ii"><div class="lbl">IVA</div><div class="val">'+(r.cuota_iva||'—')+' €</div></div>' +
        '<div class="ii"><div class="lbl">Matching</div><div class="val">'+bMatch(r.estado_matching)+'</div></div>' +
        '<div class="ii"><div class="lbl">Departamento</div><div class="val">'+(r.departamento_po||'—')+'</div></div>' +
      '</div>' +
      (r.cuenta_debe ? '<div class="cuenta-tag">📒 Cuenta ' + r.cuenta_debe + ' — ' + r.estado_asignacion + '</div>' : '') +
      alertHtml +
      (r.accion ? '<div style="margin-bottom:8px">'+bMatch('')+' <span class="badge '+(r.accion==='APROBADA'?'b-apr':'b-rec')+'">'+r.accion+'</span></div>' : '') +
      (!r.accion ? (
        '<div class="dept-row"><span class="dept-label">Departamento:</span>' +
        '<select class="dept-select" id="dept-'+i+'"><option value="">— Selecciona —</option>' +
        '<option value="F&B">F&B</option><option value="Administracion">Administración</option>' +
        '<option value="Mantenimiento">Mantenimiento</option><option value="Housekeeping">Housekeeping</option>' +
        '<option value="Seguridad">Seguridad</option></select></div>' +
        '<textarea id="c-'+i+'" placeholder="Comentario obligatorio..." oninput="chk('+i+')"></textarea>' +
        '<div class="btn-row">' +
          '<button class="btn ok disabled" id="a-'+i+'" onclick="accion('+i+',\\'APROBADA\\',\\''+r.numero_factura+'\\')">✓ Aprobar</button>' +
          '<button class="btn ko disabled" id="k-'+i+'" onclick="accion('+i+',\\'RECHAZADA\\',\\''+r.numero_factura+'\\')">✗ Rechazar</button>' +
        '</div>'
      ) : '') +
    '</div>';
  }).join('');
}

function chk(i) {
  const ok = document.getElementById('c-'+i).value.trim().length > 0 &&
             document.getElementById('dept-'+i)?.value;
  document.getElementById('a-'+i).classList.toggle('disabled', !ok);
  document.getElementById('k-'+i).classList.toggle('disabled', !ok);
}

async function accion(i, tipo, numFac) {
  const c    = document.getElementById('c-'+i).value.trim();
  const dept = document.getElementById('dept-'+i)?.value;
  if(!c || !dept) return;
  const res = await fetch('/aprobaciones-ap/api/accion',{method:'POST',headers:{'Content-Type':'application/json'},
    body: JSON.stringify({numero_factura:numFac,accion:tipo,comentario:c,departamento:dept})});
  const d = await res.json();
  if(d.ok) {
    const card = document.getElementById('card-'+i);
    card.style.transition='opacity .3s,transform .3s';
    card.style.opacity='0';
    card.style.transform='translateX('+(tipo==='APROBADA'?'20px':'-20px')+')';
    setTimeout(()=>{card.style.display='none'},300);
    historial.unshift({prov:numFac,tipo,ts:new Date().toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'})});
    renderHist();
    showToast(tipo==='APROBADA'?'✓ Aprobada y guardada':'✗ Rechazada y guardada', tipo==='APROBADA'?'#16a34a':'#dc2626');
    loadData();
  }
}

function renderHist() {
  const el=document.getElementById('hist-list');
  document.getElementById('hist-empty').style.display=historial.length?'none':'';
  el.innerHTML=historial.map(h=>'<div class="hist-item">' +
    '<div class="hist-icon '+(h.tipo==='APROBADA'?'a':'r')+'">'+(h.tipo==='APROBADA'?'✓':'✗')+'</div>' +
    '<div class="hist-info"><div class="n">'+h.prov+'</div><div class="d">'+h.ts+'</div></div>' +
    '<span class="hist-accion '+(h.tipo==='APROBADA'?'a':'r')+'">'+h.tipo+'</span></div>'
  ).join('');
}

function showToast(msg, color) {
  const t=document.getElementById('toast');
  t.textContent=msg; t.style.background=color; t.classList.add('on');
  setTimeout(()=>t.classList.remove('on'),2500);
}

loadData();
</script>
</body>
</html>"""

