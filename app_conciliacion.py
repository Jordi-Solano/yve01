"""
app_conciliacion.py — Yve.01
Dashboard de conciliacion bancaria. Puerto 5007.
"""

import os, sys, glob, json
from pathlib import Path
from flask import Blueprint, jsonify, request, Response, stream_with_context
from flask_login import login_required
import pandas as pd
import subprocess

BASE_DIR     = Path(__file__).parent
from tenant_dirs import reportes_dir as _t_rdir, datos_dir as _t_ddir, tenant_id as _t_tid
from pathlib import Path as _P
from version_estaticos import SELLO as SELLO_ESTATICOS

class _TDir:
    def __init__(self, fn): self._fn = fn
    def __truediv__(self, other): return _P(self._fn()) / other
    def __str__(self): return self._fn()

REPORTES_DIR = _TDir(_t_rdir)
REFERENCIA_DIR = _TDir(_t_ddir)

bp = Blueprint("concil", __name__, url_prefix="/conciliacion")

@bp.before_request
@login_required
def _require_login():
    """Protege todas las rutas del blueprint: exige sesión iniciada."""
    pass



def _ultimo(patron, d):
    hits = glob.glob(str(d / patron))
    if not hits:
        return None
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return hits[0]


def _sf(v):
    try:
        if v is None or str(v).strip() in ("", "nan", "None"):
            return 0.0
        return float(str(v).replace(",", "").replace("EUR", "").strip())
    except Exception:
        return 0.0


@bp.route("/api/stats")
def api_stats():
    """Movimientos del extracto REAL con el estado del informe.

    Antes leia SOLO el ultimo conciliacion_*.xlsx y la pantalla se quedaba
    congelada en la foto del dia que se concilio: los movimientos subidos
    despues no aparecian (reproducido: 2 de 4). Mismo arreglo que ya se hizo en
    el panel del dashboard, y por el mismo sitio — quien junta extracto e
    informe es almacen_datos.movimientos_banco(), aqui no se replica nada.

    Cada movimiento viaja con su `clave` (identidad fecha+concepto+importe).
    La asignacion manual usa esa clave y NO la posicion en la lista: la posicion
    cambia en cuanto el extracto crece o se rebaja con fechas anteriores.
    """
    try:
        import almacen_datos as _alm
        df, info = _alm.movimientos_banco(reportes_dir=str(REPORTES_DIR))
        if df is None or df.empty:
            return jsonify(None)

        def _s(v):
            s = str(v)
            return "" if s in ("nan", "None", "NaT", "<NA>") else s

        def _estado(v):
            e = str(v or "").strip().upper()
            return e if e else "PENDIENTE"

        rows, imp_total, imp_pend = [], 0.0, 0.0
        conc = pend = diff = 0
        for r in df.to_dict("records"):
            est = _estado(r.get("estado"))
            imp = _sf(r.get("importe", 0))
            imp_total += imp
            if est == "CONCILIADO":
                conc += 1
            elif est == "DIFERENCIA":
                diff += 1
            else:
                pend += 1
                imp_pend += abs(imp)
            rows.append({
                "clave": _alm.clave_movimiento(r),
                "fecha": _s(r.get("fecha", ""))[:10],
                "concepto": _s(r.get("concepto", "")),
                "importe": imp,
                "tipo": _s(r.get("tipo", "")) or ("ABONO" if imp > 0 else "CARGO"),
                "referencia": _s(r.get("referencia", "")),
                "saldo": _sf(r.get("saldo", 0)),
                "estado": est,
                "factura_ref": _s(r.get("factura_ref", "")),
                "origen": _s(r.get("origen", "")),
                "match_proveedor": _s(r.get("match_proveedor", "")),
                "diferencia": _sf(r.get("diferencia", 0)),
            })

        saldo = rows[-1]["saldo"] if rows else 0
        return jsonify({
            "total": len(rows), "conciliados": conc, "pendientes": pend,
            "diferencias": diff, "importe_total": round(imp_total, 2),
            "importe_pendiente": round(imp_pend, 2), "saldo": round(saldo, 2),
            "archivo": info.get("informe"),
            "extracto": info.get("extracto"),
            "movimientos": rows,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/conciliar", methods=["POST"])
def api_conciliar():
    script = str(BASE_DIR / "conciliacion_bancaria.py")
    try:
        _env = __import__('os').environ.copy()
        _env['YVE_TENANT'] = _t_tid()
        res = subprocess.run(
            [sys.executable, script],
            capture_output=True, text=True, timeout=60, cwd=str(BASE_DIR), env=_env
        )
        ok = res.returncode == 0
        return jsonify({"ok": ok, "output": res.stdout[-500:] if res.stdout else res.stderr[-500:]})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/api/upload", methods=["POST"])
def api_upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file"}), 400
    f = request.files["file"]
    name = f.filename.lower()
    if not (name.endswith(".xlsx") or name.endswith(".csv")):
        return jsonify({"ok": False, "error": "Solo .xlsx o .csv"}), 400
    save_path = str(REFERENCIA_DIR / "extracto_banco.xlsx")
    if name.endswith(".csv"):
        df = pd.read_csv(f)
        df.to_excel(save_path, index=False)
    else:
        f.save(save_path)
    return jsonify({"ok": True})


_COLS_INFORME = ["fecha", "concepto", "importe", "tipo", "referencia", "saldo",
                 "estado", "factura_ref", "origen", "match_proveedor", "diferencia"]


@bp.route("/api/asignar_manual", methods=["POST"])
def api_asignar():
    """Asigna manualmente un movimiento a una factura.

    Identifica el movimiento por su CLAVE (fecha+concepto+importe), nunca por su
    posicion en la lista: ahora la lista sale del extracto y su orden cambia en
    cuanto se vuelve a bajar con fechas anteriores. Escribir por posicion
    marcaria conciliado un movimiento distinto, y sin avisar.

    Si el movimiento aun no esta en el informe (es nuevo desde la ultima
    conciliacion) se añade una fila: si no, el boton no haria nada justo en los
    movimientos que mas falta hace poder asignar a mano.
    """
    import almacen_datos as _alm
    data = request.get_json(force=True) or {}
    clave = str(data.get("clave") or "").strip()
    factura = data.get("factura", "")
    if not clave:
        return jsonify({"ok": False, "error": "falta la clave del movimiento"}), 400

    try:
        df, _info = _alm.movimientos_banco(reportes_dir=str(REPORTES_DIR))
        mov = next((m for m in df.to_dict("records")
                    if _alm.clave_movimiento(m) == clave), None) if df is not None and not df.empty else None
        if mov is None:
            return jsonify({"ok": False, "error": "ese movimiento ya no esta en el extracto"}), 404

        ruta = _ultimo("conciliacion_*.xlsx", REPORTES_DIR)
        if ruta:
            df_rep = pd.read_excel(ruta)
        else:
            # nunca se ha conciliado: el informe se crea ahora, si no el boton
            # no haria nada en una cuenta recien subida.
            from datetime import datetime as _dt
            ruta = str(REPORTES_DIR / f"conciliacion_{_dt.now().strftime('%Y%m%d')}.xlsx")
            df_rep = pd.DataFrame(columns=_COLS_INFORME)

        for col in _COLS_INFORME:
            if col not in df_rep.columns:
                df_rep[col] = "" if col not in ("importe", "saldo", "diferencia") else 0.0
        # columnas de texto: si vienen vacías pandas las tipa como numéricas y rechaza strings
        for _col in ("estado", "factura_ref", "origen", "match_proveedor", "tipo", "referencia"):
            df_rep[_col] = df_rep[_col].astype(object)

        # Primera fila con esa clave que NO este ya conciliada. Dos movimientos
        # identicos tienen la misma clave a proposito: cada uno consume una fila.
        destino = None
        for pos, fila in enumerate(df_rep.to_dict("records")):
            if _alm.clave_movimiento(fila) == clave and \
               str(fila.get("estado", "")).strip().upper() != "CONCILIADO":
                destino = pos
                break

        if destino is None:
            nueva = {c: mov.get(c, "") for c in _COLS_INFORME}
            for c in ("importe", "saldo", "diferencia"):
                try:
                    nueva[c] = float(nueva.get(c) or 0)
                except (TypeError, ValueError):
                    nueva[c] = 0.0
            df_rep = pd.concat([df_rep, pd.DataFrame([nueva])], ignore_index=True)
            destino = len(df_rep) - 1

        df_rep.loc[destino, "estado"] = "CONCILIADO"
        df_rep.loc[destino, "factura_ref"] = factura
        df_rep.loc[destino, "origen"] = "MANUAL"
        df_rep[_COLS_INFORME].to_excel(ruta, index=False)
        return jsonify({"ok": True, "archivo": os.path.basename(ruta)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@bp.route("/")
def index():
    return HTML.replace("__ASSETS_V__", SELLO_ESTATICOS)


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
<title>Yve.01 — Conciliacion Bancaria</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;--acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--red:#ef4444;--ora:#f97316;--pur:#8b5cf6}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh;line-height:1.5}
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc);display:inline-block;margin-right:8px}
.logo-name{font-size:20px;font-weight:800;color:var(--acc2)}
.logo-tag{font-size:11px;color:var(--mut);margin-left:8px}
.nav-r{margin-left:auto;display:flex;gap:10px}
.nav-r a{color:var(--acc2);text-decoration:none;font-size:.85rem;padding:6px 14px;border:1px solid var(--s2);border-radius:8px;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s}
.nav-r a:hover{border-color:var(--acc)}
.main{max-width:1400px;margin:24px auto;padding:0 24px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px}
@media(max-width:800px){.stats{grid-template-columns:repeat(2,1fr)}}
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px}
.sc .lbl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}
.sc .val{font-size:24px;font-weight:800;letter-spacing:-1px}
.sc .sub{font-size:11px;color:var(--dim);margin-top:4px}
.actions{display:flex;gap:12px;margin-bottom:22px;flex-wrap:wrap}
.btn{padding:10px 18px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;border:none;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s}
.btn-blue{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;box-shadow:0 0 16px rgba(59,130,246,.3)}
.btn-blue:hover{box-shadow:0 0 24px rgba(59,130,246,.5)}
.btn-blue:disabled{opacity:.5;cursor:not-allowed}
.btn-sec{background:var(--s2);color:var(--tx)}
.btn-sec:hover{background:var(--s3)}
.card{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:22px;margin-bottom:22px}
.card-title{font-size:11px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:14px}
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px;min-width:900px}
th{background:rgba(51,65,85,.6);color:var(--mut);font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:10px 12px;text-align:left;border-bottom:1px solid var(--s2)}
td{padding:10px 12px;border-bottom:1px solid rgba(51,65,85,.4);white-space:nowrap}
tr:hover td{background:rgba(255,255,255,.02)}
.badge{display:inline-flex;padding:3px 9px;border-radius:20px;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.4px}
.b-ok{background:rgba(34,197,94,.12);color:#4ade80;border:1px solid rgba(34,197,94,.2)}
.b-pend{background:rgba(249,115,22,.12);color:#fb923c;border:1px solid rgba(249,115,22,.2)}
.b-diff{background:rgba(239,68,68,.12);color:#f87171;border:1px solid rgba(239,68,68,.2)}
.b-manual{background:rgba(139,92,246,.12);color:#c4b5fd;border:1px solid rgba(139,92,246,.2)}
.status-msg{font-size:.8rem;color:var(--dim);margin-bottom:16px}
</style>
<script src="/static/yve-tema.js?v=__ASSETS_V__"></script>
</head>
<body>
<nav class="nav">
  <span class="logo-dot"></span><span class="logo-name">Yve.01</span>
  <span class="logo-tag">Conciliacion Bancaria</span>
  <div class="nav-r">
    <a href="/">← Dashboard</a>
  </div>
</nav>

<div class="main">
  <div class="stats">
    <div class="sc"><div class="lbl">Movimientos</div><div class="val" id="s-total">—</div></div>
    <div class="sc"><div class="lbl">Conciliados</div><div class="val" id="s-conc" style="color:var(--grn)">—</div></div>
    <div class="sc"><div class="lbl">Pendientes</div><div class="val" id="s-pend" style="color:var(--ora)">—</div></div>
    <div class="sc"><div class="lbl">Diferencias</div><div class="val" id="s-diff" style="color:var(--red)">—</div><div class="sub" id="s-imp-pend"></div></div>
  </div>

  <div class="actions">
    <label class="btn btn-sec" for="upload-input" style="cursor:pointer">📂 Subir extracto (.xlsx/.csv)</label>
    <input type="file" id="upload-input" accept=".xlsx,.csv" style="display:none" onchange="uploadFile(this)">
    <button class="btn btn-blue" id="btn-conc" onclick="runConciliacion()">⚡ Ejecutar conciliacion</button>
    <span class="status-msg" id="status-msg"></span>
  </div>

  <div class="card">
    <div class="card-title">Movimientos Bancarios</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>Fecha</th><th>Concepto</th><th>Importe</th><th>Tipo</th><th>Ref</th><th>Saldo</th><th>Estado</th><th>Factura</th><th>Accion</th></tr></thead>
        <tbody id="tbl-body"><tr><td colspan="9" style="text-align:center;color:var(--dim);padding:40px">Cargando...</td></tr></tbody>
      </table>
    </div>
  </div>
</div>

<script>
function eur(n) {
  if (!n && n !== 0) return '—';
  return new Intl.NumberFormat('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n) + ' EUR';
}

function badge(estado) {
  if (estado === 'CONCILIADO') return '<span class="badge b-ok">' + tt('Conciliado') + '</span>';
  if (estado === 'PENDIENTE') return '<span class="badge b-pend">' + tt('Pendiente') + '</span>';
  if (estado === 'DIFERENCIA') return '<span class="badge b-diff">' + tt('Diferencia') + '</span>';
  return '<span class="badge b-manual">' + estado + '</span>';
}

async function loadData() {
  try {
    var r = await fetch('/conciliacion/api/stats');
    var d = await r.json();
    if (!d || d.error) { document.getElementById('status-msg').textContent = d ? d.error : tt('Sin datos'); return; }

    document.getElementById('s-total').textContent = d.total;
    document.getElementById('s-conc').textContent = d.conciliados;
    document.getElementById('s-pend').textContent = d.pendientes;
    document.getElementById('s-diff').textContent = d.diferencias;
    document.getElementById('s-imp-pend').textContent = tt('Pendiente: ') + eur(d.importe_pendiente);

    var tbody = document.getElementById('tbl-body');
    if (!d.movimientos || !d.movimientos.length) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--dim);padding:40px">' + tt('Sin movimientos. Sube un extracto.') + '</td></tr>';
      return;
    }
    tbody.innerHTML = d.movimientos.map(function(m, i) {
      var tipoColor = m.tipo === 'ABONO' ? 'color:var(--grn)' : 'color:var(--red)';
      // la clave identifica el movimiento; la POSICIÓN i ya no vale porque la
      // lista sale del extracto y su orden cambia al rebajarlo
      var accBtn = m.estado === 'PENDIENTE'
        ? '<button class="btn btn-sec" style="padding:4px 8px;font-size:11px" onclick="asignarManual(' + JSON.stringify(m.clave) + ')">' + tt('Asignar') + '</button>'
        : (m.factura_ref || '—');
      return '<tr>'
        + '<td style="color:var(--dim)">' + m.fecha + '</td>'
        + '<td style="max-width:280px;overflow:hidden;text-overflow:ellipsis">' + m.concepto + '</td>'
        + '<td style="font-weight:700;' + tipoColor + '">' + eur(m.importe) + '</td>'
        + '<td>' + m.tipo + '</td>'
        + '<td style="font-size:11px;color:var(--dim)">' + m.referencia + '</td>'
        + '<td>' + eur(m.saldo) + '</td>'
        + '<td>' + badge(m.estado) + '</td>'
        + '<td style="font-size:11px">' + (m.estado !== 'PENDIENTE' ? (m.factura_ref || '') : '') + '</td>'
        + '<td>' + accBtn + '</td>'
        + '</tr>';
    }).join('');
  } catch(e) {
    document.getElementById('status-msg').textContent = tt('Error cargando datos');
  }
}

async function runConciliacion() {
  var btn = document.getElementById('btn-conc');
  var msg = document.getElementById('status-msg');
  btn.disabled = true; btn.textContent = tt('Procesando...');
  msg.textContent = '';
  try {
    var r = await fetch('/conciliacion/api/conciliar', {method:'POST'});
    var d = await r.json();
    if (d.ok) {
      msg.textContent = tt('Conciliacion completada');
      loadData();
    } else {
      msg.textContent = 'Error: ' + (d.error || d.output || '');
    }
  } catch(e) { msg.textContent = tt('Error de conexion'); }
  btn.disabled = false; btn.textContent = '⚡ ' + tt('Ejecutar conciliacion');
}

async function uploadFile(input) {
  var file = input.files[0];
  if (!file) return;
  var msg = document.getElementById('status-msg');
  msg.textContent = tt('Subiendo ') + file.name + '...';
  var form = new FormData();
  form.append('file', file);
  try {
    var r = await fetch('/conciliacion/api/upload', {method:'POST', body:form});
    var d = await r.json();
    msg.textContent = d.ok ? tt('Extracto subido. Ejecuta la conciliacion.') : ('Error: ' + d.error);
  } catch(e) { msg.textContent = tt('Error subiendo archivo'); }
  input.value = '';
}

async function asignarManual(clave) {
  var factura = prompt(tt('Numero de factura para asignar a este movimiento:'));
  if (!factura) return;
  var msg = document.getElementById('status-msg');
  try {
    var r = await fetch('/conciliacion/api/asignar_manual', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({clave:clave, factura:factura})
    });
    var d = await r.json();
    if (d.ok) { loadData(); }
    else if (msg) { msg.textContent = 'Error: ' + (d.error || tt('no se pudo asignar')); }
  } catch(e) { if (msg) msg.textContent = tt('Error de conexion'); }
}

// ── i18n: mismo idioma que el dashboard (localStorage yve_lang) ──
var _lang = localStorage.getItem('yve_lang') || 'es';
var _L = {
 en: {"Conciliacion Bancaria":"Bank Reconciliation","← Dashboard":"← Dashboard","Movimientos":"Transactions","Conciliados":"Matched","Pendientes":"Pending","Diferencias":"Differences","Movimientos Bancarios":"Bank Transactions","📂 Subir extracto (.xlsx/.csv)":"📂 Upload statement (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Run reconciliation","Cargando...":"Loading...","Fecha":"Date","Concepto":"Description","Importe":"Amount","Tipo":"Type","Ref":"Ref","Saldo":"Balance","Estado":"Status","Factura":"Invoice","Accion":"Action","Conciliado":"Matched","Pendiente":"Pending","Diferencia":"Difference","Asignar":"Assign","Sin movimientos. Sube un extracto.":"No transactions. Upload a statement.","Pendiente: ":"Pending: ","Conciliacion completada":"Reconciliation completed","Error de conexion":"Connection error","Procesando...":"Processing...","Ejecutar conciliacion":"Run reconciliation","Subiendo ":"Uploading ","Extracto subido. Ejecuta la conciliacion.":"Statement uploaded. Run the reconciliation.","Error subiendo archivo":"Error uploading file","Numero de factura para asignar a este movimiento:":"Invoice number to assign to this transaction:","Sin datos":"No data","Error cargando datos":"Error loading data"},
 ca: {"Conciliacion Bancaria":"Conciliació Bancària","← Dashboard":"← Tauler","Movimientos":"Moviments","Conciliados":"Conciliats","Pendientes":"Pendents","Diferencias":"Diferències","Movimientos Bancarios":"Moviments Bancaris","📂 Subir extracto (.xlsx/.csv)":"📂 Pujar extracte (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Executar conciliació","Cargando...":"Carregant...","Fecha":"Data","Concepto":"Concepte","Importe":"Import","Tipo":"Tipus","Ref":"Ref","Saldo":"Saldo","Estado":"Estat","Factura":"Factura","Accion":"Acció","Conciliado":"Conciliat","Pendiente":"Pendent","Diferencia":"Diferència","Asignar":"Assignar","Sin movimientos. Sube un extracto.":"Sense moviments. Puja un extracte.","Pendiente: ":"Pendent: ","Conciliacion completada":"Conciliació completada","Error de conexion":"Error de connexió","Procesando...":"Processant...","Ejecutar conciliacion":"Executar conciliació","Subiendo ":"Pujant ","Extracto subido. Ejecuta la conciliacion.":"Extracte pujat. Executa la conciliació.","Error subiendo archivo":"Error pujant el fitxer","Numero de factura para asignar a este movimiento:":"Número de factura per assignar a aquest moviment:","Sin datos":"Sense dades","Error cargando datos":"Error carregant dades"},
 fr: {"Conciliacion Bancaria":"Rapprochement Bancaire","← Dashboard":"← Tableau de bord","Movimientos":"Mouvements","Conciliados":"Rapprochés","Pendientes":"En attente","Diferencias":"Écarts","Movimientos Bancarios":"Mouvements Bancaires","📂 Subir extracto (.xlsx/.csv)":"📂 Charger le relevé (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Lancer le rapprochement","Cargando...":"Chargement...","Fecha":"Date","Concepto":"Libellé","Importe":"Montant","Tipo":"Type","Ref":"Réf","Saldo":"Solde","Estado":"Statut","Factura":"Facture","Accion":"Action","Conciliado":"Rapproché","Pendiente":"En attente","Diferencia":"Écart","Asignar":"Assigner","Sin movimientos. Sube un extracto.":"Aucun mouvement. Chargez un relevé.","Pendiente: ":"En attente : ","Conciliacion completada":"Rapprochement terminé","Error de conexion":"Erreur de connexion","Procesando...":"Traitement...","Ejecutar conciliacion":"Lancer le rapprochement","Subiendo ":"Chargement ","Extracto subido. Ejecuta la conciliacion.":"Relevé chargé. Lancez le rapprochement.","Error subiendo archivo":"Erreur lors du chargement","Numero de factura para asignar a este movimiento:":"Numéro de facture à assigner à ce mouvement :","Sin datos":"Aucune donnée","Error cargando datos":"Erreur de chargement"},
 de: {"Conciliacion Bancaria":"Bankabstimmung","← Dashboard":"← Dashboard","Movimientos":"Bewegungen","Conciliados":"Abgestimmt","Pendientes":"Offen","Diferencias":"Differenzen","Movimientos Bancarios":"Bankbewegungen","📂 Subir extracto (.xlsx/.csv)":"📂 Kontoauszug hochladen (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Abstimmung starten","Cargando...":"Laden...","Fecha":"Datum","Concepto":"Beschreibung","Importe":"Betrag","Tipo":"Typ","Ref":"Ref","Saldo":"Saldo","Estado":"Status","Factura":"Rechnung","Accion":"Aktion","Conciliado":"Abgestimmt","Pendiente":"Offen","Diferencia":"Differenz","Asignar":"Zuordnen","Sin movimientos. Sube un extracto.":"Keine Bewegungen. Lade einen Kontoauszug hoch.","Pendiente: ":"Offen: ","Conciliacion completada":"Abstimmung abgeschlossen","Error de conexion":"Verbindungsfehler","Procesando...":"Verarbeite...","Ejecutar conciliacion":"Abstimmung starten","Subiendo ":"Lade hoch ","Extracto subido. Ejecuta la conciliacion.":"Kontoauszug hochgeladen. Starte die Abstimmung.","Error subiendo archivo":"Fehler beim Hochladen","Numero de factura para asignar a este movimiento:":"Rechnungsnummer für diese Bewegung:","Sin datos":"Keine Daten","Error cargando datos":"Fehler beim Laden"},
 it: {"Conciliacion Bancaria":"Riconciliazione Bancaria","← Dashboard":"← Dashboard","Movimientos":"Movimenti","Conciliados":"Riconciliati","Pendientes":"In sospeso","Diferencias":"Differenze","Movimientos Bancarios":"Movimenti Bancari","📂 Subir extracto (.xlsx/.csv)":"📂 Carica estratto (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Esegui riconciliazione","Cargando...":"Caricamento...","Fecha":"Data","Concepto":"Descrizione","Importe":"Importo","Tipo":"Tipo","Ref":"Rif","Saldo":"Saldo","Estado":"Stato","Factura":"Fattura","Accion":"Azione","Conciliado":"Riconciliato","Pendiente":"In sospeso","Diferencia":"Differenza","Asignar":"Assegna","Sin movimientos. Sube un extracto.":"Nessun movimento. Carica un estratto.","Pendiente: ":"In sospeso: ","Conciliacion completada":"Riconciliazione completata","Error de conexion":"Errore di connessione","Procesando...":"Elaborazione...","Ejecutar conciliacion":"Esegui riconciliazione","Subiendo ":"Caricamento ","Extracto subido. Ejecuta la conciliacion.":"Estratto caricato. Esegui la riconciliazione.","Error subiendo archivo":"Errore nel caricamento","Numero de factura para asignar a este movimiento:":"Numero di fattura da assegnare a questo movimento:","Sin datos":"Nessun dato","Error cargando datos":"Errore di caricamento"},
 pt: {"Conciliacion Bancaria":"Conciliação Bancária","← Dashboard":"← Painel","Movimientos":"Movimentos","Conciliados":"Conciliados","Pendientes":"Pendentes","Diferencias":"Diferenças","Movimientos Bancarios":"Movimentos Bancários","📂 Subir extracto (.xlsx/.csv)":"📂 Enviar extrato (.xlsx/.csv)","⚡ Ejecutar conciliacion":"⚡ Executar conciliação","Cargando...":"Carregando...","Fecha":"Data","Concepto":"Descrição","Importe":"Valor","Tipo":"Tipo","Ref":"Ref","Saldo":"Saldo","Estado":"Estado","Factura":"Fatura","Accion":"Ação","Conciliado":"Conciliado","Pendiente":"Pendente","Diferencia":"Diferença","Asignar":"Atribuir","Sin movimientos. Sube un extracto.":"Sem movimentos. Envie um extrato.","Pendiente: ":"Pendente: ","Conciliacion completada":"Conciliação concluída","Error de conexion":"Erro de conexão","Procesando...":"Processando...","Ejecutar conciliacion":"Executar conciliação","Subiendo ":"Enviando ","Extracto subido. Ejecuta la conciliacion.":"Extrato enviado. Execute a conciliação.","Error subiendo archivo":"Erro ao enviar arquivo","Numero de factura para asignar a este movimiento:":"Número da fatura para atribuir a este movimento:","Sin datos":"Sem dados","Error cargando datos":"Erro ao carregar dados"},
};
function tt(s) { var d = _L[_lang]; return (d && d[s]) ? d[s] : s; }
function _applyI18nConc() {
  if (_lang === 'es' || !_L[_lang]) return;
  var d = _L[_lang];
  var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
  var n, reps = [];
  while ((n = walker.nextNode())) {
    var txt = n.textContent, tr = txt.trim();
    if (tr && d[tr]) reps.push([n, txt.replace(tr, d[tr])]);
  }
  reps.forEach(function(r) { r[0].textContent = r[1]; });
}
_applyI18nConc();

loadData();
</script>
</body>
</html>"""

