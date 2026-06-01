"""
app_conciliacion.py — Yve.01
Dashboard de conciliacion bancaria. Puerto 5007.
"""

import os, sys, glob, json
from pathlib import Path
from flask import Flask, jsonify, request, Response, stream_with_context
import pandas as pd
import subprocess

BASE_DIR     = Path(__file__).parent
REPORTES_DIR = BASE_DIR / "reportes"
REFERENCIA_DIR = BASE_DIR / "datos-referencia"

app = Flask(__name__)


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


@app.route("/api/stats")
def api_stats():
    ruta = _ultimo("conciliacion_*.xlsx", REPORTES_DIR)
    if not ruta:
        return jsonify(None)
    try:
        df = pd.read_excel(ruta)
        total = len(df)
        conc = int((df["estado"] == "CONCILIADO").sum()) if "estado" in df.columns else 0
        pend = int((df["estado"] == "PENDIENTE").sum()) if "estado" in df.columns else 0
        diff = int((df["estado"] == "DIFERENCIA").sum()) if "estado" in df.columns else 0
        imp_total = df["importe"].apply(_sf).sum()
        imp_pend = df.loc[df.get("estado", pd.Series()) == "PENDIENTE", "importe"].apply(_sf).sum() if "estado" in df.columns else 0
        saldo = _sf(df["saldo"].iloc[-1]) if "saldo" in df.columns and len(df) > 0 else 0

        rows = []
        for _, r in df.iterrows():
            rows.append({
                "fecha": str(r.get("fecha", ""))[:10],
                "concepto": str(r.get("concepto", "")),
                "importe": _sf(r.get("importe", 0)),
                "tipo": str(r.get("tipo", "")),
                "referencia": str(r.get("referencia", "")),
                "saldo": _sf(r.get("saldo", 0)),
                "estado": str(r.get("estado", "PENDIENTE")),
                "factura_ref": str(r.get("factura_ref", "")),
                "origen": str(r.get("origen", "")),
                "match_proveedor": str(r.get("match_proveedor", "")),
                "diferencia": _sf(r.get("diferencia", 0)),
            })

        return jsonify({
            "total": total, "conciliados": conc, "pendientes": pend,
            "diferencias": diff, "importe_total": round(imp_total, 2),
            "importe_pendiente": round(imp_pend, 2), "saldo": round(saldo, 2),
            "archivo": os.path.basename(ruta),
            "movimientos": rows,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/conciliar", methods=["POST"])
def api_conciliar():
    script = str(BASE_DIR / "conciliacion_bancaria.py")
    try:
        res = subprocess.run(
            [sys.executable, script],
            capture_output=True, text=True, timeout=60, cwd=str(BASE_DIR)
        )
        ok = res.returncode == 0
        return jsonify({"ok": ok, "output": res.stdout[-500:] if res.stdout else res.stderr[-500:]})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/upload", methods=["POST"])
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


@app.route("/api/asignar_manual", methods=["POST"])
def api_asignar():
    """Asigna manualmente un movimiento a una factura."""
    data = request.get_json(force=True) or {}
    idx = data.get("idx")
    factura = data.get("factura", "")
    ruta = _ultimo("conciliacion_*.xlsx", REPORTES_DIR)
    if not ruta or idx is None:
        return jsonify({"ok": False}), 400
    try:
        df = pd.read_excel(ruta)
        if 0 <= idx < len(df):
            df.loc[idx, "estado"] = "CONCILIADO"
            df.loc[idx, "factura_ref"] = factura
            df.loc[idx, "origen"] = "MANUAL"
            df.to_excel(ruta, index=False)
            return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    return jsonify({"ok": False}), 400


@app.route("/")
def index():
    return HTML


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — Conciliacion Bancaria</title>
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;--acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--red:#ef4444;--ora:#f97316;--pur:#8b5cf6}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh;line-height:1.5}
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc);display:inline-block;margin-right:8px}
.logo-name{font-size:20px;font-weight:800;color:var(--acc2)}
.logo-tag{font-size:11px;color:var(--mut);margin-left:8px}
.nav-r{margin-left:auto;display:flex;gap:10px}
.nav-r a{color:var(--acc2);text-decoration:none;font-size:.85rem;padding:6px 14px;border:1px solid var(--s2);border-radius:8px;transition:.15s}
.nav-r a:hover{border-color:var(--acc)}
.main{max-width:1400px;margin:24px auto;padding:0 24px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px}
@media(max-width:800px){.stats{grid-template-columns:repeat(2,1fr)}}
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px}
.sc .lbl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}
.sc .val{font-size:24px;font-weight:800;letter-spacing:-1px}
.sc .sub{font-size:11px;color:var(--dim);margin-top:4px}
.actions{display:flex;gap:12px;margin-bottom:22px;flex-wrap:wrap}
.btn{padding:10px 18px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;border:none;transition:.15s}
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
</head>
<body>
<nav class="nav">
  <span class="logo-dot"></span><span class="logo-name">Yve.01</span>
  <span class="logo-tag">Conciliacion Bancaria</span>
  <div class="nav-r">
    <a href="http://localhost:5001">← Dashboard</a>
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
  if (estado === 'CONCILIADO') return '<span class="badge b-ok">Conciliado</span>';
  if (estado === 'PENDIENTE') return '<span class="badge b-pend">Pendiente</span>';
  if (estado === 'DIFERENCIA') return '<span class="badge b-diff">Diferencia</span>';
  return '<span class="badge b-manual">' + estado + '</span>';
}

async function loadData() {
  try {
    var r = await fetch('/api/stats');
    var d = await r.json();
    if (!d || d.error) { document.getElementById('status-msg').textContent = d ? d.error : 'Sin datos'; return; }

    document.getElementById('s-total').textContent = d.total;
    document.getElementById('s-conc').textContent = d.conciliados;
    document.getElementById('s-pend').textContent = d.pendientes;
    document.getElementById('s-diff').textContent = d.diferencias;
    document.getElementById('s-imp-pend').textContent = 'Pendiente: ' + eur(d.importe_pendiente);

    var tbody = document.getElementById('tbl-body');
    if (!d.movimientos || !d.movimientos.length) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--dim);padding:40px">Sin movimientos. Sube un extracto.</td></tr>';
      return;
    }
    tbody.innerHTML = d.movimientos.map(function(m, i) {
      var tipoColor = m.tipo === 'ABONO' ? 'color:var(--grn)' : 'color:var(--red)';
      var accBtn = m.estado === 'PENDIENTE'
        ? '<button class="btn btn-sec" style="padding:4px 8px;font-size:11px" onclick="asignarManual(' + i + ')">Asignar</button>'
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
    document.getElementById('status-msg').textContent = 'Error cargando datos';
  }
}

async function runConciliacion() {
  var btn = document.getElementById('btn-conc');
  var msg = document.getElementById('status-msg');
  btn.disabled = true; btn.textContent = 'Procesando...';
  msg.textContent = '';
  try {
    var r = await fetch('/api/conciliar', {method:'POST'});
    var d = await r.json();
    if (d.ok) {
      msg.textContent = 'Conciliacion completada';
      loadData();
    } else {
      msg.textContent = 'Error: ' + (d.error || d.output || '');
    }
  } catch(e) { msg.textContent = 'Error de conexion'; }
  btn.disabled = false; btn.textContent = 'Ejecutar conciliacion';
}

async function uploadFile(input) {
  var file = input.files[0];
  if (!file) return;
  var msg = document.getElementById('status-msg');
  msg.textContent = 'Subiendo ' + file.name + '...';
  var form = new FormData();
  form.append('file', file);
  try {
    var r = await fetch('/api/upload', {method:'POST', body:form});
    var d = await r.json();
    msg.textContent = d.ok ? 'Extracto subido. Ejecuta la conciliacion.' : ('Error: ' + d.error);
  } catch(e) { msg.textContent = 'Error subiendo archivo'; }
  input.value = '';
}

async function asignarManual(idx) {
  var factura = prompt('Numero de factura para asignar a este movimiento:');
  if (!factura) return;
  try {
    var r = await fetch('/api/asignar_manual', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({idx:idx, factura:factura})
    });
    var d = await r.json();
    if (d.ok) loadData();
  } catch(e) {}
}

loadData();
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
    print("  Yve.01 — Conciliacion Bancaria")
    print("=" * 60)
    print("  http://localhost:5007")
    print(f"  http://{ip}:5007")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5007, debug=False)
