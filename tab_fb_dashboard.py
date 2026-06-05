"""
tab_fb_dashboard.py
Blueprint Flask para integrar el módulo F&B Cost Control en dashboard.py
Añade el tab "F&B Cost" al dashboard principal de Yve (puerto 5001).

CÓMO INTEGRAR EN dashboard.py:
  1. Copiar este archivo a la carpeta raíz del proyecto (junto a dashboard.py)
  2. En dashboard.py añadir:
        from tab_fb_dashboard import fb_bp
        app.register_blueprint(fb_bp)
  3. En el HTML del dashboard, añadir el tab "F&B Cost" con href="/fb/resumen"
"""

from flask import Blueprint, render_template_string, jsonify, request
import subprocess, sys, threading
from pathlib import Path
import pandas as pd
import json

fb_bp = Blueprint("fb", __name__, url_prefix="/fb")
BASE_DIR = Path(__file__).parent
DATOS = BASE_DIR / "datos-referencia"
REPORTES = BASE_DIR / "reportes"

_lock_fb = threading.Lock()
_running_fb = False

# ─── TEMPLATE TAB F&B ──────────────────────────────────────────────────────────

TAB_FB_HTML = """
<div style="padding:0">

  <!-- BOTÓN EJECUTAR -->
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:24px">
    <div>
      <h2 style="font-size:18px;font-weight:700">F&B Cost Control</h2>
      <p style="color:#8892a4;font-size:13px;margin-top:4px">Coste real vs teórico · Inventario · Mermas · Ranking platos</p>
    </div>
    <button id="btnFB" onclick="ejecutarFB()"
      style="background:#1a73e8;color:white;border:none;padding:10px 22px;border-radius:8px;font-weight:600;font-size:14px;cursor:pointer">
      ▶ Ejecutar Análisis
    </button>
  </div>

  <!-- LOG PIPELINE -->
  <div id="fbLog" style="display:none;background:#0a0c14;border:1px solid #2e3248;border-radius:8px;padding:16px;margin-bottom:20px;font-family:monospace;font-size:12px;color:#8892a4;max-height:150px;overflow-y:auto"></div>

  <!-- KPIs ROW -->
  <div id="fbKpis" style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px">
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
      <div style="font-size:12px;color:#8892a4;margin-bottom:6px">Total Ventas F&B</div>
      <div id="fb-ventas" style="font-size:24px;font-weight:700;color:#1a73e8">—</div>
    </div>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
      <div style="font-size:12px;color:#8892a4;margin-bottom:6px">Food Cost Teórico</div>
      <div id="fb-fc-teo" style="font-size:24px;font-weight:700;color:#1db954">—</div>
    </div>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
      <div style="font-size:12px;color:#8892a4;margin-bottom:6px">Food Cost Real</div>
      <div id="fb-fc-real" style="font-size:24px;font-weight:700;color:#ff9800">—</div>
    </div>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
      <div style="font-size:12px;color:#8892a4;margin-bottom:6px">Mermas</div>
      <div id="fb-mermas" style="font-size:24px;font-weight:700;color:#e05252">—</div>
    </div>
  </div>

  <!-- CATEGORÍAS + RANKING -->
  <div style="display:grid;grid-template-columns:1.2fr 0.8fr;gap:20px;margin-bottom:24px">

    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;overflow:hidden">
      <div style="padding:16px 20px;border-bottom:1px solid #2e3248;font-size:13px;font-weight:600">Food Cost por Categoría</div>
      <table style="width:100%;border-collapse:collapse" id="fbCatTable">
        <thead><tr style="background:#252840">
          <th style="padding:8px 14px;text-align:left;font-size:11px;color:#8892a4;text-transform:uppercase">Categoría</th>
          <th style="padding:8px 14px;text-align:right;font-size:11px;color:#8892a4;text-transform:uppercase">Ventas €</th>
          <th style="padding:8px 14px;text-align:right;font-size:11px;color:#8892a4;text-transform:uppercase">FC Teórico</th>
          <th style="padding:8px 14px;text-align:right;font-size:11px;color:#8892a4;text-transform:uppercase">FC Real</th>
          <th style="padding:8px 14px;text-align:center;font-size:11px;color:#8892a4;text-transform:uppercase">Estado</th>
        </tr></thead>
        <tbody id="fbCatBody"><tr><td colspan="5" style="padding:20px;text-align:center;color:#8892a4;font-size:13px">Ejecuta el análisis para ver los datos</td></tr></tbody>
      </table>
    </div>

    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;overflow:hidden">
      <div style="padding:16px 20px;border-bottom:1px solid #2e3248;font-size:13px;font-weight:600">Ranking Platos por Food Cost %</div>
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#252840">
          <th style="padding:8px 14px;text-align:left;font-size:11px;color:#8892a4;text-transform:uppercase">Plato</th>
          <th style="padding:8px 14px;text-align:right;font-size:11px;color:#8892a4;text-transform:uppercase">FC%</th>
          <th style="padding:8px 14px;text-align:right;font-size:11px;color:#8892a4;text-transform:uppercase">Margen€</th>
        </tr></thead>
        <tbody id="fbRankBody"><tr><td colspan="3" style="padding:20px;text-align:center;color:#8892a4;font-size:13px">—</td></tr></tbody>
      </table>
    </div>
  </div>

  <!-- ALERTAS -->
  <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;overflow:hidden">
    <div style="padding:16px 20px;border-bottom:1px solid #2e3248;font-size:13px;font-weight:600">Alertas F&B</div>
    <div id="fbAlertas" style="padding:16px 20px;color:#8892a4;font-size:13px">Ejecuta el análisis para ver alertas</div>
  </div>

</div>

<script>
function ejecutarFB() {
  const btn = document.getElementById('btnFB');
  const log = document.getElementById('fbLog');
  btn.disabled = true;
  btn.textContent = '⏳ Analizando...';
  log.style.display = 'block';
  log.innerHTML = '';

  const es = new EventSource('/fb/api/ejecutar');
  es.onmessage = e => {
    if (e.data === 'FB_COMPLETO') {
      es.close();
      btn.disabled = false;
      btn.textContent = '▶ Ejecutar Análisis';
      cargarResultados();
    } else if (e.data.startsWith('ERROR:')) {
      log.innerHTML += '<span style="color:#e05252">' + e.data + '</span>\\n';
      es.close();
      btn.disabled = false;
      btn.textContent = '▶ Ejecutar Análisis';
    } else {
      log.innerHTML += e.data + '\\n';
      log.scrollTop = log.scrollHeight;
    }
  };
}

async function cargarResultados() {
  const resp = await fetch('/fb/api/resultados');
  const data = await resp.json();
  if (!data.ok) return;

  const r = data.resumen;
  document.getElementById('fb-ventas').textContent = r.total_ventas.toLocaleString('es-ES', {minimumFractionDigits:2}) + ' €';
  document.getElementById('fb-fc-teo').textContent = r.fc_teorico_pct + '%';
  document.getElementById('fb-fc-real').textContent = r.fc_real_pct + '%';
  document.getElementById('fb-fc-real').style.color = r.alerta ? '#e05252' : '#ff9800';
  document.getElementById('fb-mermas').textContent = r.coste_mermas.toLocaleString('es-ES', {minimumFractionDigits:2}) + ' €';

  // Categorías
  const catBody = document.getElementById('fbCatBody');
  catBody.innerHTML = data.categorias.map(c => `
    <tr style="border-top:1px solid #2e3248">
      <td style="padding:10px 14px;font-size:13px">${c.categoria}</td>
      <td style="padding:10px 14px;text-align:right;font-size:13px">${c.total_ventas.toLocaleString('es-ES',{minimumFractionDigits:0})} €</td>
      <td style="padding:10px 14px;text-align:right;font-size:13px">${c.fc_teorico_pct}%</td>
      <td style="padding:10px 14px;text-align:right;font-size:13px;font-weight:600;color:${c.alerta?'#e05252':'#1db954'}">${c.fc_real_pct}%</td>
      <td style="padding:10px 14px;text-align:center;font-size:12px;color:${c.alerta?'#e05252':'#1db954'}">${c.alerta?'⚠ ALERTA':'✓ OK'}</td>
    </tr>
  `).join('');

  // Ranking
  const rankBody = document.getElementById('fbRankBody');
  rankBody.innerHTML = data.ranking.map((p,i) => `
    <tr style="border-top:1px solid #2e3248">
      <td style="padding:8px 14px;font-size:13px">${i+1}. ${p.nombre}</td>
      <td style="padding:8px 14px;text-align:right;font-size:13px;font-weight:600;color:${p.fc_pct<=28?'#1db954':p.fc_pct<=35?'#ff9800':'#e05252'}">${p.fc_pct}%</td>
      <td style="padding:8px 14px;text-align:right;font-size:13px">${p.margen_bruto.toFixed(2)} €</td>
    </tr>
  `).join('');

  // Alertas
  const cont = document.getElementById('fbAlertas');
  if (!data.alertas.length) {
    cont.innerHTML = '<span style="color:#1db954">✓ Sin alertas detectadas en el período</span>';
  } else {
    cont.innerHTML = data.alertas.map(a => `
      <div style="display:flex;gap:12px;align-items:center;margin-bottom:10px">
        <span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:8px;background:${a.nivel==='CRITICO'?'rgba(224,82,82,0.2)':'rgba(255,152,0,0.2)'};color:${a.nivel==='CRITICO'?'#e05252':'#ff9800'};white-space:nowrap">${a.nivel}</span>
        <span style="font-size:13px">${a.mensaje}</span>
      </div>
    `).join('');
  }
}

// Auto-cargar si hay resultados previos
cargarResultados();
</script>
"""

# ─── ROUTES ────────────────────────────────────────────────────────────────────

@fb_bp.route("/resumen")
def resumen():
    return render_template_string(TAB_FB_HTML)

@fb_bp.route("/api/ejecutar")
def api_ejecutar():
    from flask import Response, stream_with_context
    global _running_fb

    def generar():
        global _running_fb
        with _lock_fb:
            if _running_fb:
                yield "data: Análisis F&B ya en curso\n\n"
                return
            _running_fb = True
        try:
            res = subprocess.run(
                [sys.executable, str(BASE_DIR / "fb_cost_control.py")],
                capture_output=True, text=True, cwd=str(BASE_DIR)
            )
            for line in res.stdout.splitlines():
                yield f"data: {line}\n\n"
            if res.returncode != 0:
                yield f"data: ERROR: {res.stderr[:200]}\n\n"
            else:
                yield "data: FB_COMPLETO\n\n"
        finally:
            _running_fb = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )

@fb_bp.route("/api/resultados")
def api_resultados():
    """Lee el último reporte generado y devuelve los datos para el dashboard."""
    try:
        import sys
        sys.path.insert(0, str(BASE_DIR))
        from fb_cost_control import (
            cargar_recetas, cargar_ventas, cargar_inventario, cargar_mermas,
            calcular_food_cost_teorico, calcular_food_cost_real,
            analizar_por_categoria, ranking_platos, generar_alertas
        )
        recetas = cargar_recetas()
        ventas_df = cargar_ventas()
        inventario_df = cargar_inventario()
        mermas_df = cargar_mermas()
        teorico_df = calcular_food_cost_teorico(recetas, ventas_df)
        resumen = calcular_food_cost_real(teorico_df, mermas_df)
        categorias = analizar_por_categoria(teorico_df, mermas_df)
        ranking = ranking_platos(recetas)
        inventario_data = []  # simplificado para el tab
        alertas = generar_alertas(resumen, categorias, inventario_data)
        return jsonify({"ok": True, "resumen": resumen, "categorias": categorias, "ranking": ranking, "alertas": alertas})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

