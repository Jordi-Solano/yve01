"""
dashboard_multihotel.py — Dashboard Multi-Hotel para Yve
Puerto: 5008
Muestra KPIs consolidados + benchmarking + alertas cross-hotel del grupo.
"""

from flask import Flask, render_template_string, jsonify
import pandas as pd
import json
from pathlib import Path
from datetime import datetime

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
DATOS = BASE_DIR / "datos-referencia"

# ─── DATOS ─────────────────────────────────────────────────────────────────────

def cargar_hoteles():
    with open(DATOS / "hoteles.json") as f:
        return json.load(f)

def cargar_kpis():
    df = pd.read_excel(DATOS / "kpis_hoteles.xlsx")
    return df

# ─── TEMPLATE HTML ─────────────────────────────────────────────────────────────

HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Yve — Dashboard Multi-Hotel</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #0f1117; --bg2: #1c1f2e; --bg3: #252840;
    --border: #2e3248; --text: #e8eaf6; --muted: #8892a4;
    --blue: #1a73e8; --green: #1db954; --red: #e05252;
    --orange: #ff9800; --purple: #7c4dff; --teal: #00bcd4;
    --radius: 12px; --font: 'Inter', system-ui, sans-serif;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: var(--font); min-height: 100vh; }
  
  /* HEADER */
  .header {
    background: var(--bg2); border-bottom: 1px solid var(--border);
    padding: 16px 28px; display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100;
  }
  .logo { display: flex; align-items: center; gap: 10px; }
  .logo-icon {
    width: 36px; height: 36px; background: var(--blue);
    border-radius: 8px; display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 14px; color: white;
  }
  .logo h1 { font-size: 18px; font-weight: 700; }
  .logo span { font-size: 13px; color: var(--muted); }
  .header-right { display: flex; align-items: center; gap: 16px; }
  .badge-grupo {
    background: var(--purple); color: white; padding: 4px 12px;
    border-radius: 20px; font-size: 12px; font-weight: 600;
  }
  .mes-selector {
    background: var(--bg3); border: 1px solid var(--border); color: var(--text);
    padding: 6px 12px; border-radius: 8px; font-size: 13px; cursor: pointer;
  }

  /* MAIN */
  .main { padding: 28px; max-width: 1600px; margin: 0 auto; }
  
  /* SECTION TITLE */
  .section-title {
    font-size: 13px; font-weight: 600; color: var(--muted);
    text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 16px;
  }

  /* CONSOLIDADO TOP ROW */
  .kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 16px; margin-bottom: 28px; }
  .kpi-card {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 20px;
  }
  .kpi-card .label { font-size: 12px; color: var(--muted); margin-bottom: 8px; }
  .kpi-card .value { font-size: 26px; font-weight: 700; }
  .kpi-card .sub { font-size: 12px; color: var(--muted); margin-top: 4px; }
  .kpi-card .delta { font-size: 12px; margin-top: 6px; font-weight: 600; }
  .delta.pos { color: var(--green); }
  .delta.neg { color: var(--red); }
  .kpi-azul .value { color: var(--blue); }
  .kpi-verde .value { color: var(--green); }
  .kpi-naranja .value { color: var(--orange); }
  .kpi-purple .value { color: var(--purple); }
  .kpi-teal .value { color: var(--teal); }

  /* HOTEL CARDS */
  .hotels-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 28px; }
  .hotel-card {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: var(--radius); overflow: hidden;
  }
  .hotel-header {
    padding: 16px 20px; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; justify-content: space-between;
  }
  .hotel-name { font-size: 15px; font-weight: 700; }
  .hotel-meta { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .hotel-badge {
    font-size: 11px; font-weight: 600; padding: 3px 10px;
    border-radius: 12px; background: rgba(26,115,232,0.2); color: var(--blue);
  }
  .hotel-kpis { padding: 16px 20px; display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
  .hk-item .hk-label { font-size: 11px; color: var(--muted); }
  .hk-item .hk-value { font-size: 18px; font-weight: 700; margin-top: 2px; }
  .hotel-footer { padding: 12px 20px; background: var(--bg3); border-top: 1px solid var(--border); }
  .hotel-footer-row { display: flex; justify-content: space-between; align-items: center; }
  .alerta-pill {
    font-size: 11px; font-weight: 600; padding: 3px 10px; border-radius: 12px;
  }
  .alerta-0 { background: rgba(29,185,84,0.15); color: var(--green); }
  .alerta-1, .alerta-2 { background: rgba(255,152,0,0.15); color: var(--orange); }
  .alerta-3 { background: rgba(224,82,82,0.15); color: var(--red); }

  /* CHARTS ROW */
  .charts-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-bottom: 28px; }
  .chart-card {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 20px;
  }
  .chart-title { font-size: 13px; font-weight: 600; margin-bottom: 16px; color: var(--text); }
  .chart-wrap { position: relative; height: 220px; }

  /* BENCHMARKING TABLE */
  .table-card {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: var(--radius); overflow: hidden; margin-bottom: 28px;
  }
  .table-header { padding: 16px 20px; border-bottom: 1px solid var(--border); }
  .table-header h3 { font-size: 14px; font-weight: 600; }
  table { width: 100%; border-collapse: collapse; }
  th { padding: 10px 16px; text-align: left; font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.05em; background: var(--bg3); }
  td { padding: 12px 16px; font-size: 13px; border-top: 1px solid var(--border); }
  tr:hover td { background: rgba(255,255,255,0.02); }
  .td-hotel { font-weight: 600; }
  .td-best { color: var(--green); font-weight: 700; }
  .td-worst { color: var(--red); }
  .td-mid { color: var(--orange); }

  /* ALERTAS CROSS-HOTEL */
  .alertas-grid { display: grid; grid-template-columns: 1fr; gap: 10px; margin-bottom: 28px; }
  .alerta-row {
    background: var(--bg2); border: 1px solid var(--border);
    border-radius: 8px; padding: 14px 18px;
    display: flex; align-items: center; gap: 14px;
  }
  .alerta-nivel {
    font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 8px;
    white-space: nowrap; min-width: 70px; text-align: center;
  }
  .nivel-critico { background: rgba(224,82,82,0.2); color: var(--red); }
  .nivel-aviso { background: rgba(255,152,0,0.2); color: var(--orange); }
  .nivel-info { background: rgba(26,115,232,0.2); color: var(--blue); }
  .alerta-msg { font-size: 13px; color: var(--text); }
  .alerta-hotel { font-size: 11px; color: var(--muted); margin-top: 2px; }

  /* RESPONSIVE */
  @media (max-width: 1100px) {
    .kpi-grid { grid-template-columns: repeat(3, 1fr); }
    .hotels-grid { grid-template-columns: repeat(2, 1fr); }
    .charts-grid { grid-template-columns: 1fr 1fr; }
  }
  @media (max-width: 700px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    .hotels-grid { grid-template-columns: 1fr; }
    .charts-grid { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="logo">
    <div class="logo-icon">Y</div>
    <div>
      <h1>Yve</h1>
      <span>Dashboard Multi-Hotel</span>
    </div>
  </div>
  <div class="header-right">
    <span class="badge-grupo">Grupo Calipolis — 3 hoteles</span>
    <select class="mes-selector" onchange="cambiarMes(this.value)" id="mesSel">
      <option value="2025-07">Julio 2025</option>
      <option value="2025-06">Junio 2025</option>
      <option value="2025-05">Mayo 2025</option>
    </select>
  </div>
</div>

<div class="main">

  <!-- KPIs CONSOLIDADOS -->
  <p class="section-title">Grupo Consolidado — <span id="mesLabel">Julio 2025</span></p>
  <div class="kpi-grid" id="kpiGrid">
    <div class="kpi-card kpi-azul"><div class="label">Ingresos Totales</div><div class="value" id="k-ingresos">—</div><div class="delta pos" id="k-ingresos-d">—</div></div>
    <div class="kpi-card kpi-verde"><div class="label">Ocupación Media</div><div class="value" id="k-ocup">—</div><div class="sub">3 propiedades</div></div>
    <div class="kpi-card kpi-teal"><div class="label">ADR Medio</div><div class="value" id="k-adr">—</div><div class="sub">Average Daily Rate</div></div>
    <div class="kpi-card kpi-purple"><div class="label">GOP Grupo</div><div class="value" id="k-gop">—</div><div class="delta" id="k-gop-pct">—</div></div>
    <div class="kpi-card kpi-naranja"><div class="label">Alertas Activas</div><div class="value" id="k-alertas">—</div><div class="sub">Todas las propiedades</div></div>
  </div>

  <!-- HOTEL CARDS -->
  <p class="section-title">Por Propiedad</p>
  <div class="hotels-grid" id="hotelCards">
    <!-- generado por JS -->
  </div>

  <!-- CHARTS -->
  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-title">Ingresos por Hotel (€)</div>
      <div class="chart-wrap"><canvas id="chartIngresos"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Ocupación % por Hotel</div>
      <div class="chart-wrap"><canvas id="chartOcupacion"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">GOP % por Hotel</div>
      <div class="chart-wrap"><canvas id="chartGop"></canvas></div>
    </div>
  </div>

  <!-- BENCHMARKING TABLE -->
  <div class="table-card">
    <div class="table-header"><h3>Benchmarking entre propiedades — <span id="mesTable">Julio 2025</span></h3></div>
    <table id="benchTable">
      <thead><tr>
        <th>Hotel</th><th>Hab.</th><th>Ocup%</th><th>ADR €</th><th>RevPAR €</th>
        <th>Ingresos €</th><th>Food Cost%</th><th>GOP €</th><th>GOP%</th>
        <th>AP Pend.</th><th>AR Pend.</th><th>Estado</th>
      </tr></thead>
      <tbody id="benchBody"></tbody>
    </table>
  </div>

  <!-- ALERTAS CROSS-HOTEL -->
  <p class="section-title">Alertas del Grupo</p>
  <div id="alertasContainer" class="alertas-grid">
    <div class="alerta-row"><span class="alerta-nivel nivel-info">INFO</span><div><div class="alerta-msg">Cargando alertas...</div></div></div>
  </div>

</div>

<script>
const COLORES = {
  'CAL01': '#1a73e8',
  'CAL02': '#1db954',
  'CAL03': '#7c4dff',
};
const MESES_LABEL = {
  '2025-07': 'Julio 2025',
  '2025-06': 'Junio 2025',
  '2025-05': 'Mayo 2025',
};

let chartIngresos, chartOcupacion, chartGop;

function fmt(n) { return n.toLocaleString('es-ES', {minimumFractionDigits:0, maximumFractionDigits:0}); }
function fmtE(n) { return fmt(n) + ' €'; }
function fmtP(n) { return n.toFixed(1) + '%'; }

function cambiarMes(mes) {
  document.getElementById('mesLabel').textContent = MESES_LABEL[mes];
  document.getElementById('mesTable').textContent = MESES_LABEL[mes];
  cargarDatos(mes);
}

async function cargarDatos(mes) {
  const resp = await fetch('/api/multihotel/kpis?mes=' + mes);
  const data = await resp.json();
  renderizarDatos(data, mes);
}

function renderizarDatos(data, mes) {
  const hoteles = data.hoteles;

  // KPIs CONSOLIDADOS
  const totalIngresos = hoteles.reduce((s, h) => s + h.total_ingresos, 0);
  const mediaOcup = hoteles.reduce((s, h) => s + h.ocupacion_pct, 0) / hoteles.length;
  const mediaAdr = hoteles.reduce((s, h) => s + h.adr_eur, 0) / hoteles.length;
  const totalGop = hoteles.reduce((s, h) => s + h.gop_eur, 0);
  const gopPct = totalGop / totalIngresos * 100;
  const totalAlertas = hoteles.reduce((s, h) => s + h.alertas_activas, 0);
  const totalFcPend = hoteles.reduce((s, h) => s + h.facturas_ap_pendientes, 0);

  document.getElementById('k-ingresos').textContent = fmtE(totalIngresos);
  document.getElementById('k-ingresos-d').textContent = fmtE(totalFcPend) + ' en facturas pend.';
  document.getElementById('k-ocup').textContent = fmtP(mediaOcup);
  document.getElementById('k-adr').textContent = fmtE(mediaAdr);
  document.getElementById('k-gop').textContent = fmtE(totalGop);
  document.getElementById('k-gop-pct').textContent = fmtP(gopPct) + ' GOP%';
  document.getElementById('k-gop-pct').className = 'delta ' + (gopPct >= 20 ? 'pos' : 'neg');
  document.getElementById('k-alertas').textContent = totalAlertas;

  // HOTEL CARDS
  const grid = document.getElementById('hotelCards');
  grid.innerHTML = hoteles.map(h => `
    <div class="hotel-card">
      <div class="hotel-header">
        <div>
          <div class="hotel-name">${h.hotel_nombre}</div>
          <div class="hotel-meta">${h.ciudad} · ${h.habitaciones} hab.</div>
        </div>
        <span class="hotel-badge">${h.hotel_id}</span>
      </div>
      <div class="hotel-kpis">
        <div class="hk-item"><div class="hk-label">Ocupación</div><div class="hk-value" style="color:#1a73e8">${fmtP(h.ocupacion_pct)}</div></div>
        <div class="hk-item"><div class="hk-label">ADR</div><div class="hk-value">${fmtE(h.adr_eur)}</div></div>
        <div class="hk-item"><div class="hk-label">RevPAR</div><div class="hk-value">${fmtE(h.revpar_eur)}</div></div>
        <div class="hk-item"><div class="hk-label">Ingresos</div><div class="hk-value" style="color:#1db954">${fmtE(h.total_ingresos)}</div></div>
        <div class="hk-item"><div class="hk-label">GOP</div><div class="hk-value">${fmtE(h.gop_eur)}</div></div>
        <div class="hk-item"><div class="hk-label">GOP%</div><div class="hk-value" style="color:${h.gop_pct>=20?'#1db954':'#e05252'}">${fmtP(h.gop_pct)}</div></div>
      </div>
      <div class="hotel-footer">
        <div class="hotel-footer-row">
          <span style="font-size:12px;color:#8892a4">AP: ${h.facturas_ap_pendientes} pend · AR: ${h.facturas_ar_pendientes} pend</span>
          <span class="alerta-pill alerta-${Math.min(h.alertas_activas,3)}">${h.alertas_activas === 0 ? '✓ Sin alertas' : '⚠ ' + h.alertas_activas + ' alerta' + (h.alertas_activas>1?'s':'')}</span>
        </div>
      </div>
    </div>
  `).join('');

  // CHARTS
  const labels = hoteles.map(h => h.hotel_nombre.replace('Hotel ','').replace('Calipolis ','C.'));
  const colores = hoteles.map(h => COLORES[h.hotel_id] || '#1a73e8');

  if (chartIngresos) chartIngresos.destroy();
  if (chartOcupacion) chartOcupacion.destroy();
  if (chartGop) chartGop.destroy();

  const opts = { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } },
    scales: { x: { ticks: { color:'#8892a4', font:{size:11} }, grid: { color:'#2e3248' } },
               y: { ticks: { color:'#8892a4', font:{size:11} }, grid: { color:'#2e3248' } } } };

  chartIngresos = new Chart(document.getElementById('chartIngresos'), {
    type:'bar', data:{ labels, datasets:[{ data: hoteles.map(h=>h.total_ingresos), backgroundColor: colores, borderRadius: 6 }] },
    options: {...opts, scales:{...opts.scales, y:{...opts.scales.y, ticks:{...opts.scales.y.ticks, callback: v => fmt(v/1000)+'k€'}}}}
  });
  chartOcupacion = new Chart(document.getElementById('chartOcupacion'), {
    type:'bar', data:{ labels, datasets:[{ data: hoteles.map(h=>h.ocupacion_pct), backgroundColor: colores, borderRadius: 6 }] },
    options: {...opts, scales:{...opts.scales, y:{...opts.scales.y, max:100, ticks:{...opts.scales.y.ticks, callback: v => v+'%'}}}}
  });
  chartGop = new Chart(document.getElementById('chartGop'), {
    type:'bar', data:{ labels, datasets:[{ data: hoteles.map(h=>h.gop_pct), backgroundColor: colores, borderRadius: 6 }] },
    options: {...opts, scales:{...opts.scales, y:{...opts.scales.y, ticks:{...opts.scales.y.ticks, callback: v => v+'%'}}}}
  });

  // BENCHMARKING TABLE
  // Determinar best/worst por columna
  const maxOcup = Math.max(...hoteles.map(h=>h.ocupacion_pct));
  const maxAdr = Math.max(...hoteles.map(h=>h.adr_eur));
  const maxGopPct = Math.max(...hoteles.map(h=>h.gop_pct));
  const minFcPct = Math.min(...hoteles.filter(h=>h.food_cost_pct>0).map(h=>h.food_cost_pct));

  document.getElementById('benchBody').innerHTML = hoteles.map(h => {
    const clsOcup = h.ocupacion_pct===maxOcup ? 'td-best' : '';
    const clsAdr = h.adr_eur===maxAdr ? 'td-best' : '';
    const clsGop = h.gop_pct===maxGopPct ? 'td-best' : '';
    const clsFc = h.food_cost_pct>0 && h.food_cost_pct===minFcPct ? 'td-best' : (h.food_cost_pct>33 ? 'td-worst' : '');
    const oob = h.out_of_balance_dias > 0 ? `<span style="color:#e05252">⚠ ${h.out_of_balance_dias} OOB</span>` : '<span style="color:#1db954">✓ OK</span>';
    return `<tr>
      <td class="td-hotel">${h.hotel_nombre}</td>
      <td>${h.habitaciones}</td>
      <td class="${clsOcup}">${fmtP(h.ocupacion_pct)}</td>
      <td class="${clsAdr}">${fmtE(h.adr_eur)}</td>
      <td>${fmtE(h.revpar_eur)}</td>
      <td>${fmtE(h.total_ingresos)}</td>
      <td class="${clsFc}">${h.food_cost_pct > 0 ? fmtP(h.food_cost_pct) : '—'}</td>
      <td>${fmtE(h.gop_eur)}</td>
      <td class="${clsGop}">${fmtP(h.gop_pct)}</td>
      <td>${h.facturas_ap_pendientes}</td>
      <td>${h.facturas_ar_pendientes}</td>
      <td>${oob}</td>
    </tr>`;
  }).join('');

  // ALERTAS CROSS-HOTEL
  renderizarAlertas(data.alertas);
}

function renderizarAlertas(alertas) {
  const cont = document.getElementById('alertasContainer');
  if (!alertas.length) {
    cont.innerHTML = '<div class="alerta-row"><span class="alerta-nivel nivel-info">INFO</span><div><div class="alerta-msg">No hay alertas activas en el grupo.</div></div></div>';
    return;
  }
  cont.innerHTML = alertas.map(a => `
    <div class="alerta-row">
      <span class="alerta-nivel nivel-${a.nivel.toLowerCase()}">${a.nivel}</span>
      <div>
        <div class="alerta-msg">${a.mensaje}</div>
        <div class="alerta-hotel">${a.hotel}</div>
      </div>
    </div>
  `).join('');
}

// Cargar al inicio
cargarDatos('2025-07');
</script>
</body>
</html>
"""

# ─── API ───────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/api/multihotel/kpis")
def api_kpis():
    from flask import request
    mes = request.args.get("mes", "2025-07")
    df = cargar_kpis()
    hoteles_meta = cargar_hoteles()

    df_mes = df[df["mes"] == mes].copy()
    hoteles_out = []
    alertas = []

    meta_map = {h["id"]: h for h in hoteles_meta}

    for _, row in df_mes.iterrows():
        hotel_id = row["hotel_id"]
        meta = meta_map.get(hotel_id, {})
        h = {
            "hotel_id": hotel_id,
            "hotel_nombre": row["hotel_nombre"],
            "ciudad": meta.get("ciudad", "—"),
            "habitaciones": int(row["habitaciones"]),
            "ocupacion_pct": float(row["ocupacion_pct"]),
            "adr_eur": float(row["adr_eur"]),
            "revpar_eur": float(row["revpar_eur"]),
            "total_ingresos": float(row["total_ingresos"]),
            "ingresos_rooms": float(row["ingresos_rooms"]),
            "ingresos_fb": float(row["ingresos_fb"]),
            "gop_eur": float(row["gop_eur"]),
            "gop_pct": float(row["gop_pct"]),
            "food_cost_pct": float(row["food_cost_pct"]),
            "facturas_ap_pendientes": int(row["facturas_ap_pendientes"]),
            "facturas_ar_pendientes": int(row["facturas_ar_pendientes"]),
            "alertas_activas": int(row["alertas_activas"]),
            "estado_oracle": row["estado_oracle"],
            "out_of_balance_dias": int(row["out_of_balance_dias"]),
        }
        hoteles_out.append(h)

        # Generar alertas cross-hotel
        if h["ocupacion_pct"] < 60:
            alertas.append({"nivel": "AVISO", "tipo": "OCUPACION_BAJA", "hotel": h["hotel_nombre"],
                "mensaje": f"Ocupación {h['ocupacion_pct']}% por debajo del 60%. Revisar estrategia de precios."})
        if h["gop_pct"] < 18:
            alertas.append({"nivel": "CRITICO", "tipo": "GOP_BAJO", "hotel": h["hotel_nombre"],
                "mensaje": f"GOP {h['gop_pct']}% por debajo del umbral mínimo (18%). Revisar estructura de costes."})
        if h["food_cost_pct"] > 33 and h["food_cost_pct"] > 0:
            alertas.append({"nivel": "AVISO", "tipo": "FOOD_COST_ALTO", "hotel": h["hotel_nombre"],
                "mensaje": f"Food Cost {h['food_cost_pct']}% supera el 33%. Revisar mermas y precios de compra."})
        if h["facturas_ap_pendientes"] > 5:
            alertas.append({"nivel": "AVISO", "tipo": "FACTURAS_PENDIENTES", "hotel": h["hotel_nombre"],
                "mensaje": f"{h['facturas_ap_pendientes']} facturas AP pendientes de aprobación."})
        if h["out_of_balance_dias"] > 0:
            alertas.append({"nivel": "CRITICO", "tipo": "OUT_OF_BALANCE", "hotel": h["hotel_nombre"],
                "mensaje": f"{h['out_of_balance_dias']} día(s) con Out of Balance en DRR. Revisión inmediata necesaria."})

    # Ordenar: críticos primero
    alertas.sort(key=lambda a: 0 if a["nivel"] == "CRITICO" else 1)

    return jsonify({"hoteles": hoteles_out, "alertas": alertas, "mes": mes})

@app.route("/api/multihotel/hoteles")
def api_hoteles():
    return jsonify(cargar_hoteles())

# ─── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 50)
    print("  YVE — DASHBOARD MULTI-HOTEL")
    print("  http://localhost:5008")
    print("=" * 50)
    app.run(debug=True, port=5008)
