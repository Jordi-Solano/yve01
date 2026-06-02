"""
Parche para dashboard.py — añade sección Módulo AP
"""
import re

ruta = "/sessions/brave-affectionate-hawking/mnt/yve01/dashboard.py"

with open(ruta, "r", encoding="utf-8") as f:
    contenido = f.read()

# ══════════════════════════════════════════════════════════════════════════════
# 1) NUEVO CÓDIGO PYTHON — rutas AP + helpers
# ══════════════════════════════════════════════════════════════════════════════

NUEVO_PYTHON = '''
# ── Pipeline AP ────────────────────────────────────────────────────────────
_pipeline_ap_running = False
_pipeline_ap_lock    = threading.Lock()
FACTURAS_AP_DIR      = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_AP_DIR  = os.path.join(BASE_DIR, "aprobaciones")

def cargar_datos_ap():
    """Carga facturas AP contabilizadas o procesadas."""
    df, _ = cargar_ultimo_excel("facturas_contabilizadas_*.xlsx", FACTURAS_AP_DIR)
    if df is None:
        df, _ = cargar_ultimo_excel("facturas_ap_*.xlsx", FACTURAS_AP_DIR)
    if df is None:
        return pd.DataFrame()

    # Merge con aprobaciones AP
    apro_path = os.path.join(APROBACIONES_AP_DIR, "aprobaciones_ap.xlsx")
    if os.path.exists(apro_path):
        try:
            df_apro = pd.read_excel(apro_path)
            if not df_apro.empty and "numero_factura" in df_apro.columns:
                ultimas = df_apro.sort_values("fecha_hora").groupby("numero_factura").last().reset_index()
                df = df.merge(ultimas[["numero_factura","accion","comentario"]], on="numero_factura", how="left")
        except Exception:
            pass
    return df


def calcular_stats_ap(df):
    """Calcula estadísticas del módulo AP."""
    if df.empty:
        return {"total":0,"importe":0,"matches":0,"discrepancias":0,"sin_po":0,
                "alertas_consumo":0,"manual":0,"aprobadas":0,"rechazadas":0}
    total   = len(df)
    importe = 0.0
    for c in ["total_factura","importe_total","total"]:
        if c in df.columns:
            importe = df[c].apply(safe_float).sum()
            break
    # Conteo por estado de matching
    est_col = None
    for c in ["estado_matching","estado","matching_estado"]:
        if c in df.columns:
            est_col = c
            break
    matches     = 0
    discrepancias = 0
    sin_po      = 0
    alertas     = 0
    manual      = 0
    if est_col:
        estados = df[est_col].astype(str).str.upper()
        matches        = int((estados.isin(["MATCH_CORRECTO","MATCH_3WAY_OK"])).sum())
        discrepancias  = int((estados == "DISCREPANCIA_PO").sum())
        sin_po         = int((estados == "SIN_PO").sum())
        alertas        = int((estados == "ALERTA_CONSUMO").sum())
        manual         = int((estados == "REVISAR_MANUAL").sum())
    else:
        # Fallback: revisar columna cuenta_contable
        if "cuenta_contable" in df.columns:
            manual = int((df["cuenta_contable"].astype(str).str.upper() == "REVISAR_MANUAL").sum())
    # Aprobaciones
    aprobadas  = 0
    rechazadas = 0
    if "accion" in df.columns:
        aprobadas  = int((df["accion"].astype(str).str.upper() == "APROBADA").sum())
        rechazadas = int((df["accion"].astype(str).str.upper() == "RECHAZADA").sum())
    return {"total":total,"importe":round(importe,2),"matches":matches,
            "discrepancias":discrepancias,"sin_po":sin_po,"alertas_consumo":alertas,
            "manual":manual,"aprobadas":aprobadas,"rechazadas":rechazadas}


def df_ap_a_lista(df):
    """Convierte DataFrame AP a lista de dicts."""
    rows = []
    if df.empty:
        return rows
    for _, r in df.iterrows():
        total = 0.0
        for c in ["total_factura","importe_total","total"]:
            if c in df.columns:
                total = safe_float(r.get(c, 0))
                break
        est = str(r.get("estado_matching", r.get("estado", ""))).strip().upper()
        rows.append({
            "numero_factura":    str(r.get("numero_factura","")).strip() or "N/D",
            "proveedor":         str(r.get("nombre_proveedor","")).strip() or "Desconocido",
            "tipo":              str(r.get("tipo_proveedor","")).strip().upper() or "OTRAS",
            "total":             total,
            "cuenta_contable":   str(r.get("cuenta_contable","")).strip() or "—",
            "estado":            est or "PENDIENTE",
            "accion":            str(r.get("accion","")).strip().upper() or "",
            "detalle_alerta":    str(r.get("detalle_alerta","")).strip() or "",
        })
    return rows


@app.route("/api/stats_ap")
def api_stats_ap():
    df = cargar_datos_ap()
    return jsonify(calcular_stats_ap(df))


@app.route("/api/facturas_ap")
def api_facturas_ap():
    df = cargar_datos_ap()
    return jsonify(df_ap_a_lista(df))


@app.route("/api/procesar_ap")
def api_procesar_ap():
    global _pipeline_ap_running
    scripts = [
        ("lector_facturas_ap.py",  "Leyendo facturas PDF proveedores"),
        ("matching_ap_otras.py",   "Matching facturas OTRAS vs POs"),
        ("matching_ap_fb.py",      "Matching 3-way F&B"),
        ("asignador_cuentas.py",   "Asignando cuentas contables"),
        ("generador_emails_ap.py", "Generando emails incidencias"),
    ]

    def generar():
        global _pipeline_ap_running
        with _pipeline_ap_lock:
            if _pipeline_ap_running:
                yield "data: Ya hay un proceso AP en ejecucion — espera\\n\\n"
                return
            _pipeline_ap_running = True
        try:
            yield "data: INICIO\\n\\n"
            ok_total = True
            for script, label in scripts:
                ruta = os.path.join(BASE_DIR, script)
                yield "data: >> " + label + "...\\n\\n"
                if not os.path.exists(ruta):
                    yield "data: ERROR: " + script + " no encontrado\\n\\n"
                    ok_total = False
                    continue
                try:
                    res = subprocess.run(
                        [sys.executable, ruta],
                        capture_output=True, text=True, timeout=180, cwd=BASE_DIR
                    )
                    for linea in (res.stdout + res.stderr).splitlines():
                        linea = linea.strip()
                        if linea:
                            yield "data: " + linea + "\\n\\n"
                    if res.returncode == 0:
                        yield "data: OK " + script + " completado\\n\\n"
                    else:
                        yield "data: ERROR en " + script + " (codigo " + str(res.returncode) + ")\\n\\n"
                        ok_total = False
                except subprocess.TimeoutExpired:
                    yield "data: TIMEOUT: " + script + " tardo demasiado\\n\\n"
                    ok_total = False
                except Exception as exc:
                    yield "data: ERROR: " + str(exc) + "\\n\\n"
                    ok_total = False
            yield "data: PIPELINE_COMPLETO\\n\\n" if ok_total else "data: PIPELINE_CON_ERRORES\\n\\n"
        finally:
            _pipeline_ap_running = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )

'''

# Insertar antes de @app.route("/")
marcador = '@app.route("/")\ndef index():'
if marcador not in contenido:
    print("ERROR: marcador no encontrado")
    exit(1)

contenido = contenido.replace(marcador, NUEVO_PYTHON + marcador)
print("✓ Código Python AP insertado")

# ══════════════════════════════════════════════════════════════════════════════
# 2) ACTUALIZAR TÍTULO Y AÑADIR SECCIÓN AP AL HTML
# ══════════════════════════════════════════════════════════════════════════════

# 2a) Cambiar título
contenido = contenido.replace(
    "<title>Yve.01 — Dashboard AR</title>",
    "<title>Yve.01 — Dashboard</title>"
)
print("✓ Título actualizado")

# 2b) Añadir estilos AP después del comentario de estilos del body
ESTILOS_AP = """
/* ── Tabs ─────────────────────────────────────────────── */
.tabs{display:flex;gap:8px;margin-bottom:24px;border-bottom:1px solid var(--s2);padding-bottom:0}
.tab{padding:10px 20px;background:none;border:none;color:var(--mut);cursor:pointer;font-size:.9rem;font-weight:600;border-bottom:3px solid transparent;transition:.2s}
.tab.active{color:var(--acc2);border-bottom-color:var(--acc2)}
.panel{display:none}.panel.active{display:block}
/* ── AP Cards ─────────────────────────────────────────── */
.ap-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.75rem;font-weight:700;letter-spacing:.04em}
.ap-badge.fb{background:rgba(139,92,246,.2);color:#c4b5fd}
.ap-badge.otras{background:rgba(59,130,246,.2);color:#93c5fd}
.ap-badge.ok{background:rgba(34,197,94,.2);color:#86efac}
.ap-badge.disc{background:rgba(239,68,68,.2);color:#fca5a5}
.ap-badge.alerta{background:rgba(59,130,246,.15);color:#93c5fd}
.ap-badge.sinpo{background:rgba(234,179,8,.2);color:#fde047}
.ap-badge.manual{background:rgba(249,115,22,.2);color:#fed7aa}
.alerta-box{background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.3);border-radius:6px;padding:8px 12px;margin-top:8px;font-size:.8rem;color:var(--acc3)}
.disc-box{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);border-radius:6px;padding:8px 12px;margin-top:8px;font-size:.8rem;color:#fca5a5}
"""

# Insertar estilos antes del cierre </style>
contenido = contenido.replace("</style>", ESTILOS_AP + "</style>", 1)
print("✓ Estilos AP añadidos")

# 2c) Añadir tabs + sección AP en el HTML
# Buscar dónde empieza el contenido principal (tras el <header>)
# Buscamos el div#stats para insertar los tabs antes
TABS_HTML = """
<!-- ── Tabs AR / AP ──────────────────────────────────────── -->
<div class="tabs">
  <button class="tab active" onclick="switchTab('ar',this)">📥 Módulo AR — OTAs</button>
  <button class="tab" onclick="switchTab('ap',this)">📦 Módulo AP — Proveedores</button>
</div>
<div id="panel-ar" class="panel active">
"""

# Buscamos el inicio del panel AR (div stats) y añadimos el wrapper
marcador_stats = '<div id="stats" class="grid">'
if marcador_stats not in contenido:
    # Intentar variante
    marcador_stats = '<div id="stats"'

contenido = contenido.replace(
    marcador_stats,
    TABS_HTML + marcador_stats,
    1
)
print("✓ Tabs insertados")

# 2d) Cerrar el panel AR antes de if __name__ (justo antes del cierre HTML)
# Buscamos el cierre </body> en el HTML string
contenido = contenido.replace(
    "</body>\\n</html>",
    """</div><!-- /panel-ar -->

<!-- ════════════════════════════════════════════════════════
     PANEL AP
═════════════════════════════════════════════════════════ -->
<div id="panel-ap" class="panel">

<!-- Stats AP -->
<div id="stats-ap" class="grid">
  <div class="card"><div class="lbl">Total Facturas AP</div><div class="val" id="ap-total">—</div></div>
  <div class="card"><div class="lbl">Importe Total AP</div><div class="val" id="ap-importe">—</div></div>
  <div class="card" style="border-top:3px solid var(--grn)"><div class="lbl">Matches Correctos</div><div class="val green" id="ap-matches">—</div></div>
  <div class="card" style="border-top:3px solid var(--red)"><div class="lbl">Discrepancias PO</div><div class="val red" id="ap-disc">—</div></div>
  <div class="card" style="border-top:3px solid var(--yel)"><div class="lbl">Sin Orden de Compra</div><div class="val yel" id="ap-sinpo">—</div></div>
  <div class="card" style="border-top:3px solid var(--acc)"><div class="lbl">Alertas Consumo F&B</div><div class="val" style="color:var(--acc2)" id="ap-alertas">—</div></div>
  <div class="card" style="border-top:3px solid var(--ora)"><div class="lbl">Cuenta Manual Req.</div><div class="val ora" id="ap-manual">—</div></div>
  <div class="card" style="border-top:3px solid var(--grn)"><div class="lbl">Aprobadas</div><div class="val green" id="ap-aprobadas">—</div></div>
</div>

<!-- Botón pipeline AP -->
<div style="display:flex;gap:12px;margin:20px 0;flex-wrap:wrap">
  <button class="btn" id="btnAP" onclick="procesarAP()">⚙️ Procesar Facturas AP</button>
  <a class="btn" href="http://localhost:5002" target="_blank" style="text-decoration:none">🔍 Aprobar Facturas AP</a>
</div>

<!-- Tabla AP -->
<div class="card" style="padding:0;overflow:hidden">
  <div style="padding:16px 20px;border-bottom:1px solid var(--s2);display:flex;align-items:center;justify-content:space-between">
    <span style="font-weight:700;color:var(--acc2)">Facturas AP</span>
    <span id="ap-count" style="font-size:.8rem;color:var(--mut)"></span>
  </div>
  <div style="overflow-x:auto">
  <table>
    <thead><tr>
      <th>Factura</th><th>Proveedor</th><th>Tipo</th>
      <th>Total</th><th>Cuenta</th><th>Matching</th><th>Aprobación</th>
    </tr></thead>
    <tbody id="ap-tbody"></tbody>
  </table>
  </div>
</div>

</div><!-- /panel-ap -->

</body>\\n</html>""",
    1
)
print("✓ Panel AP añadido al HTML")

# 2e) Añadir JS para el módulo AP justo antes del cierre </script>
JS_AP = """

// ══════════════════════════════════════════════════════════════
// MÓDULO AP — JavaScript
// ══════════════════════════════════════════════════════════════

function switchTab(tab, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('panel-' + tab).classList.add('active');
}

function fmtEurAP(v) {
  if (!v && v !== 0) return '—';
  return new Intl.NumberFormat('es-ES', {style:'currency',currency:'EUR',maximumFractionDigits:0}).format(v);
}

function estadoBadgeAP(est) {
  const m = {
    'MATCH_CORRECTO':'ok','MATCH_3WAY_OK':'ok',
    'DISCREPANCIA_PO':'disc','DISCREPANCIA':'disc',
    'SIN_PO':'sinpo',
    'ALERTA_CONSUMO':'alerta',
    'REVISAR_MANUAL':'manual',
    'PENDIENTE':''
  };
  const cls = m[est] || '';
  return `<span class="ap-badge ${cls}">${est || 'PENDIENTE'}</span>`;
}

async function loadAP() {
  try {
    const [stats, facts] = await Promise.all([
      fetch('/api/stats_ap').then(r=>r.json()),
      fetch('/api/facturas_ap').then(r=>r.json()),
    ]);

    document.getElementById('ap-total').textContent    = stats.total ?? '—';
    document.getElementById('ap-importe').textContent  = fmtEurAP(stats.importe);
    document.getElementById('ap-matches').textContent  = stats.matches ?? '—';
    document.getElementById('ap-disc').textContent     = stats.discrepancias ?? '—';
    document.getElementById('ap-sinpo').textContent    = stats.sin_po ?? '—';
    document.getElementById('ap-alertas').textContent  = stats.alertas_consumo ?? '—';
    document.getElementById('ap-manual').textContent   = stats.manual ?? '—';
    document.getElementById('ap-aprobadas').textContent= stats.aprobadas ?? '—';

    const tbody = document.getElementById('ap-tbody');
    tbody.innerHTML = '';
    document.getElementById('ap-count').textContent = facts.length + ' facturas';

    facts.forEach(f => {
      const tr = document.createElement('tr');
      const tipoCls = f.tipo === 'FB' ? 'fb' : 'otras';
      const accionHtml = f.accion === 'APROBADA'
        ? '<span class="badge ok">✓ Aprobada</span>'
        : f.accion === 'RECHAZADA'
          ? '<span class="badge err">✗ Rechazada</span>'
          : '<span class="badge" style="background:rgba(100,116,139,.3);color:var(--mut)">Pendiente</span>';

      let alertaHtml = '';
      if (f.estado === 'ALERTA_CONSUMO' && f.detalle_alerta) {
        alertaHtml = `<div class="alerta-box">${f.detalle_alerta}</div>`;
      } else if ((f.estado === 'DISCREPANCIA_PO' || f.estado === 'DISCREPANCIA') && f.detalle_alerta) {
        alertaHtml = `<div class="disc-box">${f.detalle_alerta}</div>`;
      }

      tr.innerHTML = `
        <td><strong>${f.numero_factura}</strong></td>
        <td>${f.proveedor}</td>
        <td><span class="ap-badge ${tipoCls}">${f.tipo}</span></td>
        <td>${fmtEurAP(f.total)}</td>
        <td><code style="font-size:.8rem;color:var(--acc3)">${f.cuenta_contable}</code></td>
        <td>${estadoBadgeAP(f.estado)}${alertaHtml}</td>
        <td>${accionHtml}</td>
      `;
      tbody.appendChild(tr);
    });
  } catch(e) {
    console.warn('Error cargando datos AP:', e);
  }
}

function procesarAP() {
  const btn  = document.getElementById('btnAP');
  const log  = document.getElementById('log');
  const spin = document.getElementById('spinner');
  const lbl  = document.getElementById('btnLabel');
  const icon = document.getElementById('modalIcon');
  const title= document.getElementById('modalTitle');
  const btnCl= document.getElementById('btnClose');

  btn.disabled = true;
  spin.style.display = 'block';
  lbl.textContent = 'Procesando AP...';
  log.innerHTML = '';
  btnCl.disabled = true;
  icon.textContent = '⚙️';
  title.textContent = 'Pipeline AP — Procesando...';
  document.getElementById('overlay').classList.add('on');

  const src = new EventSource('/api/procesar_ap');

  src.onmessage = ev => {
    const txt = ev.data;
    const p = document.createElement('p');
    if      (txt === 'PIPELINE_COMPLETO')    p.className = 'l-ok';
    else if (txt === 'PIPELINE_CON_ERRORES') p.className = 'l-err';
    else if (txt.startsWith('OK '))          p.className = 'l-ok';
    else if (txt.startsWith('ERROR') || txt.startsWith('TIMEOUT')) p.className = 'l-err';
    else if (txt.startsWith('>> ') || txt === 'INICIO') p.className = 'l-info';
    else if (txt.includes('✓'))              p.className = 'l-ok';
    else if (txt.includes('✗'))              p.className = 'l-err';
    else p.className = 'l-dim';
    p.textContent = txt;
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = ok ? 'Pipeline AP completado' : 'Pipeline AP con errores';
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = '⚙️ Procesar Facturas AP';
      btnCl.disabled = false;
      setTimeout(loadAP, 800);
    }
  };

  src.onerror = () => {
    src.close();
    const p = document.createElement('p');
    p.className = 'l-err';
    p.textContent = 'ERROR: conexión con el servidor perdida';
    log.appendChild(p);
    btn.disabled = false;
    spin.style.display = 'none';
    lbl.textContent = '⚙️ Procesar Facturas AP';
    btnCl.disabled = false;
  };
}

// Cargar datos AP al iniciar
loadAP();
setInterval(loadAP, 60000);
"""

# Insertar el JS AP antes del cierre de </script>
contenido = contenido.replace(
    "// ── Init ──────────────────────────────────────────────────────────────────\nloadAll();\nsetInterval(loadAll, 60000);\n</script>",
    "// ── Init ──────────────────────────────────────────────────────────────────\nloadAll();\nsetInterval(loadAll, 60000);\n" + JS_AP + "\n</script>",
    1
)
print("✓ JavaScript AP añadido")

# ══════════════════════════════════════════════════════════════════════════════
# 3) GUARDAR
# ══════════════════════════════════════════════════════════════════════════════

with open(ruta, "w", encoding="utf-8") as f:
    f.write(contenido)

print(f"✓ dashboard.py guardado ({len(contenido.splitlines())} líneas)")
