"""
dashboard.py — Yve.01
Dashboard principal AR. Ejecutar: python dashboard.py
Abre en: http://localhost:5001
"""

import os, glob, json, subprocess, sys, threading
from datetime import date
import pandas as pd
from flask import Flask, Response, jsonify, stream_with_context

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
REPORTES_DIR     = os.path.join(BASE_DIR, "reportes")
PROCESADAS_DIR   = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_DIR = os.path.join(BASE_DIR, "aprobaciones")
NF = "NO_ENCONTRADO"

app = Flask(__name__)
_pipeline_running = False
_pipeline_lock    = threading.Lock()

# ── Helpers ────────────────────────────────────────────────────────────────

def cargar_ultimo_excel(patron, directorio):
    archivos = sorted(glob.glob(os.path.join(directorio, patron)), reverse=True)
    if not archivos:
        return None, None
    return pd.read_excel(archivos[0]), archivos[0]

def safe_float(val):
    try:
        if val is None or str(val).strip() in ("", NF, "nan", "None"):
            return 0.0
        import re
        s = str(val).replace("EUR","").replace("€","").replace("\xa0","").replace(" ","").strip()
        if "," in s and "." in s:
            s = s.replace(",","") if s.rfind(".") > s.rfind(",") else s.replace(".","").replace(",",".")
        elif "," in s:
            s = s.replace(",","") if re.search(r",\d{3}$", s) else s.replace(",",".")
        return float(s)
    except Exception:
        return 0.0

def cargar_datos():
    df, ruta = cargar_ultimo_excel("doble_imposicion_*.xlsx", REPORTES_DIR)
    if df is None:
        df, ruta = cargar_ultimo_excel("verificacion_*.xlsx", REPORTES_DIR)
    if df is None:
        return pd.DataFrame(), {}

    apro_path = os.path.join(APROBACIONES_DIR, "aprobaciones.xlsx")
    if os.path.exists(apro_path):
        try:
            df_apro = pd.read_excel(apro_path)
            if not df_apro.empty and "numero_factura" in df_apro.columns and "accion" in df_apro.columns:
                ultimas = df_apro.sort_values("fecha").groupby("numero_factura").last().reset_index()
                df = df.merge(ultimas[["numero_factura","accion","comentario"]], on="numero_factura", how="left")
        except Exception:
            pass

    for col in ("accion", "comentario"):
        if col not in df.columns:
            df[col] = None

    meta = {"ruta": os.path.basename(ruta) if ruta else ""}
    return df, meta

def calcular_stats(df):
    if df.empty:
        return {"total":0,"importe_total":0,"correctas":0,"discrepancias":0,
                "importe_reclamable":0,"di_pendientes":0,"aprobadas":0,"rechazadas":0,"sin_accion":0}
    total = len(df)
    importe_total = sum(safe_float(v) for v in df.get("importe_bruto", pd.Series()))

    estado_col = df["estado"].fillna("") if "estado" in df.columns else pd.Series([""] * total)
    correctas     = int((estado_col == "CORRECTO").sum())
    discrepancias = int((estado_col == "DISCREPANCIA").sum())

    if "discrepancia_euros" in df.columns:
        importe_reclamable = sum(abs(safe_float(v)) for v in df.loc[estado_col == "DISCREPANCIA", "discrepancia_euros"])
    else:
        importe_reclamable = 0.0

    di_col = df["estado_di"].fillna("") if "estado_di" in df.columns else pd.Series([""] * total)
    di_pendientes = int((di_col == "FALTA_CERTIFICADO_DI").sum())

    accion_col = df["accion"].fillna("") if "accion" in df.columns else pd.Series([""] * total)
    aprobadas  = int((accion_col == "APROBADA").sum())
    rechazadas = int((accion_col == "RECHAZADA").sum())

    return {
        "total": total,
        "importe_total": round(importe_total, 2),
        "correctas": correctas,
        "discrepancias": discrepancias,
        "importe_reclamable": round(importe_reclamable, 2),
        "di_pendientes": di_pendientes,
        "aprobadas": aprobadas,
        "rechazadas": rechazadas,
        "sin_accion": total - aprobadas - rechazadas,
    }

def calcular_chart(df):
    if df.empty or "nombre_ota" not in df.columns:
        return {"labels": [], "data": []}
    counts = df["nombre_ota"].value_counts()
    counts = counts[~counts.index.isin([NF, "nan", "None", ""])]
    return {"labels": list(counts.index), "data": [int(v) for v in counts.values]}

def df_a_lista(df):
    if df.empty:
        return []
    BAD = {NF, "nan", "None", ""}
    rows = []
    for _, r in df.iterrows():
        def g(col, alts=()):
            for c in (col,) + tuple(alts):
                if c in r.index:
                    v = r[c]
                    if str(v) not in BAD and v is not None:
                        return str(v)
            return ""
        rows.append({
            "archivo":           g("archivo"),
            "numero_factura":    g("numero_factura"),
            "nombre_ota":        g("nombre_ota"),
            "nombre_hotel":      g("nombre_hotel"),
            "fecha":             g("fecha"),
            "importe_bruto":     g("importe_bruto"),
            "porcentaje_factura":g("porcentaje_factura","porcentaje_comision"),
            "estado":            g("estado"),
            "estado_di":         g("estado_di"),
            "discrepancia_euros":g("discrepancia_euros"),
            "accion":            g("accion"),
        })
    return rows

# ── Rutas API ──────────────────────────────────────────────────────────────

@app.route("/api/stats")
def api_stats():
    df, meta = cargar_datos()
    stats = calcular_stats(df)
    stats["chart"] = calcular_chart(df)
    stats["meta"]  = meta
    return jsonify(stats)

@app.route("/api/facturas")
def api_facturas():
    df, _ = cargar_datos()
    return jsonify(df_a_lista(df))

@app.route("/api/procesar")
def api_procesar():
    global _pipeline_running
    scripts = [
        ("lector_ota.py",               "Leyendo facturas PDF"),
        ("verificador_comisiones.py",   "Verificando comisiones OTA"),
        ("detector_doble_imposicion.py","Analizando certificados DI"),
    ]

    def generar():
        global _pipeline_running
        with _pipeline_lock:
            if _pipeline_running:
                yield "data: Ya hay un proceso en ejecucion — espera\n\n"
                return
            _pipeline_running = True
        try:
            yield "data: INICIO\n\n"
            ok_total = True
            for script, label in scripts:
                ruta = os.path.join(BASE_DIR, script)
                yield "data: >> " + label + "...\n\n"
                if not os.path.exists(ruta):
                    yield "data: ERROR: " + script + " no encontrado\n\n"
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
                            yield "data: " + linea + "\n\n"
                    if res.returncode == 0:
                        yield "data: OK " + script + " completado\n\n"
                    else:
                        yield "data: ERROR en " + script + " (codigo " + str(res.returncode) + ")\n\n"
                        ok_total = False
                except subprocess.TimeoutExpired:
                    yield "data: TIMEOUT: " + script + " tardo demasiado\n\n"
                    ok_total = False
                except Exception as exc:
                    yield "data: ERROR: " + str(exc) + "\n\n"
                    ok_total = False
            yield "data: PIPELINE_COMPLETO\n\n" if ok_total else "data: PIPELINE_CON_ERRORES\n\n"
        finally:
            _pipeline_running = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


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
                yield "data: Ya hay un proceso AP en ejecucion — espera\n\n"
                return
            _pipeline_ap_running = True
        try:
            yield "data: INICIO\n\n"
            ok_total = True
            for script, label in scripts:
                ruta = os.path.join(BASE_DIR, script)
                yield "data: >> " + label + "...\n\n"
                if not os.path.exists(ruta):
                    yield "data: ERROR: " + script + " no encontrado\n\n"
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
                            yield "data: " + linea + "\n\n"
                    if res.returncode == 0:
                        yield "data: OK " + script + " completado\n\n"
                    else:
                        yield "data: ERROR en " + script + " (codigo " + str(res.returncode) + ")\n\n"
                        ok_total = False
                except subprocess.TimeoutExpired:
                    yield "data: TIMEOUT: " + script + " tardo demasiado\n\n"
                    ok_total = False
                except Exception as exc:
                    yield "data: ERROR: " + str(exc) + "\n\n"
                    ok_total = False
            yield "data: PIPELINE_COMPLETO\n\n" if ok_total else "data: PIPELINE_CON_ERRORES\n\n"
        finally:
            _pipeline_ap_running = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )

@app.route("/")
def index():
    return HTML

# ── HTML ───────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Yve.01 — Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;
  --acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;
  --tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;--yel:#eab308;--pur:#8b5cf6;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh;line-height:1.5}

/* ── NAV ── */
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:200}
.logo{display:flex;align-items:baseline;gap:10px;flex-shrink:0}
.logo-name{font-size:20px;font-weight:800;color:var(--acc2);letter-spacing:-0.5px}
.logo-tag{font-size:11px;color:var(--mut);font-weight:400;white-space:nowrap}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);flex-shrink:0;box-shadow:0 0 8px var(--acc)}
.nav-mid{flex:1}
.nav-right{display:flex;align-items:center;gap:10px;flex-shrink:0}
.pill{font-size:11px;color:var(--mut);background:var(--s2);padding:4px 12px;border-radius:20px;white-space:nowrap}
.btn-ref{background:none;border:1px solid var(--s2);color:var(--mut);padding:6px 12px;border-radius:8px;font-size:12px;cursor:pointer;transition:.15s;white-space:nowrap}
.btn-ref:hover{border-color:var(--acc);color:var(--acc2)}
.btn-run{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;border:none;padding:9px 18px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:8px;white-space:nowrap;box-shadow:0 0 20px rgba(59,130,246,.35);transition:.15s}
.btn-run:hover{box-shadow:0 0 28px rgba(59,130,246,.55);transform:translateY(-1px)}
.btn-run:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
@media(max-width:640px){.logo-tag,.pill{display:none}.nav{padding:0 14px}.btn-run{padding:8px 12px;font-size:12px}}

/* ── MAIN ── */
.main{padding:24px;max-width:1440px;margin:0 auto}
@media(max-width:640px){.main{padding:14px}}

/* ── ALERT ── */
.alert{display:none;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.25);border-radius:12px;padding:12px 18px;font-size:13px;color:#fca5a5;margin-bottom:20px;align-items:center;gap:10px}
.alert.on{display:flex}

/* ── STATS ── */
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:22px}
@media(max-width:1200px){.stats{grid-template-columns:repeat(3,1fr)}}
@media(max-width:600px){.stats{grid-template-columns:repeat(2,1fr)}}
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px;transition:.2s}
.sc:hover{border-color:var(--s3)}
.sc.hl{border-color:rgba(59,130,246,.4);background:rgba(59,130,246,.05)}
.sc-lbl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px}
.sc-val{font-size:28px;font-weight:800;line-height:1;letter-spacing:-1px}
.sc-sub{font-size:11px;color:var(--dim);margin-top:6px}
.sc.c-blu .sc-val{color:var(--acc2)}
.sc.c-grn .sc-val{color:var(--grn)}
.sc.c-red .sc-val{color:var(--red)}
.sc.c-ora .sc-val{color:var(--ora)}
.sc.c-yel .sc-val{color:var(--yel)}
.sc.c-pur .sc-val{color:var(--pur)}

/* ── MID ROW ── */
.mid{display:grid;grid-template-columns:1fr 300px;gap:16px;margin-bottom:22px}
@media(max-width:960px){.mid{grid-template-columns:1fr}}

/* ── CARDS ── */
.card{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:22px}
.card-title{font-size:11px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:18px}
.chart-wrap{height:190px;position:relative}

/* ── ACTIVITY ── */
.act-item{display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--s2)}
.act-item:last-child{border-bottom:none}
.adot{width:8px;height:8px;border-radius:50%;flex-shrink:0;margin-top:5px}
.adot.g{background:var(--grn)}.adot.r{background:var(--red)}.adot.o{background:var(--ora)}
.adot.b{background:var(--acc2)}.adot.m{background:var(--mut)}
.atxt{font-size:12px;color:var(--tx);line-height:1.5}
.atxt b{color:var(--acc3);font-weight:700}

/* ── TABLE ── */
.tbl-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse;font-size:12px;min-width:800px}
th{background:rgba(51,65,85,.6);color:var(--mut);font-size:10px;text-transform:uppercase;letter-spacing:.5px;padding:10px 14px;text-align:left;white-space:nowrap;border-bottom:1px solid var(--s2)}
th:first-child{border-radius:8px 0 0 0}th:last-child{border-radius:0 8px 0 0}
td{padding:10px 14px;border-bottom:1px solid rgba(51,65,85,.4);white-space:nowrap;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.025)}
.td-dim{color:var(--dim);font-size:11px}
.td-b{font-weight:700}
.td-red{color:#f87171;font-weight:700}

/* ── BADGES ── */
.badge{display:inline-flex;align-items:center;gap:4px;font-size:9px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap}
.b-ok{background:rgba(34,197,94,.12);color:#4ade80;border:1px solid rgba(34,197,94,.2)}
.b-disc{background:rgba(239,68,68,.12);color:#f87171;border:1px solid rgba(239,68,68,.2)}
.b-fdi{background:rgba(249,115,22,.12);color:#fb923c;border:1px solid rgba(249,115,22,.2)}
.b-cok{background:rgba(59,130,246,.12);color:#60a5fa;border:1px solid rgba(59,130,246,.2)}
.b-na{background:rgba(148,163,184,.08);color:#94a3b8;border:1px solid rgba(148,163,184,.12)}
.b-unk{background:rgba(234,179,8,.12);color:#facc15;border:1px solid rgba(234,179,8,.2)}
.b-apr{background:rgba(34,197,94,.12);color:#4ade80;border:1px solid rgba(34,197,94,.2)}
.b-rec{background:rgba(239,68,68,.12);color:#f87171;border:1px solid rgba(239,68,68,.2)}
.b-pen{background:rgba(148,163,184,.07);color:#64748b;border:1px solid rgba(148,163,184,.1)}

/* ── MODAL ── */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.8);backdrop-filter:blur(4px);z-index:500;align-items:center;justify-content:center;padding:20px}
.overlay.on{display:flex}
.modal{background:var(--s1);border:1px solid var(--s2);border-radius:18px;width:100%;max-width:560px;padding:26px}
.modal-h{display:flex;align-items:center;gap:10px;margin-bottom:18px}
.modal-h h3{font-size:16px;font-weight:700;flex:1}
.log{background:#060c1a;border:1px solid var(--s2);border-radius:10px;padding:16px;height:280px;overflow-y:auto;font-family:'JetBrains Mono','Cascadia Code','Fira Code',monospace;font-size:11px;line-height:1.8;scroll-behavior:smooth}
.log p{margin:0}
.l-ok{color:#4ade80}.l-err{color:#f87171}.l-info{color:#60a5fa;font-weight:700}.l-warn{color:#facc15}.l-dim{color:#475569}
.modal-f{margin-top:16px;display:flex;justify-content:flex-end;gap:10px}
.btn-cl{background:var(--s2);color:var(--tx);border:none;padding:9px 20px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer;transition:.15s}
.btn-cl:hover:not(:disabled){background:var(--s3)}
.btn-cl:disabled{opacity:.35;cursor:not-allowed}

/* ── SPINNER ── */
.spin{width:15px;height:15px;border:2px solid rgba(255,255,255,.25);border-top-color:#fff;border-radius:50%;animation:sp .65s linear infinite;display:none}
@keyframes sp{to{transform:rotate(360deg)}}

/* ── EMPTY ── */
.empty{text-align:center;padding:48px 20px;color:var(--mut)}
.empty .ei{font-size:36px;margin-bottom:10px}
.empty p{font-size:13px;line-height:1.6}

/* ── STATUS BAR ── */
.status-bar{display:flex;align-items:center;gap:8px;font-size:11px;color:var(--dim);margin-bottom:20px}
.status-dot{width:6px;height:6px;border-radius:50%;background:var(--grn);box-shadow:0 0 6px var(--grn);animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}

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
</style>
</head>
<body>

<nav class="nav">
  <div class="logo">
    <div class="logo-dot"></div>
    <span class="logo-name">Yve.01</span>
    <span class="logo-tag">AR Dashboard</span>
  </div>
  <div class="nav-mid"></div>
  <div class="nav-right">
    <span class="pill" id="date-pill">—</span>
    <button class="btn-ref" onclick="loadAll()" title="Actualizar datos">↻ Actualizar</button>
    <button class="btn-run" id="btn-run" onclick="runPipeline()">
      <div class="spin" id="spin"></div>
      <span id="run-lbl">⚡ Procesar Facturas</span>
    </button>
  </div>
</nav>

<div class="main">

  <!-- Alerta -->
  <div class="alert" id="alert-bar">
    <span style="font-size:16px">⚠️</span>
    <span id="alert-msg"></span>
  </div>

  <!-- Barra estado -->
  <div class="status-bar">
    <div class="status-dot"></div>
    <span id="status-txt">Cargando datos...</span>
  </div>

  <!-- STATS -->
  <div class="stats">
    <div class="sc hl c-blu">
      <div class="sc-lbl">Facturas procesadas</div>
      <div class="sc-val" id="s-tot">—</div>
      <div class="sc-sub">último ciclo AR</div>
    </div>
    <div class="sc">
      <div class="sc-lbl">Importe total</div>
      <div class="sc-val" id="s-imp" style="font-size:18px;letter-spacing:-0.5px">—</div>
      <div class="sc-sub">EUR procesados</div>
    </div>
    <div class="sc c-grn">
      <div class="sc-lbl">Correctas</div>
      <div class="sc-val" id="s-ok">—</div>
      <div class="sc-sub">sin incidencias</div>
    </div>
    <div class="sc c-red">
      <div class="sc-lbl">Discrepancias</div>
      <div class="sc-val" id="s-disc">—</div>
      <div class="sc-sub" id="s-disc-sub">reclamable: —</div>
    </div>
    <div class="sc c-ora">
      <div class="sc-lbl">Certif. DI pendiente</div>
      <div class="sc-val" id="s-di">—</div>
      <div class="sc-sub">facturas extranjeras</div>
    </div>
    <div class="sc c-pur">
      <div class="sc-lbl">Pendientes firma</div>
      <div class="sc-val" id="s-pend">—</div>
      <div class="sc-sub" id="s-pend-sub">— apr · — rec</div>
    </div>
  </div>

  <!-- MID ROW -->
  <div class="mid">
    <div class="card">
      <div class="card-title">Facturas por OTA</div>
      <div class="chart-wrap"><canvas id="ota-chart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Resumen de estados</div>
      <div id="activity">
        <div class="empty"><div class="ei">📂</div><p>Sin datos.<br>Pulsa ⚡ Procesar Facturas.</p></div>
      </div>
    </div>
  </div>

  <!-- TABLE -->
  <div class="card">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
      <div class="card-title" style="margin:0">Todas las facturas del ciclo</div>
      <span id="tbl-count" style="font-size:11px;color:var(--dim)"></span>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Archivo</th>
            <th>Nº Factura</th>
            <th>OTA</th>
            <th>Hotel</th>
            <th>Fecha</th>
            <th>Importe bruto</th>
            <th>% Com.</th>
            <th>Estado</th>
            <th>Estado DI</th>
            <th>Discrepancia</th>
            <th>Aprobación</th>
          </tr>
        </thead>
        <tbody id="tbl-body">
          <tr><td colspan="11" class="empty"><p>Sin datos. Pulsa ⚡ Procesar Facturas para empezar.</p></td></tr>
        </tbody>
      </table>
    </div>
  </div>

</div><!-- /main -->

<!-- MODAL PIPELINE -->
<div class="overlay" id="overlay">
  <div class="modal">
    <div class="modal-h">
      <span id="modal-icon" style="font-size:20px">⚡</span>
      <h3 id="modal-title">Pipeline AR — Procesando...</h3>
    </div>
    <div class="log" id="log"></div>
    <div class="modal-f">
      <button class="btn-cl" id="btn-cl" onclick="closeModal()" disabled>Cerrar</button>
    </div>
  </div>
</div>

<script>
// ── Globals ─────────────────────────────────────────────────────────────
let otaChart = null;

// ── Formato ─────────────────────────────────────────────────────────────
function eur(n) {
  if (n === null || n === undefined || n === '' || n === 0) return '—';
  return new Intl.NumberFormat('es-ES', {minimumFractionDigits:2, maximumFractionDigits:2}).format(n) + ' €';
}

// ── Badges ───────────────────────────────────────────────────────────────
function bEstado(e) {
  const m = {
    CORRECTO:        ['b-ok',   '✓ Correcto'],
    DISCREPANCIA:    ['b-disc', '⚠ Discrepancia'],
    OTA_DESCONOCIDA: ['b-unk',  '? OTA desc.'],
    SIN_PORCENTAJE:  ['b-na',   '~ Sin %'],
  };
  const [c, l] = m[e] || ['b-na', e || '—'];
  return '<span class="badge ' + c + '">' + l + '</span>';
}

function bDI(e) {
  const m = {
    CERTIFICADO_OK:       ['b-cok', '✓ Cert. OK'],
    FALTA_CERTIFICADO_DI: ['b-fdi', '✗ Falta DI'],
    NO_APLICA:            ['b-na',  '— N/A'],
    OTA_DESCONOCIDA:      ['b-unk', '? Desc.'],
  };
  const [c, l] = m[e] || ['b-na', e || '—'];
  return '<span class="badge ' + c + '">' + l + '</span>';
}

function bApro(a) {
  if (!a || a === '') return '<span class="badge b-pen">· Pendiente</span>';
  if (a === 'APROBADA')  return '<span class="badge b-apr">✓ Aprobada</span>';
  if (a === 'RECHAZADA') return '<span class="badge b-rec">✗ Rechazada</span>';
  return '<span class="badge b-na">—</span>';
}

// ── Carga datos ──────────────────────────────────────────────────────────
async function loadAll() {
  document.getElementById('status-txt').textContent = 'Actualizando...';
  try {
    const [sr, fr] = await Promise.all([fetch('/api/stats'), fetch('/api/facturas')]);
    const stats = await sr.json();
    const facturas = await fr.json();

    renderStats(stats);
    renderChart(stats.chart);
    renderTable(facturas);
    renderActivity(facturas);

    const hoy = new Date();
    document.getElementById('date-pill').textContent =
      hoy.toLocaleDateString('es-ES', {day:'2-digit', month:'short', year:'numeric'}) + ' · ' +
      hoy.toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'});

    // Alert bar
    const alertBar = document.getElementById('alert-bar');
    const parts = [];
    if (stats.discrepancias > 0)
      parts.push(stats.discrepancias + ' discrepancia(s) · ' + eur(stats.importe_reclamable) + ' reclamables');
    if (stats.di_pendientes > 0)
      parts.push(stats.di_pendientes + ' factura(s) sin certificado DI');
    if (parts.length) {
      document.getElementById('alert-msg').textContent = parts.join(' — ');
      alertBar.classList.add('on');
    } else {
      alertBar.classList.remove('on');
    }

    document.getElementById('status-txt').textContent =
      'Actualizado · ' + facturas.length + ' factura' + (facturas.length !== 1 ? 's' : '') + ' cargada' + (facturas.length !== 1 ? 's' : '');
  } catch(e) {
    console.error(e);
    document.getElementById('status-txt').textContent = 'Error al cargar datos';
  }
}

function renderStats(s) {
  document.getElementById('s-tot').textContent  = s.total ?? '—';
  document.getElementById('s-imp').textContent  = s.importe_total ? eur(s.importe_total) : '—';
  document.getElementById('s-ok').textContent   = s.correctas ?? '—';
  document.getElementById('s-disc').textContent = s.discrepancias ?? '—';
  document.getElementById('s-disc-sub').textContent = 'reclamable: ' + eur(s.importe_reclamable);
  document.getElementById('s-di').textContent   = s.di_pendientes ?? '—';
  document.getElementById('s-pend').textContent = s.sin_accion ?? '—';
  document.getElementById('s-pend-sub').textContent = (s.aprobadas ?? 0) + ' apr · ' + (s.rechazadas ?? 0) + ' rec';
}

function renderChart(ch) {
  if (!ch || !ch.labels.length) return;
  const ctx = document.getElementById('ota-chart').getContext('2d');
  const palette = ['#3b82f6','#8b5cf6','#06b6d4','#10b981','#f59e0b','#ef4444','#ec4899'];
  const cols = ch.labels.map((_, i) => palette[i % palette.length]);
  if (otaChart) otaChart.destroy();
  otaChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ch.labels,
      datasets: [{
        label: 'Facturas',
        data: ch.data,
        backgroundColor: cols.map(c => c + '22'),
        borderColor: cols,
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1e293b',
          borderColor: '#334155',
          borderWidth: 1,
          callbacks: {
            label: ctx => ' ' + ctx.raw + ' factura' + (ctx.raw !== 1 ? 's' : '')
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255,255,255,.04)' },
          ticks: { color: '#94a3b8', font: { size: 11 } }
        },
        y: {
          grid: { color: 'rgba(255,255,255,.04)' },
          ticks: { color: '#94a3b8', stepSize: 1, font: { size: 11 } },
          beginAtZero: true
        }
      }
    }
  });
}

function renderTable(rows) {
  const tbody = document.getElementById('tbl-body');
  document.getElementById('tbl-count').textContent = rows.length ? rows.length + ' registros' : '';
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="11" class="empty"><p>Sin facturas. Pulsa ⚡ Procesar Facturas.</p></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => {
    const hasDisc = r.discrepancia_euros && r.discrepancia_euros !== '';
    return [
      '<tr>',
      '<td class="td-dim">' + (r.archivo || '—') + '</td>',
      '<td class="td-b">' + (r.numero_factura || '—') + '</td>',
      '<td class="td-b" style="color:var(--acc3)">' + (r.nombre_ota || '—') + '</td>',
      '<td>' + (r.nombre_hotel || '—') + '</td>',
      '<td class="td-dim">' + (r.fecha || '—') + '</td>',
      '<td class="td-b">' + (r.importe_bruto || '—') + '</td>',
      '<td>' + (r.porcentaje_factura || '—') + '</td>',
      '<td>' + bEstado(r.estado) + '</td>',
      '<td>' + bDI(r.estado_di) + '</td>',
      '<td class="' + (hasDisc ? 'td-red' : 'td-dim') + '">' + (hasDisc ? r.discrepancia_euros : '—') + '</td>',
      '<td>' + bApro(r.accion) + '</td>',
      '</tr>'
    ].join('');
  }).join('');
}

function renderActivity(rows) {
  const el = document.getElementById('activity');
  if (!rows.length) {
    el.innerHTML = '<div class="empty"><div class="ei">📂</div><p>Sin datos.<br>Pulsa ⚡ Procesar Facturas.</p></div>';
    return;
  }
  const c = {}, d = {};
  rows.forEach(r => {
    if (r.estado)    c[r.estado]    = (c[r.estado]    || 0) + 1;
    if (r.estado_di) d[r.estado_di] = (d[r.estado_di] || 0) + 1;
  });
  const items = [
    { dot:'g', n: c.CORRECTO             || 0, txt: 'correctas sin incidencias' },
    { dot:'r', n: c.DISCREPANCIA         || 0, txt: 'con discrepancia de comisión' },
    { dot:'o', n: d.FALTA_CERTIFICADO_DI || 0, txt: 'sin certificado DI' },
    { dot:'b', n: d.CERTIFICADO_OK       || 0, txt: 'con certificado DI OK' },
    { dot:'m', n: d.OTA_DESCONOCIDA      || 0, txt: 'OTA no reconocida' },
  ];
  el.innerHTML = items.map(i =>
    '<div class="act-item">' +
    '<div class="adot ' + i.dot + '"></div>' +
    '<div class="atxt"><b>' + i.n + '</b> factura' + (i.n !== 1 ? 's' : '') + ' ' + i.txt + '</div>' +
    '</div>'
  ).join('');
}

// ── Pipeline SSE ─────────────────────────────────────────────────────────
function runPipeline() {
  const btn    = document.getElementById('btn-run');
  const spin   = document.getElementById('spin');
  const lbl    = document.getElementById('run-lbl');
  const log    = document.getElementById('log');
  const btnCl  = document.getElementById('btn-cl');
  const icon   = document.getElementById('modal-icon');
  const title  = document.getElementById('modal-title');

  btn.disabled = true;
  spin.style.display = 'block';
  lbl.textContent = 'Procesando...';
  log.innerHTML = '';
  btnCl.disabled = true;
  icon.textContent = '⚡';
  title.textContent = 'Pipeline AR — Procesando...';
  document.getElementById('overlay').classList.add('on');

  const src = new EventSource('/api/procesar');

  src.onmessage = ev => {
    const txt = ev.data;
    const p = document.createElement('p');

    if      (txt === 'PIPELINE_COMPLETO')    p.className = 'l-ok';
    else if (txt === 'PIPELINE_CON_ERRORES') p.className = 'l-err';
    else if (txt.startsWith('OK '))          p.className = 'l-ok';
    else if (txt.startsWith('ERROR') || txt.startsWith('TIMEOUT')) p.className = 'l-err';
    else if (txt.startsWith('>> ') || txt === 'INICIO') p.className = 'l-info';
    else if (txt.includes('✓') || txt.includes('v]')) p.className = 'l-ok';
    else if (txt.includes('✗') || txt.includes('X]')) p.className = 'l-err';
    else p.className = 'l-dim';

    p.textContent = txt;
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = ok ? 'Pipeline completado con éxito' : 'Pipeline finalizado con errores';
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = '⚡ Procesar Facturas';
      btnCl.disabled = false;
      setTimeout(loadAll, 800);
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
    lbl.textContent = '⚡ Procesar Facturas';
    btnCl.disabled = false;
    icon.textContent = '⚠️';
    title.textContent = 'Error de conexión';
  };
}

function closeModal() {
  document.getElementById('overlay').classList.remove('on');
}

// ── Init ──────────────────────────────────────────────────────────────────
loadAll();
setInterval(loadAll, 60000);


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
    print("  Yve.01 — Dashboard Principal AR")
    print("=" * 60)
    print("  Escritorio:  http://localhost:5001")
    print(f"  Movil:       http://{ip}:5001")
    print("  Ctrl+C para detener")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5001, debug=False)
