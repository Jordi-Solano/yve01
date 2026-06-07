"""
dashboard.py — Yve.01
Dashboard principal AR. Ejecutar: python dashboard.py
Abre en: http://localhost:5001
"""

import os, glob, json, subprocess, sys, threading
from datetime import date
import pandas as pd
from flask import Flask, Response, jsonify, request, stream_with_context, redirect
from flask_login import login_required, current_user

# Ruta base del proyecto — robusta ante ejecución desde cualquier directorio
def _get_base_dir():
    """Devuelve la carpeta que contiene dashboard.py, funcione desde donde funcione."""
    # __file__ siempre apunta al script real
    _f = os.path.abspath(__file__)
    _d = os.path.dirname(_f)
    # Sanity check: la carpeta debe contener lector_ota.py
    if os.path.isfile(os.path.join(_d, "lector_ota.py")):
        return _d
    # Fallback: directorio de trabajo actual
    cwd = os.getcwd()
    if os.path.isfile(os.path.join(cwd, "lector_ota.py")):
        return cwd
    return _d

BASE_DIR         = _get_base_dir()
REPORTES_DIR     = os.path.join(BASE_DIR, "reportes")
PROCESADAS_DIR   = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_DIR = os.path.join(BASE_DIR, "aprobaciones")
NF = "NO_ENCONTRADO"

# ── Excel cache (TTL 5 min) ──────────────────────────────────────────────
import time as _time
_EXCEL_CACHE: dict = {}
_CACHE_TTL = 300  # seconds

def _excel(path, sheet_name=0, header=0, **kw):
    key = f"{path}|{sheet_name}|{header}"
    now = _time.time()
    if key in _EXCEL_CACHE:
        df, ts = _EXCEL_CACHE[key]
        if now - ts < _CACHE_TTL:
            return df
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, **kw)
    _EXCEL_CACHE[key] = (df, now)
    return df

def _invalidate_cache(path_fragment=None):
    """Invalida entradas del caché (llamar al guardar nuevos datos)."""
    global _EXCEL_CACHE
    if path_fragment:
        _EXCEL_CACHE = {k: v for k, v in _EXCEL_CACHE.items() if path_fragment not in k}
    else:
        _EXCEL_CACHE = {}


app = Flask(__name__)
DEMO_MODE = False
app.secret_key = os.environ.get("SECRET_KEY") or "yve01-dev-secret-CHANGE-IN-PROD"

# Auth + módulos: la app es UN solo proceso que sirve todo el producto en un puerto
sys.path.insert(0, BASE_DIR)
from auth import init_login, inicializar_usuarios
init_login(app)
inicializar_usuarios()

# Registrar cada módulo como blueprint (login, configuración, admin, aprobaciones, conciliación)
from login import bp as auth_bp
from onboarding import bp as config_bp
from panel_admin import bp as admin_bp
from app_aprobacion import bp as aprob_ar_bp
from app_aprobacion_ap import bp as aprob_ap_bp
from app_conciliacion import bp as concil_bp
from tab_fb_dashboard import fb_bp
from tab_ar_real import ar_real_bp
from tab_multi_hotel import multi_hotel_bp
from tab_exportador import exportador_bp
from tab_calipolis import calipolis_bp
from tab_demo import demo_bp
from tab_demo_simulador import demo_sim_bp
from tab_calipolis_analisis import calipolis_analisis_bp
from tab_reportes_pdf import reportes_pdf_bp
from rol_dashboard import get_dashboard_config
from demo_completo import generar_hoteles_demo, generar_facturas_demo_ar, generar_facturas_demo_ap, generar_alertas_demo
from landing import LANDING_HTML as LANDING_PAGE
from blog import blog_bp
from billing import billing_bp
from signup import signup_bp
from about import about_bp
from exportador_pdf import pdf_bp
for _bp in (auth_bp, config_bp, admin_bp, aprob_ar_bp, aprob_ap_bp, concil_bp, fb_bp, ar_real_bp, multi_hotel_bp, exportador_bp, calipolis_bp, demo_bp, demo_sim_bp, calipolis_analisis_bp, reportes_pdf_bp, blog_bp, billing_bp, signup_bp, about_bp, pdf_bp):
    app.register_blueprint(_bp)

_pipeline_running = False
_pipeline_lock    = threading.Lock()

# ── Helpers ────────────────────────────────────────────────────────────────

def cargar_ultimo_excel(patron, directorio):
    """Devuelve el Excel más reciente que coincida con el patrón, ordenado por fecha de modificación."""
    hits = glob.glob(os.path.join(directorio, patron))
    if not hits:
        return None, None
    # Ordenar por fecha de modificación real (más reciente primero)
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    try:
        return pd.read_excel(hits[0]), hits[0]
    except Exception as e:
        print(f"  ⚠  Error leyendo {hits[0]}: {e}")
        for ruta in hits[1:]:
            try:
                return pd.read_excel(ruta), ruta
            except Exception:
                continue
        return None, None

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
    """
    Carga los datos AR del archivo más reciente disponible.
    Prioridad:
      1. reportes/doble_imposicion_*.xlsx   (salida de detector_doble_imposicion.py)
      2. reportes/verificacion_*.xlsx        (salida de verificador_comisiones.py)
      3. facturas-procesadas/facturas_procesadas_*.xlsx  (salida de lector_ota.py)
    En todos los casos usa el archivo más reciente por fecha de modificación.
    """
    print(f"[cargar_datos] BASE_DIR={BASE_DIR}")
    print(f"[cargar_datos] REPORTES_DIR={REPORTES_DIR} (existe: {os.path.isdir(REPORTES_DIR)})")
    print(f"[cargar_datos] PROCESADAS_DIR={PROCESADAS_DIR} (existe: {os.path.isdir(PROCESADAS_DIR)})")
    df, ruta = cargar_ultimo_excel("doble_imposicion_*.xlsx", REPORTES_DIR)
    if df is None:
        df, ruta = cargar_ultimo_excel("verificacion_*.xlsx", REPORTES_DIR)
    if df is None:
        df, ruta = cargar_ultimo_excel("facturas_procesadas_*.xlsx", PROCESADAS_DIR)
    if df is None:
        print("[cargar_datos] ADVERTENCIA: no se encontró ningún Excel AR")
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

@app.route("/api/debug")
def api_debug():
    """Endpoint de diagnóstico — muestra rutas y archivos disponibles."""
    import datetime
    info = {
        "BASE_DIR":       BASE_DIR,
        "BASE_DIR_exists": os.path.isdir(BASE_DIR),
        "REPORTES_DIR":   REPORTES_DIR,
        "PROCESADAS_DIR": PROCESADAS_DIR,
        "archivos_reportes": [],
        "archivos_procesadas": [],
        "cwd": os.getcwd(),
        "python_file": __file__,
    }
    for d, key in [(REPORTES_DIR, "archivos_reportes"), (PROCESADAS_DIR, "archivos_procesadas")]:
        if os.path.isdir(d):
            for f in sorted(os.listdir(d)):
                if f.endswith(".xlsx"):
                    ruta = os.path.join(d, f)
                    mtime = os.path.getmtime(ruta)
                    info[key].append({
                        "nombre": f,
                        "bytes": os.path.getsize(ruta),
                        "modificado": datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    })
    # Try actually loading
    df, ruta_cargada = cargar_ultimo_excel("doble_imposicion_*.xlsx", REPORTES_DIR)
    if df is None:
        df, ruta_cargada = cargar_ultimo_excel("verificacion_*.xlsx", REPORTES_DIR)
    if df is None:
        df, ruta_cargada = cargar_ultimo_excel("facturas_procesadas_*.xlsx", PROCESADAS_DIR)
    info["archivo_cargado"] = os.path.basename(ruta_cargada) if ruta_cargada else None
    info["filas_cargadas"] = len(df) if df is not None else 0
    info["columnas"] = list(df.columns) if df is not None else []
    return jsonify(info)

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
            # Enviar notificaciones automáticas
            if ok_total:
                try:
                    sys.path.insert(0, BASE_DIR)
                    from exportador_final import generar_reporte_ejecutivo, generar_excel_consolidado
                    alertas = enviar_pendientes()
                    if alertas:
                        yield "data: >> Notificaciones: " + str(len(alertas)) + " alerta(s) procesada(s)\n\n"
                except Exception as e_notif:
                    yield "data: >> Notificaciones: error — " + str(e_notif)[:80] + "\n\n"
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
_pipeline_oracle_running = False
_pipeline_oracle_lock    = threading.Lock()
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
            if ok_total:
                try:
                    sys.path.insert(0, BASE_DIR)
                    from exportador_final import generar_reporte_ejecutivo, generar_excel_consolidado
                    alertas = enviar_pendientes()
                    if alertas:
                        yield "data: >> Notificaciones: " + str(len(alertas)) + " alerta(s) procesada(s)\n\n"
                except Exception as e_notif:
                    yield "data: >> Notificaciones: error — " + str(e_notif)[:80] + "\n\n"
        finally:
            _pipeline_ap_running = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/procesar_oracle")
def api_procesar_oracle():
    global _pipeline_oracle_running
    script = os.path.join(BASE_DIR, "oracle_pipeline.py")

    def generar():
        global _pipeline_oracle_running
        with _pipeline_oracle_lock:
            if _pipeline_oracle_running:
                yield "data: Ya hay un proceso Oracle en ejecucion — espera\n\n"
                return
            _pipeline_oracle_running = True
        try:
            yield "data: INICIO\n\n"
            if not os.path.exists(script):
                yield "data: ERROR: oracle_pipeline.py no encontrado\n\n"
                yield "data: PIPELINE_CON_ERRORES\n\n"
                return
            res = subprocess.run(
                [sys.executable, script],
                capture_output=True, text=True, timeout=300, cwd=BASE_DIR
            )
            for linea in (res.stdout + res.stderr).splitlines():
                linea = linea.strip()
                if linea:
                    yield "data: " + linea + "\n\n"
            if res.returncode == 0:
                yield "data: PIPELINE_COMPLETO\n\n"
            else:
                yield "data: PIPELINE_CON_ERRORES\n\n"
        finally:
            _pipeline_oracle_running = False

    return Response(
        stream_with_context(generar()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── Chat AI — Yve Copilot ──────────────────────────────────────────────

def _cargar_contexto_chat():
    """Construye el contexto financiero actual para el system prompt del chat."""
    try:
        df_ar, _  = cargar_datos()
        stats_ar  = calcular_stats(df_ar) if not df_ar.empty else {}
        stats_ap  = calcular_stats_ap(cargar_datos_ap())

        # Detalles AR
        facturas_ar = df_a_lista(df_ar)[:50]
        disc_ar  = [f for f in facturas_ar if str(f.get("estado","")).upper()
                    not in ("CORRECTO","APROBADA","")]
        ota_data = calcular_chart(df_ar)
        otas_str = "; ".join(f"{o['ota']}: {o['n']} facturas" for o in ota_data[:8]) if ota_data else "sin datos"

        # Detalles AP
        df_ap    = cargar_datos_ap()
        lista_ap = df_ap_a_lista(df_ap) if not df_ap.empty else []
        pend_ap  = [f for f in lista_ap if not f.get("accion")]
        disc_ap  = [f for f in lista_ap if f.get("estado") in
                    ("DISCREPANCIA_PO","SIN_PO","ALERTA_CONSUMO","DISCREPANCIA")]

        # Top proveedores con más errores
        from collections import Counter
        prov_err = Counter(f.get("proveedor","") for f in disc_ap)
        top_err  = "; ".join(f"{p}: {n}" for p,n in prov_err.most_common(5)) or "ninguno"

        ctx = f"""ESTADO FINANCIERO ACTUAL DEL HOTEL — Yve.01

=== MÓDULO AR (Facturas OTA) ===
Total facturas AR procesadas hoy: {stats_ar.get('total_facturas', 0)}
Importe total AR: {stats_ar.get('importe_total', 0):,.2f} EUR
Facturas correctas: {stats_ar.get('correctas', 0)}
Discrepancias AR: {stats_ar.get('discrepancias', 0)} — importe reclamable: {stats_ar.get('importe_discrepancias', 0):,.2f} EUR
DI pendientes: {stats_ar.get('di_pendientes', 0)}
Aprobadas: {stats_ar.get('aprobadas', 0)} | Rechazadas: {stats_ar.get('rechazadas', 0)}
OTAs y volumen: {otas_str}
Facturas con discrepancias AR: {'; '.join(f"{f.get('ota','?')} {f.get('importe','?')}€" for f in disc_ar[:5]) or 'ninguna'}

=== MÓDULO AP (Facturas Proveedores) ===
Total facturas AP: {stats_ap.get('total', 0)}
Importe total AP: {stats_ap.get('importe', 0):,.2f} EUR
Matches correctos (F&B+OTRAS): {stats_ap.get('matches', 0)}
Discrepancias PO: {stats_ap.get('discrepancias', 0)}
Sin Orden de Compra: {stats_ap.get('sin_po', 0)}
Alertas consumo F&B: {stats_ap.get('alertas_consumo', 0)}
Pendientes asignación manual: {stats_ap.get('manual', 0)}
Aprobadas AP: {stats_ap.get('aprobadas', 0)} | Rechazadas AP: {stats_ap.get('rechazadas', 0)}
Facturas AP pendientes de aprobar: {len(pend_ap)}
Facturas AP con discrepancias: {'; '.join(f"{f.get('proveedor','?')} ({f.get('estado','?')}) {f.get('total',0):,.0f}EUR" for f in disc_ap[:5]) or 'ninguna'}
Proveedores con más errores: {top_err}"""
        return ctx
    except Exception as e:
        return f"Error cargando contexto financiero: {e}"


@app.route("/api/chat", methods=["POST"])
def api_chat():
    """Endpoint del chat AI — llama a Claude con contexto de datos reales."""
    import json as _json

    data     = request.get_json(force=True, silent=True) or {}
    messages = data.get("messages", [])
    if not messages:
        return jsonify({"error": "No messages provided"}), 400

    # Cargar API key
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        env_path = os.path.join(BASE_DIR, ".env")
        if os.path.exists(env_path):
            for line in open(env_path).readlines():
                if line.startswith("ANTHROPIC_API_KEY="):
                    api_key = line.split("=",1)[1].strip().strip('"').strip("'")
                    break

    if not api_key:
        return jsonify({"reply":
            "El asistente IA necesita una **API key de Anthropic** para responder.\n\n"
            "Mientras tanto, puedes consultar todos los datos directamente en las pestañas del dashboard: "
            "**AR**, **AP**, **DRR** y **Banco**.\n\n"
            "Para activarme, añade `ANTHROPIC_API_KEY` en las variables de entorno de Render."}), 200

    contexto = _cargar_contexto_chat()

    system_prompt = f"""Eres Yve, el asistente financiero inteligente del hotel integrado en el dashboard Yve.01.
Tienes acceso en tiempo real a todos los datos financieros del hotel: facturas AR (OTAs), facturas AP (proveedores), aprobaciones, discrepancias, importes reclamables y estados de contabilización Oracle.

DATOS ACTUALES DEL HOTEL:
{contexto}

INSTRUCCIONES:
- Responde SIEMPRE en español, con tono profesional pero cercano
- Sé directo y específico — da números reales de los datos de arriba
- Si te preguntan por discrepancias, menciona los importes exactos y las OTAs/proveedores concretos
- Si no tienes el dato exacto, dilo claramente y sugiere qué módulo revisar
- Usa emojis con moderación para hacer las respuestas más legibles
- Nunca inventes datos que no aparezcan en el contexto financiero anterior
- Si te preguntan quién eres: "Soy Yve, tu copiloto financiero de Yve.01. Tengo acceso a todos los datos del dashboard en tiempo real."
- Para preguntas sobre facturas concretas, busca en los datos del contexto
- Respuestas concisas: máximo 4-5 líneas salvo que se pida un análisis detallado"""

    try:
        import anthropic
        client  = anthropic.Anthropic(api_key=api_key)
        # Filtrar solo mensajes user/assistant
        msgs = [{"role": m["role"], "content": m["content"]}
                for m in messages if m.get("role") in ("user","assistant")]
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            system=system_prompt,
            messages=msgs,
        )
        reply = resp.content[0].text.strip()
        return jsonify({"reply": reply})
    except Exception as e:
        err = str(e).lower()
        if "credit" in err or "balance" in err or "quota" in err or "insufficient" in err:
            msg = ("La cuenta de Anthropic se ha quedado **sin créditos**.\n\n"
                   "Recarga el saldo en console.anthropic.com para reactivar el asistente. "
                   "El resto del dashboard sigue funcionando con normalidad.")
        elif "rate" in err or "429" in err:
            msg = "Demasiadas consultas seguidas. Espera unos segundos y vuelve a intentarlo."
        else:
            msg = f"No he podido procesar la consulta ahora mismo. Inténtalo de nuevo en un momento."
        return jsonify({"reply": msg}), 200

_pipeline_oracle_running = False
_pipeline_oracle_lock    = threading.Lock()

# ── DRR (Daily Revenue Report) ────────────────────────────────────────

FACTURAS_ENTRADA_DIR = os.path.join(BASE_DIR, "facturas-entrada")
DRR_UPLOAD_DIR       = os.path.join(BASE_DIR, "facturas-entrada")

def _cargar_drr_procesado():
    """Carga el último drr_procesado_*.xlsx de reportes/."""
    hits = glob.glob(os.path.join(REPORTES_DIR, "drr_procesado_*.xlsx"))
    if not hits:
        return None
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return hits[0]

def _leer_drr_stats(ruta):
    """Lee el Excel procesado del DRR y devuelve stats para el frontend."""
    try:
        # Hoja Resumen — métricas KPI
        df_res = pd.read_excel(ruta, sheet_name="Resumen", header=None)
        metricas = {}
        KEYS = ["Total Revenue", "Occupancy %", "ADR", "Revenue PAR", "GOP", "GOP %",
                "Rooms Revenue", "F&B Revenue Total", "Rooms Occupied", "Spend PAR"]
        for _, row in df_res.iterrows():
            name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if name in KEYS:
                metricas[name] = {
                    "today": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "N/D",
                    "mtd": str(row.iloc[2]) if pd.notna(row.iloc[2]) else "N/D",
                    "forecast": str(row.iloc[3]) if pd.notna(row.iloc[3]) else "N/D",
                }

        # Hoja Alertas — días y su estado
        dias = []
        try:
            df_al = pd.read_excel(ruta, sheet_name="Alertas", header=None)
            for _, row in df_al.iterrows():
                dia_val = row.iloc[0]
                if isinstance(dia_val, (int, float)) and 1 <= dia_val <= 31:
                    estado_txt = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""
                    oob = "OUT" in estado_txt.upper()
                    diff = 0.0
                    try:
                        diff = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
                    except Exception:
                        pass
                    dias.append({
                        "dia": int(dia_val),
                        "fecha": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                        "oob": oob,
                        "diff": round(diff, 2),
                    })
        except Exception:
            pass

        # Top 3 alertas
        alertas = []
        oob_dias = [d for d in dias if d["oob"]]
        for d in oob_dias[:3]:
            alertas.append(f"Día {d['dia']}: Out of Balance — diferencia {d['diff']:,.2f} EUR")
        # Si hay menos de 3 alertas OOB, añadir métricas relevantes
        if len(alertas) < 3:
            gop_mtd = metricas.get("GOP %", {}).get("mtd", "")
            if gop_mtd and gop_mtd != "N/D":
                alertas.append(f"GOP % MTD: {gop_mtd}")
        if len(alertas) < 3:
            occ = metricas.get("Occupancy %", {}).get("forecast", "")
            if occ and occ != "N/D":
                alertas.append(f"Occupancy % Forecast: {occ}")

        return {
            "metricas": metricas,
            "dias": dias,
            "alertas": alertas[:3],
            "archivo": os.path.basename(ruta),
            "total_dias": len(dias),
            "dias_oob": len(oob_dias),
        }
    except Exception as e:
        return {"error": str(e)}


@app.route("/api/drr_daily_chart")
@login_required
def api_drr_daily_chart():
    """Devuelve serie diaria Revenue + Expenses para el gráfico del DRR."""
    ruta = _cargar_drr_procesado()
    if not ruta:
        return jsonify(None)
    try:
        df = _excel(ruta, sheet_name="Trial_Balance_Completo", header=0)
        income = df[df["Sección"] == "INCOME"].copy()
        income["Total"] = pd.to_numeric(income["Total"], errors="coerce").abs()
        expenses = df[df["Sección"] == "EXPENSES"].copy()
        expenses["Total"] = pd.to_numeric(expenses["Total"], errors="coerce").abs()
        daily_rev = income.groupby("Día")["Total"].sum()
        daily_exp = expenses.groupby("Día")["Total"].sum()
        oob_dias = set(
            int(d) for d in df[df["Out of Balance"].astype(str).str.contains("OOB", na=False)]["Día"].unique()
        )
        fechas_map = {int(row["Día"]): str(row["Fecha"])[:10]
                      for _, row in df.drop_duplicates("Día").iterrows()
                      if pd.notna(row["Fecha"])}
        dias = sorted([int(d) for d in daily_rev.index.tolist()])
        return jsonify({
            "dias": dias,
            "fechas": [fechas_map.get(d, str(d)) for d in dias],
            "revenue": [round(float(daily_rev.get(d, 0)), 0) for d in dias],
            "expenses": [round(float(daily_exp.get(d, 0)), 0) for d in dias],
            "oob": [d in oob_dias for d in dias],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/stats_drr")
def api_stats_drr():
    ruta = _cargar_drr_procesado()
    if not ruta:
        return jsonify(None)
    return jsonify(_leer_drr_stats(ruta))


@app.route("/api/upload_drr", methods=["POST"])
def api_upload_drr():
    """Recibe un .xlsm, lo guarda y ejecuta lector_drr.py."""
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsm"):
        return jsonify({"ok": False, "error": "Solo archivos .xlsm"}), 400

    os.makedirs(DRR_UPLOAD_DIR, exist_ok=True)
    save_path = os.path.join(DRR_UPLOAD_DIR, f.filename)
    f.save(save_path)

    # Ejecutar lector_drr.py
    script = os.path.join(BASE_DIR, "lector_drr.py")
    if not os.path.exists(script):
        return jsonify({"ok": False, "error": "lector_drr.py no encontrado"}), 500
    try:
        res = subprocess.run(
            [sys.executable, script, save_path],
            capture_output=True, text=True, timeout=120, cwd=BASE_DIR
        )
        if res.returncode != 0:
            return jsonify({"ok": False, "error": res.stderr[-300:] if res.stderr else "Error desconocido"}), 500
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "Timeout procesando DRR"}), 500

    # Leer el resultado
    ruta = _cargar_drr_procesado()
    if not ruta:
        return jsonify({"ok": False, "error": "No se generó el reporte"}), 500

    stats = _leer_drr_stats(ruta)
    return jsonify({"ok": True, "stats": stats})


# ── Notificaciones ────────────────────────────────────────────────────

@app.route("/api/stats_banco")
def api_stats_banco():
    """Resumen de conciliacion bancaria para el dashboard."""
    ruta = None
    hits = glob.glob(os.path.join(REPORTES_DIR, "conciliacion_*.xlsx"))
    if hits:
        hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        ruta = hits[0]
    if not ruta:
        return jsonify(None)
    try:
        df = pd.read_excel(ruta)
        total = len(df)
        conc = int((df["estado"] == "CONCILIADO").sum()) if "estado" in df.columns else 0
        pend = int((df["estado"] == "PENDIENTE").sum()) if "estado" in df.columns else 0
        diff = int((df["estado"] == "DIFERENCIA").sum()) if "estado" in df.columns else 0
        imp_pend = float(df.loc[df.get("estado", pd.Series()) == "PENDIENTE", "importe"].apply(safe_float).sum()) if "estado" in df.columns else 0
        # Alertas: pendientes con mas de 7 dias
        alertas = []
        if "estado" in df.columns and "fecha" in df.columns:
            from datetime import datetime
            hoy = datetime.now()
            for _, r in df[df["estado"] == "PENDIENTE"].iterrows():
                try:
                    f = pd.to_datetime(r["fecha"])
                    dias = (hoy - f).days
                    if dias > 7:
                        alertas.append({"concepto": str(r.get("concepto", ""))[:50], "importe": safe_float(r.get("importe", 0)), "dias": dias})
                except Exception:
                    pass
        return jsonify({"total": total, "conciliados": conc, "pendientes": pend,
                        "diferencias": diff, "importe_pendiente": round(imp_pend, 2),
                        "alertas": alertas[:10]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500



@app.route("/api/test_notif", methods=["POST"])
@login_required
def api_test_notif():
    """Envia una notificacion de prueba por todos los canales activos."""
    try:
        from notificaciones import enviar_por_canales, _load_config
        cfg = _load_config()
        asunto = "Test Yve.01 — Notificacion de prueba"
        html_body = "<p>Esta es una notificacion de prueba de <strong>Yve.01</strong>.</p><p>Si ves este mensaje, las notificaciones estan correctamente configuradas.</p>"
        txt_body = "Notificacion de prueba de Yve.01. Si ves este mensaje, las notificaciones funcionan correctamente."
        resultados = enviar_por_canales(asunto, html_body, txt_body, "test")
        canales_activos = [k for k, v in cfg.get("canales", {}).items() if v]
        return jsonify({
            "ok": True,
            "resultados": resultados,
            "canales_activos": canales_activos,
            "message": ("Enviado por: " + ", ".join(resultados.keys())) if resultados else "Sin canales configurados. Activa Email, Slack o WhatsApp en el panel de Notificaciones."
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/notif_config", methods=["GET"])
def api_notif_config_get():
    """Devuelve la configuración de notificaciones."""
    path = os.path.join(BASE_DIR, "datos-referencia", "notif_config.json")
    default = {
        "canales": {"email": True, "whatsapp": False, "telegram": False, "slack": False, "push": True},
        "email": "", "whatsapp": "", "telegram_chat": "", "slack_webhook": "",
        "alertas": {
            "ar_discrepancia": True, "ar_falta_di": True,
            "ap_discrepancia": True, "drr_oob": True,
            "banco_sin_conciliar": True, "factura_pendiente_firma": False,
        },
        "frecuencia": "inmediata",
    }
    if os.path.exists(path):
        try:
            saved = json.load(open(path))
            default.update(saved)
        except Exception:
            pass
    return jsonify(default)


@app.route("/api/notif_config", methods=["POST"])
def api_notif_config_save():
    """Guarda la configuración de notificaciones."""
    path = os.path.join(BASE_DIR, "datos-referencia", "notif_config.json")
    data = request.get_json(silent=True) or {}
    try:
        with open(path, "w") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500



@app.route("/fb/api/upload_ventas", methods=["POST"])
@login_required
def api_upload_ventas_pos():
    """Sube un Excel/CSV de ventas POS y lo appendea a ventas_fb_diarias.xlsx."""
    import pandas as pd, io
    from datetime import datetime
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "No se recibio archivo"}), 400
    fname = f.filename.lower()
    try:
        if fname.endswith('.csv'):
            df_new = pd.read_csv(f)
        elif fname.endswith(('.xlsx', '.xls')):
            df_new = pd.read_excel(f)
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Usa .xlsx o .csv"}), 400

        # Normalize columns — accept flexible naming
        col_map = {}
        for col in df_new.columns:
            cl = col.lower().replace(' ','_')
            if 'fecha' in cl or 'date' in cl:                col_map[col] = 'fecha'
            elif 'receta' in cl or 'recipe' in cl or 'id' in cl: col_map[col] = 'id_receta'
            elif 'plato' in cl or 'nombre' in cl or 'name' in cl: col_map[col] = 'nombre_plato'
            elif 'categ' in cl:                               col_map[col] = 'categoria'
            elif 'unidad' in cl or 'qty' in cl or 'cantidad' in cl: col_map[col] = 'unidades_vendidas'
            elif 'precio' in cl or 'price' in cl or 'unit' in cl:   col_map[col] = 'precio_unitario'
            elif 'total' in cl or 'venta' in cl or 'revenue' in cl: col_map[col] = 'total_venta'
        df_new = df_new.rename(columns=col_map)

        # Validate minimum required columns
        required = ['fecha', 'nombre_plato', 'total_venta']
        missing = [c for c in required if c not in df_new.columns]
        if missing:
            return jsonify({"ok": False, "error": "Faltan columnas: " + ", ".join(missing) +
                           ". El archivo debe tener: fecha, nombre_plato, total_venta"}), 400

        # Fill defaults for optional columns
        df_new['fecha'] = pd.to_datetime(df_new['fecha'], errors='coerce').dt.strftime('%Y-%m-%d')
        if 'id_receta'          not in df_new.columns: df_new['id_receta']          = 'IMPORT'
        if 'categoria'          not in df_new.columns: df_new['categoria']          = 'Importado'
        if 'unidades_vendidas'  not in df_new.columns: df_new['unidades_vendidas']  = 1
        if 'precio_unitario'    not in df_new.columns:
            df_new['precio_unitario'] = df_new['total_venta'] / df_new['unidades_vendidas'].replace(0, 1)

        # Load existing and append
        path = os.path.join(BASE_DIR, "datos-referencia", "ventas_fb_diarias.xlsx")
        df_existing = pd.read_excel(path) if os.path.exists(path) else pd.DataFrame()
        df_combined = pd.concat([df_existing, df_new[['fecha','id_receta','nombre_plato',
                                                       'categoria','unidades_vendidas',
                                                       'precio_unitario','total_venta']]], ignore_index=True)
        df_combined = df_combined.drop_duplicates()
        df_combined.to_excel(path, index=False)

        # Invalidate caches
        if 'ventas_fb_diarias.xlsx' in _EXCEL_CACHE:
            del _EXCEL_CACHE['ventas_fb_diarias.xlsx']
        try:
            from tab_fb_dashboard import _invalidate
            _invalidate()
        except Exception:
            pass

        return jsonify({
            "ok": True,
            "filas_importadas": len(df_new),
            "total_filas": len(df_combined),
            "fechas": df_new['fecha'].dropna().unique().tolist()[:5],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/cache/clear", methods=["POST"])
@login_required
def api_cache_clear():
    """Limpia el cache de Excel en memoria."""
    global _EXCEL_CACHE
    _EXCEL_CACHE = {}
    try:
        from tab_fb_dashboard import _invalidate
        _invalidate()
    except Exception:
        pass
    try:
        from dashboard_calipolis import _CAL_CACHE
        _CAL_CACHE.clear()
    except Exception:
        pass
    return jsonify({"ok": True, "message": "Cache limpiado"})


@app.route("/api/notificaciones")
def api_notificaciones():
    """Devuelve historial de notificaciones."""
    hist_path = os.path.join(BASE_DIR, "datos-referencia", "notificaciones_historial.json")
    if os.path.exists(hist_path):
        try:
            with open(hist_path, "r", encoding="utf-8") as f:
                return jsonify(json.load(f))
        except Exception:
            pass
    return jsonify([])


@app.route("/api/enviar_notificaciones", methods=["POST"])
def api_enviar_notificaciones():
    """Escanea alertas y envía pendientes."""
    try:
        sys.path.insert(0, BASE_DIR)
        from exportador_final import generar_reporte_ejecutivo, generar_excel_consolidado
        alertas = enviar_pendientes()
        return jsonify({"ok": True, "enviadas": len(alertas)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500


# ── Chat AI — Yve Copilot ──────────────────────────────────────────────

def _hotel_name():
    cfg_path = os.path.join(BASE_DIR, "datos-referencia", "hotel_config.json")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return cfg.get("hotel", {}).get("nombre", "")
        except Exception:
            pass
    return ""

@app.route("/")
def index():
    # Muestra landing si no hay sesión, dashboard si la hay
    if not current_user.is_authenticated:
        return LANDING_PAGE
    name = _hotel_name()
    tag = name if name else "AR Dashboard"
    configured = "true" if name else "false"
    user_name = current_user.nombre
    user_rol  = current_user.rol
    out = HTML.replace("__HOTEL_TAG__", tag).replace("__CONFIGURED__", configured)
    admin_display = "inline" if user_rol in ("admin", "financial_controller") else "none"
    out = out.replace("__USER_NAME__", user_name).replace("__USER_ROL__", user_rol)
    out = out.replace("__ADMIN_DISPLAY__", admin_display)
    return out

@app.route("/app")
@login_required
def app_dashboard():
    return index()

# ── HTML ───────────────────────────────────────────────────────────────────

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
<title>Yve.01 — Dashboard</title>
<script async src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;
  --acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;
  --tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;--yel:#eab308;--pur:#8b5cf6;
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  background:var(--bg);color:var(--tx);
  font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  min-height:100vh;line-height:1.5;position:relative;
}
/* Gradient glow igual que el login — sutil, solo ambiente */
body::before{
  content:'';position:fixed;inset:0;z-index:0;pointer-events:none;
  background:
    radial-gradient(900px 500px at 90% -5%,rgba(59,130,246,.1),transparent 60%),
    radial-gradient(700px 400px at -5% 105%,rgba(139,92,246,.08),transparent 55%)
}
.main{position:relative;z-index:1}

/* ── NAV ── */
.nav{
  background:rgba(15,23,42,.92);
  backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  border-bottom:1px solid var(--s2);
  padding:0 24px;height:60px;
  display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200
}
.logo{display:flex;align-items:baseline;gap:10px;flex-shrink:0}
.logo-name{font-size:20px;font-weight:800;color:#fff;letter-spacing:-0.5px}
.logo-tag{font-size:11px;color:var(--mut);font-weight:400;white-space:nowrap}
.logo-dot{width:8px;height:8px;border-radius:50%;background:var(--acc);flex-shrink:0;box-shadow:0 0 8px var(--acc)}
.logo-dot-one{color:var(--acc2)}
.logo-mark{display:none}
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
.sc:hover{border-color:var(--s3);transform:translateY(-1px)}
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
.mid{display:grid;grid-template-columns:1fr 300px;gap:16px;margin-bottom:24px}
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

/* ── TABS ────────────────────────────────────────────── */
.tabs{display:flex;gap:2px;margin-bottom:24px;border-bottom:1px solid var(--s2);padding-bottom:0;overflow-x:auto}
.tabs::-webkit-scrollbar{height:3px}
.tabs::-webkit-scrollbar-thumb{background:var(--s3);border-radius:2px}
.tab{padding:10px 18px;background:none;border:none;color:var(--mut);cursor:pointer;font-size:.85rem;font-weight:600;border-bottom:2px solid transparent;transition:.18s;white-space:nowrap}
.tab:hover{color:var(--tx)}
.tab.active{color:var(--acc2);border-bottom-color:var(--acc)}
.panel{display:none}.panel.active{display:block;animation:fadeIn .18s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:none}}
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

/* ── Chat AI — Yve Copilot ─────────────────────────────── */
#chat-fab{position:fixed;bottom:28px;right:28px;z-index:1000;
  display:flex;align-items:center;gap:10px;
  background:linear-gradient(135deg,#7c3aed,#3b82f6);
  color:#fff;border:none;border-radius:50px;padding:14px 22px 14px 18px;
  cursor:pointer;font-size:.95rem;font-weight:700;letter-spacing:.02em;
  box-shadow:0 4px 24px rgba(124,58,237,.5);transition:.2s;white-space:nowrap}
#chat-fab:hover{transform:translateY(-2px);box-shadow:0 8px 32px rgba(124,58,237,.6)}
#chat-fab .fab-dot{width:9px;height:9px;border-radius:50%;
  background:#22c55e;box-shadow:0 0 6px #22c55e;animation:pulse-dot 2s infinite}
@keyframes pulse-dot{0%,100%{opacity:1}50%{opacity:.4}}

#chat-panel{position:fixed;bottom:0;right:0;width:420px;height:100vh;
  background:#0f172a;border-left:1px solid #1e293b;z-index:999;
  display:flex;flex-direction:column;transform:translateX(100%);
  transition:transform .3s cubic-bezier(.4,0,.2,1);
  box-shadow:-8px 0 40px rgba(0,0,0,.5)}
#chat-panel.open{transform:translateX(0)}
@media(max-width:480px){#chat-panel{width:100vw}}

#chat-header{padding:18px 20px;background:#1e293b;border-bottom:1px solid #334155;
  display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
#chat-header .chat-title{display:flex;align-items:center;gap:10px}
#chat-header .chat-title span:first-child{font-size:1.5rem}
#chat-header h3{margin:0;font-size:1rem;font-weight:700;color:#f1f5f9}
#chat-header p{margin:0;font-size:.75rem;color:#60a5fa}
#chat-close{background:none;border:none;color:#64748b;cursor:pointer;
  font-size:1.4rem;padding:4px 8px;border-radius:6px;transition:.15s}
#chat-close:hover{background:#334155;color:#f1f5f9}

#chat-msgs{flex:1;overflow-y:auto;padding:20px;display:flex;
  flex-direction:column;gap:14px;scroll-behavior:smooth}
#chat-msgs::-webkit-scrollbar{width:4px}
#chat-msgs::-webkit-scrollbar-track{background:transparent}
#chat-msgs::-webkit-scrollbar-thumb{background:#334155;border-radius:2px}

.msg{max-width:90%;padding:12px 16px;border-radius:16px;font-size:.88rem;
  line-height:1.55;animation:msgIn .2s ease}
@keyframes msgIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.msg.user{align-self:flex-end;background:linear-gradient(135deg,#3b82f6,#2563eb);
  color:#fff;border-bottom-right-radius:4px}
.msg.bot{align-self:flex-start;background:#1e293b;color:#e2e8f0;
  border:1px solid #334155;border-bottom-left-radius:4px}
.msg.bot.thinking{color:#64748b;font-style:italic;border-style:dashed}

#chat-suggestions{padding:0 16px 12px;display:flex;flex-wrap:wrap;gap:7px;flex-shrink:0}
.sug{background:#1e293b;border:1px solid #334155;color:#94a3b8;
  border-radius:20px;padding:6px 13px;font-size:.78rem;cursor:pointer;
  transition:.15s;white-space:nowrap}
.sug:hover{border-color:#60a5fa;color:#60a5fa;background:#1e3a5f}

.typing{display:inline-flex;gap:4px;align-items:center;padding:4px 0}
.typing span{width:7px;height:7px;border-radius:50%;background:#60a5fa;animation:typingDot 1.2s infinite}
.typing span:nth-child(2){animation-delay:.2s}
.typing span:nth-child(3){animation-delay:.4s}
@keyframes typingDot{0%,60%,100%{opacity:.3;transform:translateY(0)}30%{opacity:1;transform:translateY(-3px)}}
#chat-input-row{padding:14px 16px;border-top:1px solid #1e293b;
  display:flex;gap:10px;align-items:center;flex-shrink:0;background:#0f172a}
#chat-input{flex:1;background:#1e293b;border:1px solid #334155;color:#f1f5f9;
  border-radius:24px;padding:11px 18px;font-size:.88rem;outline:none;
  resize:none;font-family:inherit;transition:.15s;max-height:120px}
#chat-input:focus{border-color:#60a5fa;box-shadow:0 0 0 2px rgba(96,165,250,.15)}
#chat-input::placeholder{color:#475569}
#chat-send{background:linear-gradient(135deg,#7c3aed,#3b82f6);border:none;
  color:#fff;border-radius:50%;width:42px;height:42px;cursor:pointer;
  font-size:1.1rem;flex-shrink:0;transition:.15s;display:flex;
  align-items:center;justify-content:center}
#chat-send:hover{transform:scale(1.08)}
#chat-send:disabled{opacity:.4;cursor:not-allowed;transform:none}

/* ── DRR Panel ─────────────────────────────────────────── */
.drr-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:22px}
@media(max-width:900px){.drr-metrics{grid-template-columns:repeat(2,1fr)}}
@media(max-width:500px){.drr-metrics{grid-template-columns:1fr}}
.drr-mc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px}
.drr-mc .mc-name{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:12px}
.drr-mc .mc-row{display:flex;justify-content:space-between;font-size:.8rem;padding:3px 0}
.drr-mc .mc-row .mc-k{color:var(--dim)}.drr-mc .mc-row .mc-v{color:var(--tx);font-weight:700}
.drr-upload{display:flex;align-items:center;gap:14px;margin-bottom:22px;flex-wrap:wrap}
.drr-upload label{margin:0;padding:10px 18px;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;transition:.15s;white-space:nowrap}
.drr-upload label:hover{box-shadow:0 0 20px rgba(59,130,246,.4)}
.drr-upload .drr-status{font-size:.8rem;color:var(--dim)}
.drr-days{display:grid;grid-template-columns:repeat(auto-fill,minmax(80px,1fr));gap:6px;margin-bottom:22px}
.drr-day{text-align:center;padding:10px 4px;border-radius:10px;font-size:.75rem;font-weight:700;border:1px solid var(--s2)}
.drr-day.ok{background:rgba(34,197,94,.08);color:var(--grn);border-color:rgba(34,197,94,.2)}
.drr-day.oob{background:rgba(239,68,68,.1);color:var(--red);border-color:rgba(239,68,68,.25)}
.drr-day.empty{background:var(--s1);color:var(--dim);border-color:var(--s2)}
.drr-day .day-n{font-size:1.1rem;margin-bottom:2px}
.drr-alerts .da-item{display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid var(--s2)}
.drr-alerts .da-item:last-child{border-bottom:none}
.drr-alerts .da-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;margin-top:5px;background:var(--ora)}
.drr-alerts .da-txt{font-size:.85rem;color:var(--tx)}


/* ── GUIDED TOUR ─────────────────────────────────── */
#tour-overlay{
  position:fixed;inset:0;z-index:9000;pointer-events:none;
  background:transparent;transition:background .35s;
}
#tour-overlay.active{pointer-events:all;background:rgba(4,9,20,.72)}
#tour-spotlight{
  position:fixed;z-index:9001;
  border-radius:14px;
  box-shadow:0 0 0 9999px rgba(4,9,20,.72);
  transition:left .38s cubic-bezier(.4,0,.2,1),
             top .38s cubic-bezier(.4,0,.2,1),
             width .38s cubic-bezier(.4,0,.2,1),
             height .38s cubic-bezier(.4,0,.2,1);
  pointer-events:none;
  border:2px solid rgba(59,130,246,.65);
  animation:spotGlow 2s ease-in-out infinite;
}
@keyframes spotGlow{
  0%,100%{border-color:rgba(59,130,246,.55);box-shadow:0 0 0 9999px rgba(4,9,20,.72),0 0 0 4px rgba(59,130,246,.12)}
  50%{border-color:rgba(96,165,250,.9);box-shadow:0 0 0 9999px rgba(4,9,20,.72),0 0 0 6px rgba(59,130,246,.22)}
}
#tour-card{
  position:fixed;z-index:9002;
  background:linear-gradient(160deg,rgba(30,41,59,.98),rgba(10,18,35,.98));
  border:1px solid rgba(59,130,246,.35);
  border-radius:18px;padding:24px 26px;width:330px;
  box-shadow:0 24px 64px rgba(0,0,0,.7),0 0 0 1px rgba(59,130,246,.08) inset;
  transition:left .38s cubic-bezier(.4,0,.2,1),top .38s cubic-bezier(.4,0,.2,1);
}
#tour-card.entering{animation:cardEnter .28s cubic-bezier(.2,.8,.2,1)}
@keyframes cardEnter{from{opacity:0;transform:translateY(8px) scale(.97)}to{opacity:1;transform:none}}
.tour-content-wrap{transition:opacity .18s;min-height:56px}
.tour-content-wrap.fading{opacity:0}
#tour-card h3{font-size:15px;font-weight:800;margin-bottom:8px;color:var(--tx);letter-spacing:-.2px}
#tour-card p{font-size:13.5px;color:var(--mut);line-height:1.65;margin-bottom:0}
.tour-progress-bar{height:2px;background:var(--s2);border-radius:1px;margin-bottom:16px;overflow:hidden}
.tour-progress-fill{height:100%;background:linear-gradient(90deg,var(--acc),var(--pur));border-radius:1px;transition:width .38s cubic-bezier(.4,0,.2,1)}
.tour-footer{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:18px}
.tour-counter{font-size:11px;color:var(--dim);font-weight:600;letter-spacing:.3px}
.tour-dots{display:flex;gap:5px;align-items:center}
.tour-dot{width:6px;height:6px;border-radius:50%;background:var(--s3);transition:all .25s}
.tour-dot.active{background:var(--acc);width:20px;border-radius:3px}
.tour-btns{display:flex;gap:7px;align-items:center}
.tour-btn-skip{background:none;border:none;font-size:11px;color:var(--dim);cursor:pointer;padding:5px 7px;border-radius:6px;transition:.15s;letter-spacing:.2px}
.tour-btn-skip:hover{color:var(--mut)}
.tour-btn-prev{background:rgba(51,65,85,.6);border:1px solid var(--s2);color:var(--mut);font-size:12px;font-weight:600;padding:7px 13px;border-radius:9px;cursor:pointer;transition:.15s}
.tour-btn-prev:hover{background:var(--s2);color:var(--tx)}
.tour-btn-next{background:linear-gradient(135deg,var(--acc),#2563eb);border:none;color:#fff;font-size:12px;font-weight:700;padding:8px 18px;border-radius:9px;cursor:pointer;box-shadow:0 0 14px rgba(59,130,246,.4);transition:.15s;white-space:nowrap}
.tour-btn-next:hover{box-shadow:0 0 22px rgba(59,130,246,.6);transform:translateY(-1px)}
.tour-target{outline:2px solid var(--acc)!important;outline-offset:4px!important;border-radius:10px!important;animation:tourPulse 2s ease-in-out infinite!important;position:relative!important;z-index:9001!important}
@keyframes tourPulse{0%,100%{outline-color:rgba(59,130,246,.6)}50%{outline-color:rgba(96,165,250,1)}}
@media(max-width:600px){
  #tour-card{width:calc(100vw - 24px)!important;left:12px!important;bottom:16px!important;top:auto!important}
}

/* ── F&B Sub-tabs ─────────────────────────────────── */
.fb-sub{background:none;border:none;color:var(--mut);padding:7px 14px;border-radius:7px;cursor:pointer;font-size:13px;font-weight:500;font-family:inherit;transition:.15s;white-space:nowrap}
.fb-sub:hover{color:var(--tx)}
.fb-sub.active{background:var(--s2);color:var(--tx);font-weight:600}
.stock-bar{height:8px;border-radius:4px;background:var(--s2);overflow:hidden;margin-top:4px}
.stock-fill{height:100%;border-radius:4px;transition:width .6s}

/* ── Skeleton loading ─────────────────────────────────────── */
@keyframes shimmer{0%{background-position:-400px 0}100%{background-position:400px 0}}
.skel{border-radius:10px;background:linear-gradient(90deg,var(--s1) 25%,var(--s2) 50%,var(--s1) 75%);background-size:800px 100%;animation:shimmer 1.4s infinite}
.skel-card{height:88px;border-radius:13px}
.skel-line{height:14px;border-radius:6px;margin-bottom:10px}
.skel-line.short{width:60%}
.skel-line.med{width:80%}
.skel-title{height:20px;border-radius:6px;width:40%;margin-bottom:16px}
/* ── Sparklines en stat cards ── */
.sc-spark{display:block;width:100%;height:24px;margin-top:8px;opacity:.7}
.sc:hover .sc-spark{opacity:1}

/* ── DRR Revenue chart ── */
.drr-chart-wrap{height:200px;position:relative;margin-bottom:22px}

/* ── Mobile / Responsive ──────────────────────────────────── */
@media(max-width:480px){
  /* Nav */
  .nav{height:54px;padding:0 12px;gap:8px}
  .logo-tag{display:none}
  .pill{display:none}
  .btn-run{padding:7px 12px;font-size:11px;gap:5px}
  .btn-ref{padding:5px 10px;font-size:11px}

  /* Main */
  .main{padding:12px}

  /* Stats — single column on very small */
  .stats{grid-template-columns:repeat(2,1fr)!important;gap:8px}
  .sc{padding:14px 12px}
  .sc-val{font-size:22px!important}
  .sc-spark{height:18px}

  /* Tabs — scrollable pill bar */
  .tabs{gap:2px;padding:3px}
  .tab{padding:6px 11px;font-size:.78rem}

  /* Cards */
  .card{padding:16px}

  /* Mid row → single col */
  .mid{grid-template-columns:1fr!important}
  .chart-wrap{height:160px}

  /* AR Real 2-col → 1 col */
  #panel-ar_real [style*="grid-template-columns:1fr 1fr"]{
    grid-template-columns:1fr!important
  }

  /* DRR metrics */
  .drr-metrics{grid-template-columns:1fr!important}
  .drr-days{grid-template-columns:repeat(5,1fr)!important}

  /* Calipolis hotels */
  #cal-hoteles{grid-template-columns:1fr!important}
  #panel-calipolis [style*="grid-template-columns:1fr 1fr"]{
    grid-template-columns:1fr!important
  }

  /* Multi-hotel KPIs 4col → 2col */
  #mh-kpis{grid-template-columns:repeat(2,1fr)!important}
  #mh-status{grid-template-columns:1fr!important}

  /* Modal full screen */
  .overlay{padding:0;align-items:flex-end}
  .modal{border-radius:18px 18px 0 0;max-width:100%;padding:20px}

  /* Chat panel full width */
  #chat-panel{width:100vw}

  /* Tables: always wrap */
  .tbl-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
}

@media(min-width:481px) and (max-width:768px){
  .main{padding:18px}
  .stats{grid-template-columns:repeat(3,1fr)!important;gap:10px}
  .mid{grid-template-columns:1fr!important}
  #cal-hoteles{grid-template-columns:1fr!important}
  #mh-kpis{grid-template-columns:repeat(4,1fr)!important}
  #panel-ar_real [style*="grid-template-columns:1fr 1fr"]{
    grid-template-columns:1fr!important
  }
}

/* ── Dropdown menus (navbar) — sustituye estilos inline hardcodeados ── */
.dropdown{display:inline-block;position:relative}
.menu{display:none;position:absolute;top:46px;right:0;background:var(--s1);border:1px solid var(--s2);border-radius:11px;padding:7px;z-index:1000;min-width:218px;box-shadow:0 12px 40px rgba(0,0,0,.45)}
.menu.open{display:block;animation:menuIn .14s ease}
@keyframes menuIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}
.menu-head{padding:7px 12px 5px;color:var(--dim);font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px}
.menu-item{display:flex;align-items:center;gap:9px;width:100%;text-align:left;padding:9px 12px;color:var(--tx);text-decoration:none;border:none;background:transparent;cursor:pointer;border-radius:7px;font-size:12.5px;font-family:inherit;transition:background .12s;white-space:nowrap}
.menu-item:hover{background:var(--s2)}
.menu-sep{border-top:1px solid var(--s2);margin:6px 5px}
.rol-sub{padding-left:8px;border-left:2px solid var(--s2);margin:2px 0 2px 12px}
.rol-sub .menu-item{font-size:12px;color:var(--mut);padding:7px 11px}
.rol-sub .menu-item:hover{color:var(--tx)}

</style>
</head>
<body>

<nav class="nav">
  <div class="logo">
    <div class="logo-dot"></div>
    <span class="logo-name">Yve<span style="color:var(--acc2)">.01</span></span>
    <span class="logo-tag">__HOTEL_TAG__</span>
  </div>
  <div class="nav-mid"></div>
  <div class="nav-right">
    <span class="pill" id="date-pill">—</span>
    <span class="pill" style="color:var(--acc2)">👤 __USER_NAME__</span>

    <div class="dropdown">
      <button class="btn-ref" onclick="toggleMenu('reportes-menu')" title="Descargar reportes" data-i18n="nav.reportes">📄 Reportes</button>
      <div id="reportes-menu" class="menu">
        <div class="menu-head">Reportes PDF</div>
        <a href="/api/reportes/diario" class="menu-item">📄 Diario</a>
        <a href="/api/reportes/semanal" class="menu-item">📊 Semanal</a>
        <a href="/api/reportes/mensual" class="menu-item">📈 Mensual</a>
        <div class="menu-sep"></div>
        <div class="menu-head">Ejecutivos</div>
        <a href="/api/reportes/ejecutivo.pdf" class="menu-item">🎯 Ejecutivo PDF</a>
        <a href="/api/reportes/consolidado.xlsx" class="menu-item">📊 Consolidado Excel</a>
      </div>
    </div>

    <button class="btn-ref" onclick="loadAll()" title="Actualizar datos" data-i18n="nav.actualizar">↻ Actualizar</button>

    <button class="btn-run" id="btn-run" onclick="runPipeline()">
      <div class="spin" id="spin"></div>
      <span id="run-lbl" data-i18n="nav.procesar">⚡ Procesar Facturas</span>
    </button>

    <div class="dropdown">
      <button class="btn-ref" onclick="toggleMenu('main-menu')" title="Más opciones" style="font-size:17px;line-height:1;padding:5px 12px">⋯</button>
      <div id="main-menu" class="menu">
        <div class="menu-head" data-i18n="menu.presentacion">Presentación</div>
        <button class="menu-item" data-i18n="nav.tour" onclick="tourStart();document.getElementById('main-menu').classList.remove('open')">🎯 Tour guiado</button>
        <button class="menu-item" id="btn-demo" onclick="toggleDemoMode()"><span data-i18n="nav.demo">🎭 Demo Mode</span></button>
        <div class="menu-sep"></div>
        <div class="menu-head" data-i18n="menu.cambiarRol">Cambiar rol</div>
        <button class="menu-item" id="rol-btn">👤 Admin</button>
        <div id="rol-menu" class="rol-sub" style="display:none">
          <button class="menu-item" onclick="cambiarRol('admin')">🔑 Administrador</button>
          <button class="menu-item" onclick="cambiarRol('financial_controller')">💰 Controller Financiero</button>
          <button class="menu-item" onclick="cambiarRol('income_auditor')">📊 Income Auditor</button>
          <button class="menu-item" onclick="cambiarRol('fb_manager')">🍽️ Jefe F&B</button>
          <button class="menu-item" onclick="cambiarRol('jefe_otras')">🛠️ Jefe Servicios</button>
        </div>
        <div class="menu-sep"></div>
        <a href="/configuracion/" class="menu-item" data-i18n="nav.config">⚙️ Configuración</a>
        <a href="/admin/" class="menu-item" style="display:__ADMIN_DISPLAY__" data-i18n="menu.admin">👥 Administración</a>
        <div class="menu-sep"></div>
        <a href="/logout" class="menu-item" data-i18n="nav.salir" style="color:#f87171">↩ Cerrar sesión</a>
      </div>
    </div>
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
    <span id="status-txt" data-i18n="status.cargando">Cargando datos...</span>
  </div>

  <!-- TABS -->
  <div class="tabs">
    <button class="tab active" onclick="switchTab('ar',this)" data-i18n="tab.ar">📥 AR — OTAs</button>
    <button class="tab" onclick="switchTab('ap',this)" data-i18n="tab.ap">📦 AP — Proveedores</button>
    <button class="tab" onclick="switchTab('drr',this)" data-i18n="tab.drr">📊 DRR</button>
    <button class="tab" onclick="switchTab('banco',this)" data-i18n="tab.banco">🏦 Banco</button>
    <button class="tab" onclick="switchTab('notif',this)" data-i18n="tab.notif">🔔 Notificaciones</button>
    <button class="tab" onclick="switchTab('fb',this)" id="tab-fb" data-i18n="tab.fb">🍽️ F&amp;B Cost</button>
    <button class="tab" onclick="switchTab('ar_real',this)" id="tab-ar-real" data-i18n="tab.arreal">🏢 AR Real</button>
    <button class="tab" onclick="switchTab('calipolis',this)" id="tab-calipolis" data-i18n="tab.calipolis">🏩 Calipolis</button>
    <button class="tab" onclick="switchTab('multi_hotel',this)" id="tab-multi-hotel" data-i18n="tab.multihotel">🏨 Multi-Hotel</button>
  </div>

  <div id="panel-ar" class="panel active">
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ar" class="btn-ref" style="text-decoration:none">⬇️ Excel</a><a href="/api/exportar/ar/pdf" class="btn-ref" style="text-decoration:none">📄 PDF</a><a href="/aprobaciones-ar/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AR">📲 Aprobar facturas AR</a></div>
  <!-- STATS -->
  <div class="stats">
    <div class="sc hl c-blu">
      <div class="sc-lbl" data-i18n="sc.procesadas">Facturas procesadas</div>
      <div class="sc-val" id="s-tot">—</div>
      <div class="sc-sub" data-i18n="sc.ciclo">último ciclo AR</div>
    </div>
    <div class="sc">
      <div class="sc-lbl" data-i18n="sc.importe">Importe total</div>
      <div class="sc-val" id="s-imp" style="font-size:18px;letter-spacing:-0.5px">—</div>
      <div class="sc-sub" data-i18n="sc.eurProcesados">EUR procesados</div>
    </div>
    <div class="sc c-grn">
      <div class="sc-lbl" data-i18n="sc.correctas">Correctas</div>
      <div class="sc-val" id="s-ok">—</div>
      <div class="sc-sub" data-i18n="sc.sinIncidencias">sin incidencias</div>
    </div>
    <div class="sc c-red">
      <div class="sc-lbl" data-i18n="sc.discrepancias">Discrepancias</div>
      <div class="sc-val" id="s-disc">—</div>
      <div class="sc-sub" id="s-disc-sub">reclamable: —</div>
    </div>
    <div class="sc c-ora">
      <div class="sc-lbl" data-i18n="sc.di">Certif. DI pendiente</div>
      <div class="sc-val" id="s-di">—</div>
      <div class="sc-sub" data-i18n="sc.extranjer">facturas extranjeras</div>
    </div>
    <div class="sc c-pur">
      <div class="sc-lbl" data-i18n="sc.pendiente">Pendientes firma</div>
      <div class="sc-val" id="s-pend">—</div>
      <div class="sc-sub" id="s-pend-sub">— apr · — rec</div>
    </div>
  </div>

  <!-- MID ROW -->
  <div class="mid">
    <div class="card">
      <div class="card-title" data-i18n="card.porOta">Facturas por OTA</div>
      <div class="chart-wrap"><canvas id="ota-chart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.resumen">Resumen de estados</div>
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
            <th data-i18n="th.archivo">Archivo</th>
            <th data-i18n="th.factura">Nº Factura</th>
            <th data-i18n="th.ota">OTA</th>
            <th data-i18n="th.hotel">Hotel</th>
            <th data-i18n="th.fecha">Fecha</th>
            <th data-i18n="th.importe">Importe bruto</th>
            <th data-i18n="th.comision">% Com.</th>
            <th data-i18n="th.estado">Estado</th>
            <th data-i18n="th.estadoDI">Estado DI</th>
            <th data-i18n="th.discrepancia">Discrepancia</th>
            <th data-i18n="th.aprobacion">Aprobación</th>
          </tr>
        </thead>
        <tbody id="tbl-body">
          <tr><td colspan="11" class="empty"><p>Sin datos. Pulsa ⚡ Procesar Facturas para empezar.</p></td></tr>
        </tbody>
      </table>
    </div>
  </div>

  </div><!-- /panel-ar -->

  <!-- PANEL AP -->
  <div id="panel-ap" class="panel">
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ap" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px" data-i18n="btn.downloadExcel" data-i18n="btn.downloadExcel">⬇️ Descargar Excel</a><a href="/aprobaciones-ap/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AP" data-i18n="btn.aprobarAP">📲 Aprobar facturas AP</a></div>
    <div class="stats" id="stats-ap-grid">
      <div class="sc hl c-blu"><div class="sc-lbl" data-i18n="ap.totalLabel">Total Facturas AP</div><div class="sc-val" id="ap-total">—</div><div class="sc-sub" data-i18n="ap.proveedores">proveedores</div></div>
      <div class="sc"><div class="sc-lbl" data-i18n="ap.importe">Importe Total</div><div class="sc-val" id="ap-importe" style="font-size:18px;letter-spacing:-.5px">—</div><div class="sc-sub">EUR</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-i18n="ap.matchOk">Matches OK</div><div class="sc-val" id="ap-matches">—</div><div class="sc-sub" data-i18n="ap.fbOtras">F&B + OTRAS</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-i18n="sc.discrepancias">Discrepancias</div><div class="sc-val" id="ap-disc">—</div><div class="sc-sub" data-i18n="ap.vsPo">vs PO</div></div>
      <div class="sc c-ora"><div class="sc-lbl" data-i18n="ap.sinPO">Sin PO</div><div class="sc-val" id="ap-sinpo">—</div><div class="sc-sub" data-i18n="ap.sinOrden">sin orden compra</div></div>
      <div class="sc c-pur"><div class="sc-lbl" data-i18n="ap.aprobadas">Aprobadas</div><div class="sc-val" id="ap-aprobadas">—</div><div class="sc-sub" data-i18n="ap.firmadas">firmadas</div></div>
    </div>
    <div class="card" style="margin-bottom:22px">
      <div class="card-title" data-i18n="card.facturasAP">Facturas AP</div>
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Factura</th><th data-i18n="th.proveedor">Proveedor</th><th data-i18n="th.tipo">Tipo</th><th>Total</th><th data-i18n="th.cuenta">Cuenta</th><th data-i18n="th.matching">Matching</th><th data-i18n="th.aprobacion">Aprobación</th></tr></thead>
          <tbody id="ap-tbody"><tr><td colspan="7" class="empty"><p>Sin datos AP.</p></td></tr></tbody>
        </table>
      </div>
      <span id="ap-count" style="font-size:.75rem;color:var(--dim);margin-top:8px;display:block"></span>
    </div>
  </div><!-- /panel-ap -->

  <!-- PANEL DRR -->
  <div id="panel-drr" class="panel">
  <div style="display:flex;justify-content:flex-end;margin-bottom:14px"><a href="/api/exportar/drr" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px">⬇️ Descargar Excel</a></div>
    <div class="drr-upload">
      <label for="drr-file-input" data-i18n="btn.uploadDrr">📂 Subir DRR (.xlsm)</label>
      <input type="file" id="drr-file-input" accept=".xlsm" style="display:none" onchange="uploadDRR(this)">
      <span class="drr-status" id="drr-status">Sin archivo cargado</span>
    </div>

    <!-- KPI Metrics -->
    <div class="drr-metrics" id="drr-metrics">
      <div class="empty"><div class="ei">📊</div><p>Sube un archivo DRR para ver las métricas.</p></div>
    </div>

    <!-- Revenue Chart -->
    <div class="card" style="margin-bottom:22px" id="drr-chart-card" style="display:none">
      <div class="card-title" data-i18n="card.revDiario">Revenue Diario</div>
      <div class="drr-chart-wrap"><canvas id="drr-revenue-chart"></canvas></div>
    </div>

    <!-- Days grid -->
    <div class="card" style="margin-bottom:22px">
      <div class="card-title" data-i18n="card.trialBalance">Trial Balance — Estado Diario</div>
      <div class="drr-days" id="drr-days"></div>
    </div>

    <!-- Alerts -->
    <div class="card">
      <div class="card-title" data-i18n="card.alertasDrr">Alertas DRR</div>
      <div class="drr-alerts" id="drr-alerts">
        <div class="empty"><p>Sin alertas.</p></div>
      </div>
    </div>

  </div><!-- /panel-drr -->

  <!-- PANEL BANCO -->
  <div id="panel-banco" class="panel">
    <div class="stats" id="banco-stats">
      <div class="sc hl c-blu"><div class="sc-lbl" data-i18n="sc.movimientos">Movimientos</div><div class="sc-val" id="bk-total">—</div><div class="sc-sub" data-i18n="sc.delExtracto">del extracto</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-i18n="sc.conciliados">Conciliados</div><div class="sc-val" id="bk-conc">—</div><div class="sc-sub" data-i18n="sc.conFactura">con factura</div></div>
      <div class="sc c-ora"><div class="sc-lbl" data-i18n="sc.pendientes">Pendientes</div><div class="sc-val" id="bk-pend">—</div><div class="sc-sub" id="bk-imp-pend">—</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-i18n="sc.diferencias">Diferencias</div><div class="sc-val" id="bk-diff">—</div><div class="sc-sub" data-i18n="sc.importeNoCuadra">importe no cuadra</div></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.alertasBanco">Alertas Bancarias</div>
      <div id="bk-alertas"><div class="empty"><p>Cargando...</p></div></div>
    </div>
    <div style="margin-top:16px">
      <a href="/conciliacion/" class="btn-run" style="text-decoration:none;display:inline-flex;font-size:13px;padding:10px 20px" data-i18n="btn.verConciliacion">🏦 Ver conciliación completa</a>
    </div>
  </div><!-- /panel-banco -->

  <!-- PANEL NOTIFICACIONES -->
  <div id="panel-notif" class="panel">
    <!-- Configuración de canales -->
    <div class="card" style="margin-bottom:20px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:10px">
        <div class="card-title" style="margin:0" data-i18n="notif.canales">Canales de notificación</div>
        <button class="btn-ref" onclick="guardarNotifConfig()" id="btn-save-notif" data-i18n="notif.guardar" style="font-size:12px">💾 Guardar configuración</button>
        <button class="btn-ref" onclick="probarNotif()" style="font-size:12px" data-i18n="btn.test">🔔 Probar</button>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px" id="notif-canales">
        <!-- channel cards inject here -->
      </div>
      <div id="notif-channel-fields" style="display:grid;gap:12px;margin-bottom:8px"></div>
      <div style="border-top:1px solid var(--s2);margin-top:14px;padding-top:16px">
        <div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;font-weight:600;margin-bottom:12px" data-i18n="notif.eventosLabel">Eventos que disparan alerta</div>
        <div id="notif-alertas" style="display:grid;grid-template-columns:1fr 1fr;gap:10px"></div>
      </div>
    </div>

    <!-- Historial -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
      <div>
        <span style="font-size:1.1rem;font-weight:700" data-i18n="notif.historial">Historial de Notificaciones</span>
        <span id="notif-count" style="font-size:.8rem;color:var(--dim);margin-left:8px"></span>
      </div>
      <button class="btn-run" id="btn-send-notif" onclick="enviarNotificaciones()" style="font-size:12px;padding:8px 16px">
        <span data-i18n="notif.enviar">🔔 Enviar notificaciones pendientes</span>
      </button>
    </div>
    <div class="card">
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Fecha</th><th>Tipo</th><th>Asunto</th><th>Destinatario</th><th data-i18n="th.estado">Estado</th></tr></thead>
          <tbody id="notif-tbody"><tr><td colspan="5" class="empty"><p>Sin notificaciones.</p></td></tr></tbody>
        </table>
      </div>
    </div>
  </div><!-- /panel-notif -->

  <!-- PANEL F&B -->

  <div id="panel-fb" class="panel">
    <!-- F&B Sub-tabs -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:10px">
      <div style="display:flex;gap:4px;background:var(--s1);border-radius:10px;padding:4px;border:1px solid var(--s2)">
        <button class="fb-sub active" onclick="fbSub('resumen',this)" data-i18n="fb.resumen">📊 Resumen</button>
        <button class="fb-sub" onclick="fbSub('inventario',this)" data-i18n="fb.inventario">📦 Inventario</button>
        <button class="fb-sub" onclick="fbSub('mermas',this)" data-i18n="fb.mermas">⚠️ Mermas</button>
        <button class="fb-sub" onclick="fbSub('recetas',this)" data-i18n="fb.recetas">📋 Recetas</button>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <label for="fb-upload-input" class="btn-ref" style="cursor:pointer;font-size:12px" data-i18n="btn.importarPos">📤 Importar ventas POS</label>
        <input type="file" id="fb-upload-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="fbUploadPOS(this)">
        <a href="/api/exportar/fb/pdf" class="btn-ref" style="text-decoration:none;font-size:12px">📄 PDF</a>
      </div>
    </div>
    <div id="fb-upload-msg" style="font-size:12px;margin-bottom:12px;min-height:16px"></div>
    <div id="fb-resumen"><div class="empty"><p>Cargando...</p></div></div>
    <div id="fb-inventario" style="display:none"></div>
    <div id="fb-mermas-panel" style="display:none"></div>
    <div id="fb-recetas" style="display:none"></div>
  </div><!-- /panel-fb -->

  <!-- PANEL AR REAL -->
  <div id="panel-ar_real" class="panel">
    <!-- Header -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;flex-wrap:wrap;gap:12px">
      <div>
        <h2 style="font-size:18px;font-weight:700;margin:0">🏢 AR Real — Grupos Corporativos</h2>
        <div style="font-size:12px;color:var(--mut);margin-top:4px">Clientes de crédito · Facturas corporativas · BEOs y grupos</div>
      </div>
      <button class="btn-run" onclick="procesarARReal()" id="btn-ar-real" style="font-size:13px">
        <div class="spin" id="spin-ar"></div>
        <span id="lbl-ar">▶ Procesar Archivos</span>
      </button>
    </div>

    <!-- KPI row -->
    <div id="ar-real-kpis" class="stats" style="grid-template-columns:repeat(4,1fr);margin-bottom:22px">
      <div class="sc hl c-ora"><div class="sc-lbl">Pend. Facturar</div><div class="sc-val" id="arp-pend">—</div><div class="sc-sub">en reservas</div></div>
      <div class="sc c-yel"><div class="sc-lbl">Facturado</div><div class="sc-val" id="arp-fact">—</div><div class="sc-sub">pendiente cobro</div></div>
      <div class="sc c-grn"><div class="sc-lbl">Cobrado</div><div class="sc-val" id="arp-cobr">—</div><div class="sc-sub">este período</div></div>
      <div class="sc c-red"><div class="sc-lbl">Saldo Total</div><div class="sc-val" id="arp-saldo">—</div><div class="sc-sub" id="arp-nclientes">— clientes</div></div>
    </div>

    <!-- Main grid: clients + reservations -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">
      <!-- Clients table -->
      <div class="card">
        <div class="card-title" data-i18n="card.clientesCredito">Clientes de Crédito</div>
        <div class="tbl-wrap" style="min-width:0">
          <table style="min-width:0;width:100%">
            <thead><tr>
              <th>Cliente</th><th style="text-align:right">Saldo</th><th style="text-align:center">Días</th><th style="text-align:center">Estado</th>
            </tr></thead>
            <tbody id="ar-clients-tbody"></tbody>
          </table>
        </div>
      </div>
      <!-- Reservations table -->
      <div class="card">
        <div class="card-title" data-i18n="card.reservasCorp">Reservas Corporativas</div>
        <div class="tbl-wrap" style="min-width:0">
          <table style="min-width:0;width:100%">
            <thead><tr>
              <th>Reserva</th><th>Entrada</th><th style="text-align:right">Total</th><th style="text-align:center">Estado</th>
            </tr></thead>
            <tbody id="ar-reservas-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Process log (collapsed by default) -->
    <div class="card" id="ar-log-card" style="display:none">
      <div class="card-title" data-i18n="card.logProcesamiento">Log de Procesamiento</div>
      <div id="ar-real-log" style="background:#060c1a;border:1px solid var(--s2);border-radius:10px;padding:14px;max-height:220px;overflow-y:auto;font-family:'JetBrains Mono',monospace;font-size:11px;line-height:1.8;color:var(--tx)"></div>
    </div>
    <div id="ar-real-status" style="display:none"></div>
  </div><!-- /panel-ar_real -->

  <!-- PANEL MULTI-HOTEL -->
  <div id="panel-calipolis" class="panel">
  <!-- Header -->
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;flex-wrap:wrap;gap:12px">
    <div>
      <h2 style="font-size:18px;font-weight:700;margin:0">🏩 Calipolis Hotels Group</h2>
      <div style="font-size:12px;color:var(--mut);margin-top:4px">Sitges · 3 propiedades · 307 habitaciones · Jun 2026</div>
    </div>
    <div style="display:flex;align-items:center;gap:10px">
      <span id="cal-mes-badge" style="font-size:11px;background:var(--s2);color:var(--acc2);border:1px solid var(--s2);border-radius:20px;padding:4px 12px">Junio 2026</span>
      <a href="/api/exportar/calipolis" class="btn-ref" style="text-decoration:none">⬇️ Descargar Excel</a>
    </div>
  </div>

  <!-- KPIs consolidados -->
  <div id="cal-kpis" style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px"></div>

  <!-- Trend row: GOP y Facturas pendientes últimos 6 meses -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:22px">
    <div class="card">
      <div class="card-title" data-i18n="card.gopEvolucion">GOP% — evolución 6 meses</div>
      <div style="height:120px;position:relative"><canvas id="cal-gop-chart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.apEvolucion">Facturas AP pendientes — evolución 6 meses</div>
      <div style="height:120px;position:relative"><canvas id="cal-ap-chart"></canvas></div>
    </div>
  </div>

  <!-- Hotel cards -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:22px" id="cal-hoteles"></div>

  <!-- Detail panel (hidden by default) -->
  <div id="cal-detail" style="display:none;margin-top:4px"></div>
  </div><!-- /panel-calipolis -->

  <div id="panel-multi_hotel" class="panel">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;flex-wrap:wrap;gap:12px">
      <h2 style="font-size:18px;font-weight:700;margin:0">🏨 Multi-Hotel Dashboard</h2>
      <div style="display:flex;align-items:center;gap:10px">
        <select id="grupo-filter" onchange="loadMultiHotel()" style="background:var(--s2);color:var(--tx);border:1px solid var(--s2);border-radius:8px;padding:7px 12px;font-size:12px;cursor:pointer;font-family:inherit">
          <option value="">Todos los grupos</option>
        </select>
        <a href="/api/exportar/multihotel" class="btn-ref" style="text-decoration:none">⬇️ Descargar Excel</a>
      </div>
    </div>

    <!-- KPIs Consolidados -->
    <div id="mh-kpis" style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px"></div>

    <!-- Status Summary -->
    <div id="mh-status" style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px"></div>

    <!-- Tabla de hoteles -->
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px;margin-bottom:20px;overflow-x:auto">
      <h3 style="font-size:14px;margin-bottom:16px;color:#8892a4">Performance por Hotel</h3>
      <table id="mh-table" style="width:100%;border-collapse:collapse;font-size:13px;min-width:900px">
        <thead>
          <tr style="border-bottom:1px solid #2e3248;color:#8892a4">
            <th style="text-align:left;padding:10px;font-weight:600">Hotel</th>
            <th style="text-align:left;padding:10px;font-weight:600">Ciudad</th>
            <th style="text-align:right;padding:10px;font-weight:600">Rooms</th>
            <th style="text-align:right;padding:10px;font-weight:600">Occ%</th>
            <th style="text-align:right;padding:10px;font-weight:600">ADR</th>
            <th style="text-align:right;padding:10px;font-weight:600">RevPAR</th>
            <th style="text-align:right;padding:10px;font-weight:600">Revenue MTD</th>
            <th style="text-align:right;padding:10px;font-weight:600">GOP%</th>
            <th style="text-align:right;padding:10px;font-weight:600">Fact. Pend.</th>
            <th style="text-align:center;padding:10px;font-weight:600">Status</th>
          </tr>
        </thead>
        <tbody id="mh-tbody"></tbody>
      </table>
    </div>

    <!-- Rankings y Alertas -->
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(400px,1fr));gap:16px">
      <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
        <h3 style="font-size:14px;margin-bottom:16px;color:#8892a4">🏆 Top Performers (Revenue MTD)</h3>
        <div id="mh-rankings"></div>
      </div>
      <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
        <h3 style="font-size:14px;margin-bottom:16px;color:#8892a4">⚠️ Alertas Activas</h3>
        <div id="mh-alertas"></div>
      </div>
    </div>

  </div><!-- /panel-multi_hotel -->

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

<!-- Chat AI — Yve Copilot -->
<button id="chat-fab" onclick="toggleChat()">
  <span style="font-size:1.3rem">💬</span>
  <span data-i18n="chat.pregunta" data-i18n="chat.pregunta">Pregunta a Yve</span>
  <div class="fab-dot"></div>
</button>

<div id="chat-panel">
  <div id="chat-header">
    <div class="chat-title">
      <span>🤖</span>
      <div>
        <h3>Yve — Copiloto Financiero</h3>
        <p>Acceso en tiempo real a los datos del hotel</p>
      </div>
    </div>
    <button id="chat-close" onclick="toggleChat()">✕</button>
  </div>
  <div id="chat-msgs"></div>
  <div id="chat-suggestions">
    <button class="sug" onclick="askSug(this)">📊 Resumen del estado financiero</button>
    <button class="sug" onclick="askSug(this)">⚠️ ¿Qué discrepancias hay abiertas?</button>
    <button class="sug" onclick="askSug(this)">💰 ¿Cuánto podemos reclamar?</button>
    <button class="sug" onclick="askSug(this)">📋 ¿Qué facturas faltan por firmar?</button>
  </div>
  <div id="chat-input-row">
    <textarea id="chat-input" rows="1" placeholder="Pregunta sobre el estado financiero del hotel…"
      onkeydown="chatKeydown(event)" oninput="autoResize(this)"></textarea>
    <button id="chat-send" onclick="sendChat()">➤</button>
  </div>
</div>

<!-- GUIDED TOUR -->
<div id="tour-overlay"></div>
<div id="tour-spotlight" style="display:none"></div>
<div id="tour-card" style="display:none">
  <div class="tour-progress-bar"><div class="tour-progress-fill" id="tour-progress" style="width:0%"></div></div>
  <div class="tour-content-wrap" id="tour-content">
    <h3 id="tour-title"></h3>
    <p id="tour-text"></p>
  </div>
  <div class="tour-footer">
    <div>
      <div class="tour-counter" id="tour-counter">1 / 10</div>
      <div class="tour-dots" id="tour-dots"></div>
    </div>
    <div class="tour-btns">
      <button class="tour-btn-skip" onclick="tourEnd()">✕ Salir</button>
      <button class="tour-btn-prev" id="tour-prev" onclick="tourPrev()">←</button>
      <button class="tour-btn-next" id="tour-next" onclick="tourNext()">Siguiente →</button>
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
  document.getElementById('status-txt').textContent = t('status.actualizando') || 'Actualizando...';
  try {
    // 1. Cargar y renderizar stats primero (independiente de facturas)
    const sr = await fetch('/api/stats');
    const stats = await sr.json();
    renderStats(stats);
    try { renderChart(stats.chart); } catch(ec) { console.warn('Chart no disponible:', ec); }

    // Alert bar
    const alertBar = document.getElementById('alert-bar');
    const parts = [];
    if (stats.discrepancias > 0)
      parts.push(stats.discrepancias + ' ' + (t('alert.discrepancias') || 'discrepancia(s) · ' + eur(stats.importe_reclamable) + ' reclamables'));
    if (stats.di_pendientes > 0)
      parts.push(stats.di_pendientes + ' ' + (t('alert.sinDI') || 'factura(s) sin certificado DI'));
    if (parts.length) {
      document.getElementById('alert-msg').textContent = parts.join(' — ');
      alertBar.classList.add('on');
    } else {
      alertBar.classList.remove('on');
    }

    // 2. Cargar facturas (aislado para que un error aquí no afecte las cards)
    let facturas = [];
    try {
      const fr = await fetch('/api/facturas');
      if (fr.ok) facturas = await fr.json();
    } catch(e2) { console.warn('Error cargando facturas:', e2); }

    renderTable(facturas);
    renderActivity(facturas);

    const hoy = new Date();
    document.getElementById('date-pill').textContent =
      hoy.toLocaleDateString('es-ES', {day:'2-digit', month:'short', year:'numeric'}) + ' · ' +
      hoy.toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'});

    document.getElementById('status-txt').textContent =
      (t('status.actualizado') || 'Actualizado') + ' · ' + (stats.total || 0) + ' ' + (t('status.facturas') || 'facturas cargadas');
  } catch(e) {
    console.error('Error en loadAll:', e);
    document.getElementById('status-txt').textContent = t('status.error') || 'Error al cargar datos';
  }
}

function renderStats(s) {
  console.log('[renderStats] datos recibidos:', JSON.stringify(s));
  document.getElementById('s-tot').textContent  = s.total ?? '—';
  document.getElementById('s-imp').textContent  = s.importe_total ? eur(s.importe_total) : '—';
  document.getElementById('s-ok').textContent   = s.correctas ?? '—';
  document.getElementById('s-disc').textContent = s.discrepancias ?? '—';
  document.getElementById('s-disc-sub').textContent = 'reclamable: ' + eur(s.importe_reclamable);
  document.getElementById('s-di').textContent   = s.di_pendientes ?? '—';
  document.getElementById('s-pend').textContent = s.sin_accion ?? '—';
  document.getElementById('s-pend-sub').textContent = (s.aprobadas ?? 0) + ' apr · ' + (s.rechazadas ?? 0) + ' rec';
  setTimeout(() => injectSparklines(AR_SPARKS), 60);
}

function renderChart(ch) {
  if (!ch || !ch.labels || !ch.labels.length) return;
  if (typeof Chart === 'undefined') { console.warn('Chart.js aún no cargado'); return; }
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
    { dot:'g', n: c.CORRECTO             || 0, txt: 'correctas sin incidencias',    key:'res.correctas' },
    { dot:'r', n: c.DISCREPANCIA         || 0, txt: 'con discrepancia de comisión',  key:'res.discrepancia' },
    { dot:'o', n: d.FALTA_CERTIFICADO_DI || 0, txt: 'sin certificado DI',            key:'res.sinDI' },
    { dot:'b', n: d.CERTIFICADO_OK       || 0, txt: 'con certificado DI OK',         key:'res.conDI' },
    { dot:'m', n: d.OTA_DESCONOCIDA      || 0, txt: 'OTA no reconocida',             key:'res.noReconocida' },
  ];
  el.innerHTML = items.map(i =>
    '<div class="act-item">' +
    '<div class="adot ' + i.dot + '"></div>' +
    '<div class="atxt"><b>' + i.n + '</b> factura' + (i.n !== 1 ? 's' : '') +
    ' <span data-i18n="' + i.key + '">' + i.txt + '</span></div>' +
    '</div>'
  ).join('');
  // Re-apply current language to freshly rendered spans
  if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
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
// ══════════════════════════════════════════════════════════════
// SPARKLINES — mini gráficos en stat cards
// ══════════════════════════════════════════════════════════════

function drawSparkline(canvasId, data, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || !data || data.length < 2) return;
  const ctx = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height, pad = 2;
  ctx.clearRect(0, 0, W, H);
  const min = Math.min(...data), max = Math.max(...data);
  const range = (max - min) || 1;
  const pts = data.map((v, i) => ({
    x: pad + (i / (data.length - 1)) * (W - 2 * pad),
    y: H - pad - ((v - min) / range) * (H - 2 * pad)
  }));
  // Fill
  ctx.beginPath();
  ctx.moveTo(pts[0].x, H - pad);
  pts.forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(pts[pts.length-1].x, H - pad);
  ctx.closePath();
  ctx.globalAlpha = 0.13;
  ctx.fillStyle = color;
  ctx.fill();
  ctx.globalAlpha = 1;
  // Line
  ctx.beginPath();
  pts.forEach((p, i) => i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y));
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.5;
  ctx.lineJoin = 'round';
  ctx.stroke();
  // Last dot
  const last = pts[pts.length - 1];
  ctx.beginPath();
  ctx.arc(last.x, last.y, 2.5, 0, Math.PI * 2);
  ctx.fillStyle = color;
  ctx.fill();
}

function makeSparkData(current, n = 7) {
  const num = parseFloat(String(current).replace(/[^0-9.]/g, '')) || 0;
  if (num === 0) return Array(n).fill(0);
  const data = [];
  for (let i = 0; i < n - 1; i++) {
    const t = i / (n - 1);
    const noise = (Math.random() - 0.45) * 0.35;
    data.push(Math.max(0, num * (0.65 + 0.25 * t + noise)));
  }
  data.push(num);
  return data;
}

function injectSparklines(cardIds) {
  cardIds.forEach(({ valId, color }) => {
    const valEl = document.getElementById(valId);
    if (!valEl) return;
    const card = valEl.closest('.sc');
    if (!card) return;
    let canvas = card.querySelector('.sc-spark');
    if (!canvas) {
      canvas = document.createElement('canvas');
      canvas.className = 'sc-spark';
      canvas.id = 'spark-' + valId;
      canvas.width = 160;
      canvas.height = 24;
      card.appendChild(canvas);
    }
    const val = valEl.textContent.trim();
    const data = makeSparkData(val);
    drawSparkline(canvas.id, data, color);
  });
}

const AR_SPARKS = [
  {valId:'s-tot',  color:'#60a5fa'},
  {valId:'s-imp',  color:'#60a5fa'},
  {valId:'s-ok',   color:'#22c55e'},
  {valId:'s-disc', color:'#ef4444'},
  {valId:'s-di',   color:'#f97316'},
  {valId:'s-pend', color:'#8b5cf6'},
];
const AP_SPARKS = [
  {valId:'ap-total',    color:'#60a5fa'},
  {valId:'ap-importe',  color:'#60a5fa'},
  {valId:'ap-matches',  color:'#22c55e'},
  {valId:'ap-disc',     color:'#ef4444'},
  {valId:'ap-sinpo',    color:'#f97316'},
  {valId:'ap-aprobadas',color:'#8b5cf6'},
];

// ══════════════════════════════════════════════════════════════
// DRR REVENUE CHART
// ══════════════════════════════════════════════════════════════
let _drrChart = null;

async function renderDRRChart() {
  try {
    const r = await fetch('/api/drr_daily_chart');
    const d = await r.json();
    if (!d || d.error || !d.dias) return;
    const card = document.getElementById('drr-chart-card');
    if (card) card.style.display = 'block';
    const canvas = document.getElementById('drr-revenue-chart');
    if (!canvas || !window.Chart) return;
    if (_drrChart) { _drrChart.destroy(); _drrChart = null; }
    const oobColors = d.dias.map((_, i) => d.oob[i] ? 'rgba(239,68,68,0.85)' : 'rgba(59,130,246,0.75)');
    _drrChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: d.dias.map((dia, i) => {
          const f = d.fechas[i] || '';
          return f ? f.slice(8) + '/' + f.slice(5,7) : dia;
        }),
        datasets: [{
          label: 'Revenue',
          data: d.revenue,
          backgroundColor: oobColors,
          borderColor: oobColors,
          borderWidth: 0,
          borderRadius: 3,
        }, {
          label: 'Trend 7d',
          data: (() => {
            const rev = d.revenue, trend = [];
            for (let i = 0; i < rev.length; i++) {
              const w = rev.slice(Math.max(0,i-3), i+4);
              trend.push(w.reduce((a,b)=>a+b,0)/w.length);
            }
            return trend;
          })(),
          type: 'line',
          borderColor: 'rgba(34,197,94,0.7)',
          backgroundColor: 'transparent',
          borderWidth: 1.5,
          pointRadius: 0,
          tension: 0.4,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title: (items) => {
                const i = items[0].dataIndex;
                return 'Día ' + d.dias[i] + (d.oob[i] ? ' ⚠ OOB' : '');
              },
              label: (item) => item.dataset.label === 'Revenue'
                ? '€' + item.raw.toLocaleString('es-ES', {maximumFractionDigits:0})
                : 'Trend: €' + Math.round(item.raw).toLocaleString('es-ES')
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(51,65,85,0.4)', drawBorder: false },
            ticks: { color: '#64748b', font: { size: 9 }, maxTicksLimit: 10 }
          },
          y: {
            grid: { color: 'rgba(51,65,85,0.4)', drawBorder: false },
            ticks: {
              color: '#64748b', font: { size: 9 },
              callback: v => '€' + (v/1000).toFixed(0) + 'K'
            }
          }
        }
      }
    });
  } catch(e) { console.warn('DRR chart error:', e); }
}

// ══════════════════════════════════════════════════════════════
// MULTI-HOTEL MAP
// ══════════════════════════════════════════════════════════════

// Coordenadas aproximadas en el viewBox 700x400 para Europa
const CITY_COORDS = {
  'Sitges':    { x: 155, y: 268 },
  'Barcelona': { x: 160, y: 262 },
  'Madrid':    { x: 110, y: 285 },
  'Paris':     { x: 200, y: 205 },
  'London':    { x: 162, y: 165 },
  'Berlin':    { x: 288, y: 162 },
  'Amsterdam': { x: 227, y: 168 },
  'Roma':      { x: 272, y: 280 },
  'Lisboa':    { x: 66,  y: 295 },
  'Lisbon':    { x: 66,  y: 295 },
};

function renderMHMap(hoteles) {
  const g = document.getElementById('mh-dots');
  if (!g) return;
  g.innerHTML = '';
  // Group hotels by city to avoid overlap
  const byCity = {};
  hoteles.forEach(h => {
    const key = h.ciudad;
    if (!byCity[key]) byCity[key] = [];
    byCity[key].push(h);
  });
  Object.entries(byCity).forEach(([ciudad, hs]) => {
    const coords = CITY_COORDS[ciudad];
    if (!coords) return;
    const { x, y } = coords;
    const status = hs.some(h=>h.status==='critical') ? 'critical'
                 : hs.some(h=>h.status==='warning')  ? 'warning' : 'ok';
    const color = status==='critical' ? '#ef4444' : status==='warning' ? '#f97316' : '#22c55e';
    const count = hs.length;
    const label = ciudad + (count > 1 ? ' (' + count + ')' : '');
    const dotEl = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    dotEl.setAttribute('class', 'hotel-dot');
    dotEl.setAttribute('transform', 'translate(' + x + ',' + y + ')');
    dotEl.innerHTML =
      '<circle r="7" fill="' + color + '" fill-opacity="0.2" stroke="' + color + '" stroke-width="1.5"/>' +
      '<circle r="3.5" fill="' + color + '"/>' +
      '<circle r="12" fill="' + color + '" fill-opacity="0" class="hit-area"/>' +
      '<text x="12" y="4" fill="#f1f5f9" font-size="9.5" font-family="Inter,sans-serif" font-weight="600">' + label + '</text>';
    dotEl.addEventListener('click', () => {
      if (hs.length === 1) openHotelDetail(hs[0].id);
      else openHotelDetail(hs[0].id);
    });
    dotEl.title = label;
    g.appendChild(dotEl);
  });
}


// ══════════════════════════════════════════════════════════════
// GUIDED TOUR — Demo interactiva paso a paso
// ══════════════════════════════════════════════════════════════

const TOUR_STEPS = [
  {
    title: '👋 Bienvenido a Yve.01',
    text: 'Yve automatiza todo el departamento financiero de un hotel. En 90 segundos verás los módulos principales. Pulsa Siguiente para empezar.',
    selector: null,
    tab: null,
  },
  {
    title: '📥 AR — Comisiones OTA',
    text: 'Booking, Expedia, Hotelbeds... Yve descarga las facturas, verifica que cada comisión coincide con la tarifa pactada, y detecta automáticamente las facturas extranjeras que necesitan certificado de doble imposición.',
    selector: '#tab-ar',
    tab: 'ar',
  },
  {
    title: '📊 Estado de un vistazo',
    text: 'Las tarjetas muestran el ciclo AR completo: facturas procesadas, importe total, cuántas son correctas, discrepancias reclamables y certificados DI pendientes. Sin abrir un solo Excel.',
    selector: '.stats',
    tab: 'ar',
  },
  {
    title: '📦 AP — Proveedores',
    text: 'Yve lee PDFs de facturas vía IA, hace el 3-way matching (factura + albarán + POS) para F&B, y genera emails de reclamación automáticos cuando hay diferencias. Lo que antes costaba 3 horas diarias.',
    selector: '#tab-ap',
    tab: 'ap',
  },
  {
    title: '📈 DRR — Revenue Diario',
    text: 'El Daily Revenue Report con los datos reales del hotel. Yve detecta los días Out of Balance automáticamente cada mañana. El Income Auditor ve el estado en segundos en vez de revisar 45 hojas Excel.',
    selector: '#tab-drr',
    tab: 'drr',
  },
  {
    title: '🍽️ F&B Cost Control',
    text: 'Food Cost % real calculado desde las ventas del POS y las recetas. Categorías con alerta si el coste supera el objetivo, ranking de platos, inventario con nivel de stock, y formulario para registrar mermas.',
    selector: '#tab-fb',
    tab: 'fb',
  },
  {
    title: '🏦 Conciliación Bancaria',
    text: 'Cruza el extracto bancario con las facturas pagadas en Oracle. Los movimientos sin justificar aparecen marcados. Lo que antes llevaba medio día, en segundos.',
    selector: '#tab-banco',
    tab: 'banco',
  },
  {
    title: '🏩 Calipolis Hotels Group',
    text: 'Dashboard del grupo: 3 hoteles, 307 habitaciones, €1.9M revenue en junio. GOP% del grupo subió de 16.4% a 22.4% en 6 meses. Facturas AP pendientes de 25 a 6. Alertas: 0.',
    selector: '#tab-calipolis',
    tab: 'calipolis',
  },
  {
    title: '🌍 Multi-Hotel',
    text: 'Para grupos más grandes: KPIs consolidados, ranking de propiedades por GOP%, y alertas activas en toda la cadena. Un solo dashboard para controlar todos los hoteles.',
    selector: '#tab-multi_hotel',
    tab: 'multi_hotel',
  },
  {
    title: '🔔 Notificaciones',
    text: 'Configura los canales: email, WhatsApp, Slack. Elige qué eventos disparan alerta: discrepancias AR, días OOB, facturas sin firmar. Yve te avisa en tiempo real sin que tengas que abrir el dashboard.',
    selector: '#tab-notif',
    tab: 'notif',
  },
  {
    title: '💬 Pregunta a Yve',
    text: 'El asistente IA tiene acceso a todos los datos del hotel en tiempo real. Pregúntale "¿Cuánto podemos reclamar a Booking?" o "¿Qué facturas faltan por firmar?" y responde con los datos actuales.',
    selector: '#chat-fab',
    tab: null,
  },
  {
    title: '🚀 ¿Listo para automatizar?',
    text: 'Esto es Yve. Setup en 15 minutos. Sin consultores. Sin contratos. Desde 400€/mes. El primer mes, mide cuántas horas ahorras — y decides si continúas.',
    selector: null,
    tab: null,
    isLast: true,
  },
];

let _tourStep = 0;
let _tourActive = false;

function tourStart() {
  _tourStep = 0;
  _tourActive = true;
  document.getElementById('tour-overlay').classList.add('active');
  document.getElementById('tour-card').style.display = 'block';
  document.getElementById('tour-spotlight').style.display = 'block';
  _buildDots();
  tourGo(_tourStep);
}

function tourEnd() {
  _tourActive = false;
  document.getElementById('tour-overlay').classList.remove('active');
  document.getElementById('tour-card').style.display = 'none';
  document.getElementById('tour-spotlight').style.display = 'none';
  document.querySelectorAll('.tour-target').forEach(el => el.classList.remove('tour-target'));
}

function tourNext() {
  if (_tourStep < TOUR_STEPS.length - 1) { _tourStep++; tourGo(_tourStep); }
  else tourEnd();
}
function tourPrev() {
  if (_tourStep > 0) { _tourStep--; tourGo(_tourStep); }
}

function _buildDots() {
  const cont = document.getElementById('tour-dots');
  cont.innerHTML = TOUR_STEPS.map((_,i) =>
    '<div class="tour-dot' + (i === 0 ? ' active' : '') + '"></div>'
  ).join('');
}

function tourGo(stepIdx) {
  const step = TOUR_STEPS[stepIdx];
  const total = TOUR_STEPS.length;

  // Switch tab if needed
  if (step.tab) {
    const tabEl = document.getElementById('tab-' + step.tab);
    if (tabEl) switchTab(step.tab, tabEl);
  }

  // Fade content out, update, fade in
  const content = document.getElementById('tour-content');
  content.classList.add('fading');
  setTimeout(() => {
    document.getElementById('tour-title').textContent = step.title;
    document.getElementById('tour-text').textContent  = step.text;
    document.getElementById('tour-counter').textContent = (stepIdx + 1) + ' / ' + total;
    // Progress bar
    const pct = (stepIdx / (total - 1)) * 100;
    document.getElementById('tour-progress').style.width = pct + '%';
    content.classList.remove('fading');
  }, 160);

  // Dots
  document.querySelectorAll('.tour-dot').forEach((d,i) => d.classList.toggle('active', i === stepIdx));
  // Prev / Next buttons
  document.getElementById('tour-prev').style.visibility = stepIdx === 0 ? 'hidden' : 'visible';
  document.getElementById('tour-next').textContent = step.isLast ? '¡Empezar! 🚀' : 'Siguiente →';
  // Remove previous target
  document.querySelectorAll('.tour-target').forEach(el => el.classList.remove('tour-target'));
  // Entrance animation on card
  const card = document.getElementById('tour-card');
  card.classList.remove('entering');
  void card.offsetWidth; // reflow
  card.classList.add('entering');

  // Position spotlight and card
  setTimeout(() => _positionTour(step), step.tab ? 420 : 60);
}

function _positionTour(step) {
  const spotlight = document.getElementById('tour-spotlight');
  const card      = document.getElementById('tour-card');
  const W = window.innerWidth, H = window.innerHeight;
  const PAD = 10;

  if (!step.selector) {
    // Centered card, no spotlight
    spotlight.style.cssText = 'display:none';
    card.style.cssText = 'display:block;position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);width:340px;z-index:9002';
    return;
  }

  const el = document.querySelector(step.selector);
  if (!el) {
    spotlight.style.cssText = 'display:none';
    card.style.cssText = 'display:block;position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);width:340px;z-index:9002';
    return;
  }

  el.classList.add('tour-target');
  el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  // Wait for scroll, then position
  setTimeout(() => {
    const r = el.getBoundingClientRect();
    // Spotlight
    spotlight.style.cssText = [
      'display:block',
      'left:' + Math.max(0, r.left - PAD) + 'px',
      'top:' + Math.max(0, r.top - PAD) + 'px',
      'width:' + (r.width + PAD*2) + 'px',
      'height:' + (r.height + PAD*2) + 'px',
      'z-index:9001',
    ].join(';');

    // Card: try below first, else above
    const cardH = 180;
    let cardTop  = r.bottom + PAD + 10;
    let cardLeft = r.left;
    if (cardTop + cardH > H - 10) cardTop = r.top - cardH - 20;
    if (cardTop < 70) cardTop = 80;
    if (cardLeft + 330 > W - 10) cardLeft = W - 340;
    if (cardLeft < 10) cardLeft = 10;
    card.style.cssText = [
      'display:block',
      'position:fixed',
      'top:' + cardTop + 'px',
      'left:' + cardLeft + 'px',
      'width:320px',
      'transform:none',
      'z-index:9002',
    ].join(';');
  }, 150);
}

// Close tour on overlay click (if clicking outside card and spotlight)
document.getElementById('tour-overlay').addEventListener('click', function(e) {
  if (e.target === this) tourEnd();
});



// ── Skeleton helpers ──────────────────────────────────────
function skelCards(n=4, extraStyle='') {
  return '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;' + extraStyle + '">' +
    Array(n).fill('<div class="skel skel-card"></div>').join('') + '</div>';
}
function skelTable(rows=5) {
  const header = '<div class="skel skel-line med" style="margin-bottom:16px"></div>';
  const r = Array(rows).fill('<div class="skel skel-line" style="margin-bottom:8px"></div>').join('');
  return header + r;
}
function skelSection() {
  return '<div style="padding:4px 0">' + skelCards(4) + '<div style="margin-top:18px">' + skelTable() + '</div></div>';
}

// ── Preload all tab data in background (3s after load) ────
setTimeout(async () => {
  const preloads = [
    '/api/stats_drr', '/api/drr_daily_chart',
    '/api/ar_real_data', '/api/calipolis/kpis',
    '/api/stats_banco',
  ];
  preloads.forEach(url => fetch(url).catch(() => {}));
}, 3000);

loadAll();
setInterval(loadAll, 60000);


// ══════════════════════════════════════════════════════════════
// MÓDULO AP — JavaScript
// ══════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════
// DEMO MODE
// ═══════════════════════════════════════════════════════════════════
let demoModeActive = false;

async function toggleDemoMode() {
  try {
    const res = await fetch('/api/demo/toggle', {method: 'POST'});
    const data = await res.json();
    demoModeActive = data.demo_mode;
    
    const btn = document.getElementById('btn-demo');
    if (demoModeActive) {
      btn.style.color = '#1db954';
      btn.style.borderColor = '#1db954';
      btn.textContent = '🎭 Demo ON';
      showNotification('✓ Demo Mode ACTIVADO - Datos ficticios cargados', 'success');
      setTimeout(() => location.reload(), 500);
    } else {
      btn.style.color = '#9333ea';
      btn.style.borderColor = '#9333ea';
      btn.textContent = '🎭 Demo';
      showNotification('✗ Demo Mode desactivado - Datos reales', 'info');
      setTimeout(() => location.reload(), 500);
    }
  } catch(e) {
    console.error('Error en demo:', e);
  }
}



// ═══════════════════════════════════════════════════════════════════
// SELECTOR DE ROL
// ═══════════════════════════════════════════════════════════════════
let rolActual = 'admin';
const rolLabels = {
  'admin': '🔑 Admin',
  'financial_controller': '💰 Controller',
  'income_auditor': '📊 Auditor',
  'fb_manager': '🍽️ F&B',
  'jefe_otras': '🛠️ Servicios'
};

function toggleMenu(id) {
  const m = document.getElementById(id);
  const wasOpen = m.classList.contains('open');
  document.querySelectorAll('.menu.open').forEach(el => el.classList.remove('open'));
  if (!wasOpen) m.classList.add('open');
}

// Cerrar menús al hacer click fuera
document.addEventListener('click', (e) => {
  if (!e.target.closest('.dropdown')) {
    document.querySelectorAll('.menu.open').forEach(el => el.classList.remove('open'));
  }
});

function toggleRolMenu() {
  const menu = document.getElementById('rol-menu');
  menu.style.display = menu.style.display === 'block' ? 'none' : 'block';
}

async function cambiarRol(newRole) {
  try {
    const res = await fetch(`/api/rol/cambiar/${newRole}`, {method: 'POST'});
    const data = await res.json();
    
    rolActual = newRole;
    const btn = document.getElementById('rol-btn');
    btn.textContent = '👤 ' + rolLabels[newRole];
    
    document.getElementById('rol-menu').style.display = 'none';
    showNotification(`✓ Rol cambiado a ${rolLabels[newRole]}`, 'info');
    
    // Aquí irían los cambios visuales del dashboard según rol
    // Por ahora solo cambiamos el label
  } catch(e) {
    console.error('Error cambiando rol:', e);
  }
}

// Toggle al clickear el botón
document.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('rol-btn');
  if (btn) {
    btn.addEventListener('click', toggleRolMenu);
  }
});


function switchTab(tab, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  var panel = document.getElementById('panel-' + tab);
  if (panel) panel.classList.add('active');
  if (tab === 'fb') loadFBTab();
  if (tab === 'ar_real') cargarARRealData();
  if (tab === 'drr') loadDRR();
  if (tab === 'banco') loadBanco();
  if (tab === 'notif') loadNotifConfig();
  if (tab === 'calipolis') loadCalipolis();
  if (tab === 'multi_hotel') loadMultiHotel();
}
// ══ F&B COST CONTROL ══════════════════════════════════════════════════
let _fbLoaded = {resumen:false, inventario:false, mermas:false, recetas:false};
let _fbActive = 'resumen';

async function fbUploadPOS(input) {
  const file = input.files[0];
  if (!file) return;
  const msg = document.getElementById('fb-upload-msg');
  msg.style.color = 'var(--mut)';
  msg.textContent = 'Subiendo ' + file.name + '...';
  const form = new FormData();
  form.append('file', file);
  try {
    const r = await fetch('/fb/api/upload_ventas', {method: 'POST', body: form});
    const d = await r.json();
    if (d.ok) {
      msg.style.color = 'var(--grn)';
      msg.textContent = '✓ ' + d.filas_importadas + ' ventas importadas de ' + file.name + ' — Total acumulado: ' + d.total_filas + ' registros';
      // Reload resumen
      _fbLoaded.resumen = false;
      if (_fbActive === 'resumen') loadFBResumen(); else fbSub('resumen', document.querySelector('.fb-sub'));
    } else {
      msg.style.color = 'var(--red)';
      msg.textContent = '✗ ' + (d.error || 'Error al importar');
    }
  } catch(e) {
    msg.style.color = 'var(--red)';
    msg.textContent = '✗ Error de conexion: ' + e.message;
  }
  input.value = '';
}

function fbSub(sub, el) {
  _fbActive = sub;
  document.querySelectorAll('.fb-sub').forEach(b => b.classList.remove('active'));
  if (el) el.classList.add('active');
  const panels = {resumen:'fb-resumen', inventario:'fb-inventario', mermas:'fb-mermas-panel', recetas:'fb-recetas'};
  Object.values(panels).forEach(id => { const d = document.getElementById(id); if (d) d.style.display = 'none'; });
  const active = document.getElementById(panels[sub]);
  if (active) active.style.display = 'block';
  if (!_fbLoaded[sub]) {
    if (sub === 'resumen')    loadFBResumen();
    if (sub === 'inventario') loadFBInventario();
    if (sub === 'mermas')     loadFBMermas();
    if (sub === 'recetas')    loadFBRecetas();
  }
}

async function loadFBTab() {
  if (!_fbLoaded.resumen) loadFBResumen();
}

function runFB() {
  const es = new EventSource('/fb/api/ejecutar');
  const subs = ['resumen','inventario','mermas','recetas'];
  subs.forEach(s => { _fbLoaded[s] = false; });
  es.onmessage = (ev) => {
    if (ev.data === 'FB_COMPLETO') {
      es.close();
      loadFBResumen();
      if (_fbActive !== 'resumen') {
        if (_fbActive === 'inventario') loadFBInventario();
        if (_fbActive === 'mermas')     loadFBMermas();
        if (_fbActive === 'recetas')    loadFBRecetas();
      }
    }
  };
}

async function loadFBResumen() {
  _fbLoaded.resumen = true;
  const cont = document.getElementById('fb-resumen');
  if (!cont) return;
  cont.innerHTML = skelCards(4, 'grid-template-columns:repeat(4,1fr)') +
    '<div style="margin-top:18px">' + skelTable(4) + '</div>';
  try {
    const res = await fetch('/fb/api/resultados');
    const data = await res.json();
    if (!data.ok) {
      cont.innerHTML = '<div class="empty"><p>Sin datos F&B.</p><button class="btn-run" onclick="runFB()" style="margin-top:16px;font-size:13px">▶ Ejecutar Análisis</button></div>';
      return;
    }
    const r = data.resumen;
    const fcColor = r.alerta ? 'var(--red)' : (r.fc_real_pct <= r.fc_teorico_pct ? 'var(--grn)' : 'var(--ora)');
    const fcDiff  = (r.fc_real_pct - r.fc_teorico_pct).toFixed(2);
    const fcSign  = fcDiff > 0 ? '+' : '';

    let html = '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:10px">';
    html += '<div><h2 style="font-size:17px;font-weight:700;margin:0">F&B Cost Control</h2>';
    html += '<div style="font-size:12px;color:var(--mut);margin-top:3px">' + (t('fb.datosReales')||'Datos calculados desde ventas reales') + ' · ' + data.ventas_diarias.fechas.length + ' ' + (t('fb.dias')||'días</div></div>';
    html += '<button class="btn-ref" onclick="runFB()" style="font-size:12px" data-i18n="btn.recalcular">↺ Recalcular</button></div>';

    // KPIs
    html += '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px">';
    html += _fbKpi(t('fb.ventasFb')||'Ventas F&B', '€' + Math.round(r.total_ventas).toLocaleString('es-ES'), t('fb.periodoCompleto')||'período completo', 'var(--acc2)');
    html += _fbKpi(t('fb.fcTeorico')||'FC Teórico', r.fc_teorico_pct + '%', t('fb.objetivoCalc')||'objetivo calculado', 'var(--grn)');
    html += _fbKpi(t('fb.fcReal')||'FC Real', r.fc_real_pct + '%', fcSign + fcDiff + ' ' + (t('fb.vsObjetivo')||'pp vs objetivo'), fcColor);
    html += _fbKpi(t('fb.mermasLabel')||'Mermas', '€' + r.coste_mermas.toLocaleString('es-ES'), r.alerta ? t('fb.revisar')||'⚠ Revisar' : t('fb.bajoControl')||'bajo control', r.alerta ? 'var(--red)' : 'var(--mut)');
    html += '</div>';

    // FC% gauge
    const maxG = Math.max(r.fc_teorico_pct, r.fc_real_pct) * 1.35;
    html += '<div class="card" style="margin-bottom:16px"><div class="card-title" style="margin-bottom:14px" data-i18n="fb.gaugeTitle">Food Cost % — Teórico vs Real</div>';
    html += _fcBar(t('fb.gaugeTeorico')||'Teórico', r.fc_teorico_pct, maxG, 'var(--grn)');
    html += _fcBar(t('fb.gaugeReal')||'Real',    r.fc_real_pct,    maxG, fcColor);
    html += '</div>';

    // Ventas diarias chart
    html += '<div class="card" style="margin-bottom:16px"><div class="card-title" data-i18n="card.ventasDiarias">Ventas diarias F&B</div>';
    html += '<div style="height:160px;position:relative"><canvas id="fb-ventas-chart"></canvas></div></div>';

    // Categories + ranking
    html += '<div style="display:grid;grid-template-columns:1.2fr 0.8fr;gap:16px">';
    html += '<div class="card"><div class="card-title" data-i18n="card.fcCategoria">Food Cost por Categoría</div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>' + (t(\'fb.thCategoria\')||'Categoría') + '</th><th style="text-align:right">' + (t(\'fb.thVentas\')||'Ventas') + '</th><th style="text-align:right">FC%</th><th style="text-align:center">' + (t(\'fb.thEstado\')||'Estado') + '</th>';
    html += '</tr></thead><tbody>';
    data.categorias.forEach(c => {
      const cC = c.alerta ? 'var(--red)' : 'var(--grn)';
      const badge = c.alerta ? '<span class="badge b-disc">Alerta</span>' : '<span class="badge b-ok">OK</span>';
      html += '<tr><td style="font-weight:600">' + c.nombre + '</td>' +
        '<td style="text-align:right">€' + Math.round(c.total_ventas).toLocaleString('es-ES') + '</td>' +
        '<td style="text-align:right;color:' + cC + ';font-weight:700">' + c.fc_real_pct + '%</td>' +
        '<td style="text-align:center">' + badge + '</td></tr>';
    });
    html += '</tbody></table></div></div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.topPlatos">Top Platos</div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr><th>' + (t('fb.thPlato')||'Plato') + '</th><th style="text-align:right">€</th><th style="text-align:right">FC%</th></tr></thead><tbody>';
    data.ranking_top.forEach((p, i) => {
      const pC = p.fc_real_pct > 30 ? 'var(--ora)' : 'var(--grn)';
      html += '<tr><td><span style="color:var(--dim);font-size:10px;margin-right:5px">#' + (i+1) + '</span>' + p.nombre + '</td>' +
        '<td style="text-align:right">€' + Math.round(p.total_ventas).toLocaleString('es-ES') + '</td>' +
        '<td style="text-align:right;font-weight:700;color:' + pC + '">' + p.fc_real_pct + '%</td></tr>';
    });
    html += '</tbody></table></div></div></div>';

    cont.innerHTML = html;

    // Ventas chart
    setTimeout(() => {
      const cv = document.getElementById('fb-ventas-chart');
      if (cv && window.Chart && data.ventas_diarias.fechas.length) {
        new Chart(cv, {
          type: 'bar',
          data: {
            labels: data.ventas_diarias.fechas.map(f => f.slice(8)), // day
            datasets: [{
              data: data.ventas_diarias.totales,
              backgroundColor: 'rgba(59,130,246,.6)',
              borderRadius: 3,
            }]
          },
          options: {
            responsive:true, maintainAspectRatio:false,
            plugins:{legend:{display:false}},
            scales:{
              x:{grid:{color:'rgba(51,65,85,.3)'},ticks:{color:'#64748b',font:{size:9},maxTicksLimit:10}},
              y:{grid:{color:'rgba(51,65,85,.3)'},ticks:{color:'#64748b',font:{size:9},callback:v=>'€'+(v/1000).toFixed(0)+'K'}}
            }
          }
        });
      }
    }, 100);
  } catch(e) { cont.innerHTML = '<div class="empty"><p>Error: ' + e.message + '</p></div>'; }
}

async function loadFBInventario() {
  _fbLoaded.inventario = true;
  const cont = document.getElementById('fb-inventario');
  if (!cont) return;
  cont.innerHTML = skelCards(4, 'grid-template-columns:repeat(4,1fr)') + '<div style="margin-top:18px">' + skelTable(6) + '</div>';
  try {
    const res = await fetch('/fb/api/inventario');
    const data = await res.json();
    if (!data.ok) { cont.innerHTML = '<div class="empty"><p>Error inventario</p></div>'; return; }

    const alertas = data.items.filter(i => i.alerta);
    let html = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px">';
    html += _fbKpi(t('fb.itemsStock')||'Items en Stock', data.items.length, t('fb.ingredientes')||'ingredientes', 'var(--acc2)');
    html += _fbKpi(t('fb.valorInv')||'Valor Inventario', '€' + data.valor_total.toLocaleString('es-ES'), t('fb.valorActual')||'valoración actual', 'var(--grn)');
    html += _fbKpi(t('fb.alertasStock')||'Alertas Stock Bajo', alertas.length, alertas.length > 0 ? 'revisar urgente' : t('fb.todoOk')||'todo OK', alertas.length > 0 ? 'var(--red)' : 'var(--grn)');
    html += '</div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.stockIngredientes">Stock de Ingredientes</div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>' + (t(\'fb.thIngrediente\')||'Ingrediente') + '</th><th>' + (t(\'fb.thCategoria\')||'Categoría') + '</th><th>' + (t(\'th.proveedor\')||'Proveedor') + '</th>';
    html += '<th style="text-align:right">' + (t('fb.thActual')||'Actual') + '</th><th style="text-align:right">€/u</th>';
    html += '<th style="text-align:right;width:130px">Stock</th><th style="text-align:center">Estado</th>';
    html += '</tr></thead><tbody>';
    data.items.forEach(item => {
      const fillColor = item.critico ? 'var(--red)' : item.alerta ? 'var(--ora)' : 'var(--grn)';
      const badge = item.critico ? '<span class="badge b-disc">Crítico</span>' :
                    item.alerta  ? '<span class="badge b-unk">Bajo</span>' :
                                   '<span class="badge b-ok">OK</span>';
      html += '<tr>' +
        '<td style="font-weight:600">' + item.ingrediente + '</td>' +
        '<td style="color:var(--mut)">' + item.categoria + '</td>' +
        '<td style="color:var(--dim);font-size:12px">' + item.proveedor + '</td>' +
        '<td style="text-align:right;font-weight:700">' + item.stock_actual + ' ' + item.unidad + '</td>' +
        '<td style="text-align:right;color:var(--mut)">€' + item.coste_unitario + '</td>' +
        '<td style="text-align:right">' +
          '<div class="stock-bar"><div class="stock-fill" style="width:' + Math.min(item.pct_restante,100) + '%;background:' + fillColor + '"></div></div>' +
          '<div style="font-size:10px;color:var(--dim);text-align:right">' + item.pct_restante + '%</div></td>' +
        '<td style="text-align:center">' + badge + '</td>' +
        '</tr>';
    });
    html += '</tbody></table></div></div>';
    cont.innerHTML = html;
  } catch(e) { cont.innerHTML = '<div class="empty"><p>Error: ' + e.message + '</p></div>'; }
  if (_i18nLang && _i18nLang !== "es") applyI18n(_i18nData);
}

async function loadFBMermas() {
  _fbLoaded.mermas = true;
  const cont = document.getElementById('fb-mermas-panel');
  if (!cont) return;
  cont.innerHTML = skelTable(6);
  try {
    const res = await fetch('/fb/api/mermas');
    const data = await res.json();
    if (!data.ok) { cont.innerHTML = '<div class="empty"><p>Error mermas</p></div>'; return; }

    let html = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">';
    // Mermas por causa
    html += '<div class="card"><div class="card-title" data-i18n="card.mermasCausa">Mermas por Causa</div><div style="margin-top:8px">';
    const causas = Object.entries(data.por_causa).sort((a,b) => b[1]-a[1]);
    const maxCausa = causas[0]?.[1] || 1;
    causas.forEach(([causa, coste]) => {
      const pct = Math.round(coste/maxCausa*100);
      html += '<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px">' +
        '<span>' + causa + '</span><span style="font-weight:700;color:var(--ora)">€' + coste.toFixed(2) + '</span></div>' +
        '<div class="stock-bar"><div class="stock-fill" style="width:' + pct + '%;background:var(--ora)"></div></div></div>';
    });
    html += '</div></div>';

    // Formulario registrar merma
    html += '<div class="card"><div class="card-title" data-i18n="card.registrarMerma">Registrar Merma</div>';
    html += '<div style="display:grid;gap:10px;margin-top:8px">';
    html += _fbField('mb-ing', 'Ingrediente', 'text', 'Ej: Gambas');
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">';
    html += _fbField('mb-cant', 'Cantidad', 'number', '0.5');
    html += _fbField('mb-unidad', 'Unidad', 'text', 'kg / l');
    html += '</div>';
    html += _fbField('mb-coste', 'Coste unitario (€/u)', 'number', '12.00');
    html += '<div><label style="font-size:11px;color:var(--mut);display:block;margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px">Causa</label>';
    html += '<select id="mb-causa" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:9px;padding:10px 12px;font-family:inherit;font-size:14px">';
    ['Caducidad','Deterioro','Merma cocina','No vendido','Rotura','Otro'].forEach(c => {
      html += '<option value="' + c + '">' + c + '</option>';
    });
    html += '</select></div>';
    html += '<button class="btn-run" onclick="fbRegistrarMerma()" style="font-size:13px;margin-top:4px">💾 Registrar Merma</button>';
    html += '<div id="mb-msg" style="font-size:12px;text-align:center;min-height:18px"></div>';
    html += '</div></div></div>';

    // Historial
    html += '<div class="card"><div class="card-title">Historial de Mermas · Total: <span style="color:var(--ora)">€' + data.total.toFixed(2) + '</span></div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>Fecha</th><th>Ingrediente</th><th>Categoría</th><th style="text-align:right">Cantidad</th>';
    html += '<th>Causa</th><th style="text-align:right">Coste</th></tr></thead><tbody>';
    [...data.mermas].reverse().forEach(m => {
      html += '<tr>' +
        '<td style="color:var(--dim);font-size:12px">' + m.fecha + '</td>' +
        '<td style="font-weight:600">' + m.ingrediente + '</td>' +
        '<td style="color:var(--mut)">' + m.categoria + '</td>' +
        '<td style="text-align:right">' + m.cantidad + ' ' + m.unidad + '</td>' +
        '<td><span class="badge b-unk">' + m.causa + '</span></td>' +
        '<td style="text-align:right;font-weight:700;color:var(--red)">€' + m.coste.toFixed(2) + '</td>' +
        '</tr>';
    });
    html += '</tbody></table></div></div>';
    cont.innerHTML = html;
  } catch(e) { cont.innerHTML = '<div class="empty"><p>Error: ' + e.message + '</p></div>'; }
}

function _fbField(id, label, type, ph) {
  return '<div><label style="font-size:11px;color:var(--mut);display:block;margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px">' + label + '</label>' +
    '<input id="' + id + '" type="' + type + '" placeholder="' + ph + '" ' +
    'style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:9px;padding:10px 12px;font-family:inherit;font-size:14px;outline:none"></div>';
}

async function fbRegistrarMerma() {
  const g = id => document.getElementById(id)?.value?.trim();
  const msg = document.getElementById('mb-msg');
  const data = { ingrediente:g('mb-ing'), cantidad:g('mb-cant'), unidad:g('mb-unidad'), coste_unitario:g('mb-coste'), causa:g('mb-causa') };
  if (!data.ingrediente || !data.cantidad) { if(msg) { msg.style.color='var(--red)'; msg.textContent='Rellena ingrediente y cantidad.'; } return; }
  const r = await fetch('/fb/api/registrar_merma', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
  const res = await r.json();
  if (res.ok) {
    if(msg) { msg.style.color='var(--grn)'; msg.textContent='✓ Merma registrada (€' + res.coste.toFixed(2) + ')'; }
    _fbLoaded.mermas = false; // force reload
    setTimeout(() => { loadFBMermas(); }, 800);
  } else {
    if(msg) { msg.style.color='var(--red)'; msg.textContent='Error: ' + res.error; }
  }
}

async function loadFBRecetas() {
  _fbLoaded.recetas = true;
  const cont = document.getElementById('fb-recetas');
  if (!cont) return;
  cont.innerHTML = skelTable(8);
  try {
    const res = await fetch('/fb/api/recetas');
    const data = await res.json();
    if (!data.ok) { cont.innerHTML = '<div class="empty"><p>Error recetas</p></div>'; return; }

    const avg = data.recetas.reduce((a,r)=>a+r.fc_pct,0)/data.recetas.length;
    let html = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px">';
    html += _fbKpi('Recetas activas', data.recetas.length, 'en carta', 'var(--acc2)');
    html += _fbKpi('FC% promedio', avg.toFixed(1) + '%', 'media ponderada', avg < 30 ? 'var(--grn)' : 'var(--ora)');
    html += _fbKpi('Alertas FC alto', data.recetas.filter(r=>r.alerta).length, '>35% FC', 'var(--red)');
    html += '</div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.fichaRecetas">Ficha de Recetas con Coste Teórico</div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>Receta</th><th>Categoría</th><th style="text-align:right">PVP</th>';
    html += '<th style="text-align:right">Coste</th><th style="text-align:right">FC%</th>';
    html += '<th style="text-align:right">Margen</th><th style="text-align:center">Estado</th>';
    html += '</tr></thead><tbody>';
    data.recetas.sort((a,b)=>b.fc_pct-a.fc_pct).forEach(r => {
      const fcColor = r.alerta ? 'var(--red)' : r.fc_pct < 25 ? 'var(--grn)' : 'var(--ora)';
      const margen = (r.precio_venta - r.coste_teorico).toFixed(2);
      const badge = r.alerta ? '<span class="badge b-disc">Alto FC</span>' : '<span class="badge b-ok">OK</span>';
      html += '<tr>' +
        '<td style="font-weight:600">' + r.nombre + '</td>' +
        '<td style="color:var(--mut)">' + r.categoria + '</td>' +
        '<td style="text-align:right;font-weight:700">€' + r.precio_venta + '</td>' +
        '<td style="text-align:right;color:var(--mut)">€' + r.coste_teorico + '</td>' +
        '<td style="text-align:right;font-weight:800;color:' + fcColor + '">' + r.fc_pct + '%</td>' +
        '<td style="text-align:right;color:var(--grn)">€' + margen + '</td>' +
        '<td style="text-align:center">' + badge + '</td></tr>';
    });
    html += '</tbody></table></div></div>';
    cont.innerHTML = html;
  } catch(e) { cont.innerHTML = '<div class="empty"><p>Error: ' + e.message + '</p></div>'; }
}

function _fbKpi(lbl, val, sub, color) {
  return '<div style="background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px 16px">' +
    '<div style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px;font-weight:600">' + lbl + '</div>' +
    '<div style="font-size:24px;font-weight:800;color:' + color + ';line-height:1;letter-spacing:-.5px">' + val + '</div>' +
    '<div style="font-size:11px;color:var(--dim);margin-top:7px">' + sub + '</div></div>';
}

function _fcBar(label, pct, maxG, color) {
  const w = Math.round(pct/maxG*100);
  return '<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">' +
    '<div style="font-size:12px;color:var(--mut);width:56px;flex-shrink:0">' + label + '</div>' +
    '<div style="flex:1;background:var(--s2);border-radius:6px;height:16px;overflow:hidden">' +
      '<div style="height:100%;background:' + color + ';border-radius:6px;width:' + w + '%;transition:width .8s"></div></div>' +
    '<div style="font-size:15px;font-weight:700;color:' + color + ';width:42px;text-align:right">' + pct + '%</div></div>';
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

    // Safe element access
    const el = (id) => document.getElementById(id);
    if (el('ap-total')) el('ap-total').textContent = stats.total ?? '—';
    if (el('ap-importe')) el('ap-importe').textContent = fmtEurAP(stats.importe);
    if (el('ap-matches')) el('ap-matches').textContent = stats.matches ?? '—';
    if (el('ap-disc')) el('ap-disc').textContent = stats.discrepancias ?? '—';
    if (el('ap-sinpo')) el('ap-sinpo').textContent = stats.sin_po ?? '—';
    if (el('ap-alertas')) el('ap-alertas').textContent = stats.alertas_consumo ?? '—';
    if (el('ap-manual')) el('ap-manual').textContent = stats.manual ?? '—';
    if (el('ap-aprobadas')) el('ap-aprobadas').textContent = stats.aprobadas ?? '—';
    setTimeout(() => injectSparklines(AP_SPARKS), 60);

    const tbody = el('ap-tbody');
    if (tbody) tbody.innerHTML = '';
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


// ══════════════════════════════════════════════════════════════
// MÓDULO ORACLE — JavaScript
// ══════════════════════════════════════════════════════════════

function procesarOracle() {
  const btn   = document.getElementById('btnOracle');
  const log   = document.getElementById('log');
  const spin  = document.getElementById('spinner');
  const lbl   = document.getElementById('btnLabel');
  const icon  = document.getElementById('modalIcon');
  const title = document.getElementById('modalTitle');
  const btnCl = document.getElementById('btnClose');

  btn.disabled = true;
  spin.style.display = 'block';
  lbl.textContent = 'Contabilizando...';
  log.innerHTML = '';
  btnCl.disabled = true;
  icon.textContent = '🔮';
  title.textContent = 'Oracle Pipeline — Contabilizando...';
  document.getElementById('overlay').classList.add('on');

  const src = new EventSource('/api/procesar_oracle');

  src.onmessage = ev => {
    const txt = ev.data;
    const p = document.createElement('p');
    if      (txt === 'PIPELINE_COMPLETO')    p.className = 'l-ok';
    else if (txt === 'PIPELINE_CON_ERRORES') p.className = 'l-err';
    else if (txt.startsWith('[') && txt.includes('] ✅')) p.className = 'l-ok';
    else if (txt.startsWith('[') && txt.includes('] ❌')) p.className = 'l-err';
    else if (txt.startsWith('[') && txt.includes('] ⚠'))  p.className = 'l-err';
    else if (txt.includes('SIMULACION') || txt.includes('simulaci')) p.className = 'l-info';
    else if (txt.startsWith('──') || txt.startsWith('==')) p.className = 'l-dim';
    else if (txt.includes('✓') || txt.includes('✅') || txt.startsWith('OK ')) p.className = 'l-ok';
    else if (txt.includes('✗') || txt.includes('❌') || txt.startsWith('ERROR')) p.className = 'l-err';
    else p.className = 'l-dim';
    p.textContent = txt;
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = ok ? 'Oracle: contabilización completada' : 'Oracle: pipeline con errores';
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = '🔮 Contabilizar en Oracle';
      btnCl.disabled = false;
      setTimeout(loadAP, 800);
    }
  };

  src.onerror = () => {
    src.close();
    const p = document.createElement('p');
    p.className = 'l-err';
    p.textContent = 'ERROR: conexión con servidor perdida';
    log.appendChild(p);
    btn.disabled = false;
    spin.style.display = 'none';
    lbl.textContent = '🔮 Contabilizar en Oracle';
    btnCl.disabled = false;
  };
}


// ══════════════════════════════════════════════════════════════
// CHAT AI — Yve Copilot
// ══════════════════════════════════════════════════════════════

let chatHistory  = [];
let chatOpen     = false;
let chatGreeted  = false;

function toggleChat() {
  chatOpen = !chatOpen;
  const panel = document.getElementById('chat-panel');
  const fab   = document.getElementById('chat-fab');
  panel.classList.toggle('open', chatOpen);
  fab.style.display = chatOpen ? 'none' : 'flex';
  if (chatOpen && !chatGreeted) {
    chatGreeted = true;
    addMsg('bot', '¡Hola! Soy Yve, tu copiloto financiero 👋\\nTengo acceso en tiempo real a todos los datos del hotel. ¿En qué puedo ayudarte?');
  }
  if (chatOpen) setTimeout(() => document.getElementById('chat-input').focus(), 300);
}

function renderMarkdown(text) {
  // Escape HTML first
  let html = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  // Bold **text**
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  // Bullet lists (lines starting with - or •)
  html = html.replace(/^[\-•]\s+(.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\n?)+/g, m => '<ul style="margin:6px 0;padding-left:18px">' + m + '</ul>');
  // Line breaks
  html = html.replace(/\n/g, '<br>');
  // Clean up <br> inside lists
  html = html.replace(/<\/li><br>/g, '</li>').replace(/<ul([^>]*)><br>/g, '<ul$1>');
  return html;
}

function addMsg(role, text, isMarkdown) {
  const msgs = document.getElementById('chat-msgs');
  const div  = document.createElement('div');
  div.className = `msg ${role}`;
  if (isMarkdown && role === 'bot') {
    div.innerHTML = renderMarkdown(text);
  } else {
    div.textContent = text;
  }
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
  return div;
}

async function sendChat() {
  const input = document.getElementById('chat-input');
  const send  = document.getElementById('chat-send');
  const text  = input.value.trim();
  if (!text || send.disabled) return;

  addMsg('user', text);
  chatHistory.push({ role: 'user', content: text });
  input.value = '';
  input.style.height = 'auto';
  send.disabled = true;

  const thinkDiv = addMsg('bot', '');
  thinkDiv.classList.add('thinking');
  thinkDiv.innerHTML = '<span class="typing"><span></span><span></span><span></span></span>';

  try {
    const resp = await fetch('/api/chat', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ messages: chatHistory }),
    });
    const data = await resp.json();
    const reply = data.reply || '⚠️ Sin respuesta del servidor.';

    thinkDiv.innerHTML = renderMarkdown(reply);
    thinkDiv.classList.remove('thinking');
    chatHistory.push({ role: 'assistant', content: reply });

    // Mantener historial manejable (últimas 20 interacciones)
    if (chatHistory.length > 40) chatHistory = chatHistory.slice(-40);
  } catch(e) {
    thinkDiv.textContent = '⚠️ Error de conexión con el servidor.';
    thinkDiv.classList.remove('thinking');
  } finally {
    send.disabled = false;
    document.getElementById('chat-input').focus();
  }
}

function askSug(btn) {
  const input = document.getElementById('chat-input');
  input.value = btn.textContent;
  sendChat();
}

function chatKeydown(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendChat();
  }
}

function autoResize(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight, 120) + 'px';
}

// ══════════════════════════════════════════════════════════════
// MÓDULO DRR — JavaScript
// ══════════════════════════════════════════════════════════════

async function uploadDRR(input) {
  const file = input.files[0];
  if (!file) return;
  const status = document.getElementById('drr-status');
  status.textContent = 'Procesando ' + file.name + '...';

  const form = new FormData();
  form.append('file', file);

  try {
    const resp = await fetch('/api/upload_drr', { method: 'POST', body: form });
    const data = await resp.json();
    if (data.ok) {
      status.textContent = '✓ ' + file.name + ' procesado';
      renderDRR(data.stats);
    } else {
      status.textContent = '✗ Error: ' + (data.error || 'desconocido');
    }
  } catch(e) {
    status.textContent = '✗ Error de conexión';
  }
  input.value = '';
}

function renderDRR(s) {
  if (!s || s.error) {
    document.getElementById('drr-metrics').innerHTML = '<div class="empty"><p>Error: ' + (s ? s.error : 'sin datos') + '</p></div>';
    return;
  }

  // KPI cards
  const SHOW = [
    {key:'Total Revenue', label:'Total Revenue', color:'var(--acc2)'},
    {key:'Occupancy %', label:'Occupancy %', color:'var(--grn)'},
    {key:'ADR', label:'ADR', color:'var(--tx)'},
    {key:'Revenue PAR', label:'RevPAR', color:'var(--tx)'},
    {key:'GOP', label:'GOP', color:'var(--ora)'},
    {key:'GOP %', label:'GOP %', color:'var(--pur)'},
  ];
  const metricsEl = document.getElementById('drr-metrics');
  metricsEl.innerHTML = SHOW.map(m => {
    const d = s.metricas[m.key] || {};
    return '<div class="drr-mc">'
      + '<div class="mc-name">' + m.label + '</div>'
      + '<div class="mc-row"><span class="mc-k">Today</span><span class="mc-v" style="color:' + m.color + '">' + (d.today || 'N/D') + '</span></div>'
      + '<div class="mc-row"><span class="mc-k">MTD</span><span class="mc-v">' + (d.mtd || 'N/D') + '</span></div>'
      + '<div class="mc-row"><span class="mc-k">Forecast</span><span class="mc-v">' + (d.forecast || 'N/D') + '</span></div>'
      + '</div>';
  }).join('');

  // Days grid
  const daysEl = document.getElementById('drr-days');
  const diasMap = {};
  (s.dias || []).forEach(d => { diasMap[d.dia] = d; });
  let daysHtml = '';
  for (let i = 1; i <= 31; i++) {
    const d = diasMap[i];
    if (d) {
      const cls = d.oob ? 'oob' : 'ok';
      const label = d.oob ? '⚠ OOB' : '✓ OK';
      daysHtml += '<div class="drr-day ' + cls + '"><div class="day-n">' + i + '</div>' + label + '</div>';
    } else {
      daysHtml += '<div class="drr-day empty"><div class="day-n">' + i + '</div>—</div>';
    }
  }
  daysEl.innerHTML = daysHtml;

  // Alerts
  const alertsEl = document.getElementById('drr-alerts');
  if (s.alertas && s.alertas.length) {
    alertsEl.innerHTML = s.alertas.map(a =>
      '<div class="da-item"><div class="da-dot"></div><div class="da-txt">' + a + '</div></div>'
    ).join('');
  } else {
    alertsEl.innerHTML = '<div class="empty"><p>Sin alertas — todo en balance.</p></div>';
  }

  // Update status bar
  document.getElementById('drr-status').textContent =
    s.archivo + ' · ' + s.total_dias + ' días · ' + s.dias_oob + ' OOB';
  // Render revenue chart
  renderDRRChart();
}

async function loadDRR() {
  try {
    const r = await fetch('/api/stats_drr');
    const data = await r.json();
    if (data) renderDRR(data);
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════════
// NOTIFICACIONES — JavaScript
// ══════════════════════════════════════════════════════════════

const NOTIF_CHANNELS = [
  {key:'email',    icon:'📧', name:'Email'},
  {key:'whatsapp', icon:'💬', name:'WhatsApp'},
  {key:'telegram', icon:'✈️', name:'Telegram'},
  {key:'slack',    icon:'💼', name:'Slack'},
  {key:'push',     icon:'🔔', name:'Push'},
];
const NOTIF_ALERTAS = [
  {key:'ar_discrepancia',         get label(){return t('notif.evAr')    ||'Discrepancia en comisiones OTA (AR)'}},
  {key:'ar_falta_di',             get label(){return t('notif.evDi')    ||'Falta certificado de doble imposición'}},
  {key:'ap_discrepancia',         get label(){return t('notif.evAp')    ||'Discrepancia en facturas proveedor (AP)'}},
  {key:'drr_oob',                 get label(){return t('notif.evDrr')   ||'DRR: días Out of Balance'}},
  {key:'banco_sin_conciliar',     get label(){return t('notif.evBanco') ||'Movimientos bancarios sin conciliar'}},
  {key:'factura_pendiente_firma', get label(){return t('notif.evFirma') ||'Facturas pendientes de firma'}},
];
let _notifConfig = null;

async function loadNotifConfig() {
  try {
    const ch = document.getElementById('notif-canales');
    if (ch && !ch.dataset.loaded) ch.innerHTML = skelCards(5, 'grid-template-columns:repeat(5,1fr)');
    const r = await fetch('/api/notif_config');
    _notifConfig = await r.json();
    if (ch) ch.dataset.loaded = '1';
  } catch(e) {
    _notifConfig = {canales:{email:true,push:true},email:'',whatsapp:'',telegram_chat:'',alertas:{}};
  }
  renderNotifConfig();
}

function renderNotifConfig() {
  const c = _notifConfig;
  // Channels
  const cont = document.getElementById('notif-canales');
  if (cont) {
    cont.innerHTML = NOTIF_CHANNELS.map(ch => {
      const on = c.canales && c.canales[ch.key];
      return '<div onclick="toggleNotifCanal(\'' + ch.key + '\')" style="cursor:pointer;background:' +
        (on ? 'rgba(59,130,246,.1)' : 'var(--s2)') + ';border:1px solid ' +
        (on ? 'var(--acc)' : 'var(--s2)') + ';border-radius:12px;padding:14px;text-align:center;transition:.15s">' +
        '<div style="font-size:22px;margin-bottom:6px">' + ch.icon + '</div>' +
        '<div style="font-size:13px;font-weight:600;color:' + (on ? 'var(--acc2)' : 'var(--mut)') + '">' + ch.name + '</div>' +
        '<div style="font-size:10px;color:' + (on ? 'var(--grn)' : 'var(--dim)') + ';margin-top:4px">' + (on ? '● ' + (t(\'notif.activo\')||'Activo') : '○ ' + (t(\'notif.inactivo\')||'Inactivo')) + '</div>' +
        '</div>';
    }).join('');
  }
  // Channel fields (only for active channels needing input)
  const fields = document.getElementById('notif-channel-fields');
  if (fields) {
    let html = '';
    if (c.canales && c.canales.email)
      html += notifField('email', t('notif.emailLabel') || 'Email de notificaciones', 'controller@hotel.com', c.email || '');
    if (c.canales && c.canales.whatsapp)
      html += notifField('whatsapp', 'Número WhatsApp destino (+34...)', '+34600123456', c.whatsapp || '') +
              '<div style="font-size:11px;color:var(--dim);margin-top:4px">Necesita TWILIO_ACCOUNT_SID + TWILIO_AUTH_TOKEN + TWILIO_WHATSAPP_FROM en Render</div>';
    if (c.canales && c.canales.telegram)
      html += notifField('telegram_chat', 'Telegram Chat ID', '123456789', c.telegram_chat || '');
    if (c.canales && c.canales.slack)
      html += notifField('slack_webhook', 'Slack Webhook URL', 'https://hooks.slack.com/services/...', c.slack_webhook || '');
    fields.innerHTML = html;
  }
  // Alertas
  const al = document.getElementById('notif-alertas');
  if (al) {
    al.innerHTML = NOTIF_ALERTAS.map(a => {
      const on = c.alertas && c.alertas[a.key];
      return '<label style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--tx)">' +
        '<input type="checkbox" data-alerta="' + a.key + '"' + (on ? ' checked' : '') +
        ' style="width:17px;height:17px;accent-color:var(--acc)">' + a.label + '</label>';
    }).join('');
  }
}

function notifField(key, label, ph, val) {
  return '<div><label style="display:block;font-size:11px;color:var(--mut);margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px">' + label + '</label>' +
    '<input data-field="' + key + '" value="' + val + '" placeholder="' + ph + '" ' +
    'style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:9px;padding:10px 13px;font-size:14px;outline:none;font-family:inherit"></div>';
}

function toggleNotifCanal(key) {
  if (!_notifConfig.canales) _notifConfig.canales = {};
  _notifConfig.canales[key] = !_notifConfig.canales[key];
  renderNotifConfig();
}

async function guardarNotifConfig() {
  // Collect field values
  document.querySelectorAll('[data-field]').forEach(el => {
    _notifConfig[el.dataset.field] = el.value.trim();
  });
  if (!_notifConfig.alertas) _notifConfig.alertas = {};
  document.querySelectorAll('[data-alerta]').forEach(el => {
    _notifConfig.alertas[el.dataset.alerta] = el.checked;
  });
  const btn = document.getElementById('btn-save-notif');
  btn.textContent = 'Guardando...';
  try {
    await fetch('/api/notif_config', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(_notifConfig)});
    btn.textContent = '✓ Guardado';
    setTimeout(() => { btn.textContent = '💾 Guardar configuración'; }, 2000);
  } catch(e) {
    btn.textContent = '⚠️ Error';
    setTimeout(() => { btn.textContent = '💾 Guardar configuración'; }, 2000);
  }
}

async function probarNotif() {
  const btn = document.querySelector('[onclick="probarNotif()"]');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Enviando...'; }
  try {
    const r = await fetch('/api/test_notif', {method:'POST'});
    const d = await r.json();
    if (btn) {
      btn.textContent = d.ok ? ('✓ ' + (d.message || 'Enviado')) : ('✗ ' + (d.error || 'Error'));
      btn.style.color = d.ok ? 'var(--grn)' : 'var(--red)';
      setTimeout(() => { btn.textContent = '🔔 Probar'; btn.style.color = ''; btn.disabled = false; }, 3000);
    }
  } catch(e) {
    if (btn) { btn.textContent = '✗ Error'; btn.disabled = false; }
  }
}

async function loadNotif() {
  try {
    const r = await fetch('/api/notificaciones');
    const data = await r.json();
    const tbody = document.getElementById('notif-tbody');
    const count = document.getElementById('notif-count');
    if (!data || !data.length) {
      tbody.innerHTML = '<tr><td colspan="5" class="empty"><p>Sin notificaciones.</p></td></tr>';
      count.textContent = '';
      return;
    }
    count.textContent = data.length + ' registros';
    // Mostrar en orden inverso (más reciente primero)
    const rows = data.slice().reverse();
    tbody.innerHTML = rows.map(function(n) {
      const TIPOS = {ar_discrepancia:'AR Disc.',ar_falta_di:'AR DI',drr_oob:'DRR OOB',ap_discrepancia:'AP Disc.',general:'General'};
      const tipo = TIPOS[n.tipo] || n.tipo || '—';
      const est = n.estado === 'enviado'
        ? '<span class="badge b-ok">✓ Enviado</span>'
        : '<span class="badge b-disc">✗ Error</span>';
      return '<tr>'
        + '<td class="td-dim">' + (n.fecha || '—') + '</td>'
        + '<td>' + tipo + '</td>'
        + '<td class="td-b">' + (n.asunto || '—').substring(0,50) + '</td>'
        + '<td class="td-dim">' + (n.destinatario || '—') + '</td>'
        + '<td>' + est + '</td>'
        + '</tr>';
    }).join('');
  } catch(e) {
    console.warn('Error cargando notificaciones:', e);
  }
}

async function enviarNotificaciones() {
  const btn = document.getElementById('btn-send-notif');
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  try {
    const r = await fetch('/api/enviar_notificaciones', {method:'POST'});
    const data = await r.json();
    if (data.ok) {
      btn.textContent = '✓ ' + data.enviadas + ' alerta(s) procesadas';
      setTimeout(function() { btn.textContent = '🔔 Enviar notificaciones pendientes'; btn.disabled = false; }, 3000);
      loadNotif();
    } else {
      btn.textContent = '✗ Error: ' + (data.error || '');
      setTimeout(function() { btn.textContent = '🔔 Enviar notificaciones pendientes'; btn.disabled = false; }, 3000);
    }
  } catch(e) {
    btn.textContent = '✗ Error de conexión';
    setTimeout(function() { btn.textContent = '🔔 Enviar notificaciones pendientes'; btn.disabled = false; }, 3000);
  }
}

// ── Role-based visibility ──
(function() {
  var rol = '__USER_ROL__';
  // financial_controller y admin ven todo
  if (rol === 'financial_controller' || rol === 'admin') return;
  var tabs = document.querySelectorAll('.tab');
  var VISIBLE = {
    'income_auditor': ['drr'],
    'fb_manager': ['ap', 'fb'],
    'jefe_otras': ['ap'],
  };
  var allowed = VISIBLE[rol] || [];
  tabs.forEach(function(t) {
    var onclick = t.getAttribute('onclick') || '';
    var match = onclick.match(/switchTab\('(\w+)'/);
    if (match) {
      var tabId = match[1];
      if (allowed.length && allowed.indexOf(tabId) === -1) {
        t.style.display = 'none';
      }
    }
  });
  // Activate first visible tab
  if (allowed.length) {
    switchTab(allowed[0], document.querySelector('.tab[onclick*="' + allowed[0] + '"]'));
  }
})();

// ══════════════════════════════════════════════════════════════
// BANCO — JavaScript
// ══════════════════════════════════════════════════════════════

async function loadBanco() {
  try {
    var r = await fetch('/api/stats_banco');
    var d = await r.json();
    if (!d) return;
    document.getElementById('bk-total').textContent = d.total || '—';
    document.getElementById('bk-conc').textContent = d.conciliados || '0';
    document.getElementById('bk-pend').textContent = d.pendientes || '0';
    document.getElementById('bk-diff').textContent = d.diferencias || '0';
    document.getElementById('bk-imp-pend').textContent = d.importe_pendiente ? eur(d.importe_pendiente) + ' pend.' : '—';

    var el = document.getElementById('bk-alertas');
    if (d.alertas && d.alertas.length) {
      el.innerHTML = d.alertas.map(function(a) {
        return '<div class="act-item"><div class="adot r"></div><div class="atxt"><b>' + a.dias + ' dias</b> sin conciliar: ' + a.concepto + ' — ' + eur(a.importe) + '</div></div>';
      }).join('');
    } else {
      el.innerHTML = '<div class="empty"><p>Sin alertas bancarias pendientes.</p></div>';
    }
  } catch(e) {
    console.warn('Error banco:', e);
  }
}

// Cargar datos AP e iniciar
loadAP();
setInterval(loadAP, 60000);
loadDRR();
loadNotif();
loadBanco();



// ═════════════════════════════════════════════════════════════════════
// AR REAL — Procesar grupos corporativos
// ═════════════════════════════════════════════════════════════════════
function procesarARReal() {
  const logCard = document.getElementById('ar-log-card');
  const logDiv  = document.getElementById('ar-real-log');
  const btn     = document.getElementById('btn-ar-real');
  const spin    = document.getElementById('spin-ar');
  const lbl     = document.getElementById('lbl-ar');
  if (logCard) logCard.style.display = 'block';
  if (logDiv)  logDiv.innerHTML = '';
  if (spin) spin.style.display = 'inline-block';
  if (lbl)  lbl.textContent = 'Procesando...';
  if (btn)  btn.disabled = true;

  const es = new EventSource('/api/procesar_ar_real');
  es.onmessage = function(e) {
    const d = document.createElement('div');
    d.style.padding = '1px 0';
    d.style.color = e.data.startsWith('ERROR') ? 'var(--red)' :
                    e.data.startsWith('  ✓') ? 'var(--grn)' : 'var(--tx)';
    d.textContent = e.data;
    if (logDiv) { logDiv.appendChild(d); logDiv.scrollTop = logDiv.scrollHeight; }
    if (e.data === 'AR_REAL_COMPLETO') {
      es.close();
      if (spin) spin.style.display = 'none';
      if (lbl)  lbl.textContent = '▶ Procesar Archivos';
      if (btn)  btn.disabled = false;
      setTimeout(cargarARRealData, 800);
    }
  };
  es.onerror = function() {
    es.close();
    if (spin) spin.style.display = 'none';
    if (lbl)  lbl.textContent = '▶ Procesar Archivos';
    if (btn)  btn.disabled = false;
  };
}

async function cargarStatusARReal() { await cargarARRealData(); }

async function cargarARRealData() {
  try {
    // Skeleton while loading
    const kpis = document.getElementById('ar-real-kpis');
    if (kpis && !kpis.dataset.loaded) kpis.innerHTML = skelCards(4, 'grid-template-columns:repeat(4,1fr)');
    const r = await fetch('/api/ar_real_data');
    const d = await r.json();
    if (d.error) return;
    const k = d.kpis;
    const fmt = v => '€' + Math.round(v).toLocaleString('es-ES');

    // KPIs
    const _s = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    _s('arp-pend',     fmt(k.pendiente_facturar));
    _s('arp-fact',     fmt(k.facturado));
    _s('arp-cobr',     fmt(k.cobrado));
    _s('arp-saldo',    fmt(k.saldo_total));
    _s('arp-nclientes', k.num_clientes + ' clientes activos');
    const kpisEl = document.getElementById('ar-real-kpis'); if (kpisEl) kpisEl.dataset.loaded = '1';

    // Clients table
    const ctbody = document.getElementById('ar-clients-tbody');
    if (ctbody) {
      ctbody.innerHTML = d.clientes.map(c => {
        const sc = c.status === 'critical' ? 'var(--red)' : c.status === 'warning' ? 'var(--ora)' : 'var(--grn)';
        const dot = '<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:' + sc + '"></span>';
        const pct = c.limite_credito > 0 ? Math.round(c.saldo_pendiente/c.limite_credito*100) : 0;
        return '<tr>' +
          '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + c.nombre + '">' + c.nombre.split(' ').slice(0,3).join(' ') + '</td>' +
          '<td style="text-align:right;font-weight:700;color:' + (c.saldo_pendiente>0?'var(--ora)':'var(--grn)') + '">€' + c.saldo_pendiente.toLocaleString('es-ES') + '</td>' +
          '<td style="text-align:center;color:var(--mut)">' + c.dias_pago + 'd</td>' +
          '<td style="text-align:center">' + dot + '</td>' +
          '</tr>';
      }).join('');
    }

    // Reservations table
    const rtbody = document.getElementById('ar-reservas-tbody');
    if (rtbody) {
      rtbody.innerHTML = d.reservas.map(r => {
        const badge = r.estado === 'PENDIENTE_FACTURA' ? '<span class="badge b-unk">Pend.</span>' :
                      r.estado === 'FACTURADO'         ? '<span class="badge b-cok">Fact.</span>' :
                                                          '<span class="badge b-ok">Cobr.</span>';
        return '<tr>' +
          '<td style="font-weight:600">' + r.numero + '</td>' +
          '<td style="color:var(--mut);font-size:11px">' + r.fecha_entrada + '</td>' +
          '<td style="text-align:right;font-weight:700">€' + r.total.toLocaleString('es-ES') + '</td>' +
          '<td style="text-align:center">' + badge + '</td>' +
          '</tr>';
      }).join('');
    }
  } catch(e) { console.warn('AR Real data:', e); }
}



// ═══════════════════════════════════════════════════════════════════
// MULTI-HOTEL DASHBOARD
// ═══════════════════════════════════════════════════════════════════
let _mh_loaded = false;

async function loadMultiHotel() {
  const grupoSelect = document.getElementById('grupo-filter');
  const grupo = grupoSelect ? grupoSelect.value : '';
  const mhk = document.getElementById('mh-kpis');
  if (mhk && !mhk.dataset.loaded) mhk.innerHTML = skelCards(8, 'grid-template-columns:repeat(4,1fr)');
  try {
    const overviewRes = await fetch('/api/multi_hotel/overview' + (grupo ? '?grupo=' + encodeURIComponent(grupo) : ''));
    const overview = await overviewRes.json();
    if (!_mh_loaded && grupoSelect) {
      overview.grupos.forEach(g => {
        const opt = document.createElement('option');
        opt.value = g;
        opt.textContent = g;
        grupoSelect.appendChild(opt);
      });
      _mh_loaded = true;
    }
    renderMHKpis(overview.kpis);
    renderMHStatus(overview.kpis);
    renderMHTable(overview.hoteles);
    const rankingsRes = await fetch('/api/multi_hotel/rankings' + (grupo ? '?grupo=' + encodeURIComponent(grupo) : ''));
    const rankings = await rankingsRes.json();
    renderMHRankings(rankings.top_revenue);
    const alertasRes = await fetch('/api/multi_hotel/alertas' + (grupo ? '?grupo=' + encodeURIComponent(grupo) : ''));
    const alertas = await alertasRes.json();
    renderMHAlertas(alertas.lista);
  } catch(e) {
    console.error('Error Multi-Hotel:', e);
  }
}

function renderMHKpis(kpis) {
  const cont = document.getElementById('mh-kpis');
  if (!cont) return;
  const cards = [
    {label: 'Hoteles', value: kpis.num_hoteles, color: '#1a73e8'},
    {label: 'Habitaciones', value: kpis.total_rooms.toLocaleString('es-ES'), color: '#1a73e8'},
    {label: 'Revenue MTD', value: '€' + (kpis.total_revenue_mtd / 1000000).toFixed(2) + 'M', color: '#1db954'},
    {label: 'Ocupacion Avg', value: kpis.avg_occupancy + '%', color: '#1db954'},
    {label: 'ADR Avg', value: '€' + kpis.avg_adr.toFixed(0), color: '#ff9800'},
    {label: 'RevPAR Avg', value: '€' + kpis.avg_revpar.toFixed(0), color: '#ff9800'},
    {label: 'GOP% Avg', value: kpis.avg_gop_pct + '%', color: '#1db954'},
    {label: 'Facturas Pend.', value: kpis.total_facturas_pendientes + ' (€' + (kpis.total_facturas_importe/1000).toFixed(0) + 'K)', color: '#e05252'}
  ];
  cont.innerHTML = cards.map(c =>
    '<div style="background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px 16px">' +
    '<div style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px;font-weight:600">' + c.label + '</div>' +
    '<div style="font-size:24px;font-weight:800;color:' + c.color + ';line-height:1;letter-spacing:-.5px">' + c.value + '</div>' +
    '</div>'
  ).join('');
}

function renderMHStatus(kpis) {
  const cont = document.getElementById('mh-status');
  if (!cont) return;
  const _st = (n, col, icon, lbl, bg) =>
    '<div style="background:' + bg + ';border:1px solid ' + col + '55;border-radius:13px;padding:18px 16px;' +
    'display:flex;align-items:center;gap:14px;border-left:3px solid ' + col + '">' +
    '<div style="width:36px;height:36px;border-radius:9px;background:' + col + '22;' +
    'display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:900;color:' + col + ';flex-shrink:0">' + icon + '</div>' +
    '<div><div style="font-size:30px;font-weight:800;color:' + col + ';line-height:1;letter-spacing:-1.5px">' + n + '</div>' +
    '<div style="font-size:11px;color:var(--mut);margin-top:5px;text-transform:uppercase;letter-spacing:.5px">' + lbl + '</div></div></div>';
  cont.innerHTML =
    _st(kpis.hoteles_ok,       '#22c55e', '✓', 'Hoteles OK',  'var(--s1)') +
    _st(kpis.hoteles_warning,  '#f97316', '!', 'Con alertas', 'var(--s1)') +
    _st(kpis.hoteles_criticos, '#ef4444', '✕', 'Críticos',    'var(--s1)');
}

function renderMHTable(hoteles) {
  const tbody = document.getElementById('mh-tbody');
  if (!tbody) return;
  tbody.innerHTML = '';
  hoteles.forEach(h => {
    const statusColor = h.status === 'ok' ? '#1db954' : h.status === 'warning' ? '#ff9800' : '#e05252';
    const statusIcon = h.status === 'ok' ? '*' : h.status === 'warning' ? '!' : 'X';
    const tr = document.createElement('tr');
    tr.style.cssText = 'border-bottom:1px solid #2e3248;cursor:pointer;transition:background 0.2s';
    tr.dataset.hotelId = h.id;
    tr.addEventListener('mouseover', () => tr.style.background = 'rgba(26,115,232,0.08)');
    tr.addEventListener('mouseout', () => tr.style.background = 'transparent');
    tr.addEventListener('click', () => openHotelDetail(h.id));
    tr.innerHTML = 
      '<td style="padding:10px;font-weight:600">' + h.nombre + ' <span style="color:#8892a4;font-size:11px">' + h.tier + '</span></td>' +
      '<td style="padding:10px;color:#8892a4">' + h.ciudad + ', ' + h.pais + '</td>' +
      '<td style="padding:10px;text-align:right">' + h.habitaciones + '</td>' +
      '<td style="padding:10px;text-align:right">' + h.ocupacion_pct + '%</td>' +
      '<td style="padding:10px;text-align:right">€' + h.adr.toFixed(0) + '</td>' +
      '<td style="padding:10px;text-align:right;font-weight:600">€' + h.revpar.toFixed(0) + '</td>' +
      '<td style="padding:10px;text-align:right;font-weight:600;color:#1db954">€' + (h.revenue_mtd/1000000).toFixed(2) + 'M</td>' +
      '<td style="padding:10px;text-align:right">' + h.gop_pct + '%</td>' +
      '<td style="padding:10px;text-align:right">' + h.facturas_pendientes + '</td>' +
      '<td style="padding:10px;text-align:center;color:' + statusColor + ';font-size:18px">' + statusIcon + '</td>';
    tbody.appendChild(tr);
  });
}

function renderMHRankings(top) {
  const cont = document.getElementById('mh-rankings');
  if (!cont) return;
  cont.innerHTML = '';
  top.forEach((h, i) => {
    const medalColor = i === 0 ? '#FFD700' : i === 1 ? '#C0C0C0' : i === 2 ? '#CD7F32' : '#8892a4';
    const div = document.createElement('div');
    div.style.cssText = 'display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid #2e3248;cursor:pointer;transition:background 0.2s';
    div.addEventListener('mouseover', () => div.style.background = 'rgba(26,115,232,0.08)');
    div.addEventListener('mouseout', () => div.style.background = 'transparent');
    div.addEventListener('click', () => openHotelDetail(h.id));
    div.innerHTML = 
      '<div style="display:flex;align-items:center;gap:12px">' +
        '<span style="font-size:18px;font-weight:700;color:' + medalColor + '">' + (i+1) + '</span>' +
        '<div><div style="font-weight:600;font-size:13px">' + h.nombre + '</div>' +
        '<div style="font-size:11px;color:#8892a4">' + h.ciudad + ', ' + h.pais + '</div></div>' +
      '</div>' +
      '<div style="font-weight:700;color:#1db954;font-size:15px">€' + (h.revenue_mtd/1000000).toFixed(2) + 'M</div>';
    cont.appendChild(div);
  });
}

function renderMHAlertas(alertas) {
  const cont = document.getElementById('mh-alertas');
  if (!cont) return;
  if (alertas.length === 0) {
    cont.innerHTML = '<div style="color:#1db954;text-align:center;padding:20px">Sin alertas activas</div>';
    return;
  }
  cont.innerHTML = alertas.map(a => {
    const color = a.severity === 'critical' ? '#e05252' : '#ff9800';
    return '<div style="padding:10px 12px;border-left:3px solid ' + color + ';background:rgba(255,255,255,0.02);border-radius:4px;margin-bottom:8px">' +
      '<div style="font-size:13px;color:#cdd6f4">' + a.alerta + '</div>' +
      '<div style="font-size:11px;color:#8892a4;margin-top:4px">' + a.hotel + ' &bull; ' + a.ciudad + '</div></div>';
  }).join('');
}

async function openHotelDetail(hotelId) {
  try {
    const res = await fetch('/api/multi_hotel/hotel/' + hotelId);
    const h = await res.json();
    if (h.error) { alert('Hotel no encontrado'); return; }
    const statusColor = h.status === 'ok' ? '#1db954' : h.status === 'warning' ? '#ff9800' : '#e05252';
    const statusLabel = h.status === 'ok' ? 'OK' : h.status === 'warning' ? 'WARNING' : 'CRITICO';
    let alertasHtml = '';
    if (h.alertas && h.alertas.length > 0) {
      alertasHtml = h.alertas.map(a => '<div style="padding:8px 12px;background:rgba(224,82,82,0.1);border-left:3px solid ' + statusColor + ';margin-bottom:6px;border-radius:4px;font-size:13px">! ' + a + '</div>').join('');
    } else {
      alertasHtml = '<div style="color:#1db954;padding:8px">Sin alertas activas</div>';
    }
    const gopColor = h.gop_pct >= 40 ? '#1db954' : h.gop_pct >= 35 ? '#ff9800' : '#e05252';
    const fbColor = h.fb_pct <= 18 ? '#1db954' : h.fb_pct <= 20 ? '#ff9800' : '#e05252';
    
    const modal = document.createElement('div');
    modal.id = 'hotel-modal';
    modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.8);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;overflow-y:auto';
    modal.addEventListener('click', (e) => { if (e.target === modal) closeHotelModal(); });
    
    const inner = document.createElement('div');
    inner.style.cssText = 'background:#0f1117;border:1px solid #2e3248;border-radius:16px;max-width:1000px;width:100%;max-height:90vh;overflow-y:auto;padding:32px;position:relative';
    
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'X';
    closeBtn.style.cssText = 'position:absolute;top:16px;right:16px;background:transparent;color:#8892a4;border:none;font-size:24px;cursor:pointer;width:32px;height:32px';
    closeBtn.addEventListener('click', closeHotelModal);
    inner.appendChild(closeBtn);
    
    const otasFacts = Math.floor(h.facturas_pendientes * 0.4);
    const grupoFacts = Math.floor(h.facturas_pendientes * 0.35);
    const directosFacts = Math.floor(h.facturas_pendientes * 0.25);
    
    const detailDiv = document.createElement('div');
    detailDiv.innerHTML = 
      '<div style="display:flex;align-items:center;gap:16px;margin-bottom:8px">' +
        '<h2 style="margin:0;font-size:24px">' + h.nombre + '</h2>' +
        '<span style="background:' + statusColor + ';color:white;padding:4px 12px;border-radius:6px;font-size:11px;font-weight:700">' + statusLabel + '</span>' +
      '</div>' +
      '<div style="color:#8892a4;margin-bottom:24px">' + h.tier + ' &bull; ' + h.ciudad + ', ' + h.pais + ' &bull; ' + h.grupo + '</div>' +
      '<h3 style="font-size:14px;color:#8892a4;margin:24px 0 12px 0">KPIs Operativos</h3>' +
      '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px">' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:14px"><div style="font-size:11px;color:#8892a4">Habitaciones</div><div style="font-size:22px;font-weight:700">' + h.habitaciones + '</div></div>' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:14px"><div style="font-size:11px;color:#8892a4">Ocupacion</div><div style="font-size:22px;font-weight:700;color:#1a73e8">' + h.ocupacion_pct + '%</div></div>' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:14px"><div style="font-size:11px;color:#8892a4">ADR</div><div style="font-size:22px;font-weight:700">€' + h.adr.toFixed(0) + '</div></div>' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:14px"><div style="font-size:11px;color:#8892a4">RevPAR</div><div style="font-size:22px;font-weight:700">€' + h.revpar.toFixed(0) + '</div></div>' +
      '</div>' +
      '<h3 style="font-size:14px;color:#8892a4;margin:24px 0 12px 0">Performance Financiero</h3>' +
      '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:24px">' +
        '<div style="background:var(--s1);border:1px solid rgba(34,197,94,.3);border-radius:10px;padding:16px"><div style="font-size:11px;color:var(--mut)">Revenue MTD</div><div style="font-size:26px;font-weight:700;color:#22c55e">€' + (h.revenue_mtd/1000000).toFixed(2) + 'M</div></div>' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:16px"><div style="font-size:11px;color:#8892a4">GOP%</div><div style="font-size:26px;font-weight:700;color:' + gopColor + '">' + h.gop_pct + '%</div></div>' +
        '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:16px"><div style="font-size:11px;color:#8892a4">F&B Cost %</div><div style="font-size:26px;font-weight:700;color:' + fbColor + '">' + h.fb_pct + '%</div></div>' +
      '</div>' +
      '<h3 style="font-size:14px;color:#8892a4;margin:24px 0 12px 0">AR Dashboard - Facturas Pendientes</h3>' +
      '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:10px;padding:20px;margin-bottom:24px">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">' +
          '<div>' +
            '<div style="font-size:11px;color:#8892a4">Total pendiente de cobro</div>' +
            '<div style="font-size:32px;font-weight:700;color:#ff9800">€' + h.facturas_importe.toLocaleString('es-ES') + '</div>' +
            '<div style="font-size:13px;color:#8892a4;margin-top:4px">' + h.facturas_pendientes + ' facturas activas</div>' +
          '</div>' +
        '</div>' +
        '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:16px;padding-top:16px;border-top:1px solid #2e3248">' +
          '<div><div style="font-size:11px;color:#8892a4">OTAs (Booking, Expedia)</div><div style="font-size:18px;font-weight:600;color:#1a73e8">' + otasFacts + ' facts</div></div>' +
          '<div><div style="font-size:11px;color:#8892a4">Grupos Corporativos</div><div style="font-size:18px;font-weight:600;color:#1a73e8">' + grupoFacts + ' facts</div></div>' +
          '<div><div style="font-size:11px;color:#8892a4">Clientes Directos</div><div style="font-size:18px;font-weight:600;color:#1a73e8">' + directosFacts + ' facts</div></div>' +
        '</div>' +
      '</div>' +
      '<h3 style="font-size:14px;color:#8892a4;margin:24px 0 12px 0">Alertas' + (h.alertas && h.alertas.length > 0 ? ' (' + h.alertas.length + ')' : '') + '</h3>' +
      '<div>' + alertasHtml + '</div>';
    
    inner.appendChild(detailDiv);
    modal.appendChild(inner);
    document.body.appendChild(modal);
  } catch(e) {
    console.error('Error abriendo detalle:', e);
    alert('Error cargando hotel: ' + e.message);
  }
}


// ═══════════════════════════════════════════════════════════════════
// CALIPOLIS DASHBOARD
// ═══════════════════════════════════════════════════════════════════
let _calCharts = {};

async function loadCalipolis() {
  try {
    const kEl = document.getElementById('cal-kpis');
    const hEl = document.getElementById('cal-hoteles');
    if (kEl && !kEl.dataset.loaded) kEl.innerHTML = skelCards(4, 'grid-template-columns:repeat(4,1fr)');
    if (hEl && !hEl.dataset.loaded) hEl.innerHTML = skelCards(3, 'grid-template-columns:repeat(3,1fr)');
    const res = await fetch('/api/calipolis/kpis');
    const data = await res.json();
    renderCalipolisKpis(data.consolidado);
    renderCalipolisHoteles(data.hoteles);
    if (data.tendencias) renderCalipolisTrends(data.tendencias);
  } catch(e) {
    console.error('Error Calipolis:', e);
  }
}

function renderCalipolisTrends(t) {
  if (!window.Chart) return;
  const meses = (t.meses || []).map(m => m.slice(5)); // MM only
  const cfg = (id, label, data, color, isCurrency) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (_calCharts[id]) { _calCharts[id].destroy(); }
    _calCharts[id] = new Chart(el, {
      type: 'line',
      data: {
        labels: meses,
        datasets: [{
          data: data,
          borderColor: color,
          backgroundColor: color + '18',
          fill: true,
          tension: 0.35,
          borderWidth: 2,
          pointRadius: 3,
          pointBackgroundColor: color,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false },
          tooltip: { callbacks: { label: v => isCurrency ? v.raw + '%' : v.raw + ' fact.' }}},
        scales: {
          x: { grid: { color: 'rgba(51,65,85,.3)' }, ticks: { color: '#64748b', font: { size: 9 }}},
          y: { grid: { color: 'rgba(51,65,85,.3)' }, ticks: { color: '#64748b', font: { size: 9 },
            callback: v => isCurrency ? v + '%' : v }}
        }
      }
    });
  };
  cfg('cal-gop-chart',  'GOP%',       t.gop_pct_grupo,  '#22c55e', true);
  cfg('cal-ap-chart',   'AP Pend.',   t.ap_pendientes_total, '#ef4444', false);
}

function renderCalipolisKpis(kpis) {
  const cont = document.getElementById('cal-kpis');
  if (!cont) return;
  const totalRevM = (kpis.total_revenue_mtd/1000000).toFixed(2);
  const totalGopK = (kpis.total_gop/1000).toFixed(0);
  const cards = [
    {label:'Revenue MTD',   value:'€'+totalRevM+'M', color:'var(--grn)',  sub:'grupo completo'},
    {label:'Ocupación Avg', value:kpis.avg_ocupacion+'%', color:'var(--acc2)', sub:'3 propiedades'},
    {label:'ADR Avg',       value:'€'+Math.round(kpis.avg_adr), color:'var(--ora)', sub:'precio medio'},
    {label:'GOP% Avg',      value:kpis.avg_gop_pct+'%', color:'var(--grn)', sub:'beneficio bruto'},
    {label:'RevPAR Avg',    value:'€'+Math.round(kpis.avg_revpar), color:'var(--ora)', sub:'revenue/hab'},
    {label:'GOP Total',     value:'€'+totalGopK+'K', color:'var(--grn)', sub:'este mes'},
    {label:'AP Pendientes', value:kpis.total_ap_pendientes||'—', color:'var(--red)', sub:'facturas'},
    {label:'Alertas',       value:kpis.total_alertas||'0', color: kpis.total_alertas>0?'var(--yel)':'var(--grn)', sub:'activas'},
  ];
  cont.innerHTML = cards.map(c =>
    '<div style="background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px 16px">' +
    '<div style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px;font-weight:600">' + c.label + '</div>' +
    '<div style="font-size:24px;font-weight:800;color:' + c.color + ';line-height:1;letter-spacing:-.5px">' + c.value + '</div>' +
    '<div style="font-size:10px;color:var(--dim);margin-top:6px">' + c.sub + '</div>' +
    '</div>'
  ).join('');
}

function renderCalipolisHoteles(hoteles) {
  const cont = document.getElementById('cal-hoteles');
  if (!cont) return;
  cont.innerHTML = '';
  hoteles.forEach(h => {
    const sc = h.status === 'ok' ? '#22c55e' : h.status === 'warning' ? '#f97316' : '#ef4444';
    const apTrend = h.ap_pendientes === 0 ? '\u2713 Sin pendientes' : h.ap_pendientes + ' fact. pend.';
    const apColor = h.ap_pendientes === 0 ? '#22c55e' : h.ap_pendientes <= 3 ? '#f97316' : '#ef4444';
    const gopColor = h.gop_pct >= 22 ? '#22c55e' : h.gop_pct >= 18 ? '#f97316' : '#ef4444';
    const card = document.createElement('div');
    card.style.cssText = 'background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px;cursor:pointer;transition:border-color .18s';
    card.addEventListener('mouseover', function() { card.style.borderColor = sc + '55'; });
    card.addEventListener('mouseout',  function() { card.style.borderColor = 'var(--s2)'; });
    card.addEventListener('click',     function() { abrirDetalleCalipolis(h.id); });
    card.innerHTML =
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px">' +
        '<div><div style="font-weight:700;font-size:14px;margin-bottom:3px">' + h.nombre + '</div>' +
        '<div style="font-size:11px;color:var(--mut)">' + h.categoria + ' \u00b7 ' + h.habitaciones + ' hab.</div></div>' +
        '<div style="width:8px;height:8px;border-radius:50%;background:' + sc + ';margin-top:4px;flex-shrink:0"></div>' +
      '</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px">' +
        '<div style="background:var(--s2);border-radius:8px;padding:10px"><div style="font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Ocup.</div><div style="font-size:18px;font-weight:800;color:var(--acc2)">' + h.ocupacion + '%</div></div>' +
        '<div style="background:var(--s2);border-radius:8px;padding:10px"><div style="font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">RevPAR</div><div style="font-size:18px;font-weight:800;color:var(--ora)">\u20ac' + Math.round(h.revpar) + '</div></div>' +
        '<div style="background:var(--s2);border-radius:8px;padding:10px"><div style="font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Revenue MTD</div><div style="font-size:18px;font-weight:800;color:var(--tx)">\u20ac' + Math.round(h.total_ingresos / 1000) + 'K</div></div>' +
        '<div style="background:var(--s2);border-radius:8px;padding:10px"><div style="font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">GOP%</div><div style="font-size:18px;font-weight:800;color:' + gopColor + '">' + h.gop_pct + '%</div></div>' +
      '</div>' +
      '<div style="display:flex;justify-content:space-between;align-items:center;padding-top:12px;border-top:1px solid var(--s2)">' +
        '<span style="font-size:11px;color:' + apColor + ';font-weight:600">' + apTrend + '</span>' +
        '<span style="font-size:11px;color:var(--mut)">Ver detalle \u2192</span>' +
      '</div>';
    cont.appendChild(card);
  });
}


// ═══════════════════════════════════════════════════════════════════
// INTEGRACIONES
// ═══════════════════════════════════════════════════════════════════


function closeHotelModal() {
  const modal = document.getElementById('hotel-modal');
  if (modal) modal.remove();
}


// ═══════════════════════════════
// I18N — Sistema de traducción
// ═══════════════════════════════
const _i18nCache = {};
const _i18nOriginal = {}; // textos ES originales — para restaurar al volver a español
let _i18nData = {};
let _i18nLang = localStorage.getItem('yve_lang') || 'es';

function _saveOriginals() {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const k = el.getAttribute('data-i18n');
    if (!_i18nOriginal[k]) _i18nOriginal[k] = el.textContent;
  });
}

async function loadI18n(lang) {
  _saveOriginals();
  if (lang === 'es') {
    _i18nData = {}; _i18nLang = 'es';
    applyI18n(_i18nOriginal); // restaura textos originales
    localStorage.setItem('yve_lang', 'es'); return;
  }
  if (_i18nCache[lang]) {
    _i18nData = _i18nCache[lang]; _i18nLang = lang;
    applyI18n(_i18nData); localStorage.setItem('yve_lang', lang); return;
  }
  try {
    const r = await fetch('/static/i18n/' + lang + '.json');
    const data = await r.json();
    _i18nCache[lang] = data; _i18nData = data; _i18nLang = lang;
    applyI18n(data); localStorage.setItem('yve_lang', lang);
  } catch(e) { console.warn('i18n error:', e); }
}

function t(key) { return _i18nData[key] || _i18nOriginal[key] || key; }

function applyI18n(data) {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const k = el.getAttribute('data-i18n');
    if (data[k] !== undefined) el.textContent = data[k];
  });
}

async function cambiarIdioma(lang) {
  fetch('/api/set_lang/' + lang);   // fire-and-forget, no await
  await loadI18n(lang);
  document.querySelectorAll('.lang-btn').forEach(b => {
    b.style.fontWeight = b.dataset.lang === lang ? '700' : '400';
    b.style.color = b.dataset.lang === lang ? 'var(--acc2)' : 'var(--tx)';
  });
}

// Cargar idioma actual al inicio
loadI18n(_i18nLang);
// Precargar todos los demás idiomas en segundo plano (2s delay)
setTimeout(() => {
  ['en','ca','fr','de','it','pt'].forEach(lang => {
    if (!_i18nCache[lang]) {
      fetch('/static/i18n/' + lang + '.json')
        .then(r => r.json())
        .then(d => { _i18nCache[lang] = d; })
        .catch(() => {});
    }
  });
}, 2000);

</script>
</body>
</html>"""


@app.route('/api/demo/toggle', methods=['POST'])
def toggle_demo():
    """Activa/desactiva demo mode"""
    global DEMO_MODE
    DEMO_MODE = not DEMO_MODE
    return jsonify({"demo_mode": DEMO_MODE, "status": "activado" if DEMO_MODE else "desactivado"})

@app.route('/api/demo/status')
def demo_status():
    """Retorna estado de demo mode"""
    return jsonify({
        "demo_mode": DEMO_MODE,
        "hoteles": generar_hoteles_demo()["hoteles"] if DEMO_MODE else [],
        "alertas": generar_alertas_demo() if DEMO_MODE else []
    })

@app.route('/api/demo/datos')
def demo_data():
    """Todos los datos de demo"""
    if not DEMO_MODE:
        return jsonify({"error": "Demo mode desactivado"}), 403
    
    return jsonify({
        "hoteles": generar_hoteles_demo()["hoteles"],
        "facturas_ar": generar_facturas_demo_ar(),
        "facturas_ap": generar_facturas_demo_ap(),
        "alertas": generar_alertas_demo()
    })



@app.route('/api/rol/cambiar/<new_role>', methods=['POST'])
def cambiar_rol(new_role):
    """Cambiar rol del usuario actual"""
    roles_validos = ['admin', 'financial_controller', 'income_auditor', 'fb_manager', 'jefe_otras']
    if new_role not in roles_validos:
        return jsonify({"error": "Rol inválido"}), 400
    
    session['rol'] = new_role
    return jsonify({"rol": new_role, "status": "cambiado"})

@app.route('/api/rol/actual')
def rol_actual():
    """Retorna el rol actual y todos los roles disponibles"""
    rol_actual = session.get('rol', 'admin')
    return jsonify({
        "rol_actual": rol_actual,
        "roles_disponibles": {
            "admin": "Administrador - Acceso total",
            "financial_controller": "Controller Financiero",
            "income_auditor": "Income Auditor",
            "fb_manager": "Jefe de F&B",
            "jefe_otras": "Jefe de Servicios"
        }
    })



@app.route('/api/notificaciones/chequear', methods=['POST'])
def chequear_notificaciones():
    """Escanea alertas y envía pendientes"""
    try:
        escanear_alertas()
        enviados = enviar_pendientes(solo_check=False)
        return jsonify({
            "status": "ok",
            "mensaje": f"{enviados} notificaciones procesadas"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/notificaciones/historial')
def historial_notificaciones():
    """Retorna historial de notificaciones"""
    try:
        with open('datos-referencia/notificaciones_historial.json', 'r') as f:
            historial = json.load(f)
        return jsonify(historial[-50:])  # Últimas 50
    except:
        return jsonify([])



@app.route('/api/reportes/ejecutivo.pdf')
def reporte_ejecutivo():
    """Descarga reporte ejecutivo en PDF"""
    try:
        pdf = generar_reporte_ejecutivo()
        return send_file(pdf, mimetype='application/pdf', as_attachment=True, download_name='Reporte_Ejecutivo.pdf')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/reportes/consolidado.xlsx')
def reporte_consolidado():
    """Descarga reporte consolidado en Excel"""
    try:
        xlsx = generar_excel_consolidado()
        return send_file(xlsx, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Consolidado.xlsx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/set_lang/<lang>')
def set_lang(lang):
    """Guarda preferencia de idioma"""
    allowed = ['es','en','ca','fr','de','it','pt']
    if lang in allowed:
        session['lang'] = lang
        return jsonify({'ok': True, 'lang': lang})
    return jsonify({'ok': False}), 400

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
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)

# Redeploy trigger Sat Jun  6 12:09:22 UTC 2026
# Force fresh deploy 1780748772
