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
from tab_integraciones import integraciones_bp
from rol_dashboard import get_dashboard_config
from demo_completo import generar_hoteles_demo, generar_facturas_demo_ar, generar_facturas_demo_ap, generar_alertas_demo
for _bp in (auth_bp, config_bp, admin_bp, aprob_ar_bp, aprob_ap_bp, concil_bp, fb_bp, ar_real_bp, multi_hotel_bp, exportador_bp, calipolis_bp, demo_bp, demo_sim_bp, calipolis_analisis_bp, reportes_pdf_bp, integraciones_bp):
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
                    from notificaciones import enviar_pendientes
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
                    from notificaciones import enviar_pendientes
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
            "⚠️ API key de Anthropic no configurada. Añade ANTHROPIC_API_KEY en el archivo .env."}), 200

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
        return jsonify({"reply": f"⚠️ Error al llamar a Claude: {str(e)[:120]}"}), 200

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
        from notificaciones import enviar_pendientes
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
@login_required
def index():
    name = _hotel_name()
    tag = name if name else "AR Dashboard"
    configured = "true" if name else "false"
    # User info for header
    user_name = current_user.nombre if current_user.is_authenticated else ""
    user_rol  = current_user.rol if current_user.is_authenticated else ""
    out = HTML.replace("__HOTEL_TAG__", tag).replace("__CONFIGURED__", configured)
    admin_display = "inline" if user_rol in ("admin", "financial_controller") else "none"
    out = out.replace("__USER_NAME__", user_name).replace("__USER_ROL__", user_rol)
    out = out.replace("__ADMIN_DISPLAY__", admin_display)
    return out

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
body{background:var(--bg);color:var(--tx);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;min-height:100vh;line-height:1.5}

/* ── NAV ── */
.nav{background:var(--s1);border-bottom:1px solid var(--s2);padding:0 24px;height:60px;display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:200}
.logo{display:flex;align-items:baseline;gap:10px;flex-shrink:0}
.logo-name{font-size:20px;font-weight:800;color:#fff;letter-spacing:-0.5px}
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
    <button class="btn-ref" id="btn-demo" onclick="toggleDemoMode()" title="Activar Demo Mode para inversores" style="color:#9333ea;border-color:#9333ea">🎭 Demo</button>
    <a href="/configuracion/" class="btn-ref" title="Configuración" style="text-decoration:none">⚙️</a>
    <a href="/admin/" class="btn-ref" title="Admin" style="text-decoration:none;display:__ADMIN_DISPLAY__">👥</a>
    <div style="display:inline-block;position:relative">
      <button class="btn-ref" onclick="document.getElementById('reportes-menu').style.display=document.getElementById('reportes-menu').style.display==='block'?'none':'block'" title="Descargar reportes PDF" style="color:#1db954;border-color:#1db954">📄 Reportes</button>
      <div id="reportes-menu" style="display:none;position:absolute;top:42px;right:0;background:#1c1f2e;border:1px solid #2e3248;border-radius:8px;padding:8px;z-index:1000;min-width:180px;box-shadow:0 4px 12px rgba(0,0,0,0.5)">
        <a href="/api/reportes/diario" style="display:block;padding:10px 12px;color:#fff;text-decoration:none;border-radius:6px;font-size:13px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">📄 Reporte Diario</a>
        <a href="/api/reportes/semanal" style="display:block;padding:10px 12px;color:#fff;text-decoration:none;border-radius:6px;font-size:13px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">📊 Reporte Semanal</a>
        <a href="/api/reportes/mensual" style="display:block;padding:10px 12px;color:#fff;text-decoration:none;border-radius:6px;font-size:13px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">📈 Reporte Mensual</a>
      </div>
    </div>
    <button class="btn-ref" onclick="loadAll()" title="Actualizar datos">↻ Actualizar</button>
    <button class="btn-run" id="btn-run" onclick="runPipeline()">
      <div class="spin" id="spin"></div>
      <div style="display:inline-block;position:relative;margin-right:8px">
      <button class="btn-ref" id="rol-btn" title="Cambiar rol" style="color:#ff9800;border-color:#ff9800">👤 Admin</button>
      <div id="rol-menu" style="display:none;position:absolute;top:42px;right:0;background:#1c1f2e;border:1px solid #2e3248;border-radius:8px;padding:8px;z-index:1000;min-width:220px">
        <div style="padding:8px 12px;color:#8892a4;font-size:11px;font-weight:600">CAMBIAR ROL</div>
        <button onclick="cambiarRol('admin')" style="display:block;width:100%;text-align:left;padding:8px 12px;color:#fff;border:none;background:transparent;cursor:pointer;border-radius:4px;font-size:12px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">🔑 Administrador</button>
        <button onclick="cambiarRol('financial_controller')" style="display:block;width:100%;text-align:left;padding:8px 12px;color:#fff;border:none;background:transparent;cursor:pointer;border-radius:4px;font-size:12px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">💰 Controller Financiero</button>
        <button onclick="cambiarRol('income_auditor')" style="display:block;width:100%;text-align:left;padding:8px 12px;color:#fff;border:none;background:transparent;cursor:pointer;border-radius:4px;font-size:12px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">📊 Income Auditor</button>
        <button onclick="cambiarRol('fb_manager')" style="display:block;width:100%;text-align:left;padding:8px 12px;color:#fff;border:none;background:transparent;cursor:pointer;border-radius:4px;font-size:12px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">🍽️ Jefe F&B</button>
        <button onclick="cambiarRol('jefe_otras')" style="display:block;width:100%;text-align:left;padding:8px 12px;color:#fff;border:none;background:transparent;cursor:pointer;border-radius:4px;font-size:12px" onmouseover="this.style.background='#2e3248'" onmouseout="this.style.background='transparent'">🛠️ Jefe Servicios</button>
      </div>
    </div>
    <span id="run-lbl">⚡ Procesar Facturas</span>
    </button>
    <a href="/logout" class="btn-ref" title="Cerrar sesión" style="text-decoration:none">Salir</a>
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

  <!-- TABS -->
  <div class="tabs">
    <button class="tab active" onclick="switchTab('ar',this)">📥 AR — OTAs</button>
    <button class="tab" onclick="switchTab('ap',this)">📦 AP — Proveedores</button>
    <button class="tab" onclick="switchTab('drr',this)">📊 DRR</button>
    <button class="tab" onclick="switchTab('banco',this)">🏦 Banco</button>
    <button class="tab" onclick="switchTab('notif',this)">🔔 Notificaciones</button>
    <button class="tab" onclick="switchTab('fb',this)" id="tab-fb">🍽️ F&amp;B Cost</button>
    <button class="tab" onclick="switchTab('ar_real',this)" id="tab-ar-real">🏢 AR Real</button>
    <button class="tab" onclick="switchTab('calipolis',this)" id="tab-calipolis">🏩 Calipolis</button>
    <button class="tab" onclick="switchTab('integraciones',this)" id="tab-integraciones">⚙️ Integraciones</button>
    <button class="tab" onclick="switchTab('multi_hotel',this)" id="tab-multi-hotel">🏨 Multi-Hotel</button>
  </div>

  <div id="panel-ar" class="panel active">
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ar" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px">⬇️ Descargar Excel</a><a href="/aprobaciones-ar/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AR">📲 Aprobar facturas AR</a></div>
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

  </div><!-- /panel-ar -->

  <!-- PANEL AP -->
  <div id="panel-ap" class="panel">
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ap" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px">⬇️ Descargar Excel</a><a href="/aprobaciones-ap/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AP">📲 Aprobar facturas AP</a></div>
    <div class="stats" id="stats-ap-grid">
      <div class="sc hl c-blu"><div class="sc-lbl">Total Facturas AP</div><div class="sc-val" id="ap-total">—</div><div class="sc-sub">proveedores</div></div>
      <div class="sc"><div class="sc-lbl">Importe Total</div><div class="sc-val" id="ap-importe" style="font-size:18px;letter-spacing:-.5px">—</div><div class="sc-sub">EUR</div></div>
      <div class="sc c-grn"><div class="sc-lbl">Matches OK</div><div class="sc-val" id="ap-matches">—</div><div class="sc-sub">F&B + OTRAS</div></div>
      <div class="sc c-red"><div class="sc-lbl">Discrepancias</div><div class="sc-val" id="ap-disc">—</div><div class="sc-sub">vs PO</div></div>
      <div class="sc c-ora"><div class="sc-lbl">Sin PO</div><div class="sc-val" id="ap-sinpo">—</div><div class="sc-sub">sin orden compra</div></div>
      <div class="sc c-pur"><div class="sc-lbl">Aprobadas</div><div class="sc-val" id="ap-aprobadas">—</div><div class="sc-sub">firmadas</div></div>
    </div>
    <div class="card" style="margin-bottom:22px">
      <div class="card-title">Facturas AP</div>
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Factura</th><th>Proveedor</th><th>Tipo</th><th>Total</th><th>Cuenta</th><th>Matching</th><th>Aprobación</th></tr></thead>
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
      <label for="drr-file-input">📂 Subir DRR (.xlsm)</label>
      <input type="file" id="drr-file-input" accept=".xlsm" style="display:none" onchange="uploadDRR(this)">
      <span class="drr-status" id="drr-status">Sin archivo cargado</span>
    </div>

    <!-- KPI Metrics -->
    <div class="drr-metrics" id="drr-metrics">
      <div class="empty"><div class="ei">📊</div><p>Sube un archivo DRR para ver las métricas.</p></div>
    </div>

    <!-- Days grid -->
    <div class="card" style="margin-bottom:22px">
      <div class="card-title">Trial Balance — Estado Diario</div>
      <div class="drr-days" id="drr-days"></div>
    </div>

    <!-- Alerts -->
    <div class="card">
      <div class="card-title">Alertas DRR</div>
      <div class="drr-alerts" id="drr-alerts">
        <div class="empty"><p>Sin alertas.</p></div>
      </div>
    </div>

  </div><!-- /panel-drr -->

  <!-- PANEL BANCO -->
  <div id="panel-banco" class="panel">
    <div class="stats" id="banco-stats">
      <div class="sc hl c-blu"><div class="sc-lbl">Movimientos</div><div class="sc-val" id="bk-total">—</div><div class="sc-sub">del extracto</div></div>
      <div class="sc c-grn"><div class="sc-lbl">Conciliados</div><div class="sc-val" id="bk-conc">—</div><div class="sc-sub">con factura</div></div>
      <div class="sc c-ora"><div class="sc-lbl">Pendientes</div><div class="sc-val" id="bk-pend">—</div><div class="sc-sub" id="bk-imp-pend">—</div></div>
      <div class="sc c-red"><div class="sc-lbl">Diferencias</div><div class="sc-val" id="bk-diff">—</div><div class="sc-sub">importe no cuadra</div></div>
    </div>
    <div class="card">
      <div class="card-title">Alertas Bancarias</div>
      <div id="bk-alertas"><div class="empty"><p>Cargando...</p></div></div>
    </div>
    <div style="margin-top:16px">
      <a href="/conciliacion/" class="btn-run" style="text-decoration:none;display:inline-flex;font-size:13px;padding:10px 20px">🏦 Ver conciliación completa</a>
    </div>
  </div><!-- /panel-banco -->

  <!-- PANEL NOTIFICACIONES -->
  <div id="panel-notif" class="panel">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
      <div>
        <span style="font-size:1.1rem;font-weight:700">Historial de Notificaciones</span>
        <span id="notif-count" style="font-size:.8rem;color:var(--dim);margin-left:8px"></span>
      </div>
      <button class="btn-run" id="btn-send-notif" onclick="enviarNotificaciones()" style="font-size:12px;padding:8px 16px">
        🔔 Enviar notificaciones pendientes
      </button>
    </div>
    <div class="card">
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Fecha</th><th>Tipo</th><th>Asunto</th><th>Destinatario</th><th>Estado</th></tr></thead>
          <tbody id="notif-tbody"><tr><td colspan="5" class="empty"><p>Sin notificaciones.</p></td></tr></tbody>
        </table>
      </div>
    </div>
  </div><!-- /panel-notif -->

  <!-- PANEL F&B -->
  <div id="panel-integraciones" class="panel">
    <h2 style="font-size:18px;font-weight:700;margin-bottom:20px">⚙️ Integraciones Externas</h2>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:24px;margin-bottom:24px">
      <h3 style="margin:0 0 16px 0;font-size:16px;font-weight:600">Canales disponibles</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px">
        
        <div style="background:#0f1117;border:1px solid #2e3248;border-radius:8px;padding:16px">
          <div style="font-size:28px;margin-bottom:8px">💬</div>
          <h4 style="margin:0 0 8px 0;font-size:14px;font-weight:600">Slack</h4>
          <p style="margin:0 0 12px 0;font-size:12px;color:#8892a4">Recibe alertas críticas y reportes en Slack</p>
          <div style="background:#1c1f2e;padding:8px;border-radius:6px;margin-bottom:12px;font-size:11px;color:#e05252;font-family:monospace">
            Estado: No configurado
          </div>
          <div style="font-size:11px;color:#666;line-height:1.4">
            <strong>Para habilitar:</strong><br>
            1. Crear Slack App en api.slack.com<br>
            2. Obtener Webhook URL<br>
            3. Agregar a .env: SLACK_WEBHOOK_URL=...
          </div>
        </div>

        <div style="background:#0f1117;border:1px solid #2e3248;border-radius:8px;padding:16px">
          <div style="font-size:28px;margin-bottom:8px">📱</div>
          <h4 style="margin:0 0 8px 0;font-size:14px;font-weight:600">WhatsApp</h4>
          <p style="margin:0 0 12px 0;font-size:12px;color:#8892a4">Envía alertas vía WhatsApp con Twilio</p>
          <div style="background:#1c1f2e;padding:8px;border-radius:6px;margin-bottom:12px;font-size:11px;color:#e05252;font-family:monospace">
            Estado: No configurado
          </div>
          <div style="font-size:11px;color:#666;line-height:1.4">
            <strong>Para habilitar:</strong><br>
            1. Crear cuenta en twilio.com<br>
            2. Obtener credenciales<br>
            3. Agregar a .env: TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN
          </div>
        </div>

        <div style="background:#0f1117;border:1px solid #2e3248;border-radius:8px;padding:16px">
          <div style="font-size:28px;margin-bottom:8px">📧</div>
          <h4 style="margin:0 0 8px 0;font-size:14px;font-weight:600">Email (SMTP)</h4>
          <p style="margin:0 0 12px 0;font-size:12px;color:#8892a4">Reportes y alertas por correo</p>
          <div style="background:#1c1f2e;padding:8px;border-radius:6px;margin-bottom:12px;font-size:11px;color:#1db954;font-family:monospace">
            ✓ Configurado
          </div>
          <div style="font-size:11px;color:#1db954;line-height:1.4">
            <strong>Status:</strong> SMTP configurado y funcionando<br>
            Reportes enviados automáticamente cada día
          </div>
        </div>

      </div>
    </div>

    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:24px">
      <h3 style="margin:0 0 16px 0;font-size:16px;font-weight:600">Configuración de .env</h3>
      <div style="background:#0f1117;padding:16px;border-radius:8px;font-family:monospace;font-size:12px;color:#8892a4;line-height:1.6;overflow-x:auto">
# Slack<br>
SLACK_WEBHOOK_URL=https://hooks.slack.com/services/YOUR/WEBHOOK/URL<br>
<br>
# WhatsApp (Twilio)<br>
TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxxxx<br>
TWILIO_AUTH_TOKEN=your_auth_token<br>
TWILIO_WHATSAPP_NUMBER=+1234567890<br>
<br>
# Email (ya configurado)<br>
SMTP_SERVER=smtp.gmail.com<br>
SMTP_PORT=587<br>
SMTP_USER=your_email@gmail.com<br>
SMTP_PASSWORD=your_app_password
      </div>
    </div>
  </div><!-- /panel-integraciones -->

  <div id="panel-fb" class="panel">
    <div id="fb-tab-content"><div class="empty"><p>Cargando F&amp;B...</p></div></div>
  </div><!-- /panel-fb -->

  <!-- PANEL AR REAL -->
  <div id="panel-ar_real" class="panel">
    <h2 style="font-size:18px;font-weight:700;margin-bottom:20px">🏢 AR Real — Grupos Corporativos</h2>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px;margin-bottom:20px">
      <h3 style="font-size:14px;margin-bottom:12px;color:#8892a4">Procesar Facturas de Grupos Corporativos</h3>
      <button onclick="procesarARReal()" id="btn-ar-real" style="background:#1a73e8;color:white;border:none;padding:10px 20px;border-radius:8px;cursor:pointer;font-weight:600">▶️ Procesar Archivos</button>
      <div id="ar-real-log" style="background:#0f1117;border:1px solid #2e3248;border-radius:8px;padding:12px;margin-top:16px;max-height:300px;overflow-y:auto;font-family:monospace;font-size:12px;color:#cdd6f4;min-height:60px"></div>
    </div>
    <div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px">
      <h3 style="font-size:14px;margin-bottom:12px;color:#8892a4">Último Reporte Generado</h3>
      <div id="ar-real-status">Sin reportes generados aún</div>
    </div>
  </div><!-- /panel-ar_real -->

  <!-- PANEL MULTI-HOTEL -->
  <div id="panel-calipolis" class="panel">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
    <h2 style="font-size:18px;font-weight:700;margin:0">🏩 Calipolis Hotels Group</h2>
    <a href="/api/exportar/calipolis" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px">⬇️ Descargar Excel</a>
  </div>
  
  <div id="cal-kpis" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:24px"></div>
  
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px" id="cal-hoteles"></div>
  
  <div id="cal-detail" style="display:none;margin-top:24px;background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px;"></div>
  </div><!-- /panel-calipolis -->

  <div id="panel-multi_hotel" class="panel">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px"><h2 style="font-size:18px;font-weight:700;margin:0">🏨 Multi-Hotel Dashboard</h2><a href="/api/exportar/multihotel" style="background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:8px;cursor:pointer;font-weight:600;text-decoration:none;font-size:13px">⬇️ Descargar Excel</a></div>
      <select id="grupo-filter" onchange="loadMultiHotel()" style="background:#1c1f2e;color:#cdd6f4;border:1px solid #2e3248;border-radius:8px;padding:8px 12px;font-size:13px;cursor:pointer">
        <option value="">Todos los grupos</option>
      </select>
    </div>

    <!-- KPIs Consolidados -->
    <div id="mh-kpis" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px"></div>

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
  <span>Pregunta a Yve</span>
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
    <button class="sug" onclick="askSug(this)">¿Cuánto llevamos facturado?</button>
    <button class="sug" onclick="askSug(this)">¿Qué facturas tienen discrepancias?</button>
    <button class="sug" onclick="askSug(this)">¿Qué proveedor tiene más errores?</button>
    <button class="sug" onclick="askSug(this)">¿Cuánto podemos reclamar a Booking?</button>
  </div>
  <div id="chat-input-row">
    <textarea id="chat-input" rows="1" placeholder="Pregunta sobre el estado financiero del hotel…"
      onkeydown="chatKeydown(event)" oninput="autoResize(this)"></textarea>
    <button id="chat-send" onclick="sendChat()">➤</button>
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
    // 1. Cargar y renderizar stats primero (independiente de facturas)
    const sr = await fetch('/api/stats');
    const stats = await sr.json();
    renderStats(stats);
    try { renderChart(stats.chart); } catch(ec) { console.warn('Chart no disponible:', ec); }

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
      'Actualizado · ' + (stats.total || 0) + ' factura' + (stats.total !== 1 ? 's' : '') + ' cargada' + (stats.total !== 1 ? 's' : '');
  } catch(e) {
    console.error('Error en loadAll:', e);
    document.getElementById('status-txt').textContent = 'Error al cargar datos';
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
  if (tab === 'ar_real') cargarStatusARReal();
  if (tab === 'calipolis') loadCalipolis();
  if (tab === 'integraciones') loadIntegraciones();
  if (tab === 'multi_hotel') loadMultiHotel();
}
async function loadFBTab() {
  var cont = document.getElementById('fb-tab-content');
  if (!cont || cont.dataset.loaded) return;
  cont.dataset.loaded = '1';
  try {
    var res = await fetch('/fb/api/resultados');
    var data = await res.json();
    if (!data.ok) { cont.innerHTML = '<div class="empty"><p>Pulsa Ejecutar para generar datos F&B.</p><button onclick="runFB()" style="margin-top:12px;background:#1a73e8;color:white;border:none;padding:8px 18px;border-radius:8px;cursor:pointer">Ejecutar</button></div>'; return; }
    var r2 = data.resumen;
    var html = '<h2 style="font-size:18px;font-weight:700;margin-bottom:20px">F&B Cost Control</h2>';
    html += '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px">';
    html += '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px"><div style="font-size:12px;color:#8892a4">Ventas F&B</div><div style="font-size:22px;font-weight:700;color:#1a73e8">' + r2.total_ventas.toLocaleString('es-ES') + ' €</div></div>';
    html += '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px"><div style="font-size:12px;color:#8892a4">FC Teórico</div><div style="font-size:22px;font-weight:700;color:#1db954">' + r2.fc_teorico_pct + '%</div></div>';
    html += '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px"><div style="font-size:12px;color:#8892a4">FC Real</div><div style="font-size:22px;font-weight:700;color:' + (r2.alerta ? '#e05252' : '#ff9800') + '">' + r2.fc_real_pct + '%</div></div>';
    html += '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:20px"><div style="font-size:12px;color:#8892a4">Mermas</div><div style="font-size:22px;font-weight:700;color:#e05252">' + r2.coste_mermas.toLocaleString('es-ES') + ' €</div></div>';
    html += '</div>';
    cont.innerHTML = html;
  } catch(e) { cont.innerHTML = '<div class="empty"><p>Error F&B: ' + e.message + '</p></div>'; }
}
function runFB() {
  var es = new EventSource('/fb/api/ejecutar');
  document.getElementById('fb-tab-content').innerHTML = '<div class="empty"><p>Ejecutando...</p></div>';
  es.onmessage = function(ev) {
    if (ev.data === 'FB_COMPLETO') { es.close(); var c = document.getElementById('fb-tab-content'); if(c){delete c.dataset.loaded; loadFBTab();} }
  };
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

function addMsg(role, text) {
  const msgs = document.getElementById('chat-msgs');
  const div  = document.createElement('div');
  div.className = `msg ${role}`;
  div.textContent = text;
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

  const thinkDiv = addMsg('bot', 'Consultando datos del hotel…');
  thinkDiv.classList.add('thinking');

  try {
    const resp = await fetch('/api/chat', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ messages: chatHistory }),
    });
    const data = await resp.json();
    const reply = data.reply || '⚠️ Sin respuesta del servidor.';

    thinkDiv.textContent = reply;
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
    const logDiv = document.getElementById('ar-real-log');
    logDiv.innerHTML = '<div style="color:#666;">Conectando...</div>';
    
    const eventSource = new EventSource('/api/procesar_ar_real');
    
    eventSource.onmessage = function(event) {
        const logLine = document.createElement('div');
        logLine.textContent = event.data;
        logLine.style.padding = '2px 0';
        logDiv.appendChild(logLine);
        logDiv.scrollTop = logDiv.scrollHeight;
        
        if (event.data === 'AR_REAL_COMPLETO') {
            eventSource.close();
            setTimeout(cargarStatusARReal, 1000);
        }
    };
    
    eventSource.onerror = function(err) {
        console.error('Error AR Real:', err);
        const logLine = document.createElement('div');
        logLine.textContent = 'ERROR: Conexión perdida';
        logLine.style.color = 'red';
        logDiv.appendChild(logLine);
        eventSource.close();
    };
}

function cargarStatusARReal() {
    const statusDiv = document.getElementById('ar-real-status');
    if (!statusDiv) return;
    fetch('/api/ar_real_status')
        .then(r => r.json())
        .then(data => {
            if (data.reportes && data.reportes.length > 0) {
                const rep = data.reportes[0];
                let html = '<div style="background:#e8f5e9;border-left:4px solid #4caf50;padding:15px;margin:10px 0;">';
                html += '<p><strong>Último reporte:</strong> ' + rep.filename + '</p>';
                html += '<p><strong>Tamaño:</strong> ' + rep.size_kb + ' KB</p>';
                html += '<p><strong>Actualizado:</strong> ' + new Date(rep.timestamp).toLocaleString() + '</p>';
                html += '</div>';
                statusDiv.innerHTML = html;
            } else {
                statusDiv.innerHTML = '<p style="color:#999;">Sin reportes generados aún</p>';
            }
        })
        .catch(err => console.error('Error cargar status:', err));
}





// ═══════════════════════════════════════════════════════════════════
// MULTI-HOTEL DASHBOARD
// ═══════════════════════════════════════════════════════════════════
let _mh_loaded = false;

async function loadMultiHotel() {
  const grupoSelect = document.getElementById('grupo-filter');
  const grupo = grupoSelect ? grupoSelect.value : '';
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
    '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:16px">' +
    '<div style="font-size:11px;color:#8892a4;margin-bottom:4px">' + c.label + '</div>' +
    '<div style="font-size:20px;font-weight:700;color:' + c.color + '">' + c.value + '</div>' +
    '</div>'
  ).join('');
}

function renderMHStatus(kpis) {
  const cont = document.getElementById('mh-status');
  if (!cont) return;
  cont.innerHTML = 
    '<div style="background:#0d2818;border:1px solid #1db954;border-radius:12px;padding:16px;display:flex;align-items:center;gap:12px">' +
      '<span style="font-size:28px">OK</span>' +
      '<div><div style="font-size:24px;font-weight:700;color:#1db954">' + kpis.hoteles_ok + '</div>' +
      '<div style="font-size:12px;color:#8892a4">Hoteles OK</div></div></div>' +
    '<div style="background:#2d2410;border:1px solid #ff9800;border-radius:12px;padding:16px;display:flex;align-items:center;gap:12px">' +
      '<span style="font-size:28px">!</span>' +
      '<div><div style="font-size:24px;font-weight:700;color:#ff9800">' + kpis.hoteles_warning + '</div>' +
      '<div style="font-size:12px;color:#8892a4">Warnings</div></div></div>' +
    '<div style="background:#2d1010;border:1px solid #e05252;border-radius:12px;padding:16px;display:flex;align-items:center;gap:12px">' +
      '<span style="font-size:28px">X</span>' +
      '<div><div style="font-size:24px;font-weight:700;color:#e05252">' + kpis.hoteles_criticos + '</div>' +
      '<div style="font-size:12px;color:#8892a4">Criticos</div></div></div>';
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
        '<div style="background:linear-gradient(135deg,#0d2818,#1a4a2e);border:1px solid #1db954;border-radius:10px;padding:16px"><div style="font-size:11px;color:#8892a4">Revenue MTD</div><div style="font-size:26px;font-weight:700;color:#1db954">€' + (h.revenue_mtd/1000000).toFixed(2) + 'M</div></div>' +
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
async function loadCalipolis() {
  try {
    const res = await fetch('/api/calipolis/kpis');
    const data = await res.json();
    renderCalipolisKpis(data.consolidado);
    renderCalipolisHoteles(data.hoteles);
  } catch(e) {
    console.error('Error Calipolis:', e);
  }
}

function renderCalipolisKpis(kpis) {
  const cont = document.getElementById('cal-kpis');
  if (!cont) return;
  const cards = [
    {label: 'Hoteles', value: kpis.num_hoteles, color: '#1a73e8'},
    {label: 'Habitaciones', value: kpis.total_rooms, color: '#1a73e8'},
    {label: 'Revenue MTD', value: '€' + (kpis.total_revenue_mtd/1000000).toFixed(2) + 'M', color: '#1db954'},
    {label: 'Ocupación Avg', value: kpis.avg_ocupacion + '%', color: '#1db954'},
    {label: 'ADR Avg', value: '€' + kpis.avg_adr.toFixed(0), color: '#ff9800'},
    {label: 'RevPAR Avg', value: '€' + kpis.avg_revpar.toFixed(0), color: '#ff9800'},
    {label: 'GOP', value: '€' + (kpis.total_gop/1000).toFixed(0) + 'K', color: '#1db954'},
    {label: 'GOP%', value: kpis.avg_gop_pct + '%', color: '1db954'}
  ];
  
  cont.innerHTML = cards.map(c => 
    '<div style="background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:16px">' +
    '<div style="font-size:11px;color:#8892a4;margin-bottom:4px">' + c.label + '</div>' +
    '<div style="font-size:20px;font-weight:700;color:' + c.color + '">' + c.value + '</div>' +
    '</div>'
  ).join('');
}

function renderCalipolisHoteles(hoteles) {
  const cont = document.getElementById('cal-hoteles');
  if (!cont) return;
  
  cont.innerHTML = hoteles.map(h => {
    const statusColor = h.status === 'ok' ? '#1db954' : h.status === 'warning' ? '#ff9800' : '#e05252';
    const div = document.createElement('div');
    div.style.cssText = 'background:#1c1f2e;border:1px solid #2e3248;border-radius:12px;padding:16px;cursor:pointer;transition:all 0.2s';
    div.addEventListener('mouseover', () => div.style.background = 'rgba(26,115,232,0.08)');
    div.addEventListener('mouseout', () => div.style.background = '#1c1f2e');
    div.addEventListener('click', () => abrirDetalleCalipolis(h.id));
    
    div.innerHTML = 
      '<div style="font-weight:600;font-size:14px">' + h.nombre + '</div>' +
      '<div style="font-size:11px;color:#8892a4;margin-bottom:12px">' + h.categoria + ' • ' + h.habitaciones + ' rooms</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:12px">' +
        '<div><span style="color:#8892a4">Occ:</span> <strong>' + h.ocupacion + '%</strong></div>' +
        '<div><span style="color:#8892a4">RevPAR:</span> <strong>€' + h.revpar.toFixed(0) + '</strong></div>' +
        '<div><span style="color:#8892a4">Revenue:</span> <strong>€' + (h.total_ingresos/1000).toFixed(0) + 'K</strong></div>' +
        '<div><span style="color:' + statusColor + '">GOP: ' + h.gop_pct + '%</span></div>' +
      '</div>' +
      '<div style="margin-top:8px;padding-top:8px;border-top:1px solid #2e3248;font-size:11px;color:#8892a4">' +
        'AP Pend: ' + h.ap_pendientes + ' • AR: ' + h.ar_pendientes +
      '</div>';
    
    return div;
  }).map(d => d.outerHTML).join('');
}


// ═══════════════════════════════════════════════════════════════════
// INTEGRACIONES
// ═══════════════════════════════════════════════════════════════════
async function loadIntegraciones() {
  try {
    const res = await fetch('/api/integraciones/status');
    const data = await res.json();
    
    const cont = document.getElementById('integraciones-status');
    if (!cont) return;
    
    const integraciones = [
      {nombre: 'Slack', key: 'slack_disponible', emoji: '💬'},
      {nombre: 'WhatsApp', key: 'whatsapp_disponible', emoji: '📱'},
      {nombre: 'Email', key: 'email_disponible', emoji: '📧'}
    ];
    
    cont.innerHTML = integraciones.map(i => {
      const enabled = data[i.key];
      const color = enabled ? '#1db954' : '#e05252';
      const status = enabled ? 'CONFIGURADO' : 'No configurado';
      return '<div style="background:#1c1f2e;border:2px solid ' + color + ';border-radius:12px;padding:16px">' +
        '<div style="font-size:24px;margin-bottom:8px">' + i.emoji + '</div>' +
        '<div style="font-weight:600;margin-bottom:4px">' + i.nombre + '</div>' +
        '<div style="font-size:12px;color:' + color + '">' + status + '</div></div>';
    }).join('');
  } catch(e) {
    console.error('Error cargando integraciones:', e);
  }
}


function closeHotelModal() {
  const modal = document.getElementById('hotel-modal');
  if (modal) modal.remove();
}

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
