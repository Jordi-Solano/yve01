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
from onboarding import onboarding_bp as config_bp
from panel_admin import bp as admin_bp
from app_aprobacion import bp as aprob_ar_bp
from app_aprobacion_ap import bp as aprob_ap_bp
from app_conciliacion import bp as concil_bp
from tab_fb_dashboard import fb_bp
from tab_ar_real import ar_real_bp
from oracle_export_dryrun import oracle_export_bp
from pricing import pricing_bp
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
from legal import legal_bp
from signup import signup_bp
from about import about_bp
from exportador_pdf import pdf_bp
for _bp in (auth_bp, config_bp, admin_bp, aprob_ar_bp, aprob_ap_bp, concil_bp, fb_bp, ar_real_bp, multi_hotel_bp, exportador_bp, calipolis_bp, demo_bp, demo_sim_bp, calipolis_analisis_bp, reportes_pdf_bp, blog_bp, billing_bp, signup_bp, about_bp, pdf_bp, legal_bp):
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

def _audit(accion, detalle="", usuario=None):
    """Registra una acción en el audit log."""
    import json as _json
    from datetime import datetime as _dt
    try:
        ruta = os.path.join(os.path.dirname(__file__), "datos-referencia", "audit_log.json")
        entries = []
        if os.path.exists(ruta):
            with open(ruta) as f: entries = _json.load(f)
        entries.append({
            "ts":      _dt.now().isoformat(timespec="seconds"),
            "accion":  accion,
            "detalle": detalle[:200],
            "usuario": usuario or "sistema",
        })
        entries = entries[-500:]  # Keep last 500
        with open(ruta, "w") as f: _json.dump(entries, f, ensure_ascii=False)
    except Exception:
        pass

# ── Simple rate limiting ─────────────────────────────────────────────────
import time as _time
_rate_buckets: dict = {}

def _rate_limit(key: str, max_req: int = 30, window: int = 60) -> bool:
    """Returns True if request should be blocked (rate limit exceeded)."""
    now = _time.time()
    bucket = _rate_buckets.get(key, [])
    bucket = [t for t in bucket if now - t < window]
    if len(bucket) >= max_req:
        return True
    bucket.append(now)
    _rate_buckets[key] = bucket
    return False

@app.before_request
def check_rate_limit():
    """Apply rate limiting to API endpoints."""
    from flask import request as _req
    if _req.path.startswith('/api/procesar') or _req.path.startswith('/api/procesar_ap'):
        ip = _req.headers.get('X-Forwarded-For', _req.remote_addr or 'unknown').split(',')[0].strip()
        if _rate_limit(f"process:{ip}", max_req=10, window=60):
            from flask import jsonify as _j
            return _j({"error": "Rate limit exceeded. Wait 60 seconds."}), 429

@app.route("/robots.txt")
def robots_txt():
    return Response("""User-agent: *
Allow: /
Allow: /blog
Allow: /about
Allow: /casos
Allow: /terminos
Allow: /privacidad
Disallow: /api/
Disallow: /admin/
Disallow: /onboarding
Sitemap: https://yve01.onrender.com/sitemap.xml
""", mimetype="text/plain")

@app.route("/sitemap.xml")
def sitemap_xml():
    from datetime import date
    today = date.today().isoformat()
    BLOG_SLUGS = [
        "software-gestion-financiera-hoteles-espana",
        "automatizar-cuentas-pagar-hotel",
        "revenue-management-hoteles-independientes",
        "food-cost-hotel-restaurante-como-calcularlo",
        "conciliacion-bancaria-hotel-guia-completa",
        "out-of-balance-drr-como-detectarlo",
        "integracion-oracle-fusion-hotel",
        "revenue-management-hoteles-pequenos",
        "gestion-cuentas-cobrar-hotel-grupos",
        "ap-proveedores-hotel-como-optimizar",
    ]
    urls = [
        ("https://yve01.onrender.com/",          "1.0",  "daily"),
        ("https://yve01.onrender.com/about",      "0.8",  "monthly"),
        ("https://yve01.onrender.com/casos",      "0.8",  "monthly"),
        ("https://yve01.onrender.com/blog",       "0.9",  "weekly"),
        ("https://yve01.onrender.com/signup",     "0.9",  "monthly"),
        ("https://yve01.onrender.com/terminos",   "0.3",  "yearly"),
        ("https://yve01.onrender.com/privacidad", "0.3",  "yearly"),
    ] + [(f"https://yve01.onrender.com/blog/{s}", "0.7", "monthly") for s in BLOG_SLUGS]
    xml_parts = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    for url, pri, freq in urls:
        xml_parts.append('  <url><loc>' + url + '</loc><lastmod>' + today + '</lastmod><changefreq>' + freq + '</changefreq><priority>' + pri + '</priority></url>')
    xml_parts.append('</urlset>')
    xml = chr(10).join(xml_parts)
    return Response(xml, mimetype="application/xml")

@app.route("/health")
@app.route('/api/oracle/status')
def api_oracle_status():
    mode='real' if os.environ.get('ORACLE_BASE_URL') else 'simulation'
    return jsonify({'mode':mode,'ok':True})

@app.route('/api/procesar_batch_stream')
@login_required
def api_procesar_batch_stream():
    """SSE stream — procesamiento PARALELO, 3 a la vez, timeout 45s."""
    import json as _json, concurrent.futures as _cf, threading as _th
    archivos_str = request.args.get('archivos', '[]')
    try:
        archivos = _json.loads(archivos_str)
    except Exception:
        archivos = []

    log = _load_proc_log()
    from datetime import datetime as _dt2
    results_queue = _th.Queue()

    def _mark(fname, result='OK'):
        log[fname] = {'fecha': _dt2.now().strftime('%Y-%m-%d %H:%M'), 'resultado': result}
        _save_proc_log(log)

    def _procesar_uno(fname):
        fpath = os.path.join(ENTRADA_DIR, fname)
        if not os.path.exists(fpath):
            results_queue.put(('error', fname, 'no encontrado'))
            return False, False
        try:
            size_mb = os.path.getsize(fpath) / (1024*1024)
            if size_mb > 30:
                results_queue.put(('error', fname, f'PDF demasiado grande ({size_mb:.0f}MB)'))
                _mark(fname, 'ERR:TOO_LARGE')
                return False, False
        except: pass

        tipo = _detect_file_type(fname)
        results_queue.put(('start', fname, tipo))

        try:
            import subprocess as _sp
            if tipo == 'DRR':
                import shutil as _sh
                _sh.copy2(fpath, os.path.join(BASE_DIR, 'reportes', 'drr_upload.xlsm'))
                results_queue.put(('ok', fname, 'DRR copiado'))
                _mark(fname, 'DRR_OK')
                return False, False

            is_ar = tipo == 'AR' or (tipo == 'AR_o_AP' and any(
                x in fname.lower() for x in ['booking','expedia','hotels','despegar','ota']
            ))
            cmd = ['python3', 'lector_ota.py' if is_ar else 'lector_facturas_ap.py', '--file', fpath]
            r = _sp.run(cmd, capture_output=True, text=True, cwd=BASE_DIR, timeout=45)
            ok = r.returncode == 0
            msg = 'OK' if ok else (r.stderr[:60] or r.stdout[:60] or 'error')
            results_queue.put(('ok' if ok else 'error', fname, msg))
            _mark(fname, ('AR_OK' if is_ar else 'AP_OK') if ok else f'ERR:{msg[:30]}')
            return (ok and is_ar), (ok and not is_ar)
        except Exception as e:
            msg = f'TIMEOUT (45s)' if 'Timeout' in type(e).__name__ else str(e)[:60]
            results_queue.put(('error', fname, msg))
            _mark(fname, f'CRASH:{msg[:30]}')
            return False, False

    def generar():
        global _pipeline_running
        with _pipeline_lock:
            if _pipeline_running:
                yield 'data: ℹ Ya hay un proceso activo — espera\n\n'
                yield 'data: PIPELINE_CON_ERRORES\n\n'
                return
            _pipeline_running = True
        try:
            if not archivos:
                yield 'data: ✗ No se especificaron archivos\n\n'
                yield 'data: PIPELINE_CON_ERRORES\n\n'
                return

            total = len(archivos)
            yield f'data: >> {total} archivo(s) — procesando 3 en paralelo...\n\n'
            has_ar = False; has_ap = False

            for i in range(0, total, 3):
                batch = archivos[i:i+3]
                if len(batch) > 1:
                    yield f'data: >> Lote {i//3+1}: {len(batch)} archivos...\n\n'
                with _cf.ThreadPoolExecutor(max_workers=3) as ex:
                    futures = {ex.submit(_procesar_uno, f): f for f in batch}
                    done = 0
                    while done < len(batch):
                        try:
                            kind, fname, msg = results_queue.get(timeout=50)
                            if kind == 'start':
                                yield f'data: >> {fname} ({msg})...\n\n'
                            else:
                                yield f'data: {"✓" if kind=="ok" else "✗"} {fname}: {msg}\n\n'
                            done = sum(1 for f in futures if f.done())
                        except:
                            done = sum(1 for f in futures if f.done())
                    for f in futures:
                        try:
                            ar_ok, ap_ok = f.result()
                            if ar_ok: has_ar = True
                            if ap_ok: has_ap = True
                        except: pass
                yield f'data: ✓ Lote {i//3+1} completado ({min(i+3,total)}/{total})\n\n'

            if has_ar:
                yield 'data: >> Verificando comisiones OTA...\n\n'
                try:
                    import subprocess as _sp2
                    _sp2.run(['python3','verificador_comisiones.py'], cwd=BASE_DIR, timeout=30, capture_output=True)
                    _sp2.run(['python3','detector_doble_imposicion.py'], cwd=BASE_DIR, timeout=30, capture_output=True)
                    yield 'data: ✓ Verificación completada\n\n'
                except: pass

            yield 'data: \n\n'
            yield 'data: PIPELINE_COMPLETO\n\n'
        except Exception as e:
            yield f'data: ERROR: {str(e)[:200]}\n\n'
            yield 'data: PIPELINE_CON_ERRORES\n\n'
        finally:
            _pipeline_running = False

    return Response(stream_with_context(generar()), mimetype='text/event-stream',
                    headers={'Cache-Control':'no-cache','X-Accel-Buffering':'no'})

@app.route("/api/health")
def health():
    """Health check — keeps Render awake and provides system status."""
    from datetime import datetime
    import glob as _glob
    return jsonify({
        "status":    "ok",
        "app":       "Yve.01",
        "version":   "1.0.0-beta",
        "timestamp": datetime.now().isoformat(),
        "reports":   len(_glob.glob(os.path.join(REPORTES_DIR, "*.xlsx"))),
        "uptime":    open("/proc/uptime").read().split()[0] + "s" if os.path.exists("/proc/uptime") else "n/a",
    })

@app.route("/api/ar/enviar_recordatorios_di", methods=["POST"])
@login_required
def api_enviar_recordatorios_di():
    """Envía recordatorios de certificado DI a OTAs que lo tienen pendiente."""
    try:
        from tab_ar_real import get_ar_real_data
    except ImportError:
        pass
    # Use existing notificaciones data
    try:
        from notificaciones import enviar_email, _email_html
        from verificador_comisiones import verificar_ciclo
        # Get pending DI invoices
        facturas = _cargar_facturas_verificadas()
        pendientes_di = [f for f in facturas if f.get('estado_di') == 'FALTA_CERTIFICADO_DI']
        enviados = 0
        notif_email = os.environ.get('NOTIF_EMAIL', '')
        if notif_email and pendientes_di:
            cuerpo = _email_html(
                f'Certificados DI pendientes — {len(pendientes_di)} facturas',
                [f"Factura {f.get('numero_factura','?')} — {f.get('ota','?')} — {f.get('importe_bruto','?')}€" for f in pendientes_di[:10]],
                color='#f59e0b'
            )
            if enviar_email(notif_email, f'Yve.01 — {len(pendientes_di)} certificados DI pendientes', cuerpo, 'ar_di_reminder'):
                enviados = len(pendientes_di)
        return jsonify({'ok': True, 'enviados': enviados, 'pendientes': len(pendientes_di)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]})

@app.route("/api/test_smtp", methods=["POST"])
@login_required
def api_test_smtp():
    """Test SMTP configuration."""
    from notificaciones import test_smtp
    result = test_smtp()
    return jsonify(result)

@app.route("/api/ap/aprobar_lote", methods=["POST"])
@login_required
def api_ap_aprobar_lote():
    """Aprueba en lote facturas AP con Match OK."""
    data = request.get_json(force=True, silent=True) or {}
    facturas_nums = data.get("facturas", [])
    if not facturas_nums:
        return jsonify({"ok": False, "error": "No se especificaron facturas"}), 400
    try:
        # Update matching reports
        import glob as _g
        hits = _g.glob(os.path.join(REPORTES_DIR, "matching_*.xlsx"))
        aprobadas = 0
        for ruta in hits:
            df = pd.read_excel(ruta)
            if "numero_factura" in df.columns and "aprobacion" in df.columns:
                mask = (df["numero_factura"].astype(str).isin(facturas_nums)) &                        (df["estado"].astype(str).str.contains("MATCH_3WAY_OK", na=False))
                aprobadas += int(mask.sum())
                df.loc[mask, "aprobacion"] = "APROBADA"
                df.to_excel(ruta, index=False)
        _audit("AP_LOTE_APROBADO", f"{aprobadas} facturas aprobadas", session.get("username", "sistema"))
        return jsonify({"ok": True, "aprobadas": aprobadas})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500

@app.route("/api/test_oracle", methods=["POST"])
@login_required
def api_test_oracle():
    """Test Oracle connection using oracle_auth.test_connection()."""
    try:
        from oracle_auth import test_connection
        result = test_connection()
        return jsonify({
            "ok":      result["ok"],
            "message": result["message"],
            "mode":    result.get("mode", ""),
            "ledgers": result.get("ledgers", []),
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]})

@app.route("/api/ar_real/pdf/<numero>")
@login_required
def api_invoice_pdf(numero):
    """Descarga PDF de una factura corporativa."""
    try:
        from exportador_pdf import export_invoice_pdf
        buf, err = export_invoice_pdf(numero)
        if err:
            return jsonify({"error": err}), 404
        from flask import send_file
        return send_file(buf, mimetype="application/pdf",
                        as_attachment=True,
                        download_name=f"factura_{numero}.pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/test_telegram", methods=["POST"])
@login_required
def api_test_telegram():
    """Test Telegram Bot connection."""
    import os as _os
    token = _os.environ.get("TELEGRAM_BOT_TOKEN", "")
    chat_id = _os.environ.get("TELEGRAM_CHAT_ID", "")
    if not token or not chat_id:
        return jsonify({"ok": False, "message": "TELEGRAM_BOT_TOKEN o TELEGRAM_CHAT_ID no configurados en Render"})
    try:
        from notificaciones import enviar_telegram
        ok = enviar_telegram("✅ Test de conexión desde Yve.01 — funcionando correctamente")
        return jsonify({"ok": ok, "message": "✓ Telegram Bot conectado y mensaje enviado" if ok else "✗ Error enviando mensaje"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]})

@app.route("/api/test_stripe", methods=["POST"])
@login_required
def api_test_stripe():
    """Test Stripe configuration."""
    import os as _os
    key = _os.environ.get("STRIPE_SECRET_KEY", "")
    if not key:
        return jsonify({"ok": False, "message": "STRIPE_SECRET_KEY no configurado"})
    try:
        import stripe as _stripe
        _stripe.api_key = key
        acct = _stripe.Account.retrieve()
        mode = "🧪 TEST" if key.startswith("sk_test_") else "🔴 LIVE"
        return jsonify({"ok": True, "message": f"Stripe {mode} conectado — cuenta: {acct.get('email', acct['id'])}"})
    except ImportError:
        return jsonify({"ok": False, "error": "stripe no instalado (pip install stripe)"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]})

@app.route("/api/health")
def api_health():
    """System health check — all components status."""
    import glob as _g
    health = {
        'status': 'ok',
        'components': {}
    }
    # Data files
    has_drr = bool(_g.glob(os.path.join(REPORTES_DIR, 'drr_procesado_*.xlsx')))
    has_ar  = bool(_g.glob(os.path.join(REPORTES_DIR, 'doble_imposicion_*.xlsx')) or
                   _g.glob(os.path.join(REPORTES_DIR, 'verificacion_*.xlsx')))
    has_ap  = bool(_g.glob(os.path.join(REPORTES_DIR, 'matching_*.xlsx')))
    has_cfg = os.path.exists(os.path.join(BASE_DIR, 'datos-referencia', 'hotel_config.json'))
    
    # Oracle mode
    oracle_mode = 'real' if os.environ.get('ORACLE_BASE_URL') else 'simulation'
    
    # SMTP configured
    smtp_ok = bool(os.environ.get('SMTP_USER') and os.environ.get('SMTP_PASSWORD'))
    
    # Telegram configured
    tg_ok = bool(os.environ.get('TELEGRAM_BOT_TOKEN') and os.environ.get('TELEGRAM_CHAT_ID'))
    
    health['components'] = {
        'drr':      {'ok': has_drr,   'msg': 'DRR procesado' if has_drr else 'Sin DRR'},
        'ar':       {'ok': has_ar,    'msg': 'AR datos OK' if has_ar else 'Sin datos AR'},
        'ap':       {'ok': has_ap,    'msg': 'AP datos OK' if has_ap else 'Sin datos AP'},
        'config':   {'ok': has_cfg,   'msg': 'Hotel configurado' if has_cfg else 'Sin configuración'},
        'oracle':   {'ok': True,      'msg': f'Oracle {oracle_mode}', 'mode': oracle_mode},
        'smtp':     {'ok': smtp_ok,   'msg': 'SMTP configurado' if smtp_ok else 'SMTP no configurado'},
        'telegram': {'ok': tg_ok,     'msg': 'Telegram OK' if tg_ok else 'Telegram no configurado'},
    }
    
    # Overall status
    critical = ['ar', 'ap', 'config']
    if any(not health['components'][k]['ok'] for k in critical):
        health['status'] = 'degraded'
    
    return jsonify(health)

@app.route("/api/demo_stats")
def api_demo_stats():
    """Returns curated demo statistics for the demo mode."""
    return jsonify({
        "total": 20,
        "importe_total": 109440.05,
        "importe_reclamable": 3847.50,
        "correctas": 16,
        "discrepancias": 3,
        "di_pendientes": 1,
        "sin_accion": 0,
        "aprobadas": 12,
        "rechazadas": 4,
        "pendientes_firma": 6,
        "chart": {
            "labels": ["Booking.com", "Expedia", "HotelBeds", "Hotusa"],
            "data": [45230, 38670, 15890, 9650],
        }
    })

@app.route("/api/stats")
def api_stats():
    df, meta = cargar_datos()
    stats = calcular_stats(df)
    stats["chart"] = calcular_chart(df)
    stats["meta"]  = meta
    # Enrich with AP pending count
    try:
        import glob as _g
        ap_hits = _g.glob(os.path.join(REPORTES_DIR, "matching_*.xlsx"))
        ap_pend = 0
        for ruta in ap_hits:
            df_ap = pd.read_excel(ruta)
            if "aprobacion" in df_ap.columns:
                ap_pend += int((df_ap["aprobacion"].astype(str) == "PENDIENTE").sum())
        stats["pendientes_firma"] = ap_pend
    except: pass
    return jsonify(stats)

@app.route("/api/facturas")
def api_facturas():
    df, _ = cargar_datos()
    return jsonify(df_a_lista(df))

# ── File Upload & Processing Batch ──────────────────────────────────────────

ENTRADA_DIR   = os.path.join(BASE_DIR, 'facturas-entrada')
PROCESADAS_DIR = os.path.join(BASE_DIR, 'facturas-procesadas')
PROC_LOG_PATH  = os.path.join(BASE_DIR, 'datos-referencia', 'archivos_procesados.json')
os.makedirs(ENTRADA_DIR,   exist_ok=True)
os.makedirs(PROCESADAS_DIR, exist_ok=True)

def _load_proc_log():
    """Load the processed-files log."""
    if os.path.exists(PROC_LOG_PATH):
        try:
            with open(PROC_LOG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_proc_log(log):
    with open(PROC_LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(log, f, indent=2, ensure_ascii=False)

def _detect_file_type(filename):
    """Detect what section a file belongs to."""
    name = filename.lower()
    if name.endswith('.xlsm') or 'drr' in name or 'daily' in name or 'revenue' in name:
        return 'DRR'
    if any(ota in name for ota in ['booking', 'expedia', 'hotelbeds', 'hotusa', 'ota']):
        return 'AR'
    if name.endswith('.pdf'):
        # Check if it looks like OTA or AP
        return 'AR_o_AP'  # Will need user to clarify or auto-detect
    return 'AP'

@app.route('/api/archivos_estado', methods=['GET'])
@login_required
def api_archivos_estado():
    """List files in facturas-entrada with processed status."""
    log = _load_proc_log()
    files = []
    if os.path.exists(ENTRADA_DIR):
        for fname in sorted(os.listdir(ENTRADA_DIR)):
            fpath = os.path.join(ENTRADA_DIR, fname)
            if os.path.isfile(fpath) and not fname.startswith('.'):
                fsize = os.path.getsize(fpath)
                proc_info = log.get(fname, None)
                files.append({
                    'nombre': fname,
                    'tipo': _detect_file_type(fname),
                    'tamano': fsize,
                    'tamano_str': f'{fsize/1024:.0f}KB' if fsize < 1024*1024 else f'{fsize/1024/1024:.1f}MB',
                    'procesado': proc_info is not None,
                    'fecha_proceso': proc_info.get('fecha') if proc_info else None,
                    'resultado': proc_info.get('resultado') if proc_info else None,
                })
    return jsonify({'ok': True, 'files': files, 'total': len(files),
                    'pendientes': sum(1 for f in files if not f['procesado'])})

@app.route('/api/upload_facturas', methods=['POST'])
@login_required
def api_upload_facturas():
    """Upload one or more invoice files to facturas-entrada/."""
    if 'files' not in request.files:
        return jsonify({'ok': False, 'error': 'No files provided'}), 400
    
    log = _load_proc_log()
    results = []
    files = request.files.getlist('files')
    
    for file in files:
        if not file.filename:
            continue
        fname = os.path.basename(file.filename)
        fpath = os.path.join(ENTRADA_DIR, fname)
        already_exists = os.path.exists(fpath)
        already_processed = fname in log
        
        if already_processed:
            results.append({'nombre': fname, 'status': 'ya_procesado', 
                           'fecha': log[fname].get('fecha')})
            continue
        
        file.save(fpath)
        results.append({'nombre': fname, 'status': 'subido', 
                        'tipo': _detect_file_type(fname)})
    
    return jsonify({'ok': True, 'results': results,
                    'subidos': sum(1 for r in results if r['status'] == 'subido'),
                    'ya_procesados': sum(1 for r in results if r['status'] == 'ya_procesado')})

@app.route('/api/procesar_batch', methods=['POST'])
@login_required
def api_procesar_batch():
    """Process only new (unprocessed) files from facturas-entrada/."""
    global _pipeline_running
    data = request.get_json(force=True, silent=True) or {}
    solo_nuevos = data.get('solo_nuevos', True)  # default: skip already processed
    tipos = data.get('tipos', ['AR', 'AP', 'DRR', 'AR_o_AP'])  # which types to process
    archivos_seleccionados = data.get('archivos', [])  # specific filenames to process
    
    if _pipeline_running:
        return jsonify({'ok': False, 'error': 'Ya hay un proceso en ejecución'}), 409
    
    log = _load_proc_log()
    from datetime import datetime as _dt
    
    def _mark_processed(fname, resultado='OK'):
        log[fname] = {'fecha': _dt.now().strftime('%Y-%m-%d %H:%M'), 'resultado': resultado}
        _save_proc_log(log)
    
    def generar():
        global _pipeline_running
        with _pipeline_lock:
            if _pipeline_running:
                yield 'data: Ya hay un proceso — espera\n\n'
                return
            _pipeline_running = True
        
        try:
            yield 'data: >> Iniciando procesamiento batch\n\n'
            
            # Get files to process
            if archivos_seleccionados:
                candidatos = archivos_seleccionados
            else:
                candidatos = sorted(os.listdir(ENTRADA_DIR)) if os.path.exists(ENTRADA_DIR) else []
            
            a_procesar = []
            a_saltar = []
            for fname in candidatos:
                if fname.startswith('.'): continue
                if solo_nuevos and fname in log:
                    a_saltar.append(fname)
                else:
                    tipo = _detect_file_type(fname)
                    if tipo in tipos or 'AR_o_AP' in tipos:
                        a_procesar.append((fname, tipo))
            
            if a_saltar:
                yield f'data: ℹ Saltando {len(a_saltar)} archivos ya procesados\n\n'
            
            if not a_procesar:
                yield 'data: ✓ No hay archivos nuevos que procesar\n\n'
                yield 'data: PIPELINE_COMPLETO\n\n'
                return
            
            yield f'data: >> Procesando {len(a_procesar)} archivo(s) nuevos...\n\n'
            
            has_ar = False; has_ap = False; has_drr = False
            
            # Process each file
            for fname, tipo in a_procesar:
                fpath = os.path.join(ENTRADA_DIR, fname)
                if not os.path.exists(fpath):
                    yield f'data: ✗ {fname}: archivo no encontrado\n\n'
                    continue
                
                yield f'data: >> Procesando {fname} ({tipo})...\n\n'
                
                try:
                    if tipo == 'DRR':
                        # Copy to expected DRR location and trigger reader
                        import shutil
                        drr_dest = os.path.join(BASE_DIR, 'reportes', 'drr_upload.xlsm')
                        shutil.copy2(fpath, drr_dest)
                        yield f'data: ✓ DRR {fname}: copiado para procesamiento\n\n'
                        _mark_processed(fname, 'DRR_OK')
                        has_drr = True
                    
                    elif tipo in ('AR', 'AR_o_AP', 'AP'):
                        # Run OTA reader for AR, AP reader for others
                        if tipo == 'AR' or (tipo == 'AR_o_AP' and any(x in fname.lower() for x in ['booking','expedia','ota'])):
                            import subprocess
                            result = subprocess.run(['python3', 'lector_ota.py', '--file', fpath], 
                                capture_output=True, text=True, cwd=BASE_DIR, timeout=120)
                            if result.returncode == 0:
                                yield f'data: ✓ AR OTA {fname}: procesado\n\n'
                                _mark_processed(fname, 'AR_OK')
                                has_ar = True
                            else:
                                yield f'data: ✗ AR {fname}: {result.stderr[:100]}\n\n'
                                _mark_processed(fname, f'ERROR: {result.stderr[:50]}')
                        else:
                            result = subprocess.run(['python3', 'lector_facturas_ap.py', '--file', fpath],
                                capture_output=True, text=True, cwd=BASE_DIR, timeout=120)
                            if result.returncode == 0:
                                yield f'data: ✓ AP {fname}: procesado\n\n'
                                _mark_processed(fname, 'AP_OK')
                                has_ap = True
                            else:
                                yield f'data: ✗ AP {fname}: {result.stderr[:100]}\n\n'
                                _mark_processed(fname, f'ERROR: {result.stderr[:50]}')
                except Exception as e:
                    yield f'data: ✗ {fname}: {str(e)[:100]}\n\n'
            
            # Run verification passes if we processed AR files
            if has_ar:
                yield 'data: >> Verificando comisiones OTA...\n\n'
                try:
                    import subprocess
                    result = subprocess.run(['python3', 'verificador_comisiones.py'],
                        capture_output=True, text=True, cwd=BASE_DIR, timeout=60)
                    yield f'data: ✓ Verificación comisiones completada\n\n'
                    result2 = subprocess.run(['python3', 'detector_doble_imposicion.py'],
                        capture_output=True, text=True, cwd=BASE_DIR, timeout=60)
                    yield f'data: ✓ Análisis doble imposición completado\n\n'
                except Exception as e:
                    yield f'data: ✗ Verificación: {str(e)[:80]}\n\n'
            
            yield 'data: \n\n'
            yield 'data: ✅ Batch completado\n\n'
            yield 'data: PIPELINE_COMPLETO\n\n'
        
        except Exception as e:
            yield f'data: ERROR CRÍTICO: {str(e)[:200]}\n\n'
            yield 'data: PIPELINE_CON_ERRORES\n\n'
        finally:
            _pipeline_running = False
    
    return Response(stream_with_context(generar()), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

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
                    "today":    str(row.iloc[1]) if pd.notna(row.iloc[1]) else "N/D",
                    "mtd":      str(row.iloc[2]) if pd.notna(row.iloc[2]) else "N/D",
                    "forecast": str(row.iloc[3]) if pd.notna(row.iloc[3]) else "N/D",
                    "budget":   str(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else "N/D",
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
        labels = [fechas_map.get(d, str(d))[-5:] for d in dias]  # MM-DD
        return jsonify({
            "dias":     dias,
            "labels":   labels,
            "fechas":   [fechas_map.get(d, str(d)) for d in dias],
            "revenue":  [round(float(daily_rev.get(d, 0)), 0) for d in dias],
            "expenses": [round(float(daily_exp.get(d, 0)), 0) for d in dias],
            "oob":      [d in oob_dias for d in dias],
            "oob_count": len(oob_dias),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/stats_drr")
def api_stats_drr():
    ruta = _cargar_drr_procesado()
    if not ruta:
        return jsonify(None)
    stats = _leer_drr_stats(ruta)
    if stats and ruta:
        import os as _os2
        from datetime import datetime as _dt2, date as _d2
        mtime = _os2.path.getmtime(ruta)
        upload_date = _dt2.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')
        days_ago = (date.today() - _dt2.fromtimestamp(mtime).date()).days
        stats['last_upload'] = upload_date
        stats['days_ago'] = days_ago
    return jsonify(stats)


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

@app.route("/api/conciliar", methods=["POST"])
@login_required
def api_run_conciliacion():
    """Ejecuta la conciliación bancaria automática."""
    try:
        from conciliacion_bancaria import conciliar
        import glob as _g
        # Load extracto bancario
        extracto_ruta = os.path.join(BASE_DIR, "datos-referencia", "extracto_banco.xlsx")
        if not os.path.exists(extracto_ruta):
            return jsonify({"ok": False, "error": "No se encontró extracto_banco.xlsx"}), 404
        df_extracto = pd.read_excel(extracto_ruta)
        extracto = df_extracto.to_dict('records')
        # Load facturas AP para matching
        facturas = []
        hits = _g.glob(os.path.join(REPORTES_DIR, "matching_*.xlsx"))
        for ruta in hits:
            df_f = pd.read_excel(ruta)
            if "numero_factura" in df_f.columns:
                for _, row in df_f.iterrows():
                    facturas.append({
                        "numero": str(row.get("numero_factura","")),
                        "importe": abs(float(row.get("importe_con_iva", 0) or 0)),
                        "tipo_mov": "CARGO",
                        "proveedor": str(row.get("proveedor",""))[:30],
                    })
        # Run reconciliation
        result = conciliar(extracto, facturas)
        # Save result
        if result:
            df_result = pd.DataFrame(result)
            from datetime import date
            out_ruta = os.path.join(REPORTES_DIR, f"conciliacion_{date.today().strftime('%Y%m%d')}.xlsx")
            df_result.to_excel(out_ruta, index=False)
            conciliados = sum(1 for r in result if r.get("estado") == "CONCILIADO")
            pendientes  = sum(1 for r in result if r.get("estado") != "CONCILIADO")
            _audit("CONCILIACION_RUN", f"{conciliados} conciliados, {pendientes} pendientes")
            return jsonify({"ok": True, "total": len(result), "conciliados": conciliados,
                           "pendientes": pendientes, "archivo": os.path.basename(out_ruta)})
        return jsonify({"ok": False, "error": "Sin resultados de conciliación"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:300]}), 500

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
            # Support both flat key and nested 'hotel.nombre'
            name = cfg.get("hotel_tag") or cfg.get("hotel_nombre") or cfg.get("hotel", {}).get("nombre", "")
            return name
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
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5,viewport-fit=cover">
<meta id="csrf-token-meta" name="csrf-token" content="">
<link rel="manifest" href="/static/manifest.json">
<meta name="theme-color" content="#3b82f6">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Yve.01">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%230f172a'/%3E%3Crect width='32' height='32' rx='8' fill='url(%23g)'/%3E%3Cdefs%3E%3ClinearGradient id='g' x1='0' y1='0' x2='32' y2='32' gradientUnits='userSpaceOnUse'%3E%3Cstop offset='0' stop-color='%233b82f6' stop-opacity='.15'/%3E%3Cstop offset='1' stop-color='%23a78bfa' stop-opacity='.08'/%3E%3C/linearGradient%3E%3C/defs%3E%3Ccircle cx='16' cy='10' r='3' fill='%233b82f6'/%3E%3Cpath d='M10 6 L16 16 L22 6' stroke='%233b82f6' stroke-width='2.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3Cline x1='16' y1='16' x2='16' y2='26' stroke='%2360a5fa' stroke-width='2.5' stroke-linecap='round'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Dashboard</title>
<script async src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
/* Light mode */
body.light-mode{--bg:#f8fafc;--s1:#fff;--s2:#e2e8f0;--s3:#cbd5e1;--tx:#0f172a;--mut:#475569;--dim:#64748b;--acc:#2563eb;--acc2:#1d4ed8;--grn:#16a34a;--red:#dc2626;--ora:#ea580c;--pur:#7c3aed}
body.light-mode .nav{background:rgba(248,250,252,.9);border-bottom-color:#e2e8f0}
body.light-mode .tab-btn{color:#475569}
body.light-mode .tab-btn.active{color:#2563eb}
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
/* ── Mobile lite mode ─────────────────────────────────────────────── */
body.mobile-lite .panel > .card:not(.lite-visible),
body.mobile-lite .panel > div.tbl-wrap,
body.mobile-lite .panel > div[style*="overflow-x:auto"],
body.mobile-lite .panel > table,
body.mobile-lite #activity,
body.mobile-lite #ar-chart-wrap,
body.mobile-lite #ap-chart-wrap,
body.mobile-lite #banco-alerts-card,
body.mobile-lite .hide-lite { display: none !important; }
body.mobile-lite .mobile-lite-hint { display: block !important; }
body.mobile-lite #mobile-kpi-bar { display: block !important; }
body.mobile-lite .panel > .stats { grid-template-columns: repeat(2,1fr) !important; gap: 8px !important; }
body.mobile-lite .sc-val { font-size: 22px !important; }
body.mobile-lite .panel { padding: 12px !important; }
body.mobile-lite #exec-summary { font-size: 11px; padding: 6px 12px; }
#mobile-lite-toggle {
  position: fixed; bottom: 70px; left: 12px; z-index: 700;
  background: var(--s1); border: 1px solid var(--s2); border-radius: 20px;
  padding: 6px 14px; font-size: 11px; color: var(--mut); cursor: pointer;
  font-weight: 600; display: none; box-shadow: 0 2px 12px rgba(0,0,0,.3);
}
@media(max-width:768px){ #mobile-lite-toggle { display: block; } }

@media(max-width:900px){
  #ar-real-grid{grid-template-columns:1fr}
  .metrics{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:768px){
  /* Nav */
  .nav{padding:0 10px;gap:4px;height:52px}
  .logo-name{font-size:15px}
  .logo-tag{display:none}
  .nav-right{gap:4px}
  .btn-ref{font-size:11px;padding:5px 8px}
  .btn-run{font-size:12px;padding:7px 12px}
  #btn-install-pwa{display:none}
  /* Tabs */
  .tabs{gap:0;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;padding:0 4px}
  .tabs::-webkit-scrollbar{display:none}
  .tab-btn{white-space:nowrap;flex-shrink:0;font-size:11px;padding:8px 10px;min-width:auto}
  /* Panels */
  .panel{padding:14px 12px}
  .card{padding:12px}
  /* Stats */
  .stats{grid-template-columns:repeat(2,1fr);gap:8px}
  .sc-val{font-size:20px}
  .sc-lbl{font-size:9px}
  /* AR Real */
  #ar-real-grid{grid-template-columns:1fr}
  /* DRR metrics */
  .drr-mc{padding:10px}
  /* Tables: horizontal scroll */
  .tbl-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  table{min-width:500px}
  /* Hide secondary columns on mobile */
  .hide-mobile{display:none}
  /* Calipolis cards */
  #cal-grid{grid-template-columns:1fr!important}
  /* Multi-hotel table */
  #mh-kpis .sc-lbl{font-size:9px}
  /* Status bar */
  .status-bar{font-size:10px;padding:5px 10px}
  /* Demo banner */
  #demo-banner{font-size:10px;padding:4px 8px}
  /* Back to top */
  #back-top{bottom:70px;right:12px;width:32px;height:32px}
}
@media(max-width:600px){
  .stats{grid-template-columns:repeat(2,1fr)}
  /* Tabs: icon-only on very small screens */
  .tab-btn .tab-txt{display:none}
  .tab-btn{font-size:16px;padding:8px 10px}
  /* Modal fix */
  #modal-emitir > div, #checklist-modal > div, #changelog-modal > div{
    max-width:100%;width:calc(100% - 20px);margin:10px;max-height:calc(100vh - 40px)
  }
  /* AR Real two-col → one col */
  #ef-cliente{font-size:12px}
  /* Grids to 1 col */
  .metrics{grid-template-columns:1fr}
}
@media(max-width:600px){
  .stats{grid-template-columns:repeat(2,1fr)}
  .tab-btn span{display:none}
  .tab-btn{font-size:16px;padding:8px}
}
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

/* Tooltips */
[data-tip]{position:relative}
[data-tip]::after{content:attr(data-tip);position:absolute;bottom:calc(100% + 8px);left:50%;transform:translateX(-50%);background:#1e293b;color:#f1f5f9;padding:7px 12px;border-radius:8px;font-size:11px;line-height:1.4;max-width:240px;white-space:normal;text-align:center;border:1px solid #334155;pointer-events:none;opacity:0;transition:opacity .15s;z-index:5000;box-shadow:0 4px 12px rgba(0,0,0,.3)}
[data-tip]:hover::after{opacity:1}
.sr-item{padding:14px 18px;cursor:pointer;border-bottom:1px solid var(--s2);display:flex;align-items:center;gap:12px;transition:background .15s}.sr-item:hover{background:var(--s2)}
@media print {
  .nav, .tabs, .status-bar, #top-bar, .fab-ask, #search-overlay,
  .btn-run, .btn-ref, .dropdown, #demo-banner { display: none !important; }
  body { background: #fff !important; color: #000 !important; }
  .panel { box-shadow: none !important; border: 1px solid #ccc !important; }
  .drr-mc { break-inside: avoid; }
}

* { -webkit-tap-highlight-color: transparent; }
[data-tour-active] {
  position: relative !important;
  z-index: 9950 !important;
}
#tour-box { font-family: Inter, system-ui, sans-serif; }
@keyframes tourBoxIn {
  from { opacity:0; transform:scale(.95) translateY(10px); }
  to   { opacity:1; transform:scale(1)  translateY(0); }
}
#tour-spotlight-canvas { pointer-events: none; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.6} }
button, a { touch-action: manipulation; }

</style>
</head>
<body>

<nav class="nav" id="app-header">
  <div class="logo">
    <div class="logo-dot"></div>
    <span class="logo-name">Yve<span style="color:var(--acc2)">.01</span></span>
    <span class="logo-tag">__HOTEL_TAG__</span>
  <span style="font-size:9px;color:#334155;margin-left:4px;font-weight:500">v1.0</span>
  </div>
  <div class="nav-mid"></div>
  <div id="demo-banner" style="display:none;position:fixed;top:0;left:0;right:0;z-index:8000;background:linear-gradient(90deg,#f59e0b,#d97706);color:#000;text-align:center;padding:6px 16px;font-size:13px;font-weight:700;letter-spacing:.3px">
    🎭 MODO DEMO · Grupo Calipolis Hotels · <span style="font-weight:400">Datos reales de las 3 propiedades en Sitges</span>
    <button onclick="toggleDemoMode()" style="margin-left:16px;background:rgba(0,0,0,.2);border:none;padding:3px 10px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:700">✕ Salir</button>
  </div>
  <div class="nav-right">
    <span class="pill" id="date-pill">—</span>
  <button id="btn-install-pwa" onclick="if(_deferredInstall){_deferredInstall.prompt();}" style="display:none;background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.2);color:#22c55e;padding:4px 10px;border-radius:8px;font-size:11px;cursor:pointer">📲 Instalar</button>
  <button id="btn-atajos" onclick="toggleAtajos()" class="hide-mobile" style="background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.2);color:#60a5fa;padding:5px 12px;border-radius:8px;font-size:12px;cursor:pointer;font-weight:500" title="Ver atajos (?)">⌨ Atajos</button>
    <span class="pill" style="color:var(--acc2)">👤 __USER_NAME__</span>

    <div class="dropdown">
      <button class="btn-ref" onclick="toggleMenu('reportes-menu')" title="Descargar reportes" data-i18n="nav.reportes">📄 Reportes</button>
      <div id="reportes-menu" class="menu">
        <div class="menu-head">Reportes PDF</div>
        <a href="/api/reportes/diario" class="menu-item">📄 Diario</a>
        <a href="/api/reportes/semanal" class="menu-item">📊 Semanal</a>
        <a href="/api/reportes/mensual" class="menu-item">📈 Mensual</a>
        <div class="menu-sep"></div>
        <button class="menu-item" onclick="startTour()">🎯 Tour guiado</button>
        <button class="menu-item" onclick="showSetupChecklist()">✅ Checklist setup</button>
        <button class="menu-item" onclick="showChangelog()">🆕 Novedades</button>
        <button class="menu-item" onclick="toggleLightMode()" id="btn-theme">🌙 Modo oscuro</button>
        <div class="menu-head">Ejecutivos</div>
        <a href="/api/reportes/ejecutivo.pdf" class="menu-item">🎯 Ejecutivo PDF</a>
        <a href="/api/reportes/consolidado.xlsx" class="menu-item">📊 Consolidado Excel</a>
      </div>
    </div>

    <button class="btn-ref" onclick="loadAll()" title="Actualizar datos" data-i18n="nav.actualizar">↻ Actualizar</button>

    <button class="btn-run" id="btn-run" onclick="openUploadModal()">
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

  <!-- Top loading bar -->
  <div id="top-bar" style="position:fixed;top:0;left:0;height:2px;background:linear-gradient(90deg,#3b82f6,#a78bfa);z-index:9999;transition:width .3s ease;width:0"></div>

  <!-- Barra estado -->
  <div class="status-bar">
    <div class="status-dot"></div>
    <span id="status-txt" data-i18n="status.cargando">Cargando datos...</span>
  </div>

  <!-- Alertas del día panel -->
  <div id="daily-alerts-panel" style="display:none;background:linear-gradient(135deg,rgba(239,68,68,.06),rgba(245,158,11,.04));border-bottom:1px solid rgba(239,68,68,.15);padding:10px 16px">
    <div style="display:flex;align-items:center;justify-content:space-between">
      <div style="font-size:12px;font-weight:700;color:var(--ora)">🔔 ALERTAS ACTIVAS</div>
      <button onclick="document.getElementById('daily-alerts-panel').style.display='none'" style="background:none;border:none;color:var(--dim);font-size:14px;cursor:pointer">×</button>
    </div>
    <div id="daily-alerts-list" style="display:flex;gap:12px;margin-top:8px;flex-wrap:wrap"></div>
  </div>

  <!-- Mobile KPI Quick-View -->
  <div id="mobile-kpi-bar" style="display:none;background:var(--s1);border-bottom:1px solid var(--s2)">
    <!-- KPI scroll row -->
    <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;padding:10px 14px">
      <div style="display:flex;gap:0;min-width:max-content;align-items:stretch">
        <div style="text-align:center;min-width:80px;padding:0 12px;border-right:1px solid var(--s2)">
          <div style="font-size:20px;font-weight:900;color:var(--acc2);line-height:1.2" id="mkpi-ar">—</div>
          <div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-top:3px">AR pend</div>
        </div>
        <div style="text-align:center;min-width:80px;padding:0 12px;border-right:1px solid var(--s2)">
          <div style="font-size:20px;font-weight:900;color:var(--ora);line-height:1.2" id="mkpi-ap">—</div>
          <div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-top:3px">AP firma</div>
        </div>
        <div style="text-align:center;min-width:80px;padding:0 12px;border-right:1px solid var(--s2)">
          <div style="font-size:20px;font-weight:900;color:var(--grn);line-height:1.2" id="mkpi-occ">—</div>
          <div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-top:3px">Ocup%</div>
        </div>
        <div style="text-align:center;min-width:80px;padding:0 12px">
          <div style="font-size:20px;font-weight:900;color:var(--pur);line-height:1.2" id="mkpi-gop">—</div>
          <div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-top:3px">GOP%</div>
        </div>
      </div>
    </div>
    <!-- Last update timestamp -->
    <div id="mkpi-updated" style="font-size:9px;color:var(--dim);text-align:right;padding:3px 14px 6px;display:none"></div>
  </div>

  <!-- Executive summary bar (dynamic) -->
  <div id="exec-summary" style="display:none;padding:8px 16px;background:rgba(59,130,246,.05);border-bottom:1px solid rgba(59,130,246,.1);font-size:12px;color:var(--mut);text-align:center">
    <span id="exec-txt"></span>
  </div>

  <!-- TABS -->
  <div class="tabs">
    <button class="tab active" id="tab-ar" onclick="switchTab('ar',this)" data-i18n="tab.ar">📥 AR — OTAs</button>
    <button class="tab" id="tab-ap" onclick="switchTab('ap',this)" data-i18n="tab.ap">📦 AP — Proveedores</button>
    <button class="tab" id="tab-drr" onclick="switchTab('drr',this)" data-i18n="tab.drr">📊 DRR</button>
    <button class="tab" id="tab-banco" onclick="switchTab('banco',this)" data-i18n="tab.banco">🏦 Banco</button>
    <button class="tab" id="tab-notif" onclick="switchTab('notif',this)" data-i18n="tab.notif">🔔 Notificaciones</button>
    <button class="tab" onclick="switchTab('fb',this)" id="tab-fb" data-i18n="tab.fb">🍽️ F&amp;B Cost</button>
    <button class="tab" onclick="switchTab('ar_real',this)" id="tab-ar-real" data-i18n="tab.arreal">🏢 AR Real</button>
    <button class="tab" onclick="switchTab('calipolis',this)" id="tab-calipolis" data-i18n="tab.calipolis">🏩 Calipolis</button>
    <button class="tab" onclick="switchTab('multi_hotel',this)" id="tab-multi-hotel" data-i18n="tab.multihotel">🏨 Multi-Hotel</button>
  </div>

  <div id="panel-ar" class="panel active">
    <div class="mobile-lite-hint" style="display:none;background:rgba(59,130,246,.08);border:1px solid rgba(59,130,246,.2);border-radius:8px;padding:8px 12px;margin-bottom:12px;font-size:12px;color:var(--acc2)">📱 Vista resumida · <button onclick="toggleMobileLite()" style="background:none;border:none;color:var(--acc2);text-decoration:underline;cursor:pointer;font-size:12px">Ver tabla completa →</button></div>
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ar" class="btn-ref" style="text-decoration:none">⬇️ Excel</a><a href="/api/exportar/ar/pdf" class="btn-ref" style="text-decoration:none">📄 PDF</a><a href="/aprobaciones-ar/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AR">📲 Aprobar facturas AR</a></div>
  <!-- STATS -->
  <div class="stats" id="ar-stats-section">
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
      <div id="activity" class="hide-lite">
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
            <th style="width:28px"><input type="checkbox" id="ar-select-all" onclick="toggleSelectAll(this,'ar-row-cb')" style="cursor:pointer;accent-color:var(--acc)"></th>
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
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ap" style="background:#1a73e8;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px" data-i18n="btn.downloadExcel" data-i18n="btn.downloadExcel">⬇️ Descargar Excel</a><a href="/aprobaciones-ap/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AP" data-i18n="btn.aprobarAP">📲 Aprobar facturas AP</a>
        <button class="btn-ref" onclick="aprobarMatchOK()" style="font-size:12px" title="Aprueba automáticamente todas las facturas con 3-way match correcto">✅ Aprobar Match OK</button>
        <button class="btn-ref" onclick="filtrarAPPorEstado()" id="btn-filter-ap" style="font-size:12px">🔍 Filtrar</button>
        <select id="ap-estado-filter" onchange="filtrarAPPorEstado(this.value)" style="background:var(--s1);border:1px solid var(--s2);color:var(--tx);padding:6px 10px;border-radius:8px;font-size:12px;cursor:pointer">
          <option value="">Todos</option>
          <option value="PENDIENTE">Pendientes</option>
          <option value="MATCH_3WAY_OK">Match OK</option>
          <option value="DISCREPANCIA_PO">Discrepancias</option>
          <option value="ALERTA_CONSUMO">Alertas</option>
        </select></div>
    <div class="stats" id="stats-ap-grid">
      <div class="sc hl c-blu"><div class="sc-lbl" data-i18n="ap.totalLabel">Total Facturas AP</div><div class="sc-val" id="ap-total" data-tip="Facturas AP registradas este ciclo">—</div><div class="sc-sub" data-i18n="ap.proveedores">proveedores</div></div>
      <div class="sc"><div class="sc-lbl" data-i18n="ap.importe">Importe Total</div><div class="sc-val" id="ap-importe" data-tip="Importe bruto total de facturas AP" style="font-size:18px;letter-spacing:-.5px">—</div><div class="sc-sub">EUR</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-i18n="ap.matchOk">Matches OK</div><div class="sc-val" id="ap-matches" data-tip="Facturas con 3-way match correcto">—</div><div class="sc-sub" data-i18n="ap.fbOtras">F&B + OTRAS</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-i18n="sc.discrepancias">Discrepancias</div><div class="sc-val" id="ap-disc" data-tip="Facturas con discrepancia vs PO">—</div><div class="sc-sub" data-i18n="ap.vsPo">vs PO</div></div>
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
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
      <div class="drr-upload">
        <label for="drr-file-input" class="btn-run" style="cursor:pointer;font-size:13px" data-i18n="btn.uploadDrr">📂 Subir DRR (.xlsm)</label>
        <input type="file" id="drr-file-input" accept=".xlsm,.xlsx" style="display:none" onchange="uploadDRR(this)">
        <span class="drr-status" id="drr-status" style="font-size:12px;color:var(--mut);margin-left:10px">Sin archivo cargado</span>
      </div>
      <div style="display:flex;gap:8px">
        <div id="drr-oob-badge" style="display:none;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.25);color:var(--red);border-radius:8px;padding:5px 12px;font-size:12px;font-weight:700"></div>
        <a href="/api/exportar/drr" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="btn.downloadExcel">⬇️ Excel</a>
        <a href="/api/exportar/drr/pdf" class="btn-ref" style="text-decoration:none;font-size:12px">📄 PDF</a>
      </div>
    </div>

    <!-- KPI Metrics -->
    <div class="drr-metrics" id="drr-metrics" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:12px">
      <div class="empty" id="drr-drop-zone" 
    style="border:2px dashed var(--s3);border-radius:12px;padding:32px;cursor:pointer;transition:.2s"
    ondragover="event.preventDefault();this.style.borderColor='var(--acc)';this.style.background='rgba(59,130,246,.05)'"
    ondragleave="this.style.borderColor='var(--s3)';this.style.background=''"
    ondrop="event.preventDefault();this.style.borderColor='var(--s3)';this.style.background='';uploadDRR({files:event.dataTransfer.files})"
    onclick="document.getElementById('drr-file-input').click()">
  <div class="ei">📊</div>
  <p style="margin-bottom:6px">Arrastra tu DRR aquí o</p>
  <p style="font-size:12px;color:var(--acc2);font-weight:600">haz clic para seleccionar (.xlsm/.xlsx)</p>
</div>
    </div>
    <div id="drr-budget-bar" style="display:none;background:rgba(15,23,42,.5);border-radius:10px;padding:10px 14px;margin-bottom:16px;border:1px solid var(--s2)"></div>

    <!-- Revenue Chart -->
    <div class="card" style="margin-bottom:22px" id="drr-chart-card" style="display:none">
      <div class="card-title" data-i18n="card.revDiario">Revenue Diario</div>
      <div class="drr-chart-wrap"><canvas id="drr-revenue-chart" class="hide-lite"></canvas></div>
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
    <div class="stats lite-visible" id="banco-stats">
      <div class="sc hl c-blu"><div class="sc-lbl" data-i18n="sc.movimientos">Movimientos</div><div class="sc-val" id="bk-total" data-tip="Movimientos totales en el extracto">—</div><div class="sc-sub" data-i18n="sc.delExtracto">del extracto</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-i18n="sc.conciliados">Conciliados</div><div class="sc-val" id="bk-conc" data-tip="Cruzados con factura en el sistema">—</div><div class="sc-sub" data-i18n="sc.conFactura">con factura</div></div>
      <div class="sc c-ora"><div class="sc-lbl" data-i18n="sc.pendientes">Pendientes</div><div class="sc-val" id="bk-pend" data-tip="Sin factura asociada — pendientes">—</div><div class="sc-sub" id="bk-imp-pend">—</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-i18n="sc.diferencias">Diferencias</div><div class="sc-val" id="bk-diff">—</div><div class="sc-sub" data-i18n="sc.importeNoCuadra">importe no cuadra</div></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.alertasBanco">Alertas Bancarias</div>
      <div id="bk-alertas"><div class="empty"><p>Cargando...</p></div></div>
    </div>
    <div style="margin-top:16px">
      <a href="/api/exportar/banco" class="btn-ref" style="text-decoration:none;font-size:12px">⬇️ Excel</a>
        <button class="btn-run" onclick="runConciliacion()" style="font-size:12px">⚡ Conciliar</button>
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
      <div style="display:flex;gap:8px">
        <button class="btn-ref" onclick="testNotification()" style="font-size:12px">🧪 Test</button>
        <button class="btn-run" id="btn-send-notif" onclick="enviarNotificaciones()" style="font-size:12px;padding:8px 16px">
        <span data-i18n="notif.enviar">🔔 Enviar notificaciones pendientes</span>
      </button>
      </div>
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

    <!-- Header AR Real -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
      <div>
        <h2 style="font-size:18px;font-weight:700;margin:0">🏢 AR Real — Grupos Corporativos
          <span data-tip="Gestión completa del ciclo de cobro: facturas emitidas, antigüedad, recordatorios y cobros" style="font-size:12px;color:var(--dim);cursor:help">❓</span>
        </h2>
        <div style="font-size:12px;color:var(--mut);margin-top:4px">Clientes de crédito · Facturación corporativa · Control de cobros</div>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn-ref" onclick="abrirEmitirFactura()" style="font-size:12px">📄 Nueva factura</button>
        <button class="btn-ref" onclick="loadARRealData()" style="font-size:12px">🔄 Actualizar</button>
        <a href="/api/exportar/ar_real" class="btn-ref" style="text-decoration:none;font-size:12px">⬇️ Excel</a>
        <a href="/aprobaciones-ar/" class="btn-run" style="text-decoration:none;font-size:12px;padding:8px 14px" data-i18n="btn.aprobarAR">📲 Aprobar AR</a>
      </div>
    </div>

    <!-- Stats KPIs -->
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:16px" id="ar-real-stats">
      <div class="sc hl c-ora"><div class="sc-lbl" data-tip="Facturas emitidas aún no cobradas">PENDIENTE COBRO</div><div class="sc-val" id="arp-pendiente">—</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-tip="Facturas con más de 60 días sin cobrar">VENCIDO >60d</div><div class="sc-val" id="arp-vencido">—</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-tip="Cobrado este mes">COBRADO MES</div><div class="sc-val" id="arp-cobrado">—</div></div>
      <div class="sc"><div class="sc-lbl">CLIENTES ACTIVOS</div><div class="sc-val" id="arp-nclientes">—</div></div>
    </div>

    <!-- Two-column layout: clients + invoices -->
    <div style="display:grid;grid-template-columns:1fr 1.6fr;gap:16px" id="ar-real-grid">

      <!-- Client list -->
      <div>
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);margin-bottom:10px">Clientes</div>
        <div id="ar-clientes-list" style="display:flex;flex-direction:column;gap:8px"></div>
      </div>

      <!-- Invoices -->
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)">Facturas <span id="ar-facturas-count" style="color:var(--dim)"></span></div>
          <select id="ar-filter-estado" onchange="filtrarFacturasAR(this.value)" style="background:var(--s1);border:1px solid var(--s2);color:var(--tx);padding:5px 10px;border-radius:8px;font-size:11px">
            <option value="">Todas</option>
            <option value="PENDIENTE_FACTURA">Pendiente emitir</option>
            <option value="FACTURADO">Facturadas</option>
            <option value="COBRADO">Cobradas</option>
          </select>
        </div>
        <!-- Aging bar -->
        <div id="ar-aging-bar" style="display:none;margin-bottom:12px"></div>
        <!-- Invoice table -->
        <div style="overflow-x:auto">
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="border-bottom:2px solid var(--s2)">
              <th style="text-align:left;padding:6px 8px;color:var(--mut);font-weight:600">Nº / Cliente</th>
              <th style="text-align:right;padding:6px 8px;color:var(--mut);font-weight:600">Importe</th>
              <th style="text-align:left;padding:6px 8px;color:var(--mut);font-weight:600">Estado</th>
              <th style="padding:6px 4px;color:var(--mut);font-weight:600">Acciones</th>
            </tr></thead>
            <tbody id="ar-facturas-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Emit invoice modal -->
    <div id="modal-emitir" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9000;align-items:center;justify-content:center">
      <div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;max-width:480px;width:90%;max-height:85vh;overflow-y:auto">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
          <h3 style="font-size:16px;font-weight:700;margin:0">📄 Nueva Factura Corporativa</h3>
          <button onclick="cerrarEmitirFactura()" style="background:none;border:none;color:var(--mut);font-size:20px;cursor:pointer">×</button>
        </div>
        <div style="display:flex;flex-direction:column;gap:12px">
          <div>
            <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">Cliente</label>
            <select id="ef-cliente" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px">
              <option value="">Seleccionar cliente...</option>
            </select>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
            <div>
              <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">Entrada</label>
              <input type="date" id="ef-entrada" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px">
            </div>
            <div>
              <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">Salida</label>
              <input type="date" id="ef-salida" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px">
            </div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px">
            <div>
              <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">Habitaciones</label>
              <input type="number" id="ef-hab" value="1" min="1" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px" oninput="calcularFactura()">
            </div>
            <div>
              <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">€/noche</label>
              <input type="number" id="ef-precio" value="186" step="0.01" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px" oninput="calcularFactura()">
            </div>
            <div>
              <label style="font-size:11px;color:var(--mut);font-weight:600;text-transform:uppercase">F&B</label>
              <input type="number" id="ef-fb" value="0" step="0.01" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:9px;border-radius:8px;font-size:13px;margin-top:4px" oninput="calcularFactura()">
            </div>
          </div>
          <div style="background:var(--bg);border-radius:10px;padding:12px;font-size:13px">
            <div style="display:flex;justify-content:space-between;color:var(--mut);margin-bottom:3px"><span>Habitaciones:</span><span id="ef-sub-hab">—</span></div>
            <div style="display:flex;justify-content:space-between;color:var(--mut);margin-bottom:3px"><span>F&B:</span><span id="ef-sub-fb">—</span></div>
            <div style="display:flex;justify-content:space-between;color:var(--mut);margin-bottom:6px"><span>IVA 10%:</span><span id="ef-iva">—</span></div>
            <div style="display:flex;justify-content:space-between;font-weight:800;font-size:16px;border-top:1px solid var(--s2);padding-top:8px"><span>TOTAL:</span><span id="ef-total" style="color:var(--acc2)">—</span></div>
          </div>
          <div id="ef-msg" style="font-size:12px;display:none"></div>
          <div style="display:flex;gap:10px">
            <button onclick="calcularFactura()" class="btn-ref" style="flex:1;font-size:13px">Calcular</button>
            <button onclick="emitirFactura()" class="btn-run" style="flex:2;font-size:13px">📄 Emitir</button>
          </div>
        </div>
      </div>
    </div>

    <div id="ar-real-status" style="display:none"></div>
  </div><!-- /panel-ar_real -->


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
  <div id="cal-kpis" style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px"></div>
  <div id="cal-insights"></div>

  <!-- Trend row: GOP y Facturas pendientes últimos 6 meses -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:22px">
    <div class="card">
      <div class="card-title" data-i18n="card.gopEvolucion">GOP% — evolución 6 meses</div>
      <div id="cal-tendencias" style="height:140px;position:relative"></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.apEvolucion">Facturas AP pendientes — evolución 6 meses</div>
      <div style="height:120px;position:relative"><div id="cal-ap-chart" style="height:140px;position:relative"></div></div>
    </div>
  </div>

  <!-- Hotel cards -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:22px" id="cal-hoteles"></div>

  <!-- Detail panel (hidden by default) -->
  <div id="cal-detail" style="display:none;margin-top:4px"></div>
  </div><!-- /panel-calipolis -->

  <div id="panel-multi_hotel" class="panel">

    <!-- Header -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;flex-wrap:wrap;gap:12px">
      <div>
        <h2 style="font-size:20px;font-weight:800;margin:0">🌍 Multi-Hotel Dashboard</h2>
        <div style="font-size:12px;color:var(--mut);margin-top:3px">Vista consolidada del grupo</div>
      </div>
      <a href="/api/exportar/multihotel" class="btn-ref" style="text-decoration:none;font-size:12px">⬇️ Excel</a>
    </div>

    <!-- KPI Cards 2x2 -->
    <div id="mh-kpis" style="margin-bottom:20px"></div>

    <!-- Smart Insights row -->
    <div id="mh-insights" style="margin-bottom:20px"></div>

    <!-- Trend charts 2 col -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">
      <div class="card">
        <div class="card-title" style="font-size:11px;letter-spacing:.5px;text-transform:uppercase">GOP% — 6 Month Trend</div>
        <div id="mh-gop-chart" style="height:140px;position:relative"></div>
      </div>
      <div class="card">
        <div class="card-title" style="font-size:11px;letter-spacing:.5px;text-transform:uppercase">Revenue — 6 Month Trend</div>
        <div id="mh-rev-chart" style="height:140px;position:relative"></div>
      </div>
    </div>

    <!-- Hotel cards grid -->
    <div id="mh-hotel-cards" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:16px"></div>

  </div><!-- /panel-multi_hotel -->

</div><!-- /main -->

<!-- MODAL PIPELINE -->
<!-- ── File Upload Modal ─────────────────────────────────────────────── -->
<div id="upload-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9000;align-items:center;justify-content:center">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:28px;width:min(600px,95vw);max-height:85vh;overflow-y:auto;position:relative">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px">
      <div>
        <h2 style="font-size:18px;font-weight:800;margin:0">⚡ Procesar Facturas</h2>
        <div style="font-size:12px;color:var(--mut);margin-top:4px">OTA (PDF) · Proveedores (PDF) · DRR (.xlsm) — los ya procesados se saltan automáticamente</div>
      </div>
      <button onclick="closeUploadModal()" style="background:none;border:none;color:var(--mut);font-size:24px;cursor:pointer">×</button>
    </div>
    <div id="upload-drop-zone"
         onclick="document.getElementById('upload-file-input').click()"
         ondragover="event.preventDefault();this.style.borderColor='#3b82f6';this.style.background='rgba(59,130,246,.08)'"
         ondragleave="this.style.borderColor='var(--s3)';this.style.background=''"
         ondrop="handleUploadDrop(event)"
         style="border:2px dashed var(--s3);border-radius:14px;padding:32px;text-align:center;cursor:pointer;transition:.2s;margin-bottom:16px">
      <div style="font-size:36px;margin-bottom:10px">📂</div>
      <div style="font-size:15px;font-weight:600;color:var(--tx);margin-bottom:6px">Arrastra archivos aquí o haz clic</div>
      <div style="font-size:12px;color:var(--dim);margin-bottom:14px">PDF (facturas) · XLSM (DRR)</div>
      <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
        <button onclick="event.stopPropagation();document.getElementById('upload-file-input').click()"
                style="background:var(--acc);border:none;color:#fff;padding:8px 18px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer">📄 Seleccionar archivos</button>
        <button onclick="event.stopPropagation();document.getElementById('upload-folder-input').click()"
                style="background:var(--s2);border:1px solid var(--s3);color:var(--tx);padding:8px 18px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer">📁 Seleccionar carpeta</button>
      </div>
    </div>
    <input id="upload-file-input" type="file" multiple accept=".pdf,.xlsm,.xlsx" style="display:none" onchange="handleUploadFiles(this.files)">
    <input id="upload-folder-input" type="file" multiple webkitdirectory accept=".pdf,.xlsm,.xlsx" style="display:none" onchange="handleUploadFiles(this.files)">
    <!-- Already uploaded files on server -->
    <div id="server-files-section" style="display:none;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">EN SERVIDOR (facturas-entrada)</div>
      <div id="server-files-list" style="max-height:120px;overflow-y:auto;display:flex;flex-direction:column;gap:5px"></div>
    </div>
    
    <div id="upload-file-list" style="display:none;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">ARCHIVOS SELECCIONADOS</div>
      <div id="upload-files-container" style="max-height:220px;overflow-y:auto;display:flex;flex-direction:column;gap:6px"></div>
      <div style="margin-top:10px;font-size:12px;color:var(--mut)">
        <span id="upload-count-new" style="color:var(--acc2);font-weight:700">0 nuevos</span> · <span id="upload-count-dup" style="color:var(--ora)">0 ya procesados (se saltarán)</span>
      </div>
    </div>
    <div style="display:flex;gap:10px;justify-content:space-between;align-items:center;flex-wrap:wrap">
      <button id="btn-procesar-server" onclick="procesarPendientesServidor()" style="background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.2);color:#22c55e;padding:9px 16px;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;display:none">
        ▶ Procesar pendientes del servidor
      </button>
      <div style="display:flex;gap:10px;margin-left:auto">
        <button onclick="closeUploadModal()" class="btn-ref">Cancelar</button>
        <button id="btn-upload-procesar" onclick="uploadAndProcess()" disabled
                style="background:var(--acc);border:none;color:#fff;padding:10px 22px;border-radius:10px;font-size:14px;font-weight:700;cursor:not-allowed;opacity:.4;transition:.2s">
          ⚡ Procesar archivos nuevos
        </button>
      </div>
    </div>
  </div>
</div>

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
const CHANGELOG_VER = '2026-06-v2';
const IS_MOBILE = window.innerWidth <= 768;
var otaChart = null;

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
  if (!a || a === '') return '<span class="badge b-pen">· ' + (t('lbl.pendiente')||'Pendiente') + '</span>';
  if (a === 'APROBADA')  return '<span class="badge b-apr">✓ ' + (t('lbl.aprobado')||'Aprobada') + '</span>';
  if (a === 'RECHAZADA') return '<span class="badge b-rec">✗ ' + (t('lbl.rechazado')||'Rechazada') + '</span>';
  return '<span class="badge b-na">—</span>';
}

// ── Carga datos ──────────────────────────────────────────────────────────
function generateBriefing(stats) {
  // Build a "morning briefing" summary message
  var issues = [];
  var good = [];
  if (stats.discrepancias > 0)  issues.push(stats.discrepancias + ' discrepancia(s) AR por resolver');
  if (stats.di_pendientes > 0)  issues.push(stats.di_pendientes + ' cert. DI pendiente(s)');
  if (stats.pendientes_firma > 0) issues.push(stats.pendientes_firma + ' factura(s) esperando firma');
  if (stats.rechazadas > 0)    issues.push(stats.rechazadas + ' factura(s) rechazada(s)');
  if (stats.correctas > 0)     good.push(stats.correctas + ' facturas AR correctas');
  if (stats.total > 0 && issues.length === 0) {
    good.push('Ciclo AR limpio — ' + stats.total + ' facturas sin problemas');
  }
  var statusBar = document.getElementById('status-txt');
  if (statusBar) {
    if (issues.length > 0) {
      statusBar.style.color = 'var(--ora)';
      statusBar.textContent = '⚠ ' + issues[0] + (issues.length > 1 ? ' (+' + (issues.length-1) + ' más)' : '');
    } else if (good.length > 0) {
      statusBar.style.color = 'var(--grn)';
      statusBar.textContent = '✓ ' + good[0];
      // First all-clear celebration
      if (!sessionStorage.getItem('all_clear_shown')) { sessionStorage.setItem('all_clear_shown','1'); setTimeout(()=>showNotification('🎉 ¡Todo en orden! Ciclo AR limpio.','success'),500); }
    }
    // Pulse animation on issues
    if (issues.length > 0) {
      statusBar.style.animation = 'pulse 2s infinite';
    } else {
      statusBar.style.animation = '';
    }
  }
}

async function loadAll() {
  const topBar = document.getElementById('top-bar');
  if (topBar) { topBar.style.width = '30%'; topBar.style.opacity = '1'; }
  document.getElementById('status-txt').textContent = t('status.actualizando') || 'Actualizando...';
  try {
    // 1. Cargar y renderizar stats primero (independiente de facturas)
    const sr = await fetch('/api/stats');
    const stats = await sr.json();
    renderStats(stats);
    try { renderChart(stats.chart); } catch(ec) { console.warn('Chart no disponible:', ec); }

    // Alert bar — defer until i18n is loaded
    const alertBar = document.getElementById('alert-bar');
    const _updateAlertBar = () => {
      const parts = [];
      if (stats.discrepancias > 0)
        parts.push(stats.discrepancias + ' ' + (t('alert.discrepancias') || 'discrepancia(s) reclamables'));
      if (stats.di_pendientes > 0)
        parts.push(stats.di_pendientes + ' ' + (t('alert.sinDI') || 'factura(s) sin cert. DI'));
      if (parts.length) {
        document.getElementById('alert-msg').textContent = parts.join(' — ');
        alertBar.classList.add('on');
      } else {
        alertBar.classList.remove('on');
      }
    };
    if (_i18nData && Object.keys(_i18nData).length > 0) {
      _updateAlertBar();
    } else {
      // Wait for i18n then update
      setTimeout(_updateAlertBar, 800);
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
  // Mobile KPI quick-view bar
  if (IS_MOBILE) {
    var mBar = document.getElementById('mobile-kpi-bar');
    if (mBar) mBar.style.display = 'block';
    var mkAR = document.getElementById('mkpi-ar');
    var mkAP = document.getElementById('mkpi-ap');
    if (mkAR) mkAR.textContent = (stats.discrepancias||0) + (stats.di_pendientes||0);
    if (mkAP) mkAP.textContent = stats.pendientes_firma || '0';
    // AR Real pending from quick stats
    fetch('/api/ar_real/stats').then(r=>r.json()).then(ar=>{
      if (ar.ok) {
        var pendEl = document.getElementById('mkpi-ar');
        if (pendEl && ar.pendiente > 0) {
          pendEl.textContent = '\u20AC' + Math.round(ar.pendiente/1000) + 'K';
          pendEl.parentElement.querySelector('.mkpi-lbl') && (document.querySelector('#mkpi-ar').nextSibling.textContent = 'AR pend');
        }
      }
    }).catch(()=>{});
    // GOP% from DRR if loaded
    fetch('/api/stats_drr').then(r=>r.json()).then(d=>{
      if (d && d.metricas) {
        var occ = d.metricas['Occupancy %'];
        var gop = d.metricas['GOP %'];
        var mkO = document.getElementById('mkpi-occ');
        var mkG = document.getElementById('mkpi-gop');
        if (mkO && occ) mkO.textContent = occ.today || '—';
        if (mkG && gop) mkG.textContent = gop.today || '—';
      }
    }).catch(()=>{});
    // Timestamp
    var updEl = document.getElementById('mkpi-updated');
    if (updEl) { updEl.textContent = 'Actualizado ' + new Date().toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'}); updEl.style.display='block'; }
  }

  // Tab notification badges
  function _setTabBadge(tabId, count, color) {
    const btn = document.getElementById('tab-' + tabId);
    if (!btn) return;
    let badge = btn.querySelector('.tab-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'tab-badge';
      badge.style.cssText = 'background:' + (color||'var(--red)') + ';color:#fff;border-radius:10px;font-size:9px;font-weight:700;padding:1px 5px;margin-left:4px;min-width:14px;text-align:center;display:inline-block';
      btn.appendChild(badge);
    }
    if (count > 0) { badge.textContent = count; badge.style.display = 'inline-block'; }
    else badge.style.display = 'none';
  }
  if (stats.discrepancias || stats.di_pendientes) {
    _setTabBadge('ar_otas', (stats.discrepancias||0) + (stats.di_pendientes||0), 'var(--red)');
  } else _setTabBadge('ar_otas', 0);

  // Daily alerts panel
  const alertsPanel = document.getElementById('daily-alerts-panel');
  const alertsList  = document.getElementById('daily-alerts-list');
  if (alertsPanel && alertsList) {
    const alerts = [];
    if (stats.discrepancias > 0)  alerts.push({type:'ar',    msg:'AR: ' + stats.discrepancias + ' discrepancia(s)', tab:'ar_otas'});
    if (stats.di_pendientes > 0)  alerts.push({type:'di',    msg:'DI: ' + stats.di_pendientes + ' cert. pendiente(s)', tab:'ar_otas'});
    if (stats.pendientes_firma > 0) alerts.push({type:'firma', msg:'Firma: ' + stats.pendientes_firma + ' factura(s)', tab:'ar_otas'});
    if (alerts.length > 0) {
      alertsPanel.style.display = 'block';
      alertsList.innerHTML = alerts.map(a =>
        '<div onclick="switchTab(\'' + a.tab + '\',document.getElementById(\'tab-' + a.tab + '\'))" style="background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.2);border-radius:8px;padding:6px 12px;font-size:11px;font-weight:600;cursor:pointer;color:var(--red)" onmouseover="this.style.background=\'rgba(239,68,68,.2)\'" onmouseout="this.style.background=\'rgba(239,68,68,.1)\'">⚠ ' + a.msg + '</div>'
      ).join('');
    }
  }
  // Executive summary
  const sumEl = document.getElementById('exec-summary');
  const sumTxt = document.getElementById('exec-txt');
  if (sumEl && sumTxt && stats.total > 0) {
    const parts = [];
    if (stats.discrepancias > 0) parts.push('⚠ ' + stats.discrepancias + ' discrepancia(s) AR por resolver');
    if (stats.di_pendientes > 0) parts.push('📄 ' + stats.di_pendientes + ' cert. DI pendiente(s)');
    if (stats.pendientes_firma > 0) parts.push('✍ ' + stats.pendientes_firma + ' factura(s) esperando firma');
    if (parts.length > 0) {
      sumEl.style.display = 'block';
      sumTxt.textContent = parts.join('  ·  ');
      var lnk = document.createElement('a');
      lnk.href = '#'; lnk.style.cssText = 'color:var(--acc2);font-size:11px;margin-left:8px;text-decoration:none';
      lnk.textContent = 'Ver AR →';
      lnk.onclick = function(e){ e.preventDefault(); switchTab('ar_otas',document.getElementById('tab-ar_otas')); };
      sumTxt.appendChild(lnk);
    }
    else { sumEl.style.display = 'block'; sumTxt.textContent = '✓ Todo en orden — ' + stats.total + ' facturas procesadas sin incidencias'; sumEl.style.background = 'rgba(34,197,94,.05)'; sumEl.style.borderColor = 'rgba(34,197,94,.1)'; }
  }
  if (topBar) { topBar.style.width = '100%'; setTimeout(() => { topBar.style.opacity = '0'; setTimeout(() => { topBar.style.width = '0'; topBar.style.opacity = '1'; }, 300); }, 400); }
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

var _arRows = [];
function renderTable(rows) {
  _arRows = rows;
  const tbody = document.getElementById('tbl-body');
  document.getElementById('tbl-count').textContent = rows.length ? rows.length + ' ' + (t('lbl.registros')||'registros') : '';
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="12" style="padding:32px;text-align:center"><div style="font-size:32px;margin-bottom:8px">📦</div><div style="font-weight:600;color:var(--mut);margin-bottom:4px">Sin facturas AP</div><div style="font-size:12px;color:var(--dim)">Pulsa ⚡ Procesar Facturas AP</div></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((r, i) => {
    const hasDisc = r.discrepancia_euros && r.discrepancia_euros !== '';
    return [
      '<tr style="cursor:pointer;transition:background .15s" data-idx="' + i + '" onclick="showInvoiceDetail(_arRows[parseInt(this.getAttribute(\'data-idx\'))])" onmouseover="this.style.background=\'rgba(59,130,246,.06)\';this.style.outline=\'1px solid rgba(59,130,246,.1)\'" onmouseout="this.style.background=\'\';this.style.outline=\'\'">',
      '<td style="padding:6px 4px;text-align:center"><input type="checkbox" class="ar-row-cb" style="cursor:pointer;accent-color:var(--acc)" onchange="updateSelectionCount()"></td>',
      '<td class="td-dim">' + (r.archivo || '—') + '</td>',
      '<td class="td-b" onclick="copyToClip(\'' + (r.numero_factura||'') + '\', \'Nº factura\')" style="cursor:pointer" title="Clic para copiar">' + (r.numero_factura || '—') + '</td>',
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
  const totalAmount = rows.reduce((s, r) => s + parseFloat(String(r.importe_bruto||'0').replace(/[^0-9.]/g,'')) || 0, 0);
  const totalStr = totalAmount > 0 ? '€' + totalAmount.toLocaleString('es-ES', {minimumFractionDigits:2}) : '';
  el.innerHTML = (totalStr ? '<div style="background:var(--bg);border-radius:8px;padding:8px 12px;margin-bottom:10px;font-size:12px;color:var(--mut)">Total ciclo: <strong style="color:var(--tx)">' + totalStr + '</strong></div>' : '') +
    items.map(i =>
    '<div class="act-item">' +
    '<div class="adot ' + i.dot + '"></div>' +
    '<div class="atxt"><b>' + i.n + '</b> factura' + (i.n !== 1 ? 's' : '') +
    ' <span data-i18n="' + i.key + '">' + i.txt + '</span></div>' +
    '</div>'
  ).join('');
  // Re-apply current language to freshly rendered spans
  if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
}

// ── Pipeline SSE ──────────────────────────────────────────────────────
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

    const actionsDiv = document.createElement('div');
    actionsDiv.style.cssText = 'display:flex;gap:12px;margin-top:16px;justify-content:flex-end';
    actionsDiv.innerHTML =
      '<button onclick="closeModal()" style="background:transparent;border:1px solid #444;color:#aaa;padding:10px 16px;border-radius:8px;cursor:pointer;font-size:13px">Cerrar</button>' +
      '<button onclick="closeModal();setTimeout(runPipeline,300)" style="background:#1db954;border:none;color:#fff;padding:10px 16px;border-radius:8px;cursor:pointer;font-weight:600;font-size:13px">🔄 Reintentar</button>';
    log.appendChild(actionsDiv);
    btnCl.disabled = false;
  };
}

function closeModal() {
  document.getElementById('overlay').classList.remove('on');
}

// ── Init ──────────────────────────────────────────────────────────────────
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
var _drrChart = null;

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
        labels: d.labels || d.dias.map((dia, i) => {
          const f = d.fechas[i] || '';
          return f ? f.slice(8) + '/' + f.slice(5,7) : String(dia);
        }),
        datasets: [{
          label: 'Revenue',
          backgroundColor: d.oob ? d.oob.map(isOob => isOob ? 'rgba(239,68,68,.7)' : 'rgba(59,130,246,.5)') : 'rgba(59,130,246,.5)',
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

var _tourStep = 0;
var _tourActive = false;

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

const _i18nCache = {};
const _i18nOriginal = {}; // textos ES originales — para restaurar al volver a español
var _i18nData = {};
var _i18nLang = localStorage.getItem('yve_lang') || 'es';

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

function toggleLightMode() {
  const isLight = document.body.classList.toggle('light-mode');
  localStorage.setItem('yve_theme', isLight ? 'light' : 'dark');
  const btn = document.getElementById('btn-theme');
  if (btn) btn.textContent = isLight ? '🌙 Modo oscuro' : '☀️ Modo claro';
}
// Apply saved theme on load
if (localStorage.getItem('yve_theme') === 'light') {
  document.body.classList.add('light-mode');
  const btn = document.getElementById('btn-theme');
  if (btn) btn.textContent = '🌙 Modo oscuro';
}

async function cambiarIdioma(lang) {
  fetch('/api/set_lang/' + lang);   // fire-and-forget, no await
  await loadI18n(lang);
  document.querySelectorAll('.lang-btn').forEach(b => {
    b.style.fontWeight = b.dataset.lang === lang ? '700' : '400';
    b.style.color = b.dataset.lang === lang ? 'var(--acc2)' : 'var(--tx)';
  });
}

loadAll();
setInterval(loadAll, 60000);
// Init mobile lite mode
if (IS_MOBILE) initMobileLite();
// Changelog badge
if (localStorage.getItem('changelog_seen') !== '2026-06-v2') {
  const mb = document.getElementById('menu-badge');
  if (mb) mb.style.display = 'inline-block';
}
// Show keyboard hint on first visit
if (!localStorage.getItem('kbd_shown')) {
  localStorage.setItem('kbd_shown', '1');  // auto-set, no need to show hint
}
// Show tour offer on first login
if (localStorage.getItem('tour_done') !== _TOUR_VER) {
  setTimeout(() => {
    if (confirm('¡Bienvenido a Yve.01! ¿Quieres un tour rápido de 2 minutos para conocer el sistema?')) {
      startTour();
    } else {
      localStorage.setItem('tour_skipped', '1');
    }
  }, 2000);
}

// ── Global error capture ─────────────────────────────────────────────────
window.onerror = function(msg, src, line, col, err) {
  console.error('[Yve.01 Error]', msg, 'at', src+':'+line+':'+col);
  // Don't show toast for errors during load — only after page is ready
  if (document.readyState === 'complete') {
    showNotification('⚠ Error JS: ' + String(msg).substring(0,80), 'error');
  }
  return false; // Don't suppress
};
window.addEventListener('unhandledrejection', function(e) {
  console.error('[Yve.01 Unhandled Promise]', e.reason);
});

// ── PWA Service Worker ───────────────────────────────────────────────────
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/static/sw.js').catch(() => {});
}
// Install prompt
var _deferredInstall;
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault(); _deferredInstall = e;
  const btn = document.getElementById('btn-install-pwa');
  if (btn) btn.style.display = 'inline-block';
});
window.addEventListener('appinstalled', () => {
  const btn = document.getElementById('btn-install-pwa');
  if (btn) btn.style.display = 'none';
});

// ── Toast notifications ──────────────────────────────────────────────────
var _toastTimeout;
function showNotification(msg, type = 'info') {
  let toast = document.getElementById('yve-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'yve-toast';
    toast.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(20px);opacity:0;max-width:420px;width:calc(100% - 40px);padding:12px 18px;border-radius:12px;font-size:13px;font-weight:500;z-index:9999;transition:all .25s ease;pointer-events:none;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.4)';
    document.body.appendChild(toast);
  }
  clearTimeout(_toastTimeout);
  const colors = {
    success: {bg:'rgba(34,197,94,.15)',border:'rgba(34,197,94,.4)',color:'#22c55e'},
    error:   {bg:'rgba(239,68,68,.15)',border:'rgba(239,68,68,.4)',color:'#f87171'},
    warning: {bg:'rgba(245,158,11,.15)',border:'rgba(245,158,11,.4)',color:'#f59e0b'},
    info:    {bg:'rgba(59,130,246,.15)',border:'rgba(59,130,246,.3)',color:'#60a5fa'},
  };
  const c = colors[type] || colors.info;
  toast.style.background  = c.bg;
  toast.style.border      = '1px solid ' + c.border;
  toast.style.color       = c.color;
  toast.textContent       = msg;
  toast.style.opacity     = '1';
  toast.style.transform   = 'translateX(-50%) translateY(0)';
  _toastTimeout = setTimeout(() => {
    toast.style.opacity   = '0';
    toast.style.transform = 'translateX(-50%) translateY(20px)';
  }, type === 'error' ? 5000 : 3000);
}

// ── Guided tour ─────────────────────────────────────────────────────────
var _tourActive = false, _tourStep = 0;
var _mh_loaded = false;
var _calLoaded = false;
var _drrLoaded = false;
var _fbLoaded = false;
var _arRealLoaded = false;
var _TOUR_VER = '3';  // increment to re-offer after updates
var _tourSteps = [
  {
    el: '#app-header',
    tab: null,
    title: '👋 Bienvenido a Yve.01',
    text: 'El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting para el Grupo Calipolis. Este tour te lleva por cada módulo — 3 minutos y ya lo dominas todo.'
  },
  {
    el: '#ar-stats-section',
    tab: 'ar',
    title: '💳 AR — Comisiones OTA',
    text: 'Verifica automáticamente las comisiones de Booking.com y Expedia. De un vistazo: facturas procesadas, importe total, discrepancias reclamables y certificados de Doble Imposición pendientes. El número rojo son euros que puedes recuperar.'
  },
  {
    el: '#stats-ap-grid',
    tab: 'ap',
    title: '📦 AP — 3-way Matching',
    text: 'Para cada factura de proveedor, Yve cruza 3 documentos: la factura, el pedido (PO) y el albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.'
  },
  {
    el: '#drr-metrics',
    tab: 'drr',
    title: '📊 DRR — Daily Revenue Report',
    text: 'Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.'
  },
  {
    el: '#ar-real-stats',
    tab: 'ar_real',
    title: '🏢 AR Real — Grupos Corporativos',
    text: 'Gestión completa de clientes como Telefónica o Accenture: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.'
  },
  {
    el: '#cal-kpis',
    tab: 'calipolis',
    title: '🏨 Grupo Calipolis',
    text: 'Dashboard consolidado de las 3 propiedades: Sitges, Mar y Boutique. GOP%, RevPAR, ocupación y tendencias de 6 meses. €2M de revenue gestionados desde aquí.',
    action: function() {
      if (typeof loadCalipolis === 'function' && !document.getElementById('cal-kpis').dataset.loaded) {
        loadCalipolis();
      }
    }
  },
  {
    el: '#mh-kpis',
    tab: 'multi_hotel',
    title: '🌍 Multi-Hotel — Vista de Grupo',
    text: 'Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.',
    action: function() {
      if (typeof loadMultiHotel === 'function') {
        _mh_loaded = false;
        loadMultiHotel();
      }
    }
  }
];

function _launchConfetti() {
  var canvas = document.createElement('canvas');
  canvas.id = 'confetti-canvas';
  canvas.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;z-index:9998;pointer-events:none';
  document.body.appendChild(canvas);
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;
  var ctx = canvas.getContext('2d');
  var pieces = [];
  var colors = ['#3b82f6','#22c55e','#f59e0b','#ec4899','#a78bfa','#06b6d4','#f97316','#ef4444'];
  for (var i = 0; i < 180; i++) {
    pieces.push({
      x: Math.random() * canvas.width, y: -20 - Math.random() * canvas.height * 0.4,
      w: Math.random() * 12 + 5, h: Math.random() * 7 + 3,
      r: Math.random() * Math.PI * 2, vx: (Math.random() - 0.5) * 4,
      vy: Math.random() * 5 + 2, vr: (Math.random() - 0.5) * 0.18,
      color: colors[Math.floor(Math.random() * colors.length)], opacity: 1
    });
  }
  var frame = 0;
  (function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    var alive = false;
    pieces.forEach(function(p) {
      p.x += p.vx; p.y += p.vy; p.r += p.vr;
      if (frame > 70) p.opacity -= 0.012;
      if (p.opacity > 0 && p.y < canvas.height + 30) {
        alive = true;
        ctx.save(); ctx.globalAlpha = Math.max(0, p.opacity);
        ctx.translate(p.x, p.y); ctx.rotate(p.r);
        ctx.fillStyle = p.color;
        ctx.fillRect(-p.w/2, -p.h/2, p.w, p.h);
        ctx.restore();
      }
    });
    frame++;
    if (alive && frame < 200) requestAnimationFrame(draw);
    else canvas.remove();
  })();
}

function startTour() {
  _tourActive = true; _tourStep = 0;
  // Close menu if open
  var mm = document.getElementById('main-menu');
  if (mm) mm.classList.remove('open');
  // Reset old elements
  var oldBox = document.getElementById('tour-box');
  if (oldBox) oldBox.remove();
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (canvas) { canvas.style.display='block'; var c2=canvas.getContext('2d'); c2.clearRect(0,0,canvas.width,canvas.height); }
  _showTourStep();
}

function _showTourStep() {
  if (_tourStep >= _tourSteps.length) { endTour(); return; }
  var step = _tourSteps[_tourStep];

  // ── 1. Switch tab first, then action ─────────────────────────────────
  if (step.tab) {
    var tabEl = document.getElementById('tab-' + step.tab);
    if (tabEl) switchTab(step.tab, tabEl);
  }
  if (step.action && typeof step.action === 'function') {
    setTimeout(function() { try { step.action(); } catch(e) {} }, 300);
  }

  // ── 2. Clear previous highlight ────────────────────────────────────
  document.querySelectorAll('[data-tour-active]').forEach(function(el) {
    el.style.outline = '';
    el.style.outlineOffset = '';
    el.style.zIndex = '';
    el.removeAttribute('data-tour-active');
  });

  // ── 3. Find target + highlight ─────────────────────────────────────
  var target = null;
  if (step.el) {
    // Wait a tick for tab switch to render
    setTimeout(function() {
      target = document.querySelector(step.el);
      _positionTour(target, step);
    }, step.tab ? 500 : 0);
  } else {
    _positionTour(null, step);
  }
}

function _positionTour(target, step) {
  // ── Overlay with spotlight ────────────────────────────────────────
  var overlay = document.getElementById('tour-overlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'tour-overlay';
    document.body.appendChild(overlay);
  }
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9900;pointer-events:all;cursor:pointer';
  overlay.onclick = function(e) { if (e.target === overlay || e.target === document.getElementById('tour-spotlight-bg')) nextTourStep(); };
  overlay.style.display = 'block';

  // Create spotlight canvas overlay
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (!canvas) {
    canvas = document.createElement('canvas');
    canvas.id = 'tour-spotlight-canvas';
    canvas.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:9901';
    document.body.appendChild(canvas);
  }
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;
  var ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (target) {
    target.scrollIntoView({behavior:'smooth', block:'nearest', inline:'center'});
    target.setAttribute('data-tour-active', '1');
    target.style.position = 'relative';
    target.style.zIndex = '9950';

    setTimeout(function() {
      // Re-query target in case DOM updated after tab switch
      if (step && step.el) { var freshTarget = document.querySelector(step.el); if (freshTarget) target = freshTarget; }
      var rect = target.getBoundingClientRect();
      // If rect is 0,0,0,0 (element not visible yet) — fall back to full dim
      if (rect.width === 0 && rect.height === 0) {
        ctx.fillStyle = 'rgba(0,0,0,0.78)'; ctx.fillRect(0,0,canvas.width,canvas.height);
        _placeTourBox(null, 0); return;
      }
      var pad = 10;
      var x = rect.left - pad, y = rect.top - pad;
      var w = rect.width + pad*2, h = rect.height + pad*2;

      // Dark overlay with hole
      var extraPad = (step && step.el && (step.el.includes('stats') || step.el.includes('panel') || step.el.includes('metrics') || step.el.includes('grid'))) ? 24 : 8;
      ctx.fillStyle = 'rgba(0,0,0,0.78)';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.clearRect(x - extraPad, y - extraPad, w + extraPad*2, h + extraPad*2);

      // Blue glow around hole
      // Outer glow
      ctx.shadowColor = '#3b82f6';
      ctx.shadowBlur = 30;
      ctx.strokeStyle = 'rgba(59,130,246,0.8)';
      ctx.lineWidth = 3;
      ctx.beginPath();
      if (ctx.roundRect) ctx.roundRect(x, y, w, h, 10); else ctx.rect(x, y, w, h);
      ctx.stroke();
      ctx.shadowBlur = 0;
      // Inner highlight
      ctx.strokeStyle = 'rgba(96,165,250,0.4)';
      ctx.lineWidth = 1;
      ctx.beginPath();
      if (ctx.roundRect) ctx.roundRect(x+2, y+2, w-4, h-4, 8); else ctx.rect(x+2, y+2, w-4, h-4);
      ctx.stroke();

      _placeTourBox(rect, pad);
    }, step.tab ? 800 : 150);
  } else {
    // No target → full dim, box centered
    ctx.fillStyle = 'rgba(0,0,0,0.72)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    // Show box in top-left area when no target
    var box2 = document.getElementById('tour-box');
    if (box2) { box2.style.top = '80px'; box2.style.left = '20px'; box2.style.transform = ''; }
  }

  // ── Build / update tooltip box ────────────────────────────────────
  var box = document.getElementById('tour-box');
  if (!box) {
    box = document.createElement('div');
    box.id = 'tour-box';
    document.body.appendChild(box);
  }
  box.setAttribute('style',
    'position:fixed;background:#0f172a;border:2px solid #3b82f6;border-radius:16px;' +
    'padding:22px 26px 18px;max-width:360px;width:calc(100vw - 32px);z-index:10000;' +
    'box-shadow:0 20px 60px rgba(0,0,0,.85),0 0 0 1px rgba(59,130,246,.35),0 0 60px rgba(59,130,246,.08);' +
    'pointer-events:all;font-family:Inter,system-ui,sans-serif;color:#f1f5f9;' +
    'animation:tourBoxIn .3s cubic-bezier(.34,1.56,.64,1)');
  box.innerHTML =
    // Close button
    '<button onclick="endTour()" style="position:absolute;top:10px;right:12px;background:none;border:none;' +
      'color:#475569;font-size:22px;cursor:pointer;line-height:1;padding:0;transition:.15s" ' +
      'onmouseover="this.style.color=\'#94a3b8\'" onmouseout="this.style.color=\'#475569\'">×</button>' +
    // Title  
    '<div style="font-size:17px;font-weight:800;color:#f1f5f9;padding-right:24px;margin-bottom:10px;line-height:1.3">' + step.title + '</div>' +
    // Body text
    '<div style="font-size:14px;color:#94a3b8;line-height:1.7;margin-bottom:20px">' + step.text + '</div>' +
    // Dots + buttons row
    '<div style="display:flex;align-items:center;justify-content:space-between">' +
      '<div style="display:flex;gap:6px;align-items:center">' +
        _tourSteps.map(function(_,i) {
          return '<div style="transition:.2s;border-radius:50%;background:' +
            (i === _tourStep ? '#3b82f6' : 'rgba(59,130,246,.25)') +
            ';width:' + (i === _tourStep ? '20px' : '8px') +
            ';height:8px;border-radius:' + (i === _tourStep ? '4px' : '50%') + '"></div>';
        }).join('') +
      '</div>' +
      '<div style="display:flex;gap:8px">' +
        (_tourStep > 0 ?
          '<button onclick="prevTourStep()" style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);color:#94a3b8;' +
          'padding:8px 14px;border-radius:10px;font-size:13px;cursor:pointer;font-weight:500">← Atrás</button>' : '') +
        '<button onclick="nextTourStep()" style="background:#3b82f6;border:none;color:#fff;' +
          'padding:8px 18px;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;' +
          'box-shadow:0 4px 12px rgba(59,130,246,.4)">' +
          (_tourStep === _tourSteps.length-1 ? '✓ Finalizar' : 'Siguiente →') +
        '</button>' +
      '</div>' +
    '</div>';
}

function _placeTourBox(targetRect, pad) {
  var box = document.getElementById('tour-box');
  if (!box) return;
  var bw = 336, bh = 260;
  var vw = window.innerWidth, vh = window.innerHeight;
  var top, left;

  if (targetRect) {
    var rect = targetRect;
    var isSmall = rect.width < 200;  // tab button vs panel section
    if (isSmall) {
      // Small target (tab button): show below
      top  = rect.bottom + pad + 12;
      left = Math.max(10, Math.min(rect.left, vw - bw - 10));
      if (top + bh > vh - 16) top = rect.top - bh - 16;
    } else {
      // Large target (panel section): show beside or below
      top  = Math.max(10, rect.bottom + 16);
      left = Math.max(10, Math.min(rect.left, vw - bw - 10));
      // If goes off bottom → show in top-left of screen
      if (top + bh > vh - 16) top = Math.max(10, (vh - bh) / 2);
    }
    if (top < 10) top = 10;
  } else {
    top  = Math.max(10, (vh - bh) / 2);
    left = Math.max(10, (vw - bw) / 2);
  }
  box.style.top  = Math.round(top)  + 'px';
  box.style.left = Math.round(left) + 'px';
  box.style.transform = '';
}


function nextTourStep() {
  _tourStep++;
  if (_tourStep >= _tourSteps.length) { endTour(); return; }
  _showTourStep();
}

function prevTourStep() {
  if (_tourStep > 0) { _tourStep--; _showTourStep(); }
}

function endTour() {
  _tourActive = false; _tourStep = 0;
  // Clear spotlight canvas
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (canvas) { var ctx2 = canvas.getContext('2d'); ctx2.clearRect(0,0,canvas.width,canvas.height); canvas.style.display='none'; }
  // Clear all highlights
  document.querySelectorAll('[data-tour-active]').forEach(function(el) {
    el.style.outline = '';
    el.style.outlineOffset = '';
    el.style.boxShadow = '';
    el.style.zIndex = '';
    el.removeAttribute('data-tour-active');
  });
  // Launch confetti + show completion screen
  _launchConfetti();
  var box = document.getElementById('tour-box');
  if (box) {
    box.setAttribute('style',
      'position:fixed;background:#0f172a;border:2px solid #22c55e;border-radius:20px;' +
      'padding:30px 28px 24px;max-width:340px;width:calc(100vw - 32px);z-index:10000;' +
      'box-shadow:0 24px 60px rgba(0,0,0,.85),0 0 0 1px rgba(34,197,94,.25),0 0 40px rgba(34,197,94,.1);' +
      'pointer-events:all;font-family:Inter,system-ui,sans-serif;color:#f1f5f9;' +
      'top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;' +
      'animation:tourBoxIn .4s cubic-bezier(.34,1.56,.64,1)');
    box.innerHTML =
      '<div style="font-size:56px;margin-bottom:12px;line-height:1">🎉</div>' +
      '<div style="font-size:22px;font-weight:900;color:#f1f5f9;margin-bottom:10px;letter-spacing:-.3px">\u00A1Tour completado!</div>' +
      '<div style="font-size:14px;color:#94a3b8;line-height:1.7;margin-bottom:24px">' +
        'Ya conoces los 7 m\u00F3dulos de Yve.01.<br>Empieza procesando tus primeras facturas.' +
      '</div>' +
      '<div style="display:flex;gap:10px;justify-content:center">' +
        '<button onclick="closeTourBox()" style="background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);' +
          'color:#94a3b8;padding:11px 20px;border-radius:10px;font-size:13px;cursor:pointer;font-weight:500">Cerrar</button>' +
        '<button onclick="goToARPanel()" style="background:#22c55e;border:none;color:#fff;' +
          'padding:11px 24px;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;' +
          'box-shadow:0 4px 20px rgba(34,197,94,.4);transition:.15s">' +
          'Empezar con AR \u2192' +
        '</button>' +
      '</div>';
    setTimeout(function(){ closeTourBox(); }, 9000);
  }
  localStorage.setItem('tour_done', _TOUR_VER);
}

function goToARPanel(){ closeTourBox(); switchTab('ar', document.getElementById('tab-ar')); }
function closeTourBox() {
  var overlay = document.getElementById('tour-overlay');
  if (overlay) overlay.style.display = 'none';
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (canvas) { var c2 = canvas.getContext('2d'); c2.clearRect(0,0,canvas.width,canvas.height); canvas.style.display='none'; }
  var box = document.getElementById('tour-box');
  if (box) { box.style.display = 'none'; box.style.transform = ''; }
  document.querySelectorAll('[data-tour-active]').forEach(function(el) {
    el.style.zIndex = ''; el.removeAttribute('data-tour-active');
  });
}

// ── Invoice detail modal ─────────────────────────────────────────────────
// ── Mobile lite mode ─────────────────────────────────────────────────────
var _mobileLite = true;  // default ON on mobile

function initMobileLite() {
  if (!IS_MOBILE) return;
  // Check stored preference
  var stored = localStorage.getItem('mobile_lite');
  _mobileLite = stored === null ? true : stored === '1';
  applyMobileLite();
}

function applyMobileLite() {
  var btn = document.getElementById('mobile-lite-toggle');
  if (_mobileLite) {
    document.body.classList.add('mobile-lite');
    if (btn) btn.innerHTML = '📊 Vista completa';
    if (btn) btn.style.color = 'var(--acc2)';
    if (btn) btn.style.borderColor = 'rgba(59,130,246,.3)';
  } else {
    document.body.classList.remove('mobile-lite');
    if (btn) btn.innerHTML = '📱 Vista lite';
    if (btn) btn.style.color = 'var(--mut)';
    if (btn) btn.style.borderColor = 'var(--s2)';
  }
}

function toggleMobileLite() {
  _mobileLite = !_mobileLite;
  localStorage.setItem('mobile_lite', _mobileLite ? '1' : '0');
  applyMobileLite();
  // Reload current tab data if switching to full
  if (!_mobileLite) {
    showNotification('Vista completa activada', 'info');
  } else {
    showNotification('Vista lite activada', 'info');
  }
}

function toggleAtajos() {
  var m = document.getElementById('atajos-modal');
  if (!m) return;
  m.style.display = m.style.display === 'flex' ? 'none' : 'flex';
}

// Close modals on Escape
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') {
    closeInvoiceModal();
    var am = document.getElementById('atajos-modal');
    if (am) am.style.display = 'none';
    var tm = document.getElementById('modal-emitir');
    if (tm) tm.style.display = 'none';
    if (_tourActive) endTour();
    else closeTourBox();
  }
});

async function generarEmailAR(numero) {
  if (!numero) return;
  try {
    showNotification('⏳ Generando email...', 'info');
    const r = await fetch('/api/ar_real/recordatorio', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({numero})
    });
    const d = await r.json();
    showNotification(d.ok ? '✓ Email enviado a ' + (d.email||'') : '✗ ' + (d.error||'Error'), d.ok ? 'success' : 'error');
    if (d.ok) closeInvoiceModal();
  } catch(e) {
    showNotification('✗ Error de conexión', 'error');
  }
}

function generarEmailAP(numero) {
  if (!numero) return;
  showNotification('⏳ Generando email AP...', 'info');
  fetch('/ap/api/generar_email', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({numero_factura: numero})
  }).then(function(r){ return r.json(); }).then(function(d){
    if (d.ok) {
      showNotification('✉ Email generado: ' + (d.email||'proveedor'), 'success');
      closeInvoiceModal();
    } else {
      showNotification('✗ ' + (d.error||'Error generando email'), 'error');
    }
  }).catch(function(){ showNotification('✗ Error de conexión', 'error'); });
}

function showAPDetail(row) {
  var modal = document.getElementById('invoice-modal');
  var body  = document.getElementById('inv-modal-body');
  var title = document.getElementById('inv-modal-title');
  if (!modal || !body || !row) return;
  title.textContent = row.numero_factura || 'Factura AP';
  var stC = row.aprobacion === 'APROBADA' ? '#22c55e' : row.aprobacion === 'RECHAZADA' ? '#ef4444' : '#f59e0b';
  var fields = [
    ['Proveedor', row.proveedor||'—'], ['Fecha', row.fecha_factura||'—'],
    ['Total', row.importe_con_iva ? '€'+row.importe_con_iva : '—'],
    ['Cuenta', row.cuenta_contable||'—'], ['Tipo', row.tipo||'—'],
    ['Estado', row.estado||'—'], ['Aprobación', row.aprobacion||'—'],
    ['PO', row.tiene_po ? '✅' : '❌'], ['Albarán', row.tiene_alb ? '✅' : '❌'],
  ];
  body.innerHTML =
    '<div style="background:'+stC+'20;border:1px solid '+stC+'40;border-radius:10px;padding:10px 14px;margin-bottom:16px;font-weight:700;color:'+stC+'">'+(row.aprobacion||'Pendiente')+'</div>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">' +
    fields.map(function(f){ return '<div style="background:var(--bg);border-radius:8px;padding:10px"><div style="font-size:10px;color:var(--dim);font-weight:600;text-transform:uppercase;margin-bottom:3px">'+f[0]+'</div><div style="font-size:13px;font-weight:600">'+f[1]+'</div></div>'; }).join('') +
    '</div>' +
    '<div style="display:flex;gap:10px;margin-top:16px"><button onclick="closeInvoiceModal()" class="btn-ref" style="flex:1">Cerrar</button></div>';
  modal.style.display = 'flex';
}

// ── Upload Modal ─────────────────────────────────────────────────────────
var _uploadFiles = [];        // File objects selected by user
var _processedNames = new Set(); // Names already processed (from server)

async function openUploadModal() {
  // Reset state
  _uploadFiles = [];
  document.getElementById('upload-file-list').style.display = 'none';
  document.getElementById('upload-files-container').innerHTML = '';
  document.getElementById('upload-count-new').textContent = '0 nuevos';
  document.getElementById('upload-count-dup').textContent = '0 ya procesados (se saltarán)';
  var procBtn = document.getElementById('btn-upload-procesar');
  procBtn.disabled = true; procBtn.style.opacity = '.4'; procBtn.style.cursor = 'not-allowed';

  // Load already-processed file names from server
  try {
    var r = await fetch('/api/archivos_estado');
    var d = await r.json();
    _processedNames = new Set((d.files || []).filter(f => f.procesado).map(f => f.nombre));
  } catch(e) { _processedNames = new Set(); }

  // Show files already on server
  try {
    var r2 = await fetch('/api/archivos_estado');
    var d2 = await r2.json();
    var serverFiles = d2.files || [];
    var sSection = document.getElementById('server-files-section');
    var sList = document.getElementById('server-files-list');
    if (serverFiles.length > 0 && sSection && sList) {
      sSection.style.display = 'block';
      // Show "process server pending" button
      var serverBtn = document.getElementById('btn-procesar-server');
      var pendingCount = serverFiles.filter(function(f){ return !f.procesado; }).length;
      if (serverBtn) {
        serverBtn.style.display = pendingCount > 0 ? 'block' : 'none';
        serverBtn.textContent = '▶ Procesar ' + pendingCount + ' pendiente' + (pendingCount !== 1 ? 's' : '') + ' del servidor';
      }
      sList.innerHTML = serverFiles.map(function(f) {
        return '<div style="display:flex;align-items:center;gap:8px;padding:6px 10px;background:' + 
          (f.procesado ? 'rgba(34,197,94,.05)' : 'var(--bg)') + ';border-radius:7px;border:1px solid ' +
          (f.procesado ? 'rgba(34,197,94,.2)' : 'var(--s2)') + ';font-size:12px">' +
          '<span>' + (f.nombre.endsWith('.xlsm') ? '📊' : '📄') + '</span>' +
          '<span style="flex:1;color:var(--tx)">' + f.nombre + '</span>' +
          '<span style="color:var(--dim)">' + f.tamano_str + '</span>' +
          '<span style="font-size:11px;padding:2px 7px;border-radius:5px;background:' + 
            (f.procesado ? 'rgba(34,197,94,.1)' : 'rgba(245,158,11,.1)') + ';color:' +
            (f.procesado ? '#22c55e' : '#f59e0b') + '">' +
            (f.procesado ? '✓ Procesado' : '⏳ Pendiente') + '</span>' +
          '</div>';
      }).join('');
    } else if (sSection) {
      sSection.style.display = 'none';
    }
  } catch(e) {}

  var modal = document.getElementById('upload-modal');
  modal.style.display = 'flex';
}

function procesarPendientesServidor() {
  // Process files already on server that haven't been processed yet
  fetch('/api/archivos_estado')
    .then(function(r){ return r.json(); })
    .then(function(d) {
      var pendientes = (d.files || []).filter(function(f){ return !f.procesado; }).map(function(f){ return f.nombre; });
      if (!pendientes.length) { showNotification('No hay archivos pendientes en el servidor', 'info'); return; }
      closeUploadModal();
      showNotification('⏳ Procesando ' + pendientes.length + ' archivo(s) del servidor...', 'info');
      _runBatchPipeline(pendientes);
    })
    .catch(function(e){ showNotification('Error: ' + e.message, 'error'); });
}

function closeUploadModal() {
  document.getElementById('upload-modal').style.display = 'none';
  _uploadFiles = [];
}

function handleUploadDrop(e) {
  e.preventDefault();
  var zone = document.getElementById('upload-drop-zone');
  zone.style.borderColor = 'var(--s3)'; zone.style.background = '';
  var items = e.dataTransfer.items;
  var files = [];
  if (items) {
    // Handle folders via DataTransferItemList
    for (var i = 0; i < items.length; i++) {
      var entry = items[i].webkitGetAsEntry ? items[i].webkitGetAsEntry() : null;
      if (entry && entry.isDirectory) {
        // Read directory
        _readDir(entry, files, function() { _addFilesToList(files); });
        return;
      } else if (items[i].kind === 'file') {
        files.push(items[i].getAsFile());
      }
    }
  }
  _addFilesToList(files);
}

function _readDir(dirEntry, files, done) {
  var reader = dirEntry.createReader();
  reader.readEntries(function(entries) {
    var pending = entries.length;
    if (pending === 0) { done(); return; }
    entries.forEach(function(entry) {
      if (entry.isFile) {
        entry.file(function(f) { files.push(f); if (--pending === 0) done(); });
      } else if (entry.isDirectory) {
        _readDir(entry, files, function() { if (--pending === 0) done(); });
      } else {
        if (--pending === 0) done();
      }
    });
  });
}

function handleUploadFiles(fileList) {
  var files = Array.from(fileList).filter(function(f) {
    return f.name.match(/\.(pdf|xlsm|xlsx)$/i);
  });
  _addFilesToList(files);
}

function _addFilesToList(newFiles) {
  // Merge with existing, deduplicate by name
  var existing = new Set(_uploadFiles.map(function(f){ return f.name; }));
  newFiles.forEach(function(f) { if (!existing.has(f.name)) { _uploadFiles.push(f); } });
  _renderFileList();
}

function _detectType(fname) {
  var n = fname.toLowerCase();
  if (n.endsWith('.xlsm') || n.includes('drr') || n.includes('daily')) return 'DRR';
  if (n.includes('booking') || n.includes('expedia') || n.includes('hotelbeds') || n.includes('ota')) return 'AR — OTA';
  if (n.endsWith('.pdf')) return 'AP / AR';
  return 'Otro';
}

function _typeColor(t) {
  if (t === 'DRR') return '#a78bfa';
  if (t.includes('OTA') || t.includes('AR')) return '#60a5fa';
  if (t.includes('AP')) return '#f59e0b';
  return 'var(--mut)';
}

function _renderFileList() {
  var cont = document.getElementById('upload-files-container');
  var list = document.getElementById('upload-file-list');
  if (!_uploadFiles.length) { list.style.display = 'none'; return; }
  list.style.display = 'block';

  var newCount = 0, dupCount = 0;
  cont.innerHTML = _uploadFiles.map(function(f, i) {
    var isProc = _processedNames.has(f.name);
    var tipo = _detectType(f.name);
    var size = f.size < 1024*1024 ? Math.round(f.size/1024) + 'KB' : (f.size/1024/1024).toFixed(1) + 'MB';
    if (isProc) dupCount++; else newCount++;
    return '<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:' + 
      (isProc ? 'rgba(245,158,11,.06)' : 'var(--bg)') + ';border-radius:8px;border:1px solid ' +
      (isProc ? 'rgba(245,158,11,.2)' : 'var(--s2)') + ';opacity:' + (isProc ? '.6' : '1') + '">' +
      '<div style="font-size:18px">' + (f.name.endsWith('.xlsm') ? '📊' : '📄') + '</div>' +
      '<div style="flex:1;min-width:0">' +
        '<div style="font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + f.name + '</div>' +
        '<div style="font-size:11px;color:var(--dim)">' + size + ' · <span style="color:' + _typeColor(tipo) + '">' + tipo + '</span>' + 
          (isProc ? ' · <span style="color:var(--ora)">⚠ Ya procesado</span>' : '') +
        '</div>' +
      '</div>' +
      '<button onclick="_removeUploadFile(' + i + ')" style="background:none;border:none;color:var(--dim);cursor:pointer;font-size:16px;padding:0 4px" title="Quitar">×</button>' +
      '</div>';
  }).join('');

  document.getElementById('upload-count-new').textContent = newCount + ' nuevo' + (newCount !== 1 ? 's' : '');
  document.getElementById('upload-count-dup').textContent = dupCount + ' ya procesado' + (dupCount !== 1 ? 's' : '') + ' (se saltarán)';
  
  var procBtn = document.getElementById('btn-upload-procesar');
  procBtn.textContent = '⚡ Procesar ' + (newCount > 0 ? newCount + ' archivo' + (newCount !== 1 ? 's' : '') + ' nuevo' + (newCount !== 1 ? 's' : '') : 'seleccionados');
  procBtn.disabled = newCount === 0;
  procBtn.style.opacity = newCount > 0 ? '1' : '.4';
  procBtn.style.cursor = newCount > 0 ? 'pointer' : 'not-allowed';
}

function _removeUploadFile(idx) {
  _uploadFiles.splice(idx, 1);
  _renderFileList();
}

async function uploadAndProcess() {
  var newFiles = _uploadFiles.filter(function(f) { return !_processedNames.has(f.name); });
  if (!newFiles.length) { showNotification('No hay archivos nuevos que procesar', 'info'); return; }
  
  var btn = document.getElementById('btn-upload-procesar');
  btn.disabled = true; btn.style.opacity = '.4';
  btn.textContent = '⏳ Subiendo archivos...';

  // Upload files
  var formData = new FormData();
  newFiles.forEach(function(f) { formData.append('files', f, f.name); });
  
  try {
    var r = await fetch('/api/upload_facturas', { method: 'POST', body: formData });
    var d = await r.json();
    if (!d.ok) throw new Error(d.error || 'Upload failed');
    showNotification('✓ ' + d.subidos + ' archivo(s) subidos', 'success');
  } catch(e) {
    showNotification('✗ Error subiendo archivos: ' + e.message, 'error');
    btn.disabled = false; btn.style.opacity = '1'; btn.textContent = '⚡ Reintentar';
    return;
  }

  // Close upload modal and open pipeline modal to show progress
  closeUploadModal();
  
  // Trigger batch processing via SSE
  _runBatchPipeline(newFiles.map(function(f){ return f.name; }));
}

function _runBatchPipeline(fileNames) {
  // Open the standard pipeline log modal
  var overlay = document.getElementById('overlay');
  var log = document.getElementById('log');
  var btn = document.getElementById('btn-run');
  var spin = document.getElementById('spin');
  var lbl = document.getElementById('run-lbl');
  var btnCl = document.getElementById('btn-cl');
  var icon = document.getElementById('modal-icon');
  var title = document.getElementById('modal-title');

  if (overlay) overlay.classList.add('on');
  if (log) log.innerHTML = '';
  if (btn) btn.disabled = true;
  if (spin) spin.style.display = 'block';
  if (lbl) lbl.textContent = 'Procesando...';
  if (btnCl) btnCl.disabled = true;
  if (icon) icon.textContent = '⚡';
  if (title) title.textContent = 'Procesando ' + fileNames.length + ' archivo(s)...';

  fetch('/api/procesar_batch', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({solo_nuevos: true, archivos: fileNames, tipos: ['AR','AP','DRR','AR_o_AP']})
  }).then(function() {
    // For batch, use EventSource to stream results
    var evtSrc = new EventSource('/api/procesar_batch_stream?archivos=' + encodeURIComponent(JSON.stringify(fileNames)));
    evtSrc.onmessage = function(ev) {
      var txt = ev.data;
      if (!log) return;
      var p = document.createElement('p');
      if (txt === 'PIPELINE_COMPLETO') p.className = 'l-ok';
      else if (txt === 'PIPELINE_CON_ERRORES') p.className = 'l-err';
      else if (txt.startsWith('✓') || txt.startsWith('OK')) p.className = 'l-ok';
      else if (txt.startsWith('✗') || txt.startsWith('ERROR')) p.className = 'l-err';
      else if (txt.startsWith('>>') || txt.startsWith('ℹ')) p.className = 'l-info';
      else p.className = 'l-dim';
      p.textContent = txt;
      log.appendChild(p); log.scrollTop = log.scrollHeight;
      if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
        evtSrc.close();
        var ok = txt === 'PIPELINE_COMPLETO';
        if (icon) icon.textContent = ok ? '✅' : '⚠️';
        if (title) title.textContent = ok ? 'Procesamiento completado' : 'Completado con errores';
        if (btn) btn.disabled = false;
        if (spin) spin.style.display = 'none';
        if (lbl) lbl.textContent = '⚡ Procesar Facturas';
        if (btnCl) btnCl.disabled = false;
        setTimeout(loadAll, 800);
      }
    };
    evtSrc.onerror = function() {
      evtSrc.close();
      if (icon) icon.textContent = '⚠️';
      if (title) title.textContent = 'Error de conexión';
      if (btn) btn.disabled = false;
      if (spin) spin.style.display = 'none';
      if (lbl) lbl.textContent = '⚡ Procesar Facturas';
      if (btnCl) btnCl.disabled = false;
    };
  }).catch(function(e) {
    if (icon) icon.textContent = '⚠️';
    if (title) title.textContent = 'Error: ' + e.message;
    if (btn) btn.disabled = false;
    if (spin) spin.style.display = 'none';
    if (lbl) lbl.textContent = '⚡ Procesar Facturas';
    if (btnCl) btnCl.disabled = false;
  });
}

function closeInvoiceModal() {
  var m = document.getElementById('invoice-modal');
  if (m) m.style.display = 'none';
}

function showInvoiceDetail(row) {
  if (!row) return;
  const modal = document.getElementById('invoice-modal');
  const body  = document.getElementById('inv-modal-body');
  const title = document.getElementById('inv-modal-title');
  if (!modal || !body) return;

  title.textContent = row.numero_factura || 'Factura';
  const statusColor = row.estado === 'CORRECTA' ? 'var(--grn)' :
                      row.estado === 'DISCREPANCIA' ? 'var(--red)' : 'var(--ora)';

  const fields = [
    ['OTA / Canal',           row.nombre_ota || '—'],
    ['Hotel',                 row.nombre_hotel || '—'],
    ['Fecha',                 row.fecha || '—'],
    ['Mercado',               row.mercado || '—'],
    ['Importe bruto',         row.importe_bruto ? '€' + row.importe_bruto : '—'],
    ['Comisión pactada %',    row.porcentaje_pactado ? row.porcentaje_pactado + '%' : '—'],
    ['Comisión facturada %',  row.porcentaje_factura ? row.porcentaje_factura + '%' : '—'],
    ['Diferencia €',          row.discrepancia_euros || '0'],
    ['Estado',                row.estado || '—'],
    ['Estado DI',             row.estado_di || '—'],
    ['Período',               (row.periodo_inicio || '—') + ' → ' + (row.periodo_fin || '—')],
  ];

  body.innerHTML =
    '<div style="background:' + statusColor + '20;border:1px solid ' + statusColor + '40;border-radius:10px;padding:10px 14px;margin-bottom:16px;display:flex;align-items:center;gap:8px">' +
      '<span style="font-size:18px">' + (row.estado === 'CORRECTA' ? '✅' : row.estado === 'DISCREPANCIA' ? '⚠️' : '📋') + '</span>' +
      '<span style="color:' + statusColor + ';font-weight:700">' + (row.estado || 'Sin estado') + '</span>' +
    '</div>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">' +
    fields.map(([k, v]) =>
      '<div style="background:var(--bg);border-radius:8px;padding:10px">' +
        '<div style="font-size:10px;color:var(--dim);font-weight:600;text-transform:uppercase;margin-bottom:3px">' + k + '</div>' +
        '<div style="font-size:13px;font-weight:600">' + v + '</div>' +
      '</div>'
    ).join('') +
    '</div>' +
    (row.discrepancia_euros && parseFloat(row.discrepancia_euros) !== 0 ?
      '<div style="margin-top:16px;background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);border-radius:10px;padding:12px">' +
        '<div style="font-size:12px;font-weight:700;color:var(--red);margin-bottom:6px">⚠ Discrepancia detectada</div>' +
        '<div style="font-size:12px;color:var(--mut)">Diferencia: €' + row.discrepancia_euros + ' · Acción recomendada: solicitar factura rectificativa</div>' +
      '</div>'
    : '') +
    '<div style="display:flex;gap:10px;margin-top:16px">' +
      '<button onclick="closeInvoiceModal()" class="btn-ref" style="flex:1">Cerrar</button>' +
      '<button onclick="generarEmailAR(this.getAttribute(\'data-num\'))" data-num="' + (row.numero_factura||'') + '" class="btn-run" style="flex:1;font-size:12px">&#x1F4E7; Generar email</button>' +
    '</div>';

  modal.style.display = 'flex';
}

// ── Copy to clipboard utility ────────────────────────────────────────────
function copyToClip(text, label) {
  navigator.clipboard.writeText(text).then(() => {
    showNotification('✓ Copiado: ' + (label || text.substring(0,30)), 'success');
  }).catch(() => {
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    showNotification('✓ Copiado al portapapeles', 'success');
  });
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.metaKey || e.ctrlKey) return;
  const tabKeys = {'1':'ar_otas','2':'ap','3':'drr','4':'banco','5':'notificaciones','6':'fb_cost','7':'ar_real','8':'calipolis','9':'multi_hotel'};
  if (tabKeys[e.key]) {
    const tabEl = document.getElementById('tab-' + tabKeys[e.key]);
    if (tabEl) { switchTab(tabKeys[e.key], tabEl); }
  }
  if (e.key === '?') { toggleAtajos(); }
  if (e.key === 'r' || e.key === 'R') { loadAll(); }
});

// Redraw tour spotlight on resize
window.addEventListener('resize', function() {
  if (_tourActive) {
    var canvas = document.getElementById('tour-spotlight-canvas');
    if (canvas) { canvas.width = window.innerWidth; canvas.height = window.innerHeight; }
    _showTourStep();
  }
});

// ── Swipe gestures for mobile ────────────────────────────────────────────
(function() {
  var touchStartX = 0, touchStartY = 0;
  var TABS_ORDER = ['ar','ap','drr','banco','notif','fb','ar_real','calipolis','multi_hotel'];
  var currentTabIdx = 0;

  document.addEventListener('touchstart', function(e) {
    touchStartX = e.touches[0].clientX;
    touchStartY = e.touches[0].clientY;
  }, {passive: true});

  document.addEventListener('touchend', function(e) {
    if (!IS_MOBILE) return;
    var dx = e.changedTouches[0].clientX - touchStartX;
    var dy = e.changedTouches[0].clientY - touchStartY;
    // Only horizontal swipes that are more horizontal than vertical
    if (Math.abs(dx) < 60 || Math.abs(dx) < Math.abs(dy) * 1.5) return;
    // Don't swipe if touching a scrollable element
    var el = e.target;
    while (el && el !== document.body) {
      if (el.scrollWidth > el.clientWidth + 10) return;
      el = el.parentElement;
    }
    var activeBtn = document.querySelector('.tab-btn.active');
    if (!activeBtn) return;
    var activeId = activeBtn.id.replace('tab-','');
    var idx = TABS_ORDER.indexOf(activeId);
    if (idx < 0) return;
    var nextIdx = dx < 0 ? Math.min(idx+1, TABS_ORDER.length-1) : Math.max(idx-1, 0);
    if (nextIdx === idx) return;
    var nextTab = TABS_ORDER[nextIdx];
    var nextBtn = document.getElementById('tab-' + nextTab);
    if (nextBtn) switchTab(nextTab, nextBtn);
  }, {passive: true});
})();

// ── Pull-to-refresh on mobile ───────────────────────────────────────────
(function() {
  if (!IS_MOBILE) return;
  var startY = 0, pulling = false;
  var indicator = document.createElement('div');
  indicator.style.cssText = 'position:fixed;top:0;left:50%;transform:translateX(-50%);background:var(--acc);color:#fff;padding:6px 16px;border-radius:0 0 12px 12px;font-size:12px;font-weight:600;z-index:9999;display:none;transition:.2s';
  indicator.textContent = '↓ Suelta para actualizar';
  document.body.appendChild(indicator);

  document.addEventListener('touchstart', function(e) {
    if (window.scrollY === 0) startY = e.touches[0].pageY;
  }, {passive: true});

  document.addEventListener('touchmove', function(e) {
    if (!startY) return;
    var diff = e.touches[0].pageY - startY;
    if (diff > 60) {
      indicator.style.display = 'block';
      pulling = true;
    }
  }, {passive: true});

  document.addEventListener('touchend', function() {
    if (pulling) {
      loadAll();
      showNotification('↺ Actualizando...', 'info');
    }
    pulling = false; startY = 0;
    setTimeout(function() { indicator.style.display = 'none'; }, 500);
  }, {passive: true});
})();

// ── Session timeout warning (45 min) ─────────────────────────────────────
var _sessionTimer;
function _resetSessionTimer() {
  clearTimeout(_sessionTimer);
  _sessionTimer = setTimeout(() => {
    showNotification('⚠️ Tu sesión expirará en 5 minutos por inactividad. Haz clic para renovarla.', 'warning');
    setTimeout(() => { if (confirm('Tu sesión va a expirar. ¿Quieres continuar?')) fetch('/api/health'); }, 4 * 60 * 1000);
  }, 45 * 60 * 1000);
}
['click','keypress','scroll','mousemove'].forEach(ev => document.addEventListener(ev, _resetSessionTimer, {passive:true}));
_resetSessionTimer();


// ══════════════════════════════════════════════════════════════
// MÓDULO AP — JavaScript
// ══════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════
// DEMO MODE
// ═══════════════════════════════════════════════════════════════════
var demoModeActive = false;

async function toggleDemoMode() {
  try {
    const res = await fetch('/api/demo/toggle', {method: 'POST'});
    const data = await res.json();
    demoModeActive = data.demo_mode;
    const btn = document.getElementById('btn-demo');
    const banner = document.getElementById('demo-banner');

    if (demoModeActive) {
      // Show banner + push content down
      if (banner) { banner.style.display = 'block'; document.body.style.paddingTop = '36px'; }
      if (btn) { btn.style.color = '#f59e0b'; btn.querySelector('span') && (btn.querySelector('span').textContent = '🎭 Demo ON'); }
      // Close the menu
      document.getElementById('main-menu')?.classList.remove('open');
      // Switch to Calipolis tab
      setTimeout(() => {
        const calTab = document.getElementById('tab-calipolis');
        if (calTab) switchTab('calipolis', calTab);
        // Load Calipolis data
        loadCalipolis();
        // Show tour prompt after 1.5s
        setTimeout(() => {
          const tourMsg = document.createElement('div');
          tourMsg.style.cssText = 'position:fixed;bottom:80px;right:20px;background:linear-gradient(135deg,#1e293b,#0d1827);border:1px solid rgba(245,158,11,.4);border-radius:14px;padding:18px 20px;z-index:8500;max-width:280px;box-shadow:0 8px 32px rgba(0,0,0,.5)';
          tourMsg.innerHTML = '<div style="font-size:14px;font-weight:700;margin-bottom:8px;color:#f59e0b">🎭 Demo Calipolis</div><div style="font-size:13px;color:#94a3b8;line-height:1.5;margin-bottom:14px">Estás viendo los datos reales del Grupo Calipolis Hotels (3 propiedades, 307 hab.).</div><button onclick="tourStart();this.parentElement.remove()" style="background:linear-gradient(135deg,#3b82f6,#2563eb);border:none;color:#fff;padding:8px 16px;border-radius:8px;cursor:pointer;font-size:12px;font-weight:700;width:100%">🎯 Iniciar Tour Guiado →</button><button onclick="this.parentElement.remove()" style="background:none;border:none;color:#64748b;font-size:11px;margin-top:8px;cursor:pointer;width:100%">Explorar sin tour</button>';
          document.body.appendChild(tourMsg);
          setTimeout(() => tourMsg.remove(), 12000);
        }, 1500);
      }, 300);
    } else {
      if (banner) { banner.style.display = 'none'; document.body.style.paddingTop = ''; }
      if (btn) { btn.style.color = ''; btn.querySelector('span') && (btn.querySelector('span').textContent = '🎭 Demo Mode'); }
    }
  } catch(e) {
    console.error('Error en demo:', e);
  }
}



// ═══════════════════════════════════════════════════════════════════
// SELECTOR DE ROL
// ═══════════════════════════════════════════════════════════════════
var rolActual = 'admin';
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
  window.scrollTo({top:0, behavior:'smooth'});
  // Mobile: scroll tab into view + highlight bottom nav
  if (IS_MOBILE) {
    if (el) setTimeout(function(){ el.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'}); }, 50);
    var bnBtns = document.querySelectorAll('#mobile-bottom-nav button');
    bnBtns.forEach(function(b){ b.style.color = ''; });
    var bnMap = {'ar':0,'ap':1,'drr':2,'calipolis':3};
    if (bnMap[tab] !== undefined && bnBtns[bnMap[tab]]) bnBtns[bnMap[tab]].style.color = 'var(--acc2)';
  }
  if (tab === 'fb') loadFBTab();
  if (tab === 'ar_real') cargarARRealData();
  if (tab === 'drr') loadDRR();
  if (tab === 'banco') loadBanco();
  if (tab === 'notif') loadNotifConfig();
  if (tab === 'calipolis') loadCalipolis();
  if (tab === 'multi_hotel') loadMultiHotel();
}
// ══ F&B COST CONTROL ══════════════════════════════════
function loadFB() { if (typeof cargarFB === 'function') cargarFB(); else if (typeof loadFBCost === 'function') loadFBCost(); }

var _fbLoaded = {resumen:false, inventario:false, mermas:false, recetas:false};
var _fbActive = 'resumen';

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
    html += '<div style="font-size:12px;color:var(--mut);margin-top:3px">' + (t('fb.datosReales')||'Datos calculados desde ventas reales') + ' · ' + data.ventas_diarias.fechas.length + ' ' + (t('fb.dias')||'días') + '</div>';
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
    html += '<th>' + (t('fb.thCategoria')||'Categoría') + '</th><th style="text-align:right">' + (t('fb.thVentas')||'Ventas') + '</th><th style="text-align:right">FC%</th><th style="text-align:center">' + (t('fb.thEstado')||'Estado') + '</th>';
    html += '</tr></thead><tbody id="mh-tbody">';
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
    let html = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;margin-bottom:20px">';
    html += _fbKpi(t('fb.itemsStock')||'Items en Stock', data.items.length, t('fb.ingredientes')||'ingredientes', 'var(--acc2)');
    html += _fbKpi(t('fb.valorInv')||'Valor Inventario', '€' + data.valor_total.toLocaleString('es-ES'), t('fb.valorActual')||'valoración actual', 'var(--grn)');
    html += _fbKpi(t('fb.alertasStock')||'Alertas Stock Bajo', alertas.length, alertas.length > 0 ? 'revisar urgente' : t('fb.todoOk')||'todo OK', alertas.length > 0 ? 'var(--red)' : 'var(--grn)');
    html += '</div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.stockIngredientes">Stock de Ingredientes</div>';
    html += '<div class="tbl-wrap"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>' + (t('fb.thIngrediente')||'Ingrediente') + '</th><th>' + (t('fb.thCategoria')||'Categoría') + '</th><th>' + (t('th.proveedor')||'Proveedor') + '</th>';
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

    // Add summary KPI if total available
    const totalCoste  = data.total_coste  || 0;
    const porCategoria = data.por_categoria || {};
    let html = '';
    if (totalCoste > 0) {
      const topCat = Object.keys(porCategoria)[0] || '—';
      const topVal = Object.values(porCategoria)[0] || 0;
      html += '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px">' +
        '<div class="sc"><div class="sc-lbl" data-tip="Coste total de mermas registradas">COSTE TOTAL MERMAS</div><div class="sc-val" style="color:var(--ora)">€' + totalCoste.toLocaleString('es-ES',{minimumFractionDigits:2}) + '</div></div>' +
        '<div class="sc"><div class="sc-lbl">CATEGORÍA CON MÁS MERMA</div><div class="sc-val" style="font-size:16px;font-weight:700">' + topCat + '</div><div class="sc-sub">€' + topVal.toLocaleString('es-ES',{minimumFractionDigits:2}) + '</div></div>' +
        '<div class="sc"><div class="sc-lbl">REGISTROS</div><div class="sc-val">' + data.mermas.length + '</div></div>' +
        '</div>';
      window._fbCriticalAlerts = data.criticos_count || 0;
    if (totalCoste > 200) {
        html += '<div style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.25);border-radius:10px;padding:12px 16px;font-size:13px;color:var(--ora);margin-bottom:16px">⚠ Mermas altas: €' + totalCoste.toFixed(2) + ' este período. Revisar porcionado y almacenamiento.</div>';
      }
    }
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">';
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
    html += _fbKpi((t('fb.recetas')||'Recetas activas'), data.recetas.length, (t('fb.recetas')||'recetas') + ' en carta', 'var(--acc2)');
    if (data.avg_fc_pct) html += _fbKpi('FC% Medio', data.avg_fc_pct + '%', 'media del menú', data.avg_fc_pct <= 22 ? 'var(--grn)' : 'var(--ora)');
    if (data.best_margin) html += _fbKpi('Mejor margen', data.best_margin.split(' ').slice(0,2).join(' '), 'menor FC%', 'var(--grn)');
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
    document.getElementById('ap-count').textContent = facts.length + ' ' + (t('lbl.facturas')||'facturas');

    facts.forEach(f => {
      const tr = document.createElement('tr');
      tr.setAttribute('data-estado', f.estado || '');
      tr.style.cursor = 'pointer';
      tr.addEventListener('mouseover', function(){ this.style.background='rgba(59,130,246,.04)'; });
      tr.addEventListener('mouseout',  function(){ this.style.background=''; });
      tr.addEventListener('click', function(){ showAPDetail({
        numero_factura: f.numero_factura, proveedor: f.proveedor,
        fecha_factura: f.fecha, base_imponible: f.total_sin_iva || '',
        iva_pct: f.iva_pct || '', importe_con_iva: f.total,
        cuenta_contable: f.cuenta_contable, tipo: f.tipo,
        estado: f.estado, aprobacion: f.accion,
        tiene_po: f.tiene_po, tiene_alb: f.tiene_albarán
      }); });
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

var chatHistory  = [];
var chatOpen     = false;
var chatGreeted  = false;

function toggleChat() {
  chatOpen = !chatOpen;
  const panel = document.getElementById('chat-panel');
  const fab   = document.getElementById('chat-fab');
  panel.classList.toggle('open', chatOpen);
  fab.style.display = chatOpen ? 'none' : 'flex';
  if (chatOpen && !chatGreeted) {
    chatGreeted = true;
    addMsg('bot', '¡Hola! Soy Yve, tu copiloto financiero 👋 — Tengo acceso a todos los datos del hotel. ¿En qué puedo ayudarte?');
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
  // Accept both file input and DataTransfer (drag-drop)
  const file = (input.files || input)[0];
  if (!file) return;
  
  // Hide drag-drop zone, show loading
  const dropZone = document.getElementById('drr-drop-zone');
  if (dropZone) dropZone.style.display = 'none';
  const status  = document.getElementById('drr-status');
  const label   = input.previousElementSibling;
  const origLbl = label ? label.textContent : '';
  status.textContent = '⏳ Procesando ' + file.name + '...';
  if (label) label.textContent = '⏳ Procesando...';

  const form = new FormData();
  form.append('file', file);

  try {
    const resp = await fetch('/api/upload_drr', { method: 'POST', body: form });
    const data = await resp.json();
    if (data.ok && data.stats) {
      const diasStr = data.stats.total_dias || '?';
      const oob     = data.stats.dias_oob   || 0;
      status.textContent = '✓ ' + file.name + ' · ' + diasStr + ' ' + (t('drr.diasLabel')||'días');
      const badge = document.getElementById('drr-oob-badge');
      if (badge) {
        if (oob > 0) { badge.style.display='block'; badge.textContent='⚠ ' + oob + ' OOB'; }
        else           badge.style.display='none';
      }
      renderDRR(data.stats);
      // Also reload daily chart
      const chartResp = await fetch('/api/drr_daily_chart');
      const chartData = await chartResp.json();
      if (chartData && window._drrChartInstance) {
        window._drrChartInstance.data.labels   = chartData.labels;
        window._drrChartInstance.data.datasets[0].data = chartData.revenue;
        window._drrChartInstance.update();
      }
      if (dropZone) dropZone.style.display = 'none';
      showNotification('✓ DRR procesado — ' + diasStr + ' días · ' + oob + ' OOB', 'success');
    } else {
      status.textContent = '✗ Error: ' + (data.error || 'desconocido');
      showNotification('✗ Error procesando DRR: ' + (data.error || ''), 'error');
    }
  } catch(e) {
    status.textContent = '✗ Error de conexión al procesar';
  }
  input.value = '';
  if (label) label.textContent = origLbl;
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
    {key:'GOP', label:'GOP', color:'var(--ora)', tip:'Gross Operating Profit — beneficio bruto antes de deuda e impuestos'},
    {key:'GOP %', label:'GOP %', color:'var(--pur)', tip:'GOP como porcentaje del Revenue Total. 30-45% es saludable en hoteles 4-5★'},
  ];
  const metricsEl = document.getElementById('drr-metrics');
  metricsEl.innerHTML = SHOW.map(m => {
    const d = s.metricas[m.key] || {};
    return '<div class="drr-mc">'
      + '<div class="mc-name" ' + (m.tip ? 'data-tip="' + m.tip + '"' : '') + '>' + m.label + '</div>'
      + '<div class="mc-row"><span class="mc-k">Today</span><span class="mc-v" style="color:' + m.color + '">' + (d.today || 'N/D') + '</span></div>'
      + '<div class="mc-row"><span class="mc-k">MTD</span><span class="mc-v">' + (d.mtd || 'N/D') + '</span></div>'
      + '<div class="mc-row"><span class="mc-k">Forecast</span><span class="mc-v">' + (d.forecast || 'N/D') + '</span></div>'
      + '<div class="mc-row"><span class="mc-k">Budget</span><span class="mc-v" style="color:var(--dim)">'
      + (d.budget || 'N/D') + '</span></div>'
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

  // Show last upload date
  if (s.last_upload) {
    var tag = document.getElementById('oracle-mode-badge');
    var daysMsg = s.days_ago === 0 ? 'hoy' : s.days_ago === 1 ? 'ayer' : 'hace ' + s.days_ago + 'd';
    var uploadEl = document.getElementById('drr-last-upload');
    if (uploadEl) uploadEl.textContent = '📅 Último DRR: ' + s.last_upload + ' (' + daysMsg + ')';
    if (tag && s.days_ago > 3) {
      tag.textContent = '⚠ DRR desactualizado'; tag.style.display='inline';
      tag.style.background='rgba(239,68,68,.15)'; tag.style.color='var(--red)';
    }
  }

  // Budget vs Real bar
  const budgetBarEl = document.getElementById('drr-budget-bar');
  if (budgetBarEl && s.metricas) {
    const rev = s.metricas['Total Revenue'] || {};
    if (rev.today && rev.budget && rev.today !== 'N/D' && rev.budget !== 'N/D') {
      const todayNum = parseFloat(String(rev.today).replace(/[^0-9.]/g,'')) || 0;
      const budgetNum = parseFloat(String(rev.budget).replace(/[^0-9.]/g,'')) || 0;
      if (budgetNum > 0) {
        const pct = Math.round(todayNum / budgetNum * 100);
        const color = pct >= 100 ? 'var(--grn)' : pct >= 85 ? 'var(--ora)' : 'var(--red)';
        budgetBarEl.innerHTML = '<div style="display:flex;align-items:center;gap:12px;font-size:12px">' +
          '<span style="color:var(--mut);white-space:nowrap">Revenue hoy vs Budget:</span>' +
          '<div style="flex:1;background:var(--s2);border-radius:4px;height:8px;overflow:hidden">' +
            '<div style="height:100%;border-radius:4px;background:' + color + ';width:' + Math.min(100, pct) + '%;transition:width .6s ease"></div>' +
          '</div>' +
          '<span style="color:' + color + ';font-weight:700;min-width:36px">' + pct + '%</span>' +
          '</div>';
        budgetBarEl.style.display = 'block';
      }
    }
  }
}

async function runConciliacion() {
  const btn = document.querySelector('button[onclick="runConciliacion()"]');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Conciliando...'; }
  try {
    const r = await fetch('/api/conciliar', {method:'POST'});
    const d = await r.json();
    if (d.ok) {
      showNotification(`✓ Conciliación completada: ${d.conciliados} conciliados, ${d.pendientes} pendientes`, 'success');
      loadBanco();
    } else {
      showNotification('✗ Error conciliación: ' + (d.error||''), 'error');
    }
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
  if (btn) { btn.disabled = false; btn.textContent = '⚡ Conciliar'; }
}
async function aprobarMatchOK() {
  const rows = document.querySelectorAll('#ap-tbody tr[data-estado="MATCH_3WAY_OK"]');
  if (!rows.length) { showNotification('No hay facturas con Match OK pendientes', 'info'); return; }
  // Collect invoice numbers
  const nums = [...rows].map(r => r.cells[0]?.textContent?.trim()).filter(Boolean);
  showNotification('⏳ Aprobando ' + nums.length + ' facturas con Match OK...', 'info');
  try {
    const resp = await fetch('/api/ap/aprobar_lote', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({facturas: nums})
    });
    const d = await resp.json();
    if (d.ok) {
      showNotification('✓ ' + (d.aprobadas || nums.length) + ' facturas aprobadas', 'success');
      setTimeout(loadAP, 500);
    } else {
      showNotification('✗ Error: ' + (d.error || 'desconocido'), 'error');
    }
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
}
function filtrarAPPorEstado(estado) {
  const rows = document.querySelectorAll('#ap-tbody tr[data-estado]');
  rows.forEach(row => {
    const re = row.getAttribute('data-estado') || '';
    row.style.display = (!estado || re === estado) ? '' : 'none';
  });
  const visible = [...rows].filter(r => r.style.display !== 'none').length;
  const countEl = document.getElementById('ap-count');
  if (countEl) countEl.textContent = visible + ' ' + (t('lbl.facturas')||'facturas');
}
function exportarSeleccionados() {
  const selected = [...document.querySelectorAll('.ar-row-cb:checked')]
    .map(cb => cb.closest('tr'))
    .map(row => {
      const cells = row.cells;
      return Array.from(cells).slice(1).map(c => c.textContent.trim()).join('	');
    });
  if (!selected.length) return;
  const headers = ['Archivo','Nº Factura','OTA','Hotel','Fecha','Importe Bruto','% Com.','Estado','Estado DI','Discrepancia','Aprobación'].join('	');
  const content = headers + '\n' + selected.join('\n');
  const blob = new Blob([content], {type: 'text/plain;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'facturas_ar_seleccion_' + new Date().toISOString().slice(0,10) + '.tsv';
  a.click();
}
function toggleSelectAll(master, cbClass) {
  document.querySelectorAll('.' + cbClass).forEach(cb => cb.checked = master.checked);
  updateSelectionCount();
}
function updateSelectionCount() {
  const sel = document.querySelectorAll('.ar-row-cb:checked').length;
  const btn = document.getElementById('btn-export-selected');
  if (btn) { btn.style.display = sel > 0 ? 'inline-block' : 'none'; btn.textContent = '📤 Exportar ' + sel + ' selec.'; }
}
async function loadDRR() {
  // Check Oracle mode
  fetch('/api/oracle/status').then(r=>r.json()).then(d=>{
    var badge = document.getElementById('oracle-mode-badge');
    if (badge) {
      badge.style.display = 'inline';
      badge.textContent = d.mode === 'real' ? '🟢 Oracle Live' : '🟡 Sim';
      badge.style.background = d.mode === 'real' ? 'rgba(34,197,94,.15)' : 'rgba(245,158,11,.15)';
      badge.style.color = d.mode === 'real' ? 'var(--grn)' : 'var(--ora)';
    }
  }).catch(()=>{});
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
  {key:'whatsapp', icon:'💬', name:'WhatsApp', hint:'Requiere cuenta Twilio (TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)'},
  {key:'telegram', icon:'✈️', name:'Telegram', hint:'Requiere TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID en variables de entorno'},
  {key:'slack',    icon:'💼', name:'Slack'},
  {key:'push',     icon:'🔔', name:'Push'},
];
const NOTIF_ALERTAS = [
  {key:'ar_discrepancia',         labelKey:'notif.evAr',    label:'Discrepancia en comisiones OTA (AR)'},
  {key:'ar_falta_di',             labelKey:'notif.evDi',    label:'Falta certificado de doble imposición'},
  {key:'ap_discrepancia',         labelKey:'notif.evAp',    label:'Discrepancia en facturas proveedor (AP)'},
  {key:'drr_oob',                 labelKey:'notif.evDrr',   label:'DRR: días Out of Balance'},
  {key:'banco_sin_conciliar',     labelKey:'notif.evBanco', label:'Movimientos bancarios sin conciliar'},
  {key:'factura_pendiente_firma', labelKey:'notif.evFirma', label:'Facturas pendientes de firma'},
];
var _notifConfig = null;

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
        '<div style="font-size:10px;color:' + (on ? 'var(--grn)' : 'var(--dim)') + ';margin-top:4px">' + (on ? '● ' + (t('notif.activo')||'Activo') : '○ ' + (t('notif.inactivo')||'Inactivo')) + '</div>' +
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
        ' style="width:17px;height:17px;accent-color:var(--acc)">' + (t(a.labelKey)||a.label) + '</label>';
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
    count.textContent = data.length + ' ' + (t('lbl.registros')||'registros');
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

async function testNotification() {
  try {
    const r = await fetch('/api/test_smtp', {method:'POST'});
    const d = await r.json();
    showNotification(d.ok ? '✓ SMTP funcionando: ' + (d.message||'OK') : '✗ SMTP: ' + (d.error||'Error'), d.ok ? 'success' : 'error');
  } catch(e) { showNotification('✗ Error probando SMTP', 'error'); }
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

// ══ MULTI-HOTEL ══════════════════════════════════

async function loadMultiHotel() {
  if (_mh_loaded) return;
  try {
    var r = await fetch('/api/multi_hotel/overview');
    var data = await r.json();
    if (!data.ok) throw new Error(data.error || 'Sin datos');
    var k   = data.consolidado || {};
    var hs  = data.hoteles     || [];

    // ── 4 KPI Cards (2×2 grid, mismo estilo que Calipolis) ─────
    var gopC = k.avg_gop_pct >= 22 ? '#22c55e' : k.avg_gop_pct >= 16 ? '#f59e0b' : '#ef4444';
    var kEl  = document.getElementById('mh-kpis');
    if (kEl) kEl.innerHTML =
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">' +
      [
        {l:'REVENUE MTD',     v:'€' + Math.round((k.total_revenue||0)/1000) + 'K',
         s:(k.n_hoteles||0)+' propiedades · Grupo Calipolis', c:'#60a5fa'},
        {l:'GOP TOTAL',       v:'€' + Math.round((k.total_gop||0)/1000) + 'K',
         s:'GOP% medio: '+(k.avg_gop_pct||0)+'%', c:gopC},
        {l:'OCUPACIÓN MEDIA', v:(k.avg_occ_pct||0)+'%',
         s:'ADR €'+(k.avg_adr||0), c:'#22c55e'},
        {l:'REVPAR MEDIO',    v:'€'+(k.avg_revpar||0),
         s:'Sobre '+(k.total_habitaciones||0)+' hab.', c:'#f1f5f9'},
      ].map(function(c) {
        return '<div class="sc">' +
          '<div class="sc-lbl" style="font-size:9px;letter-spacing:.5px">'+c.l+'</div>' +
          '<div class="sc-val" style="color:'+c.c+';font-size:clamp(20px,4vw,32px);font-weight:900;line-height:1.1;margin:4px 0">'+c.v+'</div>' +
          '<div class="sc-sub" style="font-size:10px;color:var(--dim)">'+c.s+'</div>' +
          '</div>';
      }).join('') + '</div>';

    // ── Smart Insights (mismo estilo que Calipolis) ─────────────
    var iEl = document.getElementById('mh-insights');
    if (iEl && hs.length >= 2) {
      var best = hs.reduce(function(a,b){ return b.gop_pct>a.gop_pct?b:a; }, hs[0]);
      iEl.innerHTML =
        '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:4px">' +
        '<div class="card" style="border-left:3px solid #22c55e;padding:14px 16px">' +
          '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#22c55e;text-transform:uppercase;margin-bottom:6px">MEJOR GOP%</div>' +
          '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:4px">' + best.nombre.split(' ').slice(-1)[0] + '</div>' +
          '<div style="font-size:24px;font-weight:900;color:#22c55e">' + best.gop_pct + '%</div>' +
        '</div>' +
        '<div class="card" style="border-left:3px solid #60a5fa;padding:14px 16px">' +
          '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#60a5fa;text-transform:uppercase;margin-bottom:6px">REVENUE</div>' +
          '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:4px">Junio grupo</div>' +
          '<div style="font-size:24px;font-weight:900;color:#60a5fa">€' + Math.round((k.total_revenue||0)/1000) + 'K</div>' +
        '</div>' +
        '<div class="card" style="border-left:3px solid #a78bfa;padding:14px 16px">' +
          '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#a78bfa;text-transform:uppercase;margin-bottom:6px">GOP% MEDIO</div>' +
          '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:4px">Grupo</div>' +
          '<div style="font-size:24px;font-weight:900;color:#a78bfa">' + (k.avg_gop_pct||0) + '%</div>' +
        '</div>' +
        '</div>';
    }

    // ── Trend charts ─────────────────────────────────────────────
    if (data.rev_trend && window.Chart) {
      var months = data.rev_trend.map(function(r){ return r.mes.slice(5); });
      var gopData = data.rev_trend.map(function(r){ return r.gop; });
      var revData = data.rev_trend.map(function(r){ return Math.round(r.revenue/1000); });

      function makeChart(elId, datasets, yLabel) {
        var el = document.getElementById(elId);
        if (!el) return;
        el.innerHTML = ''; var c = document.createElement('canvas'); el.appendChild(c);
        return new Chart(c, {
          data:{ labels: months, datasets: datasets },
          options:{ responsive:true, maintainAspectRatio:false,
            plugins:{ legend:{labels:{color:'#94a3b8',font:{size:10},boxWidth:10}},
              tooltip:{backgroundColor:'#1e293b',titleColor:'#f1f5f9',bodyColor:'#94a3b8',borderColor:'#334155',borderWidth:1}},
            scales:{
              x:{grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#64748b',font:{size:9}}},
              y:{grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#94a3b8',font:{size:9}},
                 title:{display:true,text:yLabel,color:'#64748b',font:{size:9}},beginAtZero:false}
            }}
        });
      }
      if (window._mhGopChart) { try{window._mhGopChart.destroy();}catch(e){} }
      if (window._mhRevChart) { try{window._mhRevChart.destroy();}catch(e){} }
      window._mhGopChart = makeChart('mh-gop-chart', [{
        type:'line', label:'GOP%', data:gopData,
        borderColor:'#22c55e', backgroundColor:'rgba(34,197,94,.08)',
        tension:.4, pointRadius:4, borderWidth:2.5, fill:true
      }], '%');
      window._mhRevChart = makeChart('mh-rev-chart', [{
        type:'bar', label:'Revenue (k€)', data:revData,
        backgroundColor:'rgba(96,165,250,.2)', borderColor:'#60a5fa', borderWidth:1.5, borderRadius:4
      }], 'k€');
    }

    // ── Hotel cards (MISMO ESTILO QUE SCREENSHOT Calipolis) ──────
    var cardsEl = document.getElementById('mh-hotel-cards');
    if (cardsEl && hs.length) {
      cardsEl.innerHTML = hs.map(function(h) {
        var gc = h.gop_pct>=22?'#22c55e':h.gop_pct>=16?'#f59e0b':'#ef4444';
        var d  = h.rev_delta_pct || 0;
        var spark = _calSparkline(h.gop_trend || [], '#22c55e');
        return '<div class="card" style="padding:20px;position:relative">' +
          // Hotel header
          '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px">' +
            '<div>' +
              '<div style="font-size:15px;font-weight:700">'+h.nombre+'</div>' +
              '<div style="font-size:11px;color:var(--mut);margin-top:3px">'+(h.stars||'4★')+' · '+h.habitaciones+' hab.</div>' +
            '</div>' +
            '<div style="width:9px;height:9px;border-radius:50%;margin-top:4px;background:'+(h.alertas>0?'#ef4444':'#22c55e')+'" title="'+(h.alertas>0?h.alertas+' alertas':'Sin alertas')+'"></div>' +
          '</div>' +
          // Metrics row
          '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:16px">' +
            '<div style="background:var(--bg);border-radius:8px;padding:8px;text-align:center">' +
              '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px">Ocupación</div>' +
              '<div style="font-size:15px;font-weight:700;color:#22c55e">'+h.ocupacion_pct+'%</div>' +
            '</div>' +
            '<div style="background:var(--bg);border-radius:8px;padding:8px;text-align:center">' +
              '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px">ADR</div>' +
              '<div style="font-size:15px;font-weight:700">€'+h.adr_eur+'</div>' +
            '</div>' +
            '<div style="background:var(--bg);border-radius:8px;padding:8px;text-align:center">' +
              '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px">RevPAR</div>' +
              '<div style="font-size:15px;font-weight:700">€'+h.revpar_eur+'</div>' +
            '</div>' +
          '</div>' +
          // GOP% + sparkline
          '<div style="display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:14px">' +
            '<div>' +
              '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px;margin-bottom:3px">GOP%</div>' +
              '<div style="display:flex;align-items:baseline;gap:6px">' +
                '<span style="font-size:24px;font-weight:900;color:'+gc+'">'+h.gop_pct+'%</span>' +
                (d!==0?'<span style="font-size:11px;color:'+(d>0?'#22c55e':'#ef4444')+';font-weight:600">'+(d>0?'+':'')+d+'pp6m</span>':'') +
              '</div>' +
              '<div style="font-size:10px;color:var(--dim);margin-top:2px">€'+Math.round((h.gop_eur||0)/1000)+'K este mes</div>' +
            '</div>' +
            spark +
          '</div>' +
          // Footer
          '<div style="border-top:1px solid var(--s2);padding-top:12px;display:flex;justify-content:space-between;align-items:center;font-size:12px">' +
            '<span style="color:var(--mut)">AP pendientes</span>' +
            '<span style="font-weight:700;color:'+(h.facturas_ap>3?'#f59e0b':'var(--dim)')+'">'+( h.facturas_ap||0)+' facturas</span>' +
          '</div>' +
        '</div>';
      }).join('');
    }

    _mh_loaded = true;
    if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
  } catch(e) {
    console.error('MH Error:', e);
    var el = document.getElementById('mh-kpis');
    if (el) el.innerHTML = '<div style="color:#ef4444;padding:20px;font-size:13px">⚠ Error cargando datos: '+(e.message||e)+'</div>';
  }
}


function loadNotificaciones() { if (typeof cargarNotificaciones === 'function') cargarNotificaciones(); else if (typeof loadNotif === 'function') loadNotif(); }

async function loadBanco() {
  try {
    var r = await fetch('/api/stats_banco');
    var d = await r.json();
    if (!d) return;
    document.getElementById('bk-total').textContent = d.total || '—';
    document.getElementById('bk-conc').textContent = d.conciliados || '0';
    // Conciliation progress bar
    var _bT = d.total||0, _bC = d.conciliados||0, _pEl = document.getElementById('banco-progress-bar');
    if (_pEl && _bT > 0) { var _pct = Math.round(_bC/_bT*100), _col = _pct>=80?'var(--grn)':_pct>=50?'var(--ora)':'var(--red)'; _pEl.style.display='block'; _pEl.innerHTML = '<div style="display:flex;align-items:center;gap:12px;font-size:12px"><span style="color:var(--mut);white-space:nowrap">Conciliado:</span><div style="flex:1;background:var(--s2);border-radius:4px;height:8px;overflow:hidden"><div style="height:100%;border-radius:4px;background:'+_col+';width:'+_pct+'%;transition:width .6s ease"></div></div><span style="color:'+_col+';font-weight:700;min-width:60px">'+_bC+'/'+_bT+' ('+_pct+'%)</span></div>'; }
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
function abrirEmitirFactura() {
  const modal = document.getElementById('modal-emitir');
  modal.style.display = 'flex';
  // Populate client select
  fetch('/api/ar_real/clientes').then(r=>r.json()).then(d => {
    const sel = document.getElementById('ef-cliente');
    sel.innerHTML = '<option value="">Seleccionar cliente...</option>';
    (d.clientes || []).forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.nombre; opt.textContent = c.nombre + ' — ' + (c.NIF||'');
      sel.appendChild(opt);
    });
  }).catch(()=>{});
  // Set default dates
  const hoy = new Date();
  const man = new Date(hoy); man.setDate(man.getDate()+1);
  document.getElementById('ef-entrada').value = hoy.toISOString().slice(0,10);
  document.getElementById('ef-salida').value = man.toISOString().slice(0,10);
  calcularFactura();
}
function cerrarEmitirFactura() {
  document.getElementById('modal-emitir').style.display = 'none';
}
function calcularFactura() {
  const entrada = new Date(document.getElementById('ef-entrada').value);
  const salida  = new Date(document.getElementById('ef-salida').value);
  const noches  = Math.max(1, Math.round((salida - entrada) / 86400000));
  const hab     = parseFloat(document.getElementById('ef-hab').value) || 1;
  const precio  = parseFloat(document.getElementById('ef-precio').value) || 0;
  const fb      = parseFloat(document.getElementById('ef-fb').value) || 0;
  const subHab  = noches * hab * precio;
  const iva     = (subHab + fb) * 0.10;
  const total   = subHab + fb + iva;
  const fmt = v => '€' + v.toLocaleString('es-ES', {minimumFractionDigits:2});
  document.getElementById('ef-sub-hab').textContent = fmt(subHab) + ' (' + noches + ' noche' + (noches!==1?'s':'') + ')';
  document.getElementById('ef-sub-fb').textContent  = fmt(fb);
  document.getElementById('ef-iva').textContent     = fmt(iva);
  document.getElementById('ef-total').textContent   = fmt(total);
}
async function emitirFactura() {
  const cliente = document.getElementById('ef-cliente').value;
  if (!cliente) { alert('Selecciona un cliente'); return; }
  const entrada = document.getElementById('ef-entrada').value;
  const salida  = document.getElementById('ef-salida').value;
  const hab     = parseFloat(document.getElementById('ef-hab').value) || 1;
  const precio  = parseFloat(document.getElementById('ef-precio').value) || 0;
  const fb      = parseFloat(document.getElementById('ef-fb').value) || 0;
  const noches  = Math.max(1, Math.round((new Date(salida)-new Date(entrada))/86400000));
  const subHab  = noches*hab*precio;
  const total   = Math.round((subHab+fb)*1.10*100)/100;
  const msg = document.getElementById('ef-msg');
  msg.style.display='block';msg.style.color='var(--mut)';msg.textContent='Emitiendo factura...';
  try {
    const resp = await fetch('/api/ar_real/emitir_factura', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({cliente,fecha_entrada:entrada,fecha_salida:salida,habitaciones:hab,precio_noche:precio,fb_extras:fb,total})
    });
    const d = await resp.json();
    if (d.ok) {
      msg.style.color='var(--grn)';msg.textContent='✓ Factura '+d.numero+' emitida — €'+total.toLocaleString('es-ES');
      setTimeout(()=>{ cerrarEmitirFactura(); cargarARRealData(); },1500);
    } else {
      msg.style.color='var(--red)';msg.textContent=d.error||'Error emitiendo factura';
    }
  } catch(e) { msg.style.color='var(--red)';msg.textContent='Error de conexión'; }
}

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

async function loadARRealData() {
  cargarARRealData();
}
async function cargarARRealData() {
  try {
    // Load clients and invoices in parallel
    const [rClientes, rFacturas] = await Promise.all([
      fetch('/api/ar_real/clientes'),
      fetch('/api/ar_real/facturas'),
    ]);
    const dc = await rClientes.json();
    const df = await rFacturas.json();

    // Render stats
    if (df.ok && df.stats) {
      const s = df.stats;
      const fmt = v => '\u20AC' + (v||0).toLocaleString('es-ES',{minimumFractionDigits:2});
      _setText('arp-pendiente', fmt(s.pendiente));
      _setText('arp-vencido',   fmt(s.vencido));
      _setText('arp-cobrado',   fmt(s.cobrado_mes));
      _setText('arp-nclientes', dc.ok ? dc.clientes.length + ' activos' : '—');

      // Aging bar
      const agingEl = document.getElementById('ar-aging-bar');
      if (agingEl && s.aging) {
        const total = Object.values(s.aging).reduce((a,b) => a+b, 0) || 1;
        const colors = {'0-30 días':'var(--grn)','31-60 días':'var(--ora)','61-90 días':'var(--red)','>90 días (VENCIDA)':'#7f1d1d'};
        agingEl.style.display = 'block';
        agingEl.innerHTML = '<div style="font-size:10px;color:var(--mut);font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px">Antigüedad de saldo</div>' +
          '<div style="display:flex;height:8px;border-radius:4px;overflow:hidden;gap:2px">' +
          Object.entries(s.aging).filter(([k,v]) => v > 0).map(([k,v]) =>
            '<div style="flex:' + v + ';background:' + (colors[k]||'var(--mut)') + ';border-radius:3px" title="' + k + ': \u20AC' + v.toLocaleString('es-ES',{minimumFractionDigits:0}) + '"></div>'
          ).join('') + '</div>' +
          '<div style="display:flex;gap:12px;margin-top:5px;flex-wrap:wrap">' +
          Object.entries(s.aging).filter(([k,v]) => v > 0).map(([k,v]) =>
            '<span style="font-size:10px;color:var(--dim)"><span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:' + (colors[k]||'var(--mut)') + ';margin-right:3px"></span>' + k + '</span>'
          ).join('') + '</div>';
      }
    }

    // Render client cards
    if (dc.ok && dc.clientes.length) {
      const listEl = document.getElementById('ar-clientes-list');
      if (listEl) {
        listEl.innerHTML = dc.clientes.map(c => {
          const uso = c.uso_credito_pct || 0;
          const usoPct = Math.min(100, uso);
          const usoColor = uso >= 90 ? 'var(--red)' : uso >= 70 ? 'var(--ora)' : 'var(--grn)';
          return '<div class="card" style="padding:12px;cursor:pointer;transition:.15s" ' +
            'onclick="filtrarClienteAR(\'' + c.nombre.replace(/'/g,"\\'") + '\')" ' +
            'onmouseover="this.style.borderColor=\'var(--acc)\'" onmouseout="this.style.borderColor=\'\'"><div style="display:flex;justify-content:space-between;align-items:flex-start">' +
            '<div style="flex:1;min-width:0"><div style="font-weight:600;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + c.nombre.split(' ').slice(0,3).join(' ') + '</div>' +
            '<div style="font-size:10px;color:var(--dim);margin-top:2px">' + c.dias_pago + 'd pago · ' + c.facturas_pendientes + ' fact.</div></div>' +
            '<div style="text-align:right;flex-shrink:0;margin-left:8px">' +
            '<div style="font-size:13px;font-weight:700;color:' + usoColor + '">\u20AC' + (c.saldo_pendiente||0).toLocaleString('es-ES',{minimumFractionDigits:0}) + '</div>' +
            (c.tiene_vencidas ? '<div style="font-size:10px;color:var(--red)">⚠ Vencida</div>' : '') +
            '</div></div>' +
            '<div style="background:var(--s2);border-radius:3px;height:4px;margin-top:8px;overflow:hidden"><div style="height:100%;border-radius:3px;background:' + usoColor + ';width:' + usoPct + '%"></div></div>' +
            '<div style="font-size:9px;color:var(--dim);margin-top:2px">' + uso + '% crédito usado (límite \u20AC' + (c.limite_credito||0).toLocaleString('es-ES') + ')</div>' +
            '</div>';
        }).join('');
      }
    }

    // Render invoice table
    if (df.ok) {
      _renderFacturasAR(df.facturas || [], df.stats);
    }

    if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
  } catch(e) {
    console.error('Error AR Real:', e);
    showNotification('✗ Error cargando AR Real: ' + e.message, 'error');
  }
}

function _setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

var _arAllFacturas = [];
function _renderFacturasAR(facturas, stats) {
  _arAllFacturas = facturas;
  const tbody = document.getElementById('ar-facturas-tbody');
  const countEl = document.getElementById('ar-facturas-count');
  if (!tbody) return;
  if (countEl) countEl.textContent = '(' + facturas.length + ')';

  const estado_filter = (document.getElementById('ar-filter-estado') || {}).value || '';
  const display = estado_filter ? facturas.filter(f => f.estado === estado_filter) : facturas;

  if (!display.length) {
    tbody.innerHTML = '<tr><td colspan="4" style="padding:32px;text-align:center"><div style="font-size:32px;margin-bottom:8px">📋</div><div style="color:var(--mut);font-size:13px">No hay facturas con este filtro</div></td></tr>';
    return;
  }

  tbody.innerHTML = display.map(f => {
    const isVenc = f.days_pending > 60;
    const isOk   = f.estado === 'COBRADO';
    const rowColor = isVenc ? 'rgba(239,68,68,.04)' : '';
    const stateColor = isVenc ? 'var(--red)' : isOk ? 'var(--grn)' : f.estado === 'FACTURADO' ? 'var(--ora)' : 'var(--mut)';
    const stateLabel = {'FACTURADO':'Emitida','COBRADO':'Cobrada','PENDIENTE_FACTURA':'Pendiente'}[f.estado] || f.estado;

    return '<tr style="border-bottom:1px solid var(--s2);background:' + rowColor + '">' +
      '<td style="padding:8px"><div style="font-weight:600;font-size:12px;cursor:pointer;color:var(--acc2)" onclick="copyToClip(\'' + f.numero + '\')">' + f.numero + '</div>' +
        '<div style="font-size:11px;color:var(--mut);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px">' + f.cliente.split(' ').slice(0,2).join(' ') + '</div>' +
        (f.aging_bucket && f.aging_bucket !== 'N/A' ? '<div style="font-size:10px;color:' + (f.days_pending > 60 ? 'var(--red)' : 'var(--dim)') + '">' + f.aging_bucket + '</div>' : '') +
      '</td>' +
      '<td style="text-align:right;padding:8px;font-weight:700">\u20AC' + (f.total||0).toLocaleString('es-ES',{minimumFractionDigits:2}) + '</td>' +
      '<td style="padding:8px"><span style="background:' + stateColor + '20;color:' + stateColor + ';padding:2px 8px;border-radius:6px;font-size:10px;font-weight:700">' + stateLabel + '</span></td>' +
      '<td style="padding:8px;white-space:nowrap">' +
        (f.estado === 'FACTURADO' ? '<button onclick="cobrarFacturaAR(\'' + f.numero + '\')" class="btn bsm" style="font-size:10px;margin-right:4px;background:rgba(34,197,94,.1);color:var(--grn);border-color:rgba(34,197,94,.3)" title="Marcar como cobrada">💰</button>' : '') +
        (f.estado === 'FACTURADO' ? '<a href="/api/ar_real/pdf/' + encodeURIComponent(f.numero) + '" target="_blank" class="btn bsm" style="font-size:10px;text-decoration:none;background:rgba(59,130,246,.1);color:var(--acc2);border-color:rgba(59,130,246,.3)" title="Descargar PDF">📄</a>' : '') +
        (f.estado === 'FACTURADO' ? '<button onclick="recordatorioAR(\'' + f.numero + '\')" class="btn bsm" style="font-size:10px;background:rgba(245,158,11,.1);color:var(--ora);border-color:rgba(245,158,11,.3)" title="Enviar recordatorio email">📧</button>' : '') +
      '</td>' +
    '</tr>';
  }).join('');
}

function filtrarClienteAR(nombre) {
  // Filter invoices by client
  const sel = document.getElementById('ar-filter-estado');
  if (sel) sel.value = '';
  const filtered = _arAllFacturas.filter(f => f.cliente === nombre);
  _renderFacturasAR(filtered, null);
  showNotification(nombre.split(' ').slice(0,2).join(' ') + ' — ' + filtered.length + ' facturas', 'info');
}

function filtrarFacturasAR(estado) {
  _renderFacturasAR(_arAllFacturas, null);
}

async function cobrarFacturaAR(numero) {
  if (!confirm('¿Marcar factura ' + numero + ' como COBRADA?')) return;
  try {
    const r = await fetch('/api/ar_real/cobrar', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({numero})});
    const d = await r.json();
    if (d.ok) { showNotification('✓ ' + d.message, 'success'); cargarARRealData(); }
    else showNotification('✗ ' + (d.error||'Error'), 'error');
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
}

async function recordatorioAR(numero) {
  try {
    const r = await fetch('/api/ar_real/recordatorio', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({numero})});
    const d = await r.json();
    showNotification(d.ok ? '✓ ' + d.message : '✗ ' + (d.error||'Error'), d.ok ? 'success' : 'error');
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
}

async function procesarARReal() {
  cargarARRealData();
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
var _calCharts = {};

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
    addCalipolisInsights(data.hoteles);
    if (data.tendencias) renderCalipolisTrends(data.tendencias);
    const calUpd = document.getElementById('cal-updated');
    if (calUpd) calUpd.textContent = 'Actualizado ' + new Date().toLocaleTimeString('es-ES', {hour:'2-digit',minute:'2-digit'});
  } catch(e) {
    console.error('Error Calipolis:', e);
    const kpiEl = document.getElementById('cal-kpis');
    if (kpiEl) kpiEl.innerHTML = '<div style="color:var(--red);font-size:12px;padding:10px">⚠ Error cargando datos: ' + (e.message||e) + '</div>';
  }
}

function _calSparkline(data, color) {
  if (!data || !data.length) return '<div style="width:80px;height:40px"></div>';
  var max = Math.max(...data, 0.1);
  var min = Math.min(...data);
  var pts = data.map(function(v, i) {
    var x = (i / (data.length-1)) * 80;
    var y = 40 - ((v-min)/(max-min||1)) * 36 - 2;
    return x.toFixed(1)+','+y.toFixed(1);
  }).join(' ');
  return '<svg width="80" height="40" style="overflow:visible"><polyline points="'+pts+'" fill="none" stroke="'+color+'" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
}

function renderCalipolisTrends(tendencias) {
  if (!tendencias) return;
  var meses  = tendencias.meses || [];
  var gop    = tendencias.gop_mensual  || tendencias.gop_pct_grupo || [];
  var occ    = tendencias.occ_mensual  || [];
  var rev    = (tendencias.rev_mensual || tendencias.total_revenue || []).map(function(v){ return Math.round(v/1000); });
  var apData = tendencias.ap_pendientes || tendencias.ap_pendientes_total || [];
  var gopEl = document.getElementById('cal-tendencias');
  if (gopEl && meses.length && window.Chart) {
    if (window._calChart) { try { window._calChart.destroy(); } catch(e){} }
    var ctx1 = document.createElement('canvas');
    ctx1.style.cssText = 'width:100%;height:140px';
    gopEl.innerHTML = ''; gopEl.appendChild(ctx1);
    window._calChart = new Chart(ctx1, {
      data: {
        labels: meses,
        datasets: [
          {type:'line', label:'GOP%', data:gop, borderColor:'#22c55e', backgroundColor:'rgba(34,197,94,.08)',
           yAxisID:'y', tension:.4, pointRadius:4, borderWidth:2.5, fill:true},
          {type:'line', label:'Occ%', data:occ, borderColor:'#60a5fa', backgroundColor:'transparent',
           yAxisID:'y', tension:.4, pointRadius:3, borderWidth:1.5, borderDash:[4,3]},
          {type:'bar', label:'Rev(k€)', data:rev, backgroundColor:'rgba(167,139,250,.15)',
           borderColor:'rgba(167,139,250,.4)', borderWidth:1, yAxisID:'y2', borderRadius:4},
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{legend:{labels:{color:'#94a3b8',font:{size:10},boxWidth:10}},
          tooltip:{backgroundColor:'#1e293b',titleColor:'#f1f5f9',bodyColor:'#94a3b8',borderColor:'#334155',borderWidth:1}},
        scales:{
          x:{grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#64748b',font:{size:9}}},
          y:{position:'left',grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#94a3b8',font:{size:9}},
             title:{display:true,text:'%',color:'#64748b',font:{size:9}}},
          y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:'#94a3b8',font:{size:9}},
              title:{display:true,text:'k€',color:'#64748b',font:{size:9}}}
        }
      }
    });
  }
  var apEl = document.getElementById('cal-ap-chart');
  if (apEl && apData.length && window.Chart) {
    if (window._calApChart) { try { window._calApChart.destroy(); } catch(e){} }
    var ctx2 = document.createElement('canvas');
    ctx2.style.cssText = 'width:100%;height:140px';
    apEl.innerHTML = ''; apEl.appendChild(ctx2);
    window._calApChart = new Chart(ctx2, {
      type:'bar',
      data:{labels:meses,datasets:[{label:'AP pendientes',data:apData,backgroundColor:'rgba(245,158,11,.2)',borderColor:'#f59e0b',borderWidth:1.5,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#64748b',font:{size:9}}},y:{grid:{color:'rgba(51,65,85,.2)'},ticks:{color:'#94a3b8',font:{size:9}},beginAtZero:true}}}
    });
  }
}

function addCalipolisInsights(hoteles) {
  var el = document.getElementById('cal-insights');
  if (!el || !hoteles || !hoteles.length) return;
  var top = hoteles.reduce(function(a,b){ return a.gop_pct > b.gop_pct ? a : b; });
  var avg = Math.round((hoteles.reduce(function(s,h){ return s+h.gop_pct; },0) / hoteles.length) * 10) / 10;
  var totalRev = hoteles.reduce(function(s,h){ return s + (h.total_ingresos||0); }, 0);

  el.innerHTML =
    '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px">' +
    // Panel 1: Mejor GOP%
    '<div class="card" style="border-left:3px solid #22c55e;padding:16px 18px">' +
      '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#22c55e;text-transform:uppercase;margin-bottom:8px">MEJOR GOP%</div>' +
      '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:6px">' + (top.nombre||'').split(' ').slice(-1)[0] + '</div>' +
      '<div style="font-size:26px;font-weight:900;color:#22c55e">' + top.gop_pct + '%</div>' +
    '</div>' +
    // Panel 2: Revenue
    '<div class="card" style="border-left:3px solid #60a5fa;padding:16px 18px">' +
      '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#60a5fa;text-transform:uppercase;margin-bottom:8px">REVENUE</div>' +
      '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:6px">Junio grupo</div>' +
      '<div style="font-size:26px;font-weight:900;color:#60a5fa">€' + Math.round(totalRev/1000) + 'K</div>' +
    '</div>' +
    // Panel 3: GOP% medio
    '<div class="card" style="border-left:3px solid #a78bfa;padding:16px 18px">' +
      '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#a78bfa;text-transform:uppercase;margin-bottom:8px">GOP% MEDIO</div>' +
      '<div style="font-size:14px;font-weight:700;color:var(--tx);margin-bottom:6px">Grupo</div>' +
      '<div style="font-size:26px;font-weight:900;color:#a78bfa">' + avg + '%</div>' +
    '</div>' +
    '</div>';
}


function renderCalipolisKpis(kpis) {
  var cont = document.getElementById('cal-kpis');
  if (!cont || !kpis) return;
  cont.dataset.loaded = '1';
  var totalRevM = ((kpis.total_revenue_mtd||0) / 1000000).toFixed(2);
  var totalGopK = Math.round((kpis.total_gop||0) / 1000);
  var gopColor  = kpis.avg_gop_pct >= 22 ? 'var(--grn)' : kpis.avg_gop_pct >= 18 ? 'var(--ora)' : 'var(--red)';
  var occColor  = kpis.avg_ocupacion >= 80 ? 'var(--grn)' : kpis.avg_ocupacion >= 65 ? 'var(--ora)' : 'var(--red)';
  var cards = [
    {l:'REVENUE MTD',     v:'€' + totalRevM + 'M', s:(kpis.num_hoteles||3)+' propiedades · Grupo Calipolis', c:'var(--acc2)'},
    {l:'GOP TOTAL',       v:'€' + totalGopK + 'K', s:'GOP% medio: '+(kpis.avg_gop_pct||0)+'%',              c:gopColor},
    {l:'OCUPACIÓN MEDIA', v:(kpis.avg_ocupacion||0)+'%',  s:'ADR €'+(kpis.avg_adr||0),                      c:occColor},
    {l:'REVPAR MEDIO',    v:'€'+(kpis.avg_revpar||0),      s:'Sobre '+(kpis.total_rooms||0)+' hab.',          c:'var(--tx)'},
  ];
  cont.innerHTML = cards.map(function(c) {
    return '<div class="sc">' +
      '<div class="sc-lbl" style="font-size:9px;text-transform:uppercase;letter-spacing:.5px">'+c.l+'</div>' +
      '<div class="sc-val" style="color:'+c.c+';font-size:clamp(22px,3.5vw,36px);font-weight:900;line-height:1.1;margin:5px 0">'+c.v+'</div>' +
      '<div class="sc-sub" style="font-size:10px;color:var(--dim)">'+c.s+'</div>' +
      '</div>';
  }).join('');
  if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
}

function renderCalipolisHoteles(hoteles) {
  const cont = document.getElementById('cal-hoteles');
  if (!cont) return;
  cont.innerHTML = '';
  hoteles.forEach((h, idx) => {
    const sc = h.status === 'ok' ? 'var(--grn)' : h.status === 'warning' ? 'var(--ora)' : 'var(--red)';
    const gopColor = h.gop_pct >= 22 ? 'var(--grn)' : h.gop_pct >= 18 ? 'var(--ora)' : 'var(--red)';
    const apColor = h.ap_pendientes === 0 ? 'var(--grn)' : h.ap_pendientes <= 3 ? 'var(--ora)' : 'var(--red)';
    const card = document.createElement('div');
    card.style.cssText = 'background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:20px;transition:border-color .18s,transform .18s;cursor:pointer';
    card.addEventListener('mouseover', () => { card.style.borderColor='rgba(59,130,246,.4)'; card.style.transform='translateY(-2px)'; });
    card.addEventListener('mouseout',  () => { card.style.borderColor='var(--s2)'; card.style.transform=''; });

    // Build GOP sparkline SVG if trend data available
    let sparkSvg = '';
    if (h.gop_trend && h.gop_trend.length > 1) {
      const vals = h.gop_trend;
      const min = Math.min(...vals), max = Math.max(...vals) || 1;
      const w = 80, ht = 30;
      const pts = vals.map((v, i) => {
        const x = (i / (vals.length-1)) * w;
        const y = ht - ((v - min) / (max - min + 0.001)) * (ht - 4);
        return x + ',' + y;
      }).join(' ');
      const lastColor = vals[vals.length-1] >= 20 ? '#22c55e' : '#f97316';
      sparkSvg = '<svg width="' + w + '" height="' + ht + '" viewBox="0 0 ' + w + ' ' + ht + '" style="overflow:visible">' +
        '<polyline points="' + pts + '" fill="none" stroke="' + lastColor + '" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>' +
        '<circle cx="' + (w) + '" cy="' + (ht - ((vals[vals.length-1]-min)/(max-min+0.001))*(ht-4)) + '" r="3" fill="' + lastColor + '"/>' +
        '</svg>';
    }
    const gopDelta = h.gop_trend && h.gop_trend.length > 1
      ? (h.gop_trend[h.gop_trend.length-1] - h.gop_trend[0]).toFixed(1)
      : null;
    const deltaEl = gopDelta !== null
      ? '<span style="font-size:10px;color:' + (gopDelta > 0 ? 'var(--grn)' : 'var(--red)') + ';margin-left:6px">' + (gopDelta > 0 ? '+' : '') + gopDelta + 'pp 6m</span>'
      : '';

    card.innerHTML =
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px">' +
        '<div>' +
          '<div style="font-weight:700;font-size:14px;margin-bottom:3px">' + h.nombre + '</div>' +
          '<div style="font-size:11px;color:var(--mut)">' + h.categoria + ' · ' + h.habitaciones + ' hab.</div>' +
        '</div>' +
        '<div style="width:9px;height:9px;border-radius:50%;background:' + sc + ';margin-top:4px;flex-shrink:0;box-shadow:0 0 6px ' + sc + '"></div>' +
      '</div>' +
      // KPI row
      '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px">' +
        _calKpi('Ocupación', h.ocupacion + '%', h.ocupacion >= 80 ? 'var(--grn)' : 'var(--ora)') +
        _calKpi('ADR', '€' + h.adr, 'var(--acc2)') +
        _calKpi('RevPAR', '€' + h.revpar, 'var(--tx)') +
      '</div>' +
      // GOP with sparkline
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:10px;background:var(--bg);border-radius:10px;margin-bottom:10px">' +
        '<div>' +
          '<div style="font-size:10px;color:var(--mut);font-weight:600;text-transform:uppercase;letter-spacing:.4px">GOP%</div>' +
          '<div style="font-size:22px;font-weight:800;color:' + gopColor + ';letter-spacing:-1px;line-height:1">' + h.gop_pct + '%' + deltaEl + '</div>' +
          '<div style="font-size:10px;color:var(--dim);margin-top:2px">€' + Math.round(h.gop/1000) + 'K este mes</div>' +
        '</div>' +
        '<div>' + sparkSvg + '</div>' +
      '</div>' +
      // AP pendientes
      '<div style="display:flex;align-items:center;justify-content:space-between;font-size:12px">' +
        '<span style="color:var(--mut)">AP pendientes</span>' +
        '<span style="font-weight:700;color:' + apColor + '">' + h.ap_pendientes + (h.ap_pendientes === 0 ? ' ✓' : ' facturas') + '</span>' +
      '</div>';
    cont.appendChild(card);
  });
}

function _calKpi(label, val, color) {
  return '<div style="background:var(--bg);border-radius:8px;padding:8px;text-align:center">' +
    '<div style="font-size:10px;color:var(--mut);font-weight:600;margin-bottom:3px">' + label + '</div>' +
    '<div style="font-size:15px;font-weight:700;color:' + color + '">' + val + '</div>' +
    '</div>';
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

<!-- Search result keyboard hint -->
<div id="search-hint-bar" style="display:none;position:fixed;bottom:0;left:0;right:0;background:#0f172a;border-top:1px solid #1e293b;padding:8px 16px;font-size:11px;color:#64748b;z-index:8999;display:flex;gap:16px">
  <span>↑↓ Navegar</span>
  <span>↵ Abrir</span>
  <span>Esc Cerrar</span>
</div>

<!-- Keyboard Shortcuts Modal -->
<div id="atajos-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9500;align-items:center;justify-content:center">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;max-width:480px;width:90%;max-height:85vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <h3 style="font-size:16px;font-weight:700;margin:0">⌨ Atajos de teclado</h3>
      <button onclick="toggleAtajos()" style="background:none;border:none;color:var(--mut);font-size:22px;cursor:pointer">×</button>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px">
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">1 – 9</kbd><div style="color:var(--mut);margin-top:5px">Cambiar de pestaña</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">R</kbd><div style="color:var(--mut);margin-top:5px">Actualizar datos</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">Ctrl+K</kbd><div style="color:var(--mut);margin-top:5px">Búsqueda global</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">Ctrl+/</kbd><div style="color:var(--mut);margin-top:5px">Abrir chat Yve</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">F1</kbd><div style="color:var(--mut);margin-top:5px">Tour guiado</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">?</kbd><div style="color:var(--mut);margin-top:5px">Mostrar atajos</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">Esc</kbd><div style="color:var(--mut);margin-top:5px">Cerrar modales</div></div>
      <div style="background:var(--bg);border-radius:8px;padding:10px"><kbd style="background:var(--s2);border-radius:4px;padding:2px 7px;font-size:11px;color:var(--tx)">←  →</kbd><div style="color:var(--mut);margin-top:5px">Tour: paso anterior/siguiente</div></div>
    </div>
    <div style="margin-top:16px;background:var(--bg);border-radius:8px;padding:12px;font-size:12px;color:var(--dim)">
      📱 En móvil: desliza izquierda/derecha para cambiar pestaña · Tira hacia abajo para actualizar
    </div>
  </div>
</div>

<!-- Invoice Detail Modal -->
<div id="invoice-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9000;align-items:center;justify-content:center">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;max-width:500px;width:90%;max-height:85vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <h3 id="inv-modal-title" style="font-size:16px;font-weight:700;margin:0">Detalle factura</h3>
      <button onclick="document.getElementById('invoice-modal').style.display='none'" style="background:none;border:none;color:var(--mut);font-size:22px;cursor:pointer">×</button>
    </div>
    <div id="inv-modal-body"></div>
  </div>
</div>

<!-- Global Search Overlay -->
<div id="search-overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9500;padding:80px 20px 20px;backdrop-filter:blur(4px)" onclick="if(event.target===this)closeSearch()">
  <div style="max-width:600px;margin:0 auto">
    <div style="display:flex;align-items:center;gap:12px;background:var(--s1);border:1px solid var(--acc);border-radius:14px;padding:14px 18px;margin-bottom:12px">
      <span style="font-size:18px">🔍</span>
      <input id="search-input" placeholder="Buscar facturas, proveedores, métricas..." 
        style="flex:1;background:none;border:none;color:var(--tx);font-size:15px;outline:none"
        oninput="runSearch(this.value)" onkeydown="if(event.key==='Escape')closeSearch()">
      <kbd style="background:var(--s2);border:1px solid var(--s2);color:var(--dim);border-radius:5px;padding:2px 7px;font-size:11px">ESC</kbd>
    </div>
    <div id="search-results" style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;overflow:hidden;max-height:400px;overflow-y:auto"></div>
    <div style="font-size:11px;color:var(--dim);margin-top:8px;text-align:center">
      Ctrl+K para abrir · ESC para cerrar · Enter para ir a la sección
    </div>
  </div>
</div>

<script>
// Global search
const SEARCH_INDEX = [
  {q:['ar','ota','booking','expedia','comision','factura ar'],tab:'ar_otas',   label:'AR — OTAs',           desc:'Facturas y comisiones de agencias online'},
  {q:['ap','proveedor','supplier','matching','pago','factura ap'],tab:'ap',    label:'AP — Proveedores',    desc:'Facturas de proveedores y 3-way matching'},
  {q:['drr','revenue','ingresos','oob','out of balance','gop'],tab:'drr',     label:'DRR',                 desc:'Daily Revenue Report y métricas del hotel'},
  {q:['banco','bank','conciliacion','extracto','pago'],tab:'banco',           label:'Banco',               desc:'Conciliación bancaria y movimientos'},
  {q:['notif','alerta','email','slack','whatsapp'],tab:'notificaciones',       label:'Notificaciones',      desc:'Canales y configuración de alertas'},
  {q:['fb','food cost','restaurante','merma','receta','inventario'],tab:'fb_cost', label:'F&B Cost',        desc:'Control de coste de alimentos y bebidas'},
  {q:['ar real','grupo','corporativo','cliente','beo'],tab:'ar_real',         label:'AR Real',             desc:'Facturación a clientes corporativos'},
  {q:['calipolis','sitges','multi','grupo hotelero'],tab:'calipolis',         label:'Calipolis',           desc:'Dashboard del Grupo Calipolis Hotels'},
  {q:['multi','hotel','consolidado','grupo'],tab:'multi_hotel',              label:'Multi-Hotel',         desc:'Vista consolidada de todos los hoteles'},
  {q:['admin','usuario','configuracion','cuenta'],link:'/admin/',            label:'Administración',      desc:'Panel de administración y usuarios'},
  {q:['precio','plan','stripe','pago','billing'],link:'/checkout/starter',  label:'Planes y precios',    desc:'Contratar o cambiar el plan'},
  {q:['terminos','privacidad','legal','gdpr'],link:'/terminos',             label:'Términos legales',    desc:'Términos de uso y política de privacidad'},
];

function openSearch() {
  document.getElementById('search-overlay').style.display = 'block';
  setTimeout(() => document.getElementById('search-input').focus(), 50);
}
function closeSearch() {
  document.getElementById('search-overlay').style.display = 'none';
  document.getElementById('search-input').value = '';
  document.getElementById('search-results').innerHTML = '';
}
function runSearch(q) {
  const results = document.getElementById('search-results');
  if (!q.trim()) { results.innerHTML = ''; return; }
  const ql = q.toLowerCase();
  const matches = SEARCH_INDEX.filter(item => item.q.some(kw => kw.includes(ql) || ql.includes(kw)));
  if (!matches.length) {
    results.innerHTML = '<div style="padding:20px;text-align:center;color:var(--dim);font-size:13px">Sin resultados para "' + q + '"</div>';
    return;
  }
  results.innerHTML = matches.map(m => {
    const action = m.tab
      ? "switchTab('" + m.tab + "',document.getElementById('tab-" + m.tab + "'));closeSearch()"
      : "location.href='" + m.link + "'";
    return '<div onclick="' + action + '" class="sr-item">' +
      '<div style="font-size:20px">🔹</div>' +
      '<div><div style="font-weight:600;font-size:14px">' + m.label + '</div>' +
      '<div style="font-size:12px;color:var(--mut)">' + m.desc + '</div></div>' +
      '</div>';
  }).join('');
}
// Keyboard shortcut Ctrl+K
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'k') { e.preventDefault(); openSearch(); }
  if ((e.ctrlKey || e.metaKey) && e.key === '/') { e.preventDefault(); toggleChat(); }
  if (e.key === 'F1' || (e.shiftKey && e.key === 'T')) { e.preventDefault(); startTour(); }
  if (_tourActive) {
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { e.preventDefault(); nextTourStep(); }
    if (e.key === 'ArrowLeft'  || e.key === 'ArrowUp')   { e.preventDefault(); prevTourStep(); }
  }
});
</script>
<!-- /Global Search -->
<!-- Changelog Modal -->
<div id="changelog-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9000;align-items:center;justify-content:center;padding:20px" onclick="if(event.target===this)this.style.display='none'">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;max-width:500px;width:100%;max-height:80vh;overflow-y:auto">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
      <h3 style="font-size:16px;font-weight:700">🆕 Novedades en Yve.01</h3>
      <button onclick="document.getElementById('changelog-modal').style.display='none'" style="background:none;border:none;color:var(--mut);font-size:20px;cursor:pointer">×</button>
    </div>
    <div style="display:flex;flex-direction:column;gap:16px;font-size:13px">
      <div>
        <div style="color:#60a5fa;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px">Junio 2026</div>
        <ul style="color:#94a3b8;padding-left:16px;line-height:1.8">
          <li>🏢 AR Real: flujo completo — clientes, aging, cobro, recordatorios</li>
          <li>📄 AR Real: generación de PDF por factura</li>
          <li>📱 Mobile: barra de navegación inferior, swipe de pestañas</li>
          <li>🔄 Mobile: pull-to-refresh y KPI quick-view</li>
          <li>⚡ Banco: botón conciliación automática</li>
          <li>🎯 Tour guiado de 6 pasos (menú ⋯)</li>
          <li>📊 Calipolis: gráfico multi-métrica (GOP%+Occ+Revenue)</li>
          <li>💡 Smart Insights en panel Calipolis</li>
          <li>🔔 Status bar con briefing matutino y animación</li>
          <li>🔍 Búsqueda global Ctrl+K</li>
          <li>⌨ Atajos de teclado 1-9 + swipe</li>
          <li>⚖️ Páginas legales RGPD completas</li>
          <li>📝 Blog SEO con 10 artículos</li>
        </ul>
      </div>
      <div>
        <div style="color:#a78bfa;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px">Mayo 2026</div>
        <ul style="color:#94a3b8;padding-left:16px;line-height:1.8">
          <li>🌍 Multi-idioma: ES/EN/CA/FR/DE/IT/PT</li>
          <li>🎭 Demo Mode con tour guiado Calipolis</li>
          <li>💳 Integración Stripe (simulación)</li>
          <li>📊 Blog SEO con 6 artículos</li>
          <li>📱 Responsive móvil mejorado</li>
        </ul>
      </div>
    </div>
  </div>
</div>
<script>function showChangelog() {
  localStorage.setItem('changelog_seen', '2026-06-v2');
  const badge = document.getElementById('menu-badge');
  if (badge) badge.style.display = 'none';
  document.getElementById('changelog-modal').style.display='flex'; document.getElementById('main-menu').classList.remove('open'); }</script>

<!-- Setup Checklist Modal -->
<div id="checklist-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9000;align-items:center;justify-content:center;padding:20px">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;max-width:460px;width:100%">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
      <h3 style="font-size:16px;font-weight:700">✅ Checklist de configuración</h3>
      <button onclick="document.getElementById('checklist-modal').style.display='none'" style="background:none;border:none;color:var(--mut);font-size:20px;cursor:pointer">×</button>
    </div>
    <div id="checklist-items" style="display:flex;flex-direction:column;gap:10px;font-size:13px"></div>
    <p style="font-size:11px;color:var(--dim);margin-top:16px">Completa estos pasos para sacar el máximo partido a Yve.01</p>
  </div>
</div>
<script>
function showSetupChecklist() {
  document.getElementById('checklist-modal').style.display = 'flex';
  document.getElementById('main-menu').classList.remove('open');
  const items = [
    {label:'Configurar SMTP para notificaciones email', check: () => true, link:'/admin/', action:'Ir a Admin → Conexiones'},
    {label:'Subir primer DRR (.xlsm)', check: () => document.getElementById('drr-status')?.textContent?.includes('días'), link:null, action:'Tab DRR → Subir DRR'},
    {label:'Procesar facturas AR (⚡)', check: () => (parseInt(document.getElementById('sc-procesadas')?.textContent)||0) > 0, link:null, action:'Tab AR → Procesar Facturas'},
    {label:'Revisar discrepancias AP', check: () => document.getElementById('tab-ap')?.classList?.contains('active'), link:null, action:'Tab AP'},
    {label:'Configurar canal de notificaciones', check: () => localStorage.getItem('notif_configured'), link:null, action:'Tab Notificaciones'},
    {label:'Probar el tour guiado', check: () => localStorage.getItem('tour_done'), link:null, action:'Menú ⋯ → Demo Mode → Tour'},
  ];
  document.getElementById('checklist-items').innerHTML = items.map(it => {
    const done = it.check();
    return '<div style="display:flex;align-items:center;gap:12px;padding:10px;background:var(--bg);border-radius:8px;border:1px solid ' + (done ? 'rgba(34,197,94,.2)' : 'var(--s2)') + '">' +
      '<span style="font-size:18px">' + (done ? '✅' : '⬜') + '</span>' +
      '<div style="flex:1"><div style="font-weight:' + (done ? '400' : '600') + ';color:' + (done ? 'var(--mut)' : 'var(--tx)') + '">' + it.label + '</div>' +
      (done ? '' : '<div style="font-size:11px;color:var(--dim)">' + it.action + '</div>') +
      '</div>' +
      '</div>';
  }).join('');
}
</script>

<!-- Floating Action Button - Ask Yve AI -->

<!-- Mobile bottom nav -->
<div id="mobile-bottom-nav" style="display:none;position:fixed;bottom:0;left:0;right:0;background:var(--s1);border-top:1px solid var(--s2);z-index:600;padding:6px 0 env(safe-area-inset-bottom,0)">
  <div style="display:flex;justify-content:space-around;max-width:500px;margin:0 auto">
    <button onclick="switchTab('ar',document.getElementById('tab-ar'))" style="background:none;border:none;color:var(--mut);padding:6px 12px;cursor:pointer;flex:1;text-align:center;font-size:10px">
      <div style="font-size:18px">💳</div><div>AR</div>
    </button>
    <button onclick="switchTab('ap',document.getElementById('tab-ap'))" style="background:none;border:none;color:var(--mut);padding:6px 12px;cursor:pointer;flex:1;text-align:center;font-size:10px">
      <div style="font-size:18px">📦</div><div>AP</div>
    </button>
    <button onclick="switchTab('drr',document.getElementById('tab-drr'))" style="background:none;border:none;color:var(--mut);padding:6px 12px;cursor:pointer;flex:1;text-align:center;font-size:10px">
      <div style="font-size:18px">📊</div><div>DRR</div>
    </button>
    <button onclick="switchTab('calipolis',document.getElementById('tab-calipolis'))" style="background:none;border:none;color:var(--mut);padding:6px 12px;cursor:pointer;flex:1;text-align:center;font-size:10px">
      <div style="font-size:18px">🏨</div><div>Grupo</div>
    </button>
    <button onclick="toggleChat()" style="background:none;border:none;color:var(--mut);padding:6px 12px;cursor:pointer;flex:1;text-align:center;font-size:10px">
      <div style="font-size:18px">💬</div><div>Yve</div>
    </button>
  </div>
</div>
<script>
// Show mobile nav on small screens
if (window.innerWidth <= 768) {
  var mbn = document.getElementById('mobile-bottom-nav');
  if (mbn) mbn.style.display = 'block';
  // Add padding to main content to avoid overlap
  var mainContent = document.querySelector('.tabs');
  document.body.style.paddingBottom = '60px';
}
</script>

<button id="mobile-lite-toggle" onclick="toggleMobileLite()" title="Cambiar entre vista resumida y completa">
  📱 Vista completa
</button>
<button id="back-top" onclick="window.scrollTo({top:0,behavior:'smooth'})" 
  style="display:none;position:fixed;bottom:88px;right:20px;background:var(--s1);border:1px solid var(--s2);color:var(--mut);width:36px;height:36px;border-radius:50%;font-size:16px;cursor:pointer;z-index:500;transition:.2s;box-shadow:0 2px 8px rgba(0,0,0,.3)"
  onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
  onmouseout="this.style.borderColor='var(--s2)';this.style.color='var(--mut)'">↑</button>
<script>
window.addEventListener('scroll', () => {
  const btn = document.getElementById('back-top');
  if (btn) btn.style.display = window.scrollY > 300 ? 'flex' : 'none';
  if (btn) btn.style.alignItems = 'center'; if (btn) btn.style.justifyContent = 'center';
}, {passive: true});
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


# ── Error handlers ───────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>404 · Yve.01</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{background:#0f172a;color:#f1f5f9;font-family:-apple-system,'Inter',sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}}
.e{{max-width:400px;padding:40px}}.dot{{width:10px;height:10px;border-radius:50%;background:#3b82f6;box-shadow:0 0 12px #3b82f6;display:inline-block;margin-right:8px}}
.code{{font-size:96px;font-weight:900;background:linear-gradient(135deg,#3b82f6,#a78bfa);-webkit-background-clip:text;-webkit-text-fill-color:transparent;line-height:1}}
h1{{font-size:20px;font-weight:700;margin:16px 0 8px}}p{{font-size:14px;color:#64748b;line-height:1.6;margin-bottom:24px}}
a{{background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;padding:11px 24px;border-radius:10px;text-decoration:none;font-weight:600;font-size:14px}}</style>
</head>
<body><div class="e">
  <div><div class="dot"></div><span style="font-weight:800;font-size:18px">Yve<span style="color:#60a5fa">.01</span></span></div>
  <div class="code">404</div>
  <h1>Página no encontrada</h1>
  <p>La página que buscas no existe o ha sido movida.</p>
  <a href="/">← Volver al dashboard</a>
</div></body></html>""", 404

@app.errorhandler(500)
def server_error(e):
    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>500 · Yve.01</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{background:#0f172a;color:#f1f5f9;font-family:-apple-system,'Inter',sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}}
.e{{max-width:400px;padding:40px}}.dot{{width:10px;height:10px;border-radius:50%;background:#ef4444;box-shadow:0 0 12px #ef4444;display:inline-block;margin-right:8px}}
.code{{font-size:96px;font-weight:900;background:linear-gradient(135deg,#ef4444,#f97316);-webkit-background-clip:text;-webkit-text-fill-color:transparent;line-height:1}}
h1{{font-size:20px;font-weight:700;margin:16px 0 8px}}p{{font-size:14px;color:#64748b;line-height:1.6;margin-bottom:24px}}
a{{background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;padding:11px 24px;border-radius:10px;text-decoration:none;font-weight:600;font-size:14px}}</style>
</head>
<body><div class="e">
  <div><div class="dot"></div><span style="font-weight:800;font-size:18px">Yve<span style="color:#60a5fa">.01</span></span></div>
  <div class="code">500</div>
  <h1>Error del servidor</h1>
  <p>Algo ha salido mal. El equipo ha sido notificado.<br>Inténtalo de nuevo en unos minutos.</p>
  <a href="/">← Volver al inicio</a>
</div></body></html>""", 500


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
