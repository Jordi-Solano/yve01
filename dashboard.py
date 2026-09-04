"""
dashboard.py — Yve.01
Dashboard principal AR. Ejecutar: python dashboard.py
Abre en: http://localhost:5001
"""

import os, glob, json, subprocess, sys, threading
from datetime import date
import pandas as pd
from flask import Flask, Response, jsonify, request, stream_with_context, redirect, send_file, session
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
from tenant_dirs import (datos_dir as _ddir, reportes_dir as _rdir,
                         entrada_dir as _edir, procesadas_dir as _pdir,
                         aprobaciones_dir as _adir, tenant_id as _tenant_id)

def _env_tenant():
    """Entorno para subprocess con el tenant Y el hotel de la sesión.

    Los scripts que se lanzan como subproceso no tienen sesión de Flask: leen
    YVE_TENANT para saber en qué árbol escribir y YVE_HOTEL para saber con qué
    hotel etiquetar lo que guarden.
    """
    e = os.environ.copy()
    e["YVE_TENANT"] = _tenant_id()
    try:
        e["YVE_HOTEL"] = censo_hoteles.para_guardar() or ""
    except Exception:
        e["YVE_HOTEL"] = ""
    return e


def _falta_hotel():
    """El candado del servidor: la respuesta 409 si falta elegir hotel, o None.

    Se pone en los endpoints que ESTAMPAN el hotel, no en los que solo dejan un
    fichero en `facturas-entrada/`: subir un PDF no crea nada sin hotel, lo crea
    procesarlo.

    Por que en el servidor y no solo en el modal: el modal es la experiencia, no
    la garantia. Se puede llamar a la API a pelo, se puede tener una pestaña
    vieja abierta, y la foto del movil entra por otra puerta. La invariante
    —"no existe documento nuevo sin hotel"— solo es verdad si la sostiene el
    servidor.

    Quien decide es `censo_hoteles.exige_hotel()`, en un sitio y nada mas. En AR
    ya se aprendio lo que pasa cuando la misma regla vive en cinco.
    """
    try:
        motivo = censo_hoteles.exige_hotel()
    except Exception:
        return None          # ante la duda, no bloquear a nadie
    if not motivo:
        return None
    return jsonify({'ok': False, 'error': motivo, 'hotel_requerido': True,
                    'hoteles': censo_hoteles.para_selector()}), 409


REPORTES_DIR_LEGACY     = os.path.join(BASE_DIR, "reportes")
PROCESADAS_DIR_LEGACY   = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_DIR_LEGACY = os.path.join(BASE_DIR, "aprobaciones")
NF = "NO_ENCONTRADO"

# ── Excel cache (TTL 5 min) ──────────────────────────────────────────────
import time as _time
_EXCEL_CACHE: dict = {}
_CACHE_TTL = 300  # seconds

def _huella_fichero(path):
    """(mtime, tamaño) del fichero, o None si no existe.

    Los dos y no solo el mtime: su resolucion puede ser de 1 s y una reescritura
    dentro del mismo segundo pasaria desapercibida.
    """
    try:
        st = os.stat(path)
        return (st.st_mtime, st.st_size)
    except OSError:
        return None


def _excel(path, sheet_name=0, header=0, **kw):
    """Lee un Excel con cache que se invalida SOLA cuando el fichero cambia.

    Antes caducaba solo por TTL (5 min) y _invalidate_cache() no la llamaba
    nadie en todo el proyecto, asi que tras regenerar un fichero se seguian
    sirviendo los numeros viejos. Mismo criterio que en tab_fb_dashboard._xlsx:
    la correccion vive AQUI, en un sitio, en vez de repartir invalidaciones por
    cada escritor — que es lo que siempre se acaba olvidando.
    El TTL se queda como red de seguridad (relojes raros, ficheros en red).
    """
    key = f"{path}|{sheet_name}|{header}"
    now = _time.time()
    huella = _huella_fichero(path)
    if key in _EXCEL_CACHE:
        df, ts, huella_cache = _EXCEL_CACHE[key]
        if huella == huella_cache and now - ts < _CACHE_TTL:
            return df
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, **kw)
    _EXCEL_CACHE[key] = (df, now, huella)
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
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(32).hex()

# ── CSRF ────────────────────────────────────────────────────────────────────
import hmac as _hmac, secrets as _sec_mod

def _csrf_token():
    from flask import session
    if 'csrf_token' not in session:
        session['csrf_token'] = _sec_mod.token_hex(32)
    return session['csrf_token']

@app.before_request
def _csrf_check():
    from flask import request as _req, session, jsonify as _jfy
    if _req.method not in ('POST','PUT','PATCH','DELETE'): return
    if not _req.path.startswith('/api/'): return
    if _req.path == '/api/login': return  # entrar con otra cuenta teniendo sesión abierta
    if 'user_id' not in session and '_user_id' not in session: return
    if _req.content_type and 'multipart' in _req.content_type: return
    tok = (_req.headers.get('X-CSRF-Token') or
           (_req.get_json(silent=True) or {}).get('csrf_token') or '')
    sess_tok = session.get('csrf_token', '')
    if not tok or not _hmac.compare_digest(tok, sess_tok):
        return _jfy({'error': 'CSRF inválido', 'csrf_error': True}), 403

@app.before_request
def _require_auth_api():
    """Exige sesión iniciada para los endpoints /api de datos (evita lectura anónima)."""
    from flask import request as _rq
    p = _rq.path
    if not p.startswith('/api/'):
        return
    _OPEN = ('/api/login', '/api/csrf_token', '/api/health', '/api/oracle/status',
             '/api/push/public_key', '/api/set_lang', '/api/demo')
    if any(p.startswith(o) for o in _OPEN):
        return
    from flask_login import current_user as _cu
    if not _cu.is_authenticated:
        from flask import jsonify as _j2
        return _j2({'error': 'No autorizado', 'auth_required': True}), 401

@app.route('/api/csrf_token')
def api_csrf_token():
    from flask_login import current_user
    if not current_user.is_authenticated:
        return __import__('flask').jsonify({'error': 'No auth'}), 401
    return __import__('flask').jsonify({'token': _csrf_token()})
# ─────────────────────────────────────────────────────────────────────────────

# Auth + módulos: la app es UN solo proceso que sirve todo el producto en un puerto
sys.path.insert(0, BASE_DIR)
from auth import init_login, inicializar_usuarios
init_login(app)
inicializar_usuarios()

# Registrar cada módulo como blueprint (login, configuración, admin, aprobaciones, conciliación)
from login import bp as auth_bp
from onboarding import onboarding_bp as config_bp
from panel_admin import bp as admin_bp
import censo_hoteles
from version_estaticos import SELLO as SELLO_ESTATICOS
from app_aprobacion import bp as aprob_ar_bp
from app_aprobacion_ap import bp as aprob_ap_bp
from app_conciliacion import bp as concil_bp
from tab_fb_dashboard import fb_bp
from tab_ar_real import ar_real_bp
from reclamaciones_ota import recl_ota_bp
from reclamaciones_ap import recl_ap_bp
from oracle_export_dryrun import oracle_export_bp   # exporta SOLO lo producido por el pipeline
from tab_cierre import cierre_bp
from tab_albaranes import albaranes_bp
from oracle_export_dryrun import oracle_export_bp
from pricing import pricing_bp
from tab_multi_hotel import multi_hotel_bp
from tab_self_service import self_service_bp
from tab_exportador import exportador_bp
# calipolis quitado
from tab_demo import demo_bp
from tab_demo_simulador import demo_sim_bp
from tab_reportes_pdf import reportes_pdf_bp
from rol_dashboard import get_dashboard_config
from demo_completo import generar_hoteles_demo, generar_facturas_demo_ar, generar_facturas_demo_ap, generar_alertas_demo
from landing import LANDING_HTML as LANDING_PAGE
from blog import blog_bp
from billing import billing_bp
from exportador_asientos import asientos_bp
from legal import legal_bp
from signup import signup_bp
from about import about_bp
from exportador_pdf import pdf_bp
# pricing_bp estaba importado pero NO registrado: /precios daba 404 mientras la
# landing, el blog y "Quienes somos" enlazaban a el (Ola A).
for _bp in (auth_bp, config_bp, admin_bp, aprob_ar_bp, aprob_ap_bp, concil_bp, fb_bp, ar_real_bp, recl_ota_bp, recl_ap_bp, oracle_export_bp, cierre_bp, albaranes_bp, multi_hotel_bp, self_service_bp, exportador_bp, demo_bp, demo_sim_bp, reportes_pdf_bp, blog_bp, billing_bp, asientos_bp, signup_bp, about_bp, pdf_bp, legal_bp, pricing_bp):
    app.register_blueprint(_bp)


# Apaño temporal mientras no haya persistencia: si el censo esta vacio (cada
# despliegue lo deja asi, porque Render no tiene disco) y hay una variable de
# entorno con los hoteles, se recrean. No siembra nada si ya hay hoteles.
# El porque completo esta en `censo_hoteles.sembrar_desde_entorno`.
try:
    import censo_hoteles as _censo_arranque
    _censo_arranque.sembrar_desde_entorno()
except Exception as _e:
    print(f"[censo] siembra desde entorno omitida: {_e}")

# Igual que el censo: la config del banco (grupo vs por hotel) se siembra desde
# YVE_BANCO_MODO para sobrevivir a los despliegues (disco efimero de Render).
try:
    import config_banco as _cfgbanco_arranque
    _cfgbanco_arranque.sembrar_desde_entorno()
except Exception as _e:
    print(f"[config_banco] siembra desde entorno omitida: {_e}")


def _estampar_hotel_banco(df):
    """Marca las filas NUEVAS del extracto con el hotel activo — SOLO en modo
    'por_hotel'. En 'grupo' (o sin elegir aún) no toca nada: el banco va junto,
    como siempre. Es el mismo patron que AP/AR (`df['hotel_id'] = para_guardar()`),
    pero gobernado por la eleccion del usuario (config_banco). El clasificador no
    se toca: esto ocurre DESPUES de clasificar, al guardar. Si no hay hotel para
    asignar, se deja sin marcar y queda 'sin asignar' (visible, no escondido)."""
    try:
        import config_banco as _cfgb
        if not _cfgb.por_hotel():
            return df
        import censo_hoteles as _c
        hid = _c.para_guardar()
        if hid:
            df['hotel_id'] = hid
    except Exception as _e:
        print(f"[config_banco] no se pudo estampar el hotel en el extracto: {_e}")
    return df

_pipeline_running = False
_pipeline_lock    = threading.Lock()

# ── Las escrituras del Excel, en fila ────────────────────────────────
# Render corre con `--workers 1 --threads 8`: un proceso y ocho hilos. Los
# guardadores hacen leer -> concatenar -> escribir, y eso a la vez es una
# carrera de libro. MEDIDO: con tres guardados simultaneos se pierden 3 de 6
# facturas, y ademas salta `BadZipFile` porque un hilo lee el xlsx mientras
# otro lo esta escribiendo.
#
# El candado NO toca ni un dato: solo pone las llamadas en fila. Uno solo para
# los tres guardadores; lo que protege es el patron, no un fichero concreto, y
# escribir un xlsx son milisegundos frente a los ~7 s que tarda la IA.
_guardado_lock = threading.Lock()


def _en_fila(_fn):
    """Un guardador cada vez. Envuelve sin reindentar el cuerpo."""
    def _envuelta(*a, **kw):
        with _guardado_lock:
            return _fn(*a, **kw)
    _envuelta.__name__ = getattr(_fn, '__name__', '_envuelta')
    _envuelta.__doc__ = getattr(_fn, '__doc__', None)
    _envuelta.__wrapped__ = _fn
    return _envuelta

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

# Un numero escrito a la española o a la inglesa, y nada mas: se usa para
# decidir si una columna de texto es en realidad numerica.
_re_num_es = __import__('re').compile(r'^-?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?$|^-?\d+(?:[.,]\d+)?$')


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

def _plegar(s):
    """Texto comparable: sin acentos, sin mayusculas, sin dobles espacios."""
    import unicodedata
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


def _mismo_hotel(nombre_doc, nombre_asignado):
    """Si el hotel que nombra el documento es el que le hemos asignado.

    TOLERANTE a proposito: `nombre_hotel` sale de una expresion regular sobre
    el PDF de la OTA, asi que llega con mayusculas raras, acentos comidos y
    coletillas ("HOTEL SOL MAR S.L.", "Hotel Sol Mar - Sitges"). Se pliegan los
    acentos y basta con que uno contenga al otro. Un aviso que salta cada dos
    por tres se ignora, y entonces ya no avisa de nada.

    Ante la duda (falta cualquiera de los dos nombres) dice que SI cuadra: no
    se molesta al usuario con lo que no se sabe.
    """
    a, b = _plegar(nombre_doc), _plegar(nombre_asignado)
    if not a or not b or a in ("no_encontrado", "nan", "none"):
        return True
    return a in b or b in a


def _aviso_otro_hotel(nombre_doc):
    """El aviso, PELADO, si el papel nombra un hotel distinto del elegido.

    Devuelve '' cuando todo cuadra o cuando no hay nada que comparar.

    Sin adornos a proposito: cada sitio lo presenta como le toca. En el lote
    tiene que ser su PROPIA linea empezando por ⚠ — pegado al final del "✓ AR
    fichero: OK" se pintaba de verde, se iba al final de una linea larga y en
    un log que va scrolleando no lo veia nadie. Se dio por hecho que con que
    el texto estuviera bastaba; no basta.

    Se le pregunta al CENSO con que hotel encaja el nombre del documento en vez
    de compararlo solo con el asignado: "Hotel Sol Mar" contiene "Hotel Sol",
    asi que entre hermanos la comparacion simple se quedaba callada — el error
    exacto que esto existe para cazar. `_mismo_hotel` se queda como red de
    seguridad para los nombres que el censo no reconoce.
    """
    nombre_doc = str(nombre_doc or "").strip()
    hid = censo_hoteles.activo()
    if not hid or not nombre_doc:
        return ''
    asignado = censo_hoteles.nombre_de(hid)
    otro = censo_hoteles.encaje(nombre_doc)
    if (otro and otro != hid) or (not otro and not _mismo_hotel(nombre_doc, asignado)):
        return (f'ojo: el documento nombra otro hotel ({nombre_doc}) '
                f'y se ha guardado en {asignado}')
    return ''


def _filtrar_hotel_activo(df, cols=("hotel", "nombre_hotel")):
    """Si hay un hotel activo en sesión y el df tiene columna de hotel, filtra por él.
    Si el df no tiene columna de hotel, se devuelve tal cual (datos de grupo).

    FASE 0: la sesión guarda el ID; el nombre se resuelve del censo. El cruce
    sigue siendo por NOMBRE contra las mismas columnas — a propósito, para que
    esta fase no cambie ni un número. En la fase 1 esto se muda a
    `almacen_datos` y pasa a cruzarse por `hotel_id`, que es lo que de verdad
    arregla el problema de los nombres parecidos.
    """
    try:
        hid = censo_hoteles.activo()
    except Exception:
        return df
    if not hid or df is None or getattr(df, "empty", True):
        return df
    nombre = censo_hoteles.nombre_de(hid)
    if not nombre:
        # El hotel activo ya no está en el censo (lo han dado de baja). Vista de
        # grupo, que enseña de más y no de menos.
        return df
    for c in cols:
        if c in df.columns:
            mask = df[c].astype(str).str.contains(nombre, case=False, na=False, regex=False)
            return df[mask].copy()
    return df


@app.route("/api/hotel_activo", methods=["GET", "POST"])
@login_required
def api_hotel_activo():
    """Hotel activo de la sesión: filtra AR/AP/AR Real. Vacío = vista de grupo.

    Habla en IDs, no en nombres (fase 0). `normalizar` acepta también un nombre
    para no romper las páginas que ya estuvieran abiertas al desplegar.
    """
    if request.method == "POST":
        data = request.get_json(force=True) or {}
        hid = censo_hoteles.normalizar(data.get("hotel"))
        if hid:
            session["hotel_activo"] = hid
        else:
            session.pop("hotel_activo", None)
        actual = session.get("hotel_activo", "")
        return jsonify({"ok": True, "hotel": actual,
                        "nombre": censo_hoteles.nombre_de(actual)})
    hid = censo_hoteles.activo()
    return jsonify({"ok": True, "hotel": hid,
                    "nombre": censo_hoteles.nombre_de(hid),
                    "hoteles": censo_hoteles.para_selector()})


def cargar_datos_ar_sin_filtrar():
    """Lo mismo que `cargar_datos()` pero SIN acotar al hotel elegido.

    Es el cuerpo de toda la vida; lo unico que se ha sacado fuera es la ultima
    linea, el filtro. Existe para el agregador del grupo, que necesita las
    filas de TODOS los hoteles para partirlas el mismo (fase A).

    El orden importa y por eso el corte esta justo aqui: el enriquecimiento con
    las aprobaciones va ANTES del filtro. Si el agregador se leyera el almacen
    por su cuenta se dejaria fuera las columnas `accion` y `comentario`, y sus
    contadores de aprobadas/rechazadas no cuadrarian con los del panel — que es
    justo lo que la fase A viene a demostrar que si cuadra.
    """
    from almacen_datos import facturas_ar as _facturas_ar, resumen_fuentes as _fuentes
    df = _facturas_ar(_pdir(), _rdir())
    if df is None or df.empty:
        print("[cargar_datos] ADVERTENCIA: no se encontró ningún dato AR")
        return pd.DataFrame(), {}
    try:
        ruta = ", ".join(_fuentes(_pdir(), _rdir())["ar"]) or "consolidado"
    except Exception:
        ruta = "consolidado"

    apro_path = os.path.join(_adir(), "aprobaciones.xlsx")
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

    return df, {"ruta": ruta}


def cargar_datos():
    """
    Carga los datos AR de TODOS los dias, ya consolidados.

    La lectura y el deduplicado viven en almacen_datos (punto unico a cambiar
    el dia de la migracion a persistencia). Aqui solo queda el enriquecimiento.
    """
    # FASE 3: AR filtra por `hotel_id`, igual que AP, y no por el nombre que
    # venia del PDF. `_filtrar_hotel_activo` se queda solo para AR Real, que
    # todavia no lleva etiqueta.
    df, meta = cargar_datos_ar_sin_filtrar()
    return _solo_hotel_activo(df), meta

def calcular_stats(df):
    if df.empty:
        return {"total":0,"importe_total":0,"correctas":0,"discrepancias":0,
                "cobro_debajo":0,"sin_tarifa":0,
                "importe_reclamable":0,"di_pendientes":0,"aprobadas":0,"rechazadas":0,"sin_accion":0}
    total = len(df)
    importe_total = sum(safe_float(v) for v in df.get("importe_bruto", pd.Series()))

    estado_col = df["estado"].fillna("") if "estado" in df.columns else pd.Series([""] * total)
    correctas     = int((estado_col == "CORRECTO").sum())
    discrepancias = int((estado_col == "DISCREPANCIA").sum())
    # Los dos estados que NO son ni correcto ni reclamable, y que antes no se
    # contaban en ningun sitio: una factura asi desaparecia del resumen.
    cobro_debajo  = int((estado_col == "COBRO_POR_DEBAJO").sum())
    sin_tarifa    = int((estado_col.isin(["SIN_TARIFA_HOTEL", "OTA_DESCONOCIDA"])).sum())

    if "discrepancia_euros" in df.columns:
        # SIN abs(). El valor absoluto es justo como se perdia el signo: una
        # comision cobrada POR DEBAJO de lo pactado da un importe negativo, y
        # sumada en valor absoluto salia en el panel como dinero a devolver.
        # Ahora esas filas ni llegan aqui —tienen su propio estado— y el total
        # solo suma lo que de verdad se ha cobrado de mas.
        importe_reclamable = sum(
            v for v in (safe_float(x) for x in
                        df.loc[estado_col == "DISCREPANCIA", "discrepancia_euros"])
            if v > 0)
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
        "cobro_debajo": cobro_debajo,
        "sin_tarifa": sin_tarifa,
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
        "reportes_dir":   _rdir(),
        "procesadas_dir": _pdir(),
        "archivos_reportes": [],
        "archivos_procesadas": [],
        "cwd": os.getcwd(),
        "python_file": __file__,
        "python_version": __import__("sys").version.split()[0],
    }
    for d, key in [(_rdir(), "archivos_reportes"), (_pdir(), "archivos_procesadas")]:
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
    df, ruta_cargada = cargar_ultimo_excel("doble_imposicion_*.xlsx", _rdir())
    if df is None:
        df, ruta_cargada = cargar_ultimo_excel("verificacion_*.xlsx", _rdir())
    if df is None:
        df, ruta_cargada = cargar_ultimo_excel("facturas_procesadas_*.xlsx", _pdir())
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
        # 100 req/min: el frontend hace 1 petición por archivo, así que esto
        # permite procesar carpetas de hasta ~100 archivos sin cortar el stream
        if _rate_limit(f"process:{ip}", max_req=100, window=60):
            from flask import jsonify as _j
            return _j({"error": "Demasiados archivos a la vez. Espera 60 segundos."}), 429

@app.route("/robots.txt")
def robots_txt():
    return Response("""User-agent: *
Allow: /
Allow: /blog
Allow: /about
Allow: /casos
Allow: /precios
Allow: /unirse
Allow: /terminos
Allow: /privacidad
Disallow: /api/
Disallow: /admin/
Disallow: /onboarding
Sitemap: https://yve01.onrender.com/sitemap.xml
""", mimetype="text/plain")

@app.route("/og-image.png")
def og_image():
    """Social share card (SVG served as image)."""
    svg = """<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="630" viewBox="0 0 1200 630">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1200" y2="630" gradientUnits="userSpaceOnUse">
      <stop offset="0" stop-color="#0f172a"/>
      <stop offset="1" stop-color="#1e293b"/>
    </linearGradient>
    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#3b82f6"/>
      <stop offset="1" stop-color="#7c3aed"/>
    </linearGradient>
  </defs>
  <rect width="1200" height="630" fill="url(#bg)"/>
  <circle cx="80" cy="90" r="14" fill="url(#accent)"/>
  <text x="110" y="100" font-family="Arial,sans-serif" font-size="40" font-weight="900" fill="#f1f5f9">Yve<tspan fill="#60a5fa">.01</tspan></text>
  <text x="80" y="290" font-family="Arial,sans-serif" font-size="68" font-weight="900" fill="#f1f5f9">El sistema operativo</text>
  <text x="80" y="370" font-family="Arial,sans-serif" font-size="68" font-weight="900" fill="url(#accent)">AI para hoteles</text>
  <text x="80" y="450" font-family="Arial,sans-serif" font-size="32" fill="#94a3b8">Automatiza AP · AR · DRR · Conciliación bancaria</text>
  <text x="80" y="500" font-family="Arial,sans-serif" font-size="32" fill="#94a3b8">Setup en 15 minutos · Sin consultores</text>
  <rect x="80" y="540" width="280" height="56" rx="28" fill="url(#accent)"/>
  <text x="220" y="577" font-family="Arial,sans-serif" font-size="26" font-weight="700" fill="#fff" text-anchor="middle">yve01.onrender.com</text>
</svg>"""
    return Response(svg, mimetype="image/svg+xml")

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
        ("https://yve01.onrender.com/unirse",     "0.9",  "monthly"),
        ("https://yve01.onrender.com/precios",    "0.9",  "monthly"),
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
    """Si Oracle esta en simulacion o de verdad.

    BOMBA 2 — antes decidia por `ORACLE_BASE_URL` a secas, pero `oracle_auth`
    (lo que ejecuta el pipeline) considera simulacion mientras falten
    `ORACLE_CLIENT_ID` / `ORACLE_CLIENT_SECRET` o la URL sea la de plantilla.
    Dos criterios = una pantalla que dice "real" mientras el pipeline simula.
    Aqui manda `oracle_auth.is_simulation()`, y se consulta EN CADA peticion.
    """
    try:
        from oracle_auth import is_simulation as _is_sim
        sim = bool(_is_sim())
    except Exception:
        sim = True                       # ante la duda, nunca decir "real"
    return jsonify({'mode': 'simulation' if sim else 'real', 'simulacion': sim, 'ok': True})

# ── Provisiones de cierre (Ola A) ────────────────────────────────────────────
def _provisiones_args():
    mes = (request.args.get("mes") or "").strip()[:7] or None
    hotel = censo_hoteles.activo() or None
    return mes, hotel

@app.route("/api/provisiones")
@login_required
def api_provisiones():
    """Las dos provisiones del cierre: albaranes sin factura y comisiones OTA.

    Solo lee (provisiones.py). Respeta el hotel activo como los paneles.
    """
    import provisiones as _pv
    mes, hotel = _provisiones_args()
    try:
        alb = _pv.provision_albaranes(mes, hotel, _pdir(), _rdir(), _ddir())
        com = _pv.provision_comisiones(mes, hotel, _rdir(), _ddir())
        return jsonify({"ok": True, "mes": alb["mes"], "albaranes": alb, "comisiones": com})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500

@app.route("/api/exportar/provisiones")
@login_required
def api_exportar_provisiones():
    import provisiones as _pv
    mes, hotel = _provisiones_args()
    buf, nombre = _pv.exportar_excel(mes, hotel, procesadas_dir=_pdir(), reportes_dir=_rdir(),
                                     datos_dir=_ddir())
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Aging AP: a quien debemos y desde cuando (Ola A) ──────────────────────────
def _aging_ap_calcular():
    import aging_ap as _ag
    from almacen_datos import movimientos_banco as _mb
    df_ap = cargar_datos_ap()
    try:
        df_ar, _ = cargar_datos()
    except Exception:
        df_ar = pd.DataFrame()
    try:
        df_b, _ = _mb(reportes_dir=_rdir())
    except Exception:
        df_b = None
    return _ag.calcular_aging(df_ap, df_ar, df_b)

@app.route("/api/aging_ap")
@login_required
def api_aging_ap():
    try:
        return jsonify({"ok": True, **_aging_ap_calcular()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500

@app.route("/api/exportar/aging_ap")
@login_required
def api_exportar_aging_ap():
    import aging_ap as _ag
    buf, nombre = _ag.exportar_excel(_aging_ap_calcular())
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── PWA: service worker servido desde la raíz para controlar todo el origen ──
@app.route("/sw.js")
def pwa_service_worker():
    """Sirve el SW desde / (scope '/') en lugar de /static/ (scope '/static/')."""
    resp = app.send_static_file("sw.js")
    resp.headers["Service-Worker-Allowed"] = "/"
    resp.headers["Content-Type"] = "application/javascript; charset=utf-8"
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/offline")
def pwa_offline():
    """Página offline propia (fallback del SW en navegaciones sin red)."""
    return app.send_static_file("offline.html")

# ── Web Push (VAPID) ────────────────────────────────────────────────────────
@app.route("/api/push/public_key")
def api_push_public_key():
    """Clave pública VAPID para que el navegador se suscriba (no es secreta)."""
    try:
        import push_service
        return jsonify({"publicKey": push_service.VAPID_PUBLIC_KEY,
                        "enabled": push_service.push_enabled()})
    except Exception as e:
        return jsonify({"publicKey": "", "enabled": False, "error": str(e)[:120]}), 500

@app.route("/api/push/subscribe", methods=["POST"])
@login_required
def api_push_subscribe():
    """Registra la suscripción push del navegador (con rol + tenant del usuario)."""
    try:
        import push_service
        from flask import session as _sess
        data = request.get_json(silent=True) or {}
        sub = data.get("subscription") or data
        ok = push_service.add_subscription(
            sub,
            rol=getattr(current_user, "rol", None),
            tenant=_sess.get("tenant_id", "default"),
            username=getattr(current_user, "username", None),
        )
        return jsonify({"ok": ok, "total": push_service.count_subscriptions()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:160]}), 500

@app.route("/api/push/unsubscribe", methods=["POST"])
def api_push_unsubscribe():
    """Elimina una suscripción push."""
    try:
        import push_service
        data = request.get_json(silent=True) or {}
        endpoint = data.get("endpoint") or (data.get("subscription") or {}).get("endpoint")
        ok = push_service.remove_subscription(endpoint)
        return jsonify({"ok": ok, "total": push_service.count_subscriptions()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:160]}), 500

@app.route("/api/push/test", methods=["POST"])
@login_required
def api_push_test():
    """Envía una notificación push de prueba a todos los dispositivos suscritos."""
    try:
        import push_service
        if not push_service.push_enabled():
            return jsonify({"ok": False, "error": "Push no configurado en el servidor "
                            "(falta la variable VAPID_PRIVATE_KEY en Render)."}), 200
        from flask import session as _sess
        res = push_service.send_push(
            title="🔔 Yve.01 — Prueba de push",
            body="Las notificaciones push funcionan correctamente.",
            url="/app?tab=notif", tag="yve-test",
            tenant=_sess.get("tenant_id", "default"),
        )
        res["ok"] = True if res.get("sent", 0) > 0 else res.get("ok", False)
        if res.get("sent", 0) == 0 and res.get("total", 0) == 0:
            res["message"] = "No hay dispositivos suscritos todavía. Activa el canal Push en este dispositivo."
        else:
            res["message"] = f"Enviado a {res.get('sent',0)} de {res.get('total',0)} dispositivo(s)."
        return jsonify(res)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500



# ── Normalización de columnas para datos extraídos por IA ────────────
def _clave_col(nombre):
    """El nombre de una columna, comparable: sin acentos, en minusculas y con
    guion bajo donde hubiera espacios, guiones o puntos.

    Asi "Unidades Vendidas", "unidades-vendidas" y "UNIDADES VENDIDAS" son la
    misma columna, que es lo que las listas de alias querian decir siempre.
    """
    s = _plegar(nombre)
    for ch in ('-', '.', '/', '(', ')', ':'):
        s = s.replace(ch, ' ')
    return '_'.join(s.split())


def _normalize_cols(df, expected_map):
    """Renombra columnas de un DataFrame para que coincidan con el esquema esperado.
    expected_map: {'nombre_esperado': ['alternativa1', 'alternativa2', ...]}

    Dos pasadas: primero el nombre EXACTO —lo de siempre, sin cambiar nada— y
    solo si ninguna alternativa casa, el nombre NORMALIZADO (`_clave_col`). La
    segunda pasada nunca reasigna una columna que la primera ya haya colocado,
    asi que solo puede recuperar columnas que antes se quedaban sin mapear.

    La comparacion es SIEMPRE de nombre COMPLETO, jamas por subcadena, y de eso
    hay cicatriz: `/fb/api/upload_ventas` tenia su propia cadena de `elif` con
    `'id' in cl`, e "id" esta dentro de `cantidad`, de `unidad` y hasta de
    `unidades_vendidas` —su propio nombre canonico—. Las cantidades acababan en
    `id_receta`, `unidades_vendidas` se rellenaba con 1 por defecto y el food
    cost salia 0,06% en vez de 16,25%.

    Y una columna de origen no se usa dos veces: dos columnas cayendo en el
    mismo nombre canonico dejaban el DataFrame con nombres duplicados, que es
    un fallo mucho mas raro de encontrar que un renombrado que no ocurre.
    """
    por_clave = {}
    for c in df.columns:
        por_clave.setdefault(_clave_col(c), c)
    canonicos = {_clave_col(k) for k in expected_map}

    rename = {}
    for expected, alternatives in expected_map.items():
        if expected in df.columns:
            continue
        for alt in alternatives:
            if alt in df.columns and alt not in rename:
                rename[alt] = expected
                break
        else:
            # segunda pasada: los mismos alias, comparando nombres normalizados
            for alt in list(alternatives) + [expected]:
                clave = _clave_col(alt)
                c = por_clave.get(clave)
                if c is None or c in rename:
                    continue
                # y no se le roba a otro campo su propio nombre canonico
                if clave != _clave_col(expected) and clave in canonicos:
                    continue
                rename[c] = expected
                break
    if rename:
        df = df.rename(columns=rename)
    return df

# Los tres ficheros de F&B que son de CADA hotel, con la clave que identifica
# una fila DENTRO de su hotel. El inventario es el que dolia: el "Tomate" del
# hotel B borraba el del hotel A, y cada hotel tiene su stock y su precio de
# compra. Mermas y ventas no tienen clave propia —una merma es un hecho con
# fecha, y dos mermas identicas el mismo dia son dos mermas— asi que se
# deduplican por la fila entera, que ya incluye el hotel.
_CLAVE_FB = {
    'inventario.xlsx': ('ingrediente',),
    'mermas.xlsx': (),
    'ventas_fb_diarias.xlsx': (),
}


def _guardar_fb_del_hotel(df, fichero):
    """Añade `df` a un fichero de F&B del hotel activo.

    Devuelve (filas DE ESTE HOTEL, filas entrantes) — no el fichero entero. Es
    a proposito: quien llama lo usa para el mensaje de la pantalla, y decir "20
    en el hotel" cuando 20 son las filas de los dos hoteles juntos es la clase
    de numero que hace perder una tarde. Al fichero completo se va por el
    fichero.

    UN solo sitio donde se estampa el hotel y se deduplica, porque estos tres
    ficheros los escriben TRES puertas distintas —la capa 1 por nombre de
    fichero, el clasificador de IA y la foto— y solo la de la IA lo hacia bien.

    Lo que pasaba por las otras dos: la fila se guardaba sin `hotel_id`, y los
    paneles de F&B leen con `_xlsx_hotel`, que filtra por el hotel de la sesion
    y FALLA CERRADO. Resultado medido en la prueba de integracion: el lote
    cantaba "14 items integrados" y el panel de inventario mostraba 0. Un cero
    en silencio es peor que un error.

    Y el inventario se deduplicaba por `ingrediente` a secas, sin hotel: subir
    el inventario de dos hoteles con ingredientes en comun dejaba 16 filas
    donde tenian que haber 20, porque el stock de uno borraba el del otro.

    Los ficheros de antes de la separacion no traen la columna: sus filas se
    marcan como "sin asignar" en vez de heredar el hotel de quien esta mirando,
    que seria inventarse a quien pertenece un dato viejo.
    """
    import pandas as _pd
    df = df.copy()
    hid = censo_hoteles.para_guardar()
    df['hotel_id'] = hid
    # M3: las mermas sin fecha salian sin fecha en el historial. Las ventas ya
    # traen la suya; las mermas muchas veces no, y una merma sin fecha no se
    # puede ordenar ni comparar con el mes anterior. Se pone la de HOY solo
    # donde falta: si el fichero trae fecha, manda la suya.
    if fichero == 'mermas.xlsx':
        _hoy = date.today().strftime('%Y-%m-%d')
        if 'fecha' not in df.columns:
            df['fecha'] = _hoy
        else:
            df['fecha'] = df['fecha'].astype(object)
            _vacia = df['fecha'].isna() | (df['fecha'].map(lambda v: str(v).strip()) == '')
            df.loc[_vacia, 'fecha'] = _hoy
    entrantes = len(df)
    ruta = os.path.join(_ddir(), fichero)
    if os.path.exists(ruta):
        viejo = _pd.read_excel(ruta)
        if not viejo.empty:
            viejo = viejo.copy()
            if 'hotel_id' not in viejo.columns:
                viejo['hotel_id'] = ''
            # Un hotel vacio vuelve de Excel como NaN y NaN != '': el
            # deduplicado no lo veia igual y cada subida DUPLICABA la fila
            # (visto con la hoja de recuento del cierre, 0 hoteles).
            viejo['hotel_id'] = viejo['hotel_id'].map(safe_str)
            df = _pd.concat([viejo, df], ignore_index=True)
    df['hotel_id'] = df['hotel_id'].map(safe_str)
    clave = [c for c in _CLAVE_FB.get(fichero, ()) if c in df.columns]
    if clave:
        df = df.drop_duplicates(subset=clave + ['hotel_id'], keep='last')
    else:
        df = df.drop_duplicates(keep='last')
    df.to_excel(ruta, index=False)
    if fichero in _EXCEL_CACHE:
        del _EXCEL_CACHE[fichero]
    try:
        from tab_fb_dashboard import _invalidate as _inv_fb
        _inv_fb()
    except Exception:
        pass
    del_hotel = df[df['hotel_id'].astype(str) == str(hid)]
    return del_hotel, entrantes


# ── Capa 1: el nombre PROPONE, las cabeceras CONFIRMAN ───────────────────
# Antes el nombre del fichero mandaba y el archivo no se abria nunca: un CSV
# llamado "extracto_movimientos_junio" con ventas dentro acababa en el libro de
# banco, metiendo 276 EUR de comida como abonos y añadiendo columnas 'plato' y
# 'unidades' al extracto. Medido: 4 de 9 nombres plausibles acababan en la
# seccion equivocada, dos de ellos por el ORDEN de las reglas y no por nombres
# engañosos ('movimientos' ganaba a 'stock', 'food' ganaba a 'waste').
#
# Ahora cada regla de nombre tiene que superar una comprobacion de CABECERAS
# (leer solo la fila de titulos: milisegundos, cero llamadas a la IA). Si no la
# supera, se sigue probando el resto de reglas — asi el orden deja de importar —
# y si ninguna encaja, el fichero cae al clasificador de IA, que es el camino
# que ya sabemos que acierta (probado en produccion: banco, F&B, inventario y
# mermas, 4 de 4 con nombres neutros).
#
# El fallo es SIEMPRE hacia el lado seguro: dudar cuesta una llamada a la IA;
# fiarse del nombre costaba contaminar una seccion.

_CAB_TIPOS = {
    'BANCO': {
        'requiere': [
            {'fecha', 'date', 'dia', 'f_valor', 'fecha_valor', 'fecha_operacion'},
            {'concepto', 'descripcion', 'description', 'detalle', 'referencia',
             'movimiento', 'operacion'},
            {'importe', 'cantidad', 'amount', 'monto', 'valor', 'cargo', 'abono',
             'debe', 'haber'},
        ],
        'excluye': {'plato', 'producto', 'ingrediente', 'articulo', 'receta',
                    'stock', 'merma', 'habitacion', 'huesped'},
    },
    'F&B': {
        'requiere': [
            {'plato', 'nombre_plato', 'producto', 'articulo', 'item', 'dish',
             'menu', 'receta'},
            {'total', 'importe', 'venta', 'ventas', 'revenue', 'precio',
             'unidades', 'cantidad', 'qty', 'units'},
        ],
        # numero_po / importe_aprobado / orden_compra: una orden de compra NO es
        # una hoja de ventas. Un fichero llamado "POs_julio.csv" propone F&B
        # porque "POs" contiene "pos", y sin esto pasaba el guard y metia el
        # pedido en ventas de restaurante (medido). Ojo: 'orden' y 'pedido' a
        # secas NO se pueden excluir — una comanda de restaurante las lleva.
        'excluye': {'saldo', 'concepto', 'stock_actual', 'stock_inicial',
                    'existencias', 'merma', 'habitacion',
                    'numero_po', 'importe_aprobado', 'orden_compra',
                    'purchase_order', 'num_po', 'n_po'},
    },
    'INVENTARIO': {
        'requiere': [
            {'ingrediente', 'producto', 'articulo', 'material', 'item', 'nombre'},
            {'stock', 'existencias', 'cantidad_actual', 'cantidad_inicial'},
        ],
        'excluye': {'saldo', 'concepto', 'plato', 'merma', 'habitacion'},
    },
    'MERMAS': {
        'requiere': [
            {'ingrediente', 'producto', 'articulo', 'item', 'nombre'},
            {'merma', 'waste', 'desperdicio', 'perdida', 'causa', 'motivo'},
        ],
        'excluye': {'saldo', 'concepto', 'plato', 'stock_inicial'},
    },
    'ROOMING': {
        'requiere': [
            {'grupo', 'group', 'cliente', 'evento'},
            {'habitacion', 'habitaciones', 'rooms', 'pax', 'huesped', 'guest',
             'entrada', 'checkin', 'check_in', 'salida', 'checkout', 'noches'},
        ],
        'excluye': {'saldo', 'concepto', 'plato', 'stock', 'merma'},
    },
}

# nombre -> tipo que PROPONE, en el orden en que se prueban
_CAB_KEYWORDS = [
    ('BANCO',      ['extracto', 'bank', 'statement', 'movimientos', 'bancario']),
    ('F&B',        ['pos', 'ventas', 'sales', 'tpv', 'food', 'beverage', 'f&b', 'fnb',
                    'restaurante', 'bar ', 'menu_mix', 'product_mix', 'ticket']),
    ('INVENTARIO', ['inventario', 'inventory', 'stock', 'almacen']),
    ('MERMAS',     ['merma', 'waste', 'pérdida', 'perdida']),
    ('ROOMING',    ['rooming', 'room list', 'guest list', 'room block']),
]

_DRR_KEYWORDS = ['drr', 'revenue report', 'daily report', 'daily_report']


def _norm_cab(v):
    """Cabecera normalizada: minusculas, sin acentos, espacios y puntos a '_'."""
    import unicodedata as _u
    s = '' if v is None else str(v)
    s = _u.normalize('NFKD', s)
    s = ''.join(c for c in s if not _u.combining(c))
    s = s.strip().lower()
    for ch in (' ', '.', '-', '/', '\\'):
        s = s.replace(ch, '_')
    while '__' in s:
        s = s.replace('__', '_')
    return s.strip('_')


def _leer_cabeceras(fpath):
    """Solo la fila de titulos. Devuelve [] si no se puede leer.

    Barato a proposito: nrows=0 no parsea los datos. Si falla, quien llama debe
    tratarlo como "sin evidencia" y mandar el fichero a la IA, nunca aceptarlo
    a ciegas.
    """
    try:
        ext = os.path.splitext(fpath)[1].lower()
        if ext == '.csv':
            try:
                df = pd.read_csv(fpath, nrows=0, sep=None, engine='python')
            except Exception:
                df = pd.read_csv(fpath, nrows=0)
        elif ext in ('.xlsx', '.xls'):
            df = pd.read_excel(fpath, nrows=0)
        else:
            return []
        return [_norm_cab(c) for c in df.columns]
    except Exception:
        return []


def _cabeceras_encajan(tipo, cabeceras):
    """True si las cabeceras confirman que el fichero ES de ese tipo."""
    regla = _CAB_TIPOS.get(tipo)
    if not regla or not cabeceras:
        return False
    unidas = '|'.join(cabeceras)
    # ninguna señal que contradiga
    for prohibido in regla['excluye']:
        if prohibido in unidas:
            return False
    # al menos una cabecera de cada grupo obligatorio
    for grupo in regla['requiere']:
        if not any(t in unidas for t in grupo):
            return False
    return True


def _destino_capa1(fname, fpath):
    """A donde manda la capa 1: 'DRR', 'BANCO', 'F&B', 'INVENTARIO', 'MERMAS',
    'ROOMING' o 'IA' (que el clasificador lo decida abriendo el fichero).

    El DRR es la excepcion y sigue mandando por NOMBRE: el prompt del
    clasificador no tenia tipo DRR y, sobre todo, un DRR de 33 hojas no cabe en
    la ventana del clasificador (medido: solo ve DAILY_MASTER y dia y medio, no
    ve CtaCble ni ningun 'Out of Balance'), y lo que si ve son filas de debe y
    haber con fechas — o sea, exactamente el aspecto de un extracto bancario.
    """
    fl = (fname or '').lower()
    ext = os.path.splitext(fl)[1]
    if ext == '.xlsm' and any(k in fl for k in _DRR_KEYWORDS):
        return 'DRR'
    if ext not in ('.xlsx', '.xls', '.csv'):
        # fotos y PDFs no pasan por aqui: van al lector universal / Vision.
        return 'IA'

    candidatos = [t for t, kws in _CAB_KEYWORDS if any(k in fl for k in kws)]
    if not candidatos:
        return 'IA'

    cabeceras = _leer_cabeceras(fpath)
    for tipo in candidatos:                 # el orden solo desempata entre iguales
        if _cabeceras_encajan(tipo, cabeceras):
            return tipo
    return 'IA'


_INV_COL_MAP = {
    'ingrediente': ['producto', 'nombre', 'item', 'articulo', 'material'],
    'categoria': ['tipo', 'category', 'grupo', 'familia'],
    'coste_unitario': ['precio', 'coste', 'precio_unitario', 'cost', 'precio_kg'],
    'stock_actual_kg_l': ['stock_actual', 'stock', 'cantidad', 'cantidad_actual'],
    'stock_inicial_kg_l': ['stock_inicial', 'cantidad_inicial'],
    'unidad': ['unit', 'medida', 'uom'],
    'proveedor': ['supplier', 'vendor', 'distribuidor'],
}

_MER_COL_MAP = {
    # M3: sin esta entrada, un fichero de mermas con 'Fecha' o 'dia' la perdia
    # por el camino y el historial salia sin fecha.
    'fecha': ['dia', 'día', 'fecha_merma', 'date', 'day'],
    'ingrediente': ['producto', 'nombre', 'item', 'articulo'],
    'categoria': ['tipo', 'category', 'grupo'],
    'cantidad_merma': ['cantidad', 'amount', 'qty', 'kilos'],
    'causa': ['motivo', 'reason', 'causa_merma'],
    'coste_merma': ['coste', 'coste_total', 'valor', 'importe'],
    'coste_unitario': ['precio', 'coste_kg'],
    'unidad': ['unit', 'medida'],
}

# OJO al ORDEN: los mapas se recorren de arriba abajo y la primera columna que
# case se la queda ese nombre canonico. `unidades_vendidas` va ANTES de
# `precio_unitario` a proposito, porque "unidades" y "unitario" comparten
# familia de nombres y las cantidades son lo que no se puede perder.
_VEN_COL_MAP = {
    'fecha': ['dia', 'date', 'day', 'fecha_venta'],
    'nombre_plato': ['plato', 'nombre', 'producto', 'item', 'dish', 'articulo',
                     'descripcion'],
    'categoria': ['tipo', 'category', 'grupo', 'familia'],
    'unidades_vendidas': ['cantidad', 'qty', 'units', 'unidades', 'uds',
                          'cant', 'n_unidades'],
    'precio_unitario': ['precio', 'pvp', 'price', 'precio_venta'],
    'total_venta': ['total', 'importe', 'revenue', 'ventas', 'total_ventas',
                    'importe_total'],
    'id_receta': ['receta', 'recipe', 'id', 'codigo', 'cod', 'sku',
                  'id_plato', 'ref'],
}

_BANK_COL_MAP = {
    'fecha': ['date', 'dia'],
    'concepto': ['descripcion', 'description', 'detalle', 'referencia'],
    'importe': ['cantidad', 'amount', 'monto', 'valor'],
    'saldo': ['balance', 'saldo_final'],
}

# Facturas de comision OTA -> esquema que lee verificador_comisiones.py
# (el mismo que produce lector_ota.py en facturas_procesadas_*.xlsx)
_OTA_COL_MAP = {
    'numero_factura': ['num_factura', 'invoice', 'invoice_number', 'numero', 'factura', 'n_factura', 'nº_factura'],
    'nombre_ota': ['ota', 'plataforma', 'canal', 'portal', 'agencia'],
    'nombre_hotel': ['hotel', 'establecimiento', 'propiedad', 'property'],
    'fecha': ['fecha_factura', 'date', 'invoice_date'],
    'periodo_inicio': ['inicio', 'desde', 'fecha_inicio', 'periodo_desde', 'from'],
    'periodo_fin': ['fin', 'hasta', 'fecha_fin', 'periodo_hasta', 'to'],
    'importe_bruto': ['bruto', 'importe', 'base', 'reservas', 'gross', 'importe_reservas'],
    'porcentaje_comision': ['porcentaje_factura', 'porcentaje', 'comision_pct', 'pct_comision',
                            'rate', 'commission_rate', 'porcentaje_aplicado', 'comision_porcentaje'],
    'importe_comision': ['comision', 'commission', 'comision_eur', 'importe_comision_factura'],
    'importe_neto': ['neto', 'a_pagar', 'net', 'importe_a_pagar', 'net_payout'],
}

# Tarifas pactadas -> esquema de datos-referencia/comisiones_pactadas.xlsx
# OJO: las columnas van en Mayuscula porque asi las lee verificador_comisiones.
_PACT_COL_MAP = {
    'OTA': ['ota', 'nombre_ota', 'plataforma', 'canal', 'portal', 'agencia'],
    'Hotel': ['nombre_hotel', 'hotel', 'establecimiento', 'propiedad', 'property'],
    'Porcentaje_Comision': ['porcentaje_pactado', 'porcentaje', 'comision_pactada',
                            'pct_pactado', 'porcentaje_comision', 'rate', 'comision'],
    'Mercado': ['mercado', 'market', 'zona', 'ambito', 'region'],
}

def _clave_ota_hotel(ota, hotel):
    """Clave (OTA, hotel) para deduplicar tarifas pactadas.

    Hecho con Python plano y NO con el accesor .str a proposito: en pandas 3
    astype(str) YA NO convierte los nulos a la cadena 'nan' -- los deja como
    NaN -- y .str.lower() los propaga. Con eso, TODAS las filas sin hotel
    acababan compartiendo la clave NaN y drop_duplicates se llevaba por delante
    la tabla de tarifas entera menos una fila. Cazado al probar el cruce por
    hotel, no en revision.
    """
    def _t(v):
        s = '' if v is None else str(v)
        return '' if s.strip().lower() in ('', 'nan', 'none', '<na>', 'nat') else ' '.join(s.split()).lower()
    return _t(ota) + '|' + _t(hotel)


def _ota_es_lista(v):
    """True si el campo OTA trae VARIAS OTAs en una sola cadena.

    Un acuerdo de distribucion tipico cubre Booking Y Expedia. El schema del
    clasificador tiene la OTA arriba, en singular, asi que cuando el contrato es
    multi-OTA la IA no tiene donde poner la OTA de cada tarifa y las junta:
    "Booking.com / Expedia". Esa cadena no es una OTA — es una lista — y
    estamparla en las filas colapsa el contrato y hace que ninguna factura
    cruce. Se detecta por separador tipico o por dos nombres conocidos dentro.
    """
    s = str(v or '').strip()
    if not s:
        return False
    if any(sep in s for sep in ('/', ',', ' y ', ' & ', ' + ', '+')):
        return True
    try:
        from lector_facturas_ap import OTAS_CONOCIDAS
    except Exception:
        return False
    sl = s.lower()
    return sum(1 for o in OTAS_CONOCIDAS if o in sl) >= 2


def _ap_tiene_datos(reg):
    """True si de una factura AP se ha extraido algo aprovechable.

    Basta con UNO de proveedor / numero de factura / total. Si no hay ninguno,
    la fila es una hilera de NO_ENCONTRADO: no aporta nada y hace que el
    resumen cuente una factura que en realidad no existe.
    """
    _NF = 'NO_ENCONTRADO'
    return any(reg.get(k) not in (_NF, None, '', 0)
               for k in ('nombre_proveedor', 'numero_factura', 'total_factura'))


def _tipo_documento(reg):
    """El tipo que ha dicho la IA, con FACTURA como caso especial.

    De los 13 esquemas del prompt, FACTURA es el UNICO que no lleva
    `tipo_documento`: se reconoce por `es_factura`. El camino de fotos ya hacia
    esta traduccion por su cuenta; el de hojas de calculo no, y por eso una
    planilla de facturas de proveedor perfectamente leida volvia como "hoja de
    calculo sin clasificar" y se tiraba. Ahora la traduccion esta en un sitio.
    """
    if not isinstance(reg, dict):
        return ''
    t = str(reg.get('tipo_documento') or '').strip().upper()
    if t:
        return t
    return 'FACTURA' if reg.get('es_factura') else ''


@_en_fila
def _guardar_factura_ap(filas):
    """Guarda 1..N facturas AP en el Excel del dia. UNICO sitio que las guarda.

    Antes habia tres copias de esto —PDF, foto y la entrada --file— y se habian
    desincronizado: la de las fotos usaba guardar_excel(), que SOBRESCRIBE el
    fichero, asi que escanear una factura con el movil borraba las que el lote
    hubiera guardado ese mismo dia. Reproducido antes de arreglarlo.

    Devuelve cuantas filas se han guardado.
    """
    from lector_facturas_ap import guardar_excel as _gx
    filas = [filas] if isinstance(filas, dict) else list(filas or [])
    _crudas = [f for f in filas if isinstance(f, dict)]
    # Una cuenta y un tipo POR PROVEEDOR (cuentas_proveedor): el lector decide
    # por una palabra del concepto de cada factura y salian dos cuentas para el
    # mismo distribuidor y una lavanderia en 600. Se decide aqui, con las
    # lineas a la vista, y se aprende para la siguiente factura del proveedor.
    try:
        import cuentas_proveedor as _CP
        _CP.normalizar(_crudas, str(_ddir()))
    except Exception as _e_cp:
        print(f"[cuentas_proveedor] no aplicado: {_e_cp}")
    # las claves internas (_facturas, _skip, _lineas...) no son columnas de la
    # factura: se quitan de la hoja plana. `_crudas` las conserva porque las
    # lineas de la Fase 3c viajan justamente en una de ellas.
    filas = [{k: v for k, v in f.items() if not str(k).startswith('_')}
             for f in _crudas]
    if not filas:
        return 0
    # ── El hotel al que pertenece lo que entra (fase 2) ──────────────────
    # Se estampa AQUI, despues de clasificar y leer el documento: el papel no
    # dice de que hotel es, lo dice la sesion. El clasificador no se entera.
    _hid = censo_hoteles.para_guardar()
    for _f in filas:
        _f['hotel_id'] = _hid
    ruta = os.path.join(_pdir(), f'facturas_ap_{date.today().strftime("%Y%m%d")}.xlsx')
    _lineas = [dict(l, hotel_id=_hid) for f in _crudas for l in (f.get('_lineas') or [])]
    _claves = {str(f.get('archivo', '')) for f in filas}
    if os.path.exists(ruta):
        _df = pd.concat([pd.read_excel(ruta), pd.DataFrame(filas)], ignore_index=True)
        # La identidad dentro del fichero es (archivo, hotel). Solo con el
        # nombre del fichero, dos hoteles del mismo grupo que subieran
        # `factura_enero.pdf` el mismo dia se pisaban: el segundo borraba al
        # primero. Con el hotel dentro, cada uno tiene la suya.
        _sub = [c for c in ('archivo', 'hotel_id') if c in _df.columns]
        _df.drop_duplicates(subset=_sub, keep='last', inplace=True)
        # Las lineas viejas de los MISMOS documentos se van con ellos (Fase 3c).
        # Si solo se sustituyera la cabecera quedarian lineas huerfanas de una
        # version anterior mezcladas con las nuevas, y el cruce sumaria
        # mercancia que no se facturo. Misma leccion que en _guardar_albaran.
        try:
            _lv = pd.read_excel(ruta, sheet_name='Lineas')
        except Exception:
            _lv = pd.DataFrame()
        if not _lv.empty and 'archivo' in _lv.columns:
            # Solo se van las lineas del MISMO hotel: reprocesar un documento
            # en el hotel B no puede llevarse por delante las lineas que el
            # hotel A tenia para un fichero que se llama igual.
            _fuera = _lv['archivo'].astype(str).isin(_claves)
            if 'hotel_id' in _lv.columns:
                _fuera = _fuera & (_lv['hotel_id'].astype(str) == str(_hid))
            _lv = _lv[~_fuera]
        _dfl = pd.concat([_lv, pd.DataFrame(_lineas)], ignore_index=True) \
            if (not _lv.empty or _lineas) else pd.DataFrame()
    else:
        _df = pd.DataFrame(filas)
        _dfl = pd.DataFrame(_lineas)
    # La hoja de facturas va PRIMERA: todo el mundo lee este fichero con
    # pd.read_excel(ruta) sin nombre de hoja, o sea la primera. Si se colara
    # delante otra hoja, cada consumidor de AP leeria otra cosa.
    with pd.ExcelWriter(ruta, engine='openpyxl') as _w:
        _df.to_excel(_w, sheet_name='Facturas', index=False)
        if not _dfl.empty:
            _dfl.to_excel(_w, sheet_name='Lineas', index=False)
    return len(filas)


_COLS_LINEA_ALB = ('clave', 'numero_albaran', 'nombre_proveedor', 'n_linea',
                   'descripcion', 'cantidad', 'unidad', 'precio_unitario', 'importe')


@_en_fila
def _guardar_albaran(cabecera, lineas):
    """Guarda un albaran en albaranes_YYYYMMDD.xlsx, en DOS hojas.

    Un albaran es una cabecera con N lineas y un Excel es plano. Dos hojas
    (`Albaranes` y `Lineas`) unidas por `clave` es justo lo que necesitara el
    cruce factura-albaran, y ademas se puede abrir y auditar a ojo. La
    alternativa —las lineas serializadas en una columna JSON— obliga a parsear
    en cada lectura y no se ve en Excel.

    Reprocesar el mismo albaran lo ACTUALIZA: se quitan antes su cabecera Y sus
    lineas viejas. Si solo se quitara la cabecera quedarian lineas huerfanas de
    una version anterior mezcladas con las nuevas, y el cruce sumaria mercancia
    que no se entrego.
    """
    ruta = os.path.join(_pdir(), f'albaranes_{date.today().strftime("%Y%m%d")}.xlsx')
    _hid = censo_hoteles.para_guardar()
    cabecera = dict(cabecera, hotel_id=_hid)
    lineas = [dict(l, hotel_id=_hid) for l in (lineas or [])]
    df_cab = pd.DataFrame([cabecera])
    df_lin = pd.DataFrame(lineas) if lineas else pd.DataFrame(columns=list(_COLS_LINEA_ALB) + ['hotel_id'])
    if os.path.exists(ruta):
        try:
            _cab_old = pd.read_excel(ruta, sheet_name='Albaranes')
        except Exception:
            _cab_old = pd.DataFrame()
        try:
            _lin_old = pd.read_excel(ruta, sheet_name='Lineas')
        except Exception:
            _lin_old = pd.DataFrame()
        _cl = str(cabecera.get('clave', ''))
        # Se sustituye SOLO la version de este hotel. Sin el hotel en la
        # condicion, reprocesar un albaran en el hotel B borraba el del A.
        def _quitar(_d):
            if _d.empty or 'clave' not in _d.columns:
                return _d
            _m = _d['clave'].astype(str) == _cl
            if 'hotel_id' in _d.columns:
                _m = _m & (_d['hotel_id'].astype(str) == str(_hid))
            return _d[~_m]
        _cab_old = _quitar(_cab_old)
        _lin_old = _quitar(_lin_old)
        if not _cab_old.empty:
            df_cab = pd.concat([_cab_old, df_cab], ignore_index=True)
        if not _lin_old.empty:
            df_lin = pd.concat([_lin_old, df_lin], ignore_index=True)
    with pd.ExcelWriter(ruta, engine='openpyxl') as _w:
        df_cab.to_excel(_w, sheet_name='Albaranes', index=False)
        df_lin.to_excel(_w, sheet_name='Lineas', index=False)
    return len(lineas)


_COLS_LINEA_PO = ('clave', 'numero_po', 'nombre_proveedor', 'n_linea',
                  'descripcion', 'cantidad', 'unidad', 'precio_unitario', 'importe')


@_en_fila
def _guardar_orden_compra(cabecera, lineas):
    """Guarda un PO en ordenes_compra_YYYYMMDD.xlsx, en DOS hojas.

    Clonado de `_guardar_albaran` a proposito: mismo problema (una cabecera con
    N lineas en un Excel plano), misma solucion probada (hojas `Ordenes` y
    `Lineas` unidas por `clave`, auditables a ojo desde Excel).

    Reprocesar el mismo pedido lo ACTUALIZA y se lleva sus lineas viejas. Si
    solo se sustituyera la cabecera quedarian lineas huerfanas de una version
    anterior y el cruce sumaria un compromiso que ya no existe.
    """
    ruta = os.path.join(_pdir(), f'ordenes_compra_{date.today().strftime("%Y%m%d")}.xlsx')
    _hid = censo_hoteles.para_guardar()
    cabecera = dict(cabecera, hotel_id=_hid)
    lineas = [dict(l, hotel_id=_hid) for l in (lineas or [])]
    df_cab = pd.DataFrame([cabecera])
    df_lin = pd.DataFrame(lineas) if lineas else pd.DataFrame(columns=list(_COLS_LINEA_PO) + ['hotel_id'])
    if os.path.exists(ruta):
        try:
            _cab_old = pd.read_excel(ruta, sheet_name='Ordenes')
        except Exception:
            _cab_old = pd.DataFrame()
        try:
            _lin_old = pd.read_excel(ruta, sheet_name='Lineas')
        except Exception:
            _lin_old = pd.DataFrame()
        _cl = str(cabecera.get('clave', ''))
        # Se sustituye SOLO la version de este hotel. Sin el hotel en la
        # condicion, reprocesar un albaran en el hotel B borraba el del A.
        def _quitar(_d):
            if _d.empty or 'clave' not in _d.columns:
                return _d
            _m = _d['clave'].astype(str) == _cl
            if 'hotel_id' in _d.columns:
                _m = _m & (_d['hotel_id'].astype(str) == str(_hid))
            return _d[~_m]
        _cab_old = _quitar(_cab_old)
        _lin_old = _quitar(_lin_old)
        if not _cab_old.empty:
            df_cab = pd.concat([_cab_old, df_cab], ignore_index=True)
        if not _lin_old.empty:
            df_lin = pd.concat([_lin_old, df_lin], ignore_index=True)
    with pd.ExcelWriter(ruta, engine='openpyxl') as _w:
        df_cab.to_excel(_w, sheet_name='Ordenes', index=False)
        df_lin.to_excel(_w, sheet_name='Lineas', index=False)
    return len(lineas)


def _resumen_po(cabecera, lineas):
    """Texto honesto de lo guardado. No promete pantalla: todavia no hay."""
    num = str(cabecera.get('numero_po') or '').strip() or 's/n'
    prov = str(cabecera.get('nombre_proveedor') or '').strip() or 'proveedor sin identificar'
    imp = cabecera.get('importe_aprobado')
    eur = f' — €{imp:,.2f}' if isinstance(imp, (int, float)) else ''
    dep = str(cabecera.get('departamento') or '').strip()
    # el IVA se dice SOLO cuando se sabe: adivinarlo es lo que hace falsos
    _iva = cabecera.get('iva_incluido')
    iva = ' (IVA incluido)' if _iva is True else ''
    extra = f' · {dep}' if dep else ''
    lin = f' · {len(lineas)} línea(s)' if lineas else ''
    return f'{num} · {prov}{extra}{lin}{eur}{iva}'


BONOS_FILE = 'bonos_agencia.xlsx'


def _guardar_bono(fila):
    """Guarda un bono de agencia/empresa en datos-referencia/bonos_agencia.xlsx.

    Un fichero unico (no uno por dia): un bono es una autorizacion de cobro que
    se coteja contra la factura direct bill cuando esta exista, dias o semanas
    despues. Reprocesar el mismo bono lo ACTUALIZA (misma clave y mismo hotel).
    """
    ruta = os.path.join(_ddir(), BONOS_FILE)
    _hid = censo_hoteles.para_guardar()
    fila = dict(fila, hotel_id=_hid, fecha_procesado=date.today().isoformat())
    df_new = pd.DataFrame([fila])
    if os.path.exists(ruta):
        try:
            _old = pd.read_excel(ruta)
        except Exception:
            _old = pd.DataFrame()
        if not _old.empty and 'clave' in _old.columns:
            _m = _old['clave'].map(safe_str) == str(fila.get('clave', ''))
            if 'hotel_id' in _old.columns:
                # safe_str y no astype(str): un hotel vacio vuelve de Excel como NaN
                _m = _m & (_old['hotel_id'].map(safe_str) == str(_hid))
            _old = _old[~_m]
        df_new = pd.concat([_old, df_new], ignore_index=True)
    df_new.to_excel(ruta, index=False)


def _resumen_bono(fila):
    num = str(fila.get('numero_bono') or '').strip() or 's/n'
    ag = str(fila.get('agencia') or '').strip() or 'pagador sin identificar'
    fe = str(fila.get('fecha_entrada') or '').strip()
    tot = fila.get('importe_total')
    eur = f' — €{tot:,.2f}' if isinstance(tot, (int, float)) else ''
    return f'{num} · {ag}' + (f' · entrada {fe}' if fe else '') + eur


def _resumen_albaran(cabecera, lineas):
    """Texto honesto de lo guardado. No promete pantalla: todavia no hay."""
    num = str(cabecera.get('numero_albaran') or '').strip() or 's/n'
    prov = str(cabecera.get('nombre_proveedor') or '').strip() or 'proveedor sin identificar'
    tot = cabecera.get('total_albaran')
    eur = f' — €{tot:,.2f}' if isinstance(tot, (int, float)) else ''
    return f'{num} · {prov} · {len(lineas)} línea(s){eur}'


def _resumen_factura_ap(filas):
    """Texto honesto de lo guardado: una factura, o cuantas y por cuanto."""
    def _importe(f):
        v = f.get('total_factura')
        return float(v) if isinstance(v, (int, float)) else 0.0
    if not filas:
        return 'sin datos aprovechables'
    if len(filas) == 1:
        prov = str(filas[0].get('nombre_proveedor', '') or '')
        v = filas[0].get('total_factura')
        return f'{prov} — €{v:,.2f}' if isinstance(v, (int, float)) else prov
    provs = []
    for f in filas:
        p = str(f.get('nombre_proveedor') or '').strip()
        if p and p != 'NO_ENCONTRADO' and p not in provs:
            provs.append(p)
    quien = provs[0] if len(provs) == 1 else f'{len(provs)} proveedores'
    return f'{len(filas)} facturas · {quien} — €{sum(_importe(f) for f in filas):,.2f}'


def _hoja_a_texto(fpath, max_filas=200, max_chars=8000):
    """Vuelca una hoja de calculo a texto plano para que la IA pueda clasificarla.

    Lee TODAS las hojas de un xlsx/xlsm (no solo la primera) y las concatena en
    CSV, que es lo mas compacto y legible para el modelo. Recorta a max_filas por
    hoja y max_chars en total. Devuelve '' si no se puede leer: quien llama debe
    tratar la cadena vacia como "no clasificable".
    """
    try:
        ext = os.path.splitext(fpath)[1].lower()
        partes = []
        if ext == '.csv':
            try:
                # sep=None deja que pandas olfatee el separador (los CSV
                # espanoles suelen venir con ';' en vez de ',')
                _df = pd.read_csv(fpath, nrows=max_filas, dtype=str,
                                  keep_default_na=False, sep=None, engine='python')
            except Exception:
                _df = pd.read_csv(fpath, nrows=max_filas, dtype=str, keep_default_na=False)
            partes.append(_df.to_csv(index=False))
        else:
            _xl = pd.ExcelFile(fpath)
            for _hoja in _xl.sheet_names[:10]:
                _df = _xl.parse(_hoja, nrows=max_filas, dtype=str)
                if _df.empty:
                    continue
                partes.append(f'--- hoja: {_hoja} ---\n' + _df.to_csv(index=False))
        return '\n'.join(partes).strip()[:max_chars]
    except Exception as _ehoja:
        print(f'[_hoja_a_texto] {os.path.basename(fpath)}: {_ehoja}')
        return ''


def _procesar_drr(fpath, fname):
    """Ejecuta lector_drr.py sobre un .xlsm y devuelve (mensaje, marca).

    ANTES: el fichero se copiaba a reportes/drr_upload.xlsm, se decia
    "✓ DRR cargado" y ahi moria — nadie lo procesaba, asi que /api/stats_drr se
    quedaba en null para siempre. El lector ya existia y genera exactamente el
    drr_procesado_*.xlsx que el dashboard sabe leer; solo faltaba llamarlo.

    NUNCA decimos "cargado" si no se ha procesado. El lector esta probado con un
    .xlsm sintetico con la estructura que espera (DAILY_MASTER, hojas 1-31,
    CtaCble), NO con un fichero real de un hotel: es perfectamente posible que el
    primero que llegue no encaje. Por eso el fallo se cuenta como recibido-sin-
    procesar y con el motivo delante, no como exito.
    """
    import subprocess as _spd, glob as _gd, shutil as _shd
    rdir = _rdir()
    os.makedirs(rdir, exist_ok=True)
    # copia de trabajo dentro del arbol del tenant (antes iba a la ruta base)
    destino = os.path.join(rdir, 'drr_upload.xlsm')
    try:
        _shd.copy2(fpath, destino)
    except Exception as e:
        return f'⚠ DRR {fname}: recibido, no se ha podido guardar — {str(e)[:60]}', 'DRR_RECIBIDO'

    # El lector escribe SIEMPRE drr_procesado_<hoy>.xlsx, asi que una segunda
    # subida del mismo dia machaca la primera ANTES de que podamos comprobar si
    # ha salido bien. Guardamos copia y la devolvemos si el intento falla: subir
    # un fichero que no es un DRR no puede costarte el DRR bueno de esta mañana.
    import tempfile as _tf
    previos = set(_gd.glob(os.path.join(rdir, 'drr_procesado_*.xlsx')))
    respaldo = _tf.mkdtemp(prefix='drr_prev_')
    for _p in previos:
        try:
            _shd.copy2(_p, os.path.join(respaldo, os.path.basename(_p)))
        except OSError:
            pass

    def _restaurar():
        """Deja los informes como estaban antes del intento."""
        for _q in _gd.glob(os.path.join(rdir, 'drr_procesado_*.xlsx')):
            if _q not in previos:
                try:
                    os.remove(_q)          # lo creo este intento fallido
                except OSError:
                    pass
        for _b in _gd.glob(os.path.join(respaldo, '*.xlsx')):
            try:
                _shd.copy2(_b, os.path.join(rdir, os.path.basename(_b)))
            except OSError:
                pass
        _shd.rmtree(respaldo, ignore_errors=True)

    # OJO: el lector escribe siempre el MISMO nombre, asi que "ha aparecido un
    # fichero nuevo" no sirve para saber si ha funcionado — al reprocesar el
    # mismo dia no aparece ninguno. Nos fijamos en la HUELLA: vale como salida
    # todo informe nuevo O que haya cambiado de contenido.
    def _huellas():
        h = {}
        for _q in _gd.glob(os.path.join(rdir, 'drr_procesado_*.xlsx')):
            h[_q] = _huella_fichero(_q)
        return h

    antes_h = _huellas()
    try:
        r = _spd.run([sys.executable, 'lector_drr.py', destino],
                     cwd=BASE_DIR, capture_output=True, text=True,
                     timeout=180, env=_env_tenant())
    except _spd.TimeoutExpired:
        _restaurar()
        return (f'⚠ DRR {fname}: recibido, no se ha podido procesar — '
                f'tarda demasiado (>180 s)'), 'DRR_RECIBIDO'
    except Exception as e:
        _restaurar()
        return (f'⚠ DRR {fname}: recibido, no se ha podido procesar — '
                f'{str(e)[:60]}'), 'DRR_RECIBIDO'

    ahora_h = _huellas()
    salida = sorted(q for q, h in ahora_h.items() if antes_h.get(q) != h)
    if r.returncode != 0 or not salida:
        # el motivo util suele estar en la ultima linea del stderr
        motivo = ''
        for linea in reversed((r.stderr or '').strip().splitlines()):
            if linea.strip():
                motivo = linea.strip()[:70]
                break
        if not motivo:
            motivo = 'el archivo no tiene la estructura de un DRR (DAILY_MASTER, hojas 1-31)'
        _restaurar()
        return f'⚠ DRR {fname}: recibido, no se ha podido procesar — {motivo}', 'DRR_RECIBIDO'

    # Contar lo que de verdad se ha extraido, para no decir un exito vacio
    dias = oob = 0
    try:
        _df = pd.read_excel(salida[-1], sheet_name='Trial_Balance_Completo')
        dias = int(_df['Día'].nunique())
        oob = int(_df[_df['Out of Balance'].astype(str).str.contains('OOB', na=False)]['Día'].nunique())
    except Exception:
        pass
    if dias == 0:
        # Un informe vacio NO se deja en disco: _cargar_drr_procesado() coge el
        # mas reciente, asi que taparia un DRR bueno anterior y el panel
        # enseñaria ceros como si hubiera datos. Justo lo que estamos quitando.
        _restaurar()
        return (f'⚠ DRR {fname}: recibido, procesado sin datos — '
                f'no se ha extraido ningun dia'), 'DRR_RECIBIDO'
    _shd.rmtree(respaldo, ignore_errors=True)
    extra = f' · {oob} día(s) descuadrado(s)' if oob else ' · todo cuadrado'
    return f'✓ DRR {fname}: {dias} día(s) procesado(s){extra}', 'DRR_OK'


def _enrutar_tipo_doc(reg, fname, fpath=None):
    """Enruta un documento YA clasificado (reg['tipo_documento']) al modulo que toca.

    Extraido tal cual del bucle de /api/procesar_batch_stream para poder
    reutilizarlo desde otros puntos de entrada (p.ej. hojas de calculo).
    NO cambia comportamiento: cada rama produce exactamente un mensaje y una marca.

    Devuelve (mensaje_sin_prefijo_sse, marca_para_el_log, flags).
    """
    _msg = ''
    _marca = 'SKIP'
    _flags = {}
    # Claude clasificó el documento como otro tipo — enrutar
    _tipo_doc = reg['tipo_documento']
    if _tipo_doc == 'FACTURA':
        # Una factura de proveedor tambien es un tipo clasificado: hasta ahora
        # solo el camino de PDF sabia guardarla, asi que la misma factura en
        # una hoja de calculo se perdia. Un documento puede traer varias.
        from lector_facturas_ap import facturas_de_respuesta, cargar_proveedores
        _filas = [f for f in facturas_de_respuesta(reg, fname, cargar_proveedores())
                  if _ap_tiene_datos(f)]
        if not _filas:
            _msg = f'⚠ {fname}: no se pudo extraer ningún dato de la factura — revisar manualmente'
            _marca = 'SKIP'
        else:
            _guardar_factura_ap(_filas)
            _msg = f'✓ AP {fname}: {_resumen_factura_ap(_filas)}'
            _marca = 'AP_OK'
            _flags['has_ap'] = True
            _flags['ap_n'] = len(_filas)
    elif _tipo_doc == 'ORDEN_COMPRA':
        # Un PEDIDO no es una factura por pagar, y esta es la fuga que esta rama
        # tapa: sin tipo ORDEN_COMPRA lo mas parecido que encontraba el prompt
        # era FACTURA, asi que el pedido entraba en facturas_ap, se le asignaba
        # cuenta y acababa esperando aprobacion de PAGO. Se pagaria el pedido Y
        # luego su factura. Es el mismo agujero que tenia el albaran.
        # NO marca has_ap, NO lanza el asignador, NO escribe en facturas_ap.
        from lector_facturas_ap import orden_compra_de_respuesta, po_tiene_datos
        _cab_po, _lin_po = orden_compra_de_respuesta(reg, fname)
        if not po_tiene_datos(_cab_po, _lin_po):
            _msg = (f'⚠ {fname}: parece una orden de compra, pero no se ha podido '
                    f'extraer ni proveedor ni importe — revisar manualmente')
            _marca = 'SKIP'
        else:
            _guardar_orden_compra(_cab_po, _lin_po)
            _msg = f'✓ Orden de compra {_resumen_po(_cab_po, _lin_po)}'
            _marca = 'PO_OK'
            _flags['orden_compra'] = True
    elif _tipo_doc == 'ALBARAN':
        # Una nota de entrega NO es una factura por pagar. Antes el prompt no
        # sabia que existia el albaran, asi que el mas parecido que encontraba
        # era FACTURA: se guardaba en facturas_ap, se le asignaba cuenta y
        # acababa esperando aprobacion de PAGO. Reproducido antes de arreglarlo.
        from lector_facturas_ap import albaran_de_respuesta, albaran_tiene_datos
        _cab_alb, _lin_alb = albaran_de_respuesta(reg, fname)
        if not albaran_tiene_datos(_cab_alb, _lin_alb):
            _msg = (f'⚠ {fname}: albarán detectado, pero no se ha podido extraer '
                    f'ninguna línea — revisar manualmente')
            _marca = 'SKIP'
        else:
            _guardar_albaran(_cab_alb, _lin_alb)
            _msg = f'✓ Albarán {_resumen_albaran(_cab_alb, _lin_alb)}'
            _marca = 'ALBARAN_OK'
            _flags['albaran'] = True
    elif _tipo_doc == 'BONO':
        # Un bono de agencia/empresa NO es una factura ni una rooming: es la
        # autorizacion para facturar a credito. Se guarda para cotejarlo con
        # la factura direct bill (AR Real). NO toca AP.
        from lector_facturas_ap import bono_de_respuesta, bono_tiene_datos
        _bono = bono_de_respuesta(reg, fname)
        if not bono_tiene_datos(_bono):
            _msg = (f'⚠ {fname}: parece un bono de agencia, pero no se ha podido '
                    f'extraer ni quien paga ni importe/fechas — revisar manualmente')
            _marca = 'SKIP'
        else:
            _guardar_bono(_bono)
            _msg = f'✓ Bono {_resumen_bono(_bono)}'
            _marca = 'BONO_OK'
            _flags['bono'] = True
    elif _tipo_doc == 'EXTRACTO_BANCO' and reg.get('movimientos'):
        try:
            _movs = reg['movimientos']
            _df_movs = _normalize_cols(pd.DataFrame(_movs), _BANK_COL_MAP)
            _estampar_hotel_banco(_df_movs)   # modo por_hotel: marca las filas nuevas
            banco_path = os.path.join(_ddir(), 'extracto_banco.xlsx')
            if os.path.exists(banco_path):
                _df_exist = pd.read_excel(banco_path)
                _df_movs = pd.concat([_df_exist, _df_movs], ignore_index=True)
            _df_movs.to_excel(banco_path, index=False)
            n_movs = len(reg['movimientos'])
            total_cargo = sum(float(m.get('importe',0) or 0) for m in reg['movimientos'] if float(m.get('importe',0) or 0) < 0)
            total_abono = sum(float(m.get('importe',0) or 0) for m in reg['movimientos'] if float(m.get('importe',0) or 0) > 0)
            _msg = f'✓ Banco {fname}: {n_movs} movimientos (cargos €{abs(total_cargo):,.0f} / abonos €{total_abono:,.0f}) integrados'
            _marca = 'BANK_OK'
        except Exception as _eb2:
            _msg = f'⚠ {fname}: extracto detectado pero error al guardar — {str(_eb2)[:60]}'
            _marca = 'SKIP'
    elif _tipo_doc == 'VENTAS_POS':
        total = reg.get('total_ventas', 0)
        platos = reg.get('platos', [])
        # Integrar ventas detalladas en ventas_fb_diarias
        try:
            if platos:
                _df_ventas = _normalize_cols(pd.DataFrame(platos), _VEN_COL_MAP)
                fecha = reg.get('fecha', date.today().isoformat())
                if 'fecha' not in _df_ventas.columns:
                    _df_ventas['fecha'] = fecha
                _df_ventas, _ = _guardar_fb_del_hotel(_df_ventas, 'ventas_fb_diarias.xlsx')
                _msg = f'✓ F&B {fname}: {len(platos)} platos, €{total} integrados por IA'
            else:
                _msg = f'✓ F&B {fname}: ventas detectadas — €{total}'
        except Exception as _efb2:
            _msg = f'✓ F&B {fname}: ventas — €{total} (detalle no integrable: {str(_efb2)[:40]})'
        _marca = 'FB_OK'
    elif _tipo_doc == 'COMISIONES_OTA':
        # Antes esta rama SOLO imprimia un mensaje: el dato se perdia y el
        # verificador de comisiones no tenia nada que cruzar. Ahora escribe en
        # facturas_procesadas_*.xlsx, que es de donde lee verificador_comisiones.
        ota = reg.get('ota', '?')
        comision = reg.get('comision', 0)
        _facts_pre = reg.get('facturas') or []
        # Misma regla de producto que el albaran y el PO: si no ha salido nada
        # aprovechable NO se dice "✓". Sin nombre de OTA y sin una sola cifra
        # esto no es una factura de comisiones — antes se guardaba una fila
        # fantasma "1 factura(s) de ? — €0.00" y salia en verde.
        def _cifra(v):
            try:
                return float(str(v).replace(',', '.')) not in (0.0,)
            except (TypeError, ValueError):
                return False
        _ota_ok = str(ota or '').strip() not in ('', '?', '—', 'None')
        _hay_cifras = _cifra(comision) or _cifra(reg.get('importe_bruto')) or any(
            _cifra(f.get('importe_comision')) or _cifra(f.get('importe_bruto'))
            for f in _facts_pre if isinstance(f, dict))
        if not _ota_ok and not _hay_cifras:
            _msg = (f'⚠ {fname}: parece un informe de comisiones de OTA, pero no se ha '
                    f'podido leer ni la OTA ni ningún importe — revisar manualmente')
            _marca = 'SKIP'
            return _msg, _marca, _flags
        try:
            _NFO = 'NO_ENCONTRADO'
            _facts = _facts_pre
            if _facts:
                _df_ota = _normalize_cols(pd.DataFrame(_facts), _OTA_COL_MAP)
            else:
                # La IA solo dio el agregado: una fila con lo que haya
                _df_ota = pd.DataFrame([{
                    'numero_factura': reg.get('numero_factura'),
                    'periodo_inicio': reg.get('periodo'),
                    'importe_bruto': reg.get('importe_bruto'),
                    'porcentaje_comision': reg.get('porcentaje'),
                    'importe_comision': comision,
                }])
            _df_ota['archivo'] = fname
            # ── El hotel al que pertenece (fase 3) ───────────────────────
            _hid = censo_hoteles.para_guardar()
            _df_ota['hotel_id'] = _hid
            if 'nombre_ota' not in _df_ota.columns:
                _df_ota['nombre_ota'] = ota
            _df_ota['nombre_ota'] = _df_ota['nombre_ota'].astype(object).fillna(ota)
            # El verificador espera estas columnas; las que falten, NO_ENCONTRADO
            for _c in ('numero_factura', 'fecha', 'nombre_hotel', 'periodo_inicio',
                       'periodo_fin', 'importe_bruto', 'porcentaje_comision',
                       'importe_comision', 'importe_neto'):
                if _c not in _df_ota.columns:
                    _df_ota[_c] = _NFO
            _df_ota = _df_ota.astype(object).fillna(_NFO)
            _ota_xlsx = os.path.join(_pdir(), f'facturas_procesadas_{date.today().strftime("%Y%m%d")}.xlsx')
            if os.path.exists(_ota_xlsx):
                _df_prev = pd.read_excel(_ota_xlsx)
                _df_ota = pd.concat([_df_prev, _df_ota], ignore_index=True)
                if 'numero_factura' in _df_ota.columns:
                    # Solo deduplicar filas con numero de factura real:
                    # drop_duplicates considera NaN==NaN, asi que dos facturas
                    # del mismo fichero sin numero se fusionaban en una.
                    _num = _df_ota['numero_factura'].map(
                        lambda v: str(v).strip().lower() not in ('', 'nan', 'none', 'no_encontrado'))
                    _con, _sin = _df_ota[_num], _df_ota[~_num]
                    # El hotel entra en la identidad: dos hoteles del mismo
                    # grupo reciben liquidaciones con el mismo numero de la
                    # misma OTA, y sin el hotel una borraba a la otra.
                    _sub = [c for c in ('archivo', 'numero_factura', 'hotel_id') if c in _con.columns]
                    _con = _con.drop_duplicates(subset=_sub, keep='last')
                    _df_ota = pd.concat([_con, _sin], ignore_index=True)
            os.makedirs(_pdir(), exist_ok=True)
            _df_ota.to_excel(_ota_xlsx, index=False)
            _n_f = len(_facts) if _facts else 1
            _com_txt = f'{float(comision):,.2f}' if isinstance(comision, (int, float)) else str(comision or '—')
            _msg = f'✓ AR {fname}: {_n_f} factura(s) de {ota} — €{_com_txt} en comisiones, guardadas para verificar'
            # El documento trae su propio nombre de hotel. NO se usa para
            # asignar (una regex sobre un PDF no puede decidir la contabilidad
            # de nadie), pero si para avisar de que quiza te has equivocado de
            # hotel al subirlo.
            # La comparacion vive en `_aviso_otro_hotel`, compartida con la
            # rama del lote: tener dos copias es como se llego a que este
            # camino avisara y el otro —el que se usa de verdad— no.
            _nom_doc = ''
            if 'nombre_hotel' in _df_ota.columns and len(_df_ota):
                _nom_doc = str(_df_ota['nombre_hotel'].iloc[0] or '')
            _av = _aviso_otro_hotel(_nom_doc)
            if _av:
                _msg += ' · ⚠ ' + _av
            _marca = 'AR_OK'
            _flags['has_ar'] = True
        except Exception as _eota:
            _msg = f'⚠ {fname}: comisiones {ota} detectadas pero no se pudieron guardar — {str(_eota)[:60]}'
            _marca = 'SKIP'
    elif _tipo_doc == 'CONTRATO_OTA':
        # Tarifas PACTADAS con la OTA -> comisiones_pactadas.xlsx. Es el otro
        # lado del cruce: sin esto el verificador no sabe que porcentaje deberia
        # haberse aplicado. NO va a facturas: aqui no hay nada devengado.
        ota = reg.get('ota', '?')
        _tarifas_pre = reg.get('tarifas') or []
        # Un contrato de OTA SON sus tarifas pactadas: sin ninguna no hay nada
        # que cruzar despues. Antes decia "0 tarifa(s) pactada(s) de ? guardadas"
        # en verde, que es un exito con las manos vacias.
        if not _tarifas_pre and str(ota or '').strip() in ('', '?', '—', 'None'):
            _msg = (f'⚠ {fname}: parece un contrato de OTA, pero no se ha podido leer '
                    f'ni la OTA ni ninguna tarifa pactada — revisar manualmente')
            _marca = 'SKIP'
            return _msg, _marca, _flags
        try:
            _tarifas = _tarifas_pre
            if _tarifas:
                _df_p = _normalize_cols(pd.DataFrame(_tarifas), _PACT_COL_MAP)
            else:
                _df_p = pd.DataFrame([{'Porcentaje_Comision': reg.get('porcentaje_pactado',
                                                                      reg.get('porcentaje'))}])
            # La OTA de CADA fila manda. La de arriba (`ota`) solo se usa para
            # rellenar la fila que no traiga la suya, y SOLO si es una OTA
            # limpia: si es una lista ("Booking.com / Expedia") NO se estampa,
            # porque no se sabe que porcentaje es de cual. Antes se estampaba y
            # el dedup por (OTA, hotel) colapsaba el contrato a la mitad.
            if 'OTA' not in _df_p.columns:
                _df_p['OTA'] = None
            _df_p['OTA'] = _df_p['OTA'].astype(object)
            if not _ota_es_lista(ota):
                _df_p['OTA'] = _df_p['OTA'].fillna(ota)
            if 'Mercado' not in _df_p.columns:
                _df_p['Mercado'] = None
            if 'Hotel' not in _df_p.columns:
                _df_p['Hotel'] = None
            # El hotel se CONSERVA: el verificador cruza por (OTA, hotel) y un
            # grupo puede tener porcentajes distintos por establecimiento. Fila
            # sin hotel = tarifa generica de esa OTA.
            _df_p = _df_p[['OTA', 'Hotel', 'Porcentaje_Comision', 'Mercado']]
            _df_p = _df_p[_df_p['Porcentaje_Comision'].notna()]
            # Una tarifa sin OTA no se puede cruzar con ninguna factura. Se
            # aparta en vez de guardarla con una OTA inventada. Con OTA por fila
            # (el schema arreglado) esto no aparta nada; con la OTA combinada y
            # sin OTA por fila, aparta el contrato entero — y entonces se avisa
            # en vez de cantar un ✓ sobre datos que no van a cruzar.
            _sin_ota = _df_p['OTA'].isna() | (_df_p['OTA'].astype(str).str.strip().isin(['', 'nan', 'None']))
            _n_sin_ota = int(_sin_ota.sum())
            _df_p = _df_p[~_sin_ota]
            if _df_p.empty:
                _msg = (f'⚠ {fname}: el contrato cubre varias OTAs ({ota}) pero las '
                        f'tarifas no dicen la OTA de cada una — no puedo separarlas con '
                        f'seguridad. Súbelo como un contrato por OTA, o revisa el documento.')
                _marca = 'SKIP'
                return _msg, _marca, _flags
            _n_t = len(_df_p)
            _pact_path = os.path.join(_ddir(), 'comisiones_pactadas.xlsx')
            if os.path.exists(_pact_path):
                _df_old_p = pd.read_excel(_pact_path)
                if 'Hotel' not in _df_old_p.columns:
                    _df_old_p['Hotel'] = None   # ficheros antiguos: todo generico
                # Si la nueva fila no trae Mercado, conservar el que ya habia
                # para esa OTA en vez de dejarlo en blanco.
                _merc = {str(r['OTA']).strip().lower(): r.get('Mercado')
                         for _, r in _df_old_p.iterrows() if str(r.get('OTA', '')).strip()}
                _df_p['Mercado'] = [
                    m if (m is not None and str(m).strip() and str(m) != 'nan')
                    else _merc.get(str(o).strip().lower(), 'NO_ENCONTRADO')
                    for o, m in zip(_df_p['OTA'], _df_p['Mercado'])]
                _df_p = pd.concat([_df_old_p, _df_p], ignore_index=True)
            else:
                _df_p['Mercado'] = _df_p['Mercado'].astype(object).fillna('NO_ENCONTRADO')
            # Una fila por (OTA, hotel): dos hoteles de la misma OTA con
            # condiciones distintas son DOS tarifas, no una que pisa a la otra.
            _df_p['_k'] = [_clave_ota_hotel(o, h)
                           for o, h in zip(_df_p['OTA'], _df_p['Hotel'])]
            _df_p.drop_duplicates(subset=['_k'], keep='last', inplace=True)
            _df_p = _df_p.drop(columns=['_k'])
            _df_p.to_excel(_pact_path, index=False)
            # El resumen lista lo que SE HA GUARDADO, por (OTA, hotel). Antes
            # filtraba por `r["OTA"] == ota`, y con la OTA de arriba combinada no
            # casaba ninguna fila: el mensaje salia "de Booking.com / Expedia (—)"
            # aun habiendo guardado 4 tarifas.
            _recien = _df_p.tail(_n_t)
            _pcts = ', '.join(
                f'{r["OTA"]} {r["Hotel"]} {r["Porcentaje_Comision"]}%'.replace('  ', ' ')
                if str(r.get('Hotel') or '').strip() not in ('', 'nan', 'None')
                else f'{r["OTA"]} {r["Porcentaje_Comision"]}%'
                for _, r in _recien.iterrows())
            _otas_txt = ', '.join(sorted({str(r['OTA']).strip() for _, r in _recien.iterrows()}))
            _aviso_sin = (f' · {_n_sin_ota} sin OTA identificable, no guardada(s)'
                          if _n_sin_ota else '')
            _msg = (f'✓ Contrato OTA {fname}: {_n_t} tarifa(s) pactada(s) de '
                    f'{_otas_txt} ({_pcts or "—"}) guardadas{_aviso_sin}')
            _marca = 'CONTRATO_OTA_OK'
            # Cambiar lo pactado obliga a re-verificar las facturas ya cargadas
            _flags['has_ar'] = True
        except Exception as _epact:
            _msg = f'⚠ {fname}: contrato de {ota} detectado pero no se pudo guardar — {str(_epact)[:60]}'
            _marca = 'SKIP'
    elif _tipo_doc == 'INVENTARIO':
        # Integrar datos de inventario en F&B
        try:
            inv_items = reg.get('items', reg.get('productos', []))
            if inv_items:
                _df_inv = _normalize_cols(pd.DataFrame(inv_items), _INV_COL_MAP)
                # El stock es de CADA hotel. Se estampa despues de leer: el
                # clasificador no se entera de nada. La clave es
                # (ingrediente, hotel), y quien la aplica es `_guardar_fb_del_hotel`
                # — este camino ya lo hacia bien y ahora lo hace en el mismo
                # sitio que los otros dos, que es lo que impide que vuelvan a
                # separarse.
                _df_inv, _ = _guardar_fb_del_hotel(_df_inv, 'inventario.xlsx')
                nombres = [str(i.get('ingrediente', i.get('producto', '?')))[:20] for i in inv_items[:5]]
                _msg = f'✓ Inventario {fname}: {len(inv_items)} productos ({", ".join(nombres)}{"..." if len(inv_items)>5 else ""}) integrados'
            else:
                _msg = f'ℹ {fname}: inventario detectado (sin items extraíbles)'
        except Exception as _einv:
            _msg = f'ℹ {fname}: inventario detectado — {str(_einv)[:60]}'
        _marca = 'INV_OK'
    elif _tipo_doc == 'MERMAS':
        # Integrar datos de mermas en F&B
        try:
            merma_items = reg.get('items', reg.get('mermas', []))
            if merma_items:
                _df_mer = _normalize_cols(pd.DataFrame(merma_items), _MER_COL_MAP)
                _df_mer, _ = _guardar_fb_del_hotel(_df_mer, 'mermas.xlsx')
                total_merma = sum(float(m.get('coste_merma', m.get('coste', 0)) or 0) for m in merma_items)
                _msg = f'✓ Mermas {fname}: {len(merma_items)} registros — €{total_merma:.2f} extraídos por IA'
            else:
                _msg = f'ℹ {fname}: mermas detectadas (sin items extraíbles)'
        except Exception as _emer:
            _msg = f'ℹ {fname}: mermas detectadas — {str(_emer)[:60]}'
        _marca = 'INV_OK'
    elif _tipo_doc in ('BEO', 'TM', 'CONTRATO'):
        # Guardar como documento de referencia para matching
        try:
            ref_path = os.path.join(_ddir(), 'eventos_referencia.json')
            refs = json.load(open(ref_path)) if os.path.exists(ref_path) else []

            evento = reg.get('evento', reg.get('cliente', fname))
            cliente = reg.get('cliente', '—')
            total = reg.get('total_estimado', reg.get('importe_total', 0))
            items = reg.get('items', reg.get('requisitos', []))

            # Buscar si ya existe un evento con el mismo nombre
            evento_key = evento.lower().strip()[:50]
            found = False
            for ref in refs:
                if ref.get('evento_key') == evento_key:
                    ref['documentos'][_tipo_doc] = {
                        'archivo': fname,
                        'total': total,
                        'items': items,
                        'fecha': date.today().isoformat(),
                        'raw': {k:v for k,v in reg.items() if k != 'items' and k != 'requisitos'}
                    }
                    found = True
                    break

            if not found:
                refs.append({
                    'evento': evento,
                    'evento_key': evento_key,
                    'cliente': cliente,
                    'documentos': {
                        _tipo_doc: {
                            'archivo': fname,
                            'total': total,
                            'items': items,
                            'fecha': date.today().isoformat(),
                            'raw': {k:v for k,v in reg.items() if k != 'items' and k != 'requisitos'}
                        }
                    }
                })

            json.dump(refs, open(ref_path, 'w'), indent=2, ensure_ascii=False)

            n_items = len(items)
            docs_evento = [ref for ref in refs if ref.get('evento_key') == evento_key]
            n_docs = len(docs_evento[0]['documentos']) if docs_evento else 1
            tipos_docs = ', '.join(docs_evento[0]['documentos'].keys()) if docs_evento else _tipo_doc

            total_str = f' — €{total:,.2f}' if total else ''
            _msg = f'✓ {_tipo_doc} {fname}: {evento} ({cliente}){total_str} · {n_items} items · Evento tiene {n_docs} docs ({tipos_docs})'
            _marca = f'{_tipo_doc}_OK'
        except Exception as _eref:
            _msg = f'⚠ {fname}: {_tipo_doc} detectado pero error: {str(_eref)[:60]}'
            _marca = 'SKIP'
    elif _tipo_doc == 'ROOMING':
        grupo = reg.get('grupo', '—')
        habs = reg.get('num_habitaciones', '?')
        checkin = reg.get('checkin', '')
        checkout = reg.get('checkout', '')
        tarifa = reg.get('tarifa_media', '')
        info_parts = [f'{habs} hab.']
        if checkin: info_parts.append(f'{checkin}→{checkout}')
        if tarifa: info_parts.append(f'€{tarifa}/noche')
        # Guardar datos de rooming
        try:
            rooming_path = os.path.join(_ddir(), 'rooming_grupos.json')
            rooming_data = json.load(open(rooming_path)) if os.path.exists(rooming_path) else []
            rooming_data.append({
                'archivo': fname, 'grupo': grupo,
                'habitaciones': habs, 'checkin': checkin,
                'checkout': checkout, 'tarifa': tarifa,
                'fecha_procesado': date.today().isoformat()
            })
            json.dump(rooming_data, open(rooming_path, 'w'), indent=2, ensure_ascii=False)
        except Exception:
            pass
        _msg = f'✓ Rooming {fname}: {grupo} — {", ".join(info_parts)} (IA)'
        _marca = 'ROOMING'
    elif _tipo_doc == 'OTRO':
        desc = reg.get('descripcion', 'no clasificable')
        _msg = f'⚠ {fname}: {desc}'
        _marca = 'SKIP'
    elif _tipo_doc == 'DRR':
        # Un DRR con nombre neutro cae al clasificador. Antes no existia el tipo
        # y lo mas probable era que dijese EXTRACTO_BANCO (filas con fechas,
        # debe y haber): 31 dias de contabilidad entrando en el libro de banco.
        # Ahora se identifica y se manda al lector de DRR, que si sabe leerlo.
        if fpath:
            _msg, _marca = _procesar_drr(fpath, fname)
        else:
            _msg = (f'ℹ {fname}: parece un DRR — subelo desde el boton "Subir DRR" '
                    f'del tab DRR para procesarlo')
            _marca = 'DRR_RECIBIDO'
    else:
        _msg = f'ℹ {fname}: tipo {_tipo_doc} detectado por IA'
        _marca = 'SKIP'
    return _msg, _marca, _flags


@app.route('/api/procesar_batch_stream')
@login_required
def api_procesar_batch_stream():
    """SSE stream — procesa archivos en serie, timeout 60s por archivo."""
    # ANTES de abrir el stream: una vez empieza el SSE ya no se puede devolver
    # un 409, iria dentro del cuerpo y el cliente lo leeria como una linea mas.
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
    import json as _json
    archivos_str = request.args.get('archivos', '[]')
    try:
        archivos = _json.loads(archivos_str)
    except Exception:
        archivos = []

    log = _load_proc_log()
    from datetime import datetime as _dt2

    lote = {}   # solo lo procesado en ESTA tanda (el resumen final se cuenta de aqui)

    def _mark(fname, result='OK'):
        log[fname] = _entrada_proc(result)                      # M8
        lote[fname] = {'resultado': result}
        _save_proc_log(log)

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
            yield f'data: >> Procesando {total} archivo(s)...\n\n'
            has_ar = False; has_ap = False; has_ar_real = False
            has_albaran = False
            # un fichero puede traer VARIAS facturas: el resumen cuenta facturas
            ap_extra = 0

            # Las fotos valen como cualquier documento digital: primero se clasifican.
            # Si son un contrato de grupo (multipágina) -> AR Real. Si no, cada foto se
            # procesa como un documento suelto (factura, albarán...) igual que un PDF.
            _IMG_EXT = ('.jpg', '.jpeg', '.png', '.webp', '.heic')
            imgs = [a for a in archivos if os.path.splitext(a)[1].lower() in _IMG_EXT]
            docs = [a for a in archivos if a not in imgs]
            _contrato_res = None
            if imgs:
                yield f'data: >> Analizando {len(imgs)} foto(s)...\n\n'
                yield ': ping\n\n'
                try:
                    from lector_contratos_grupo import procesar_contrato_grupo as _pcg
                    _cpaths = [os.path.join(_edir(), a) for a in imgs if os.path.exists(os.path.join(_edir(), a))]
                    _contrato_res = _pcg(_cpaths)
                    # Se mira la MARCA, no el texto del error: fiarse de una
                    # frase concreta deja fuera cualquier motivo nuevo (paso con
                    # el contrato del que no se lee ningun dato aprovechable).
                    if not _contrato_res.get('ok') and (
                            _contrato_res.get('reprocesar')
                            or 'no parecen un contrato' in str(_contrato_res.get('error', ''))):
                        yield 'data: >> No es un contrato — proceso las fotos como documentos individuales\n\n'
                        docs = docs + imgs
                        _contrato_res = None
                except Exception as _ecc:
                    yield f'data: ⚠ No se pudieron clasificar las fotos: {str(_ecc)[:60]}\n\n'
                    _contrato_res = None

            for i, fname in enumerate(docs):
                fpath = os.path.join(_edir(), fname)
                if not os.path.exists(fpath):
                    yield f'data: ✗ {fname}: no encontrado\n\n'
                    continue

                # Límite 30MB
                try:
                    size_mb = os.path.getsize(fpath) / (1024*1024)
                    if size_mb > 30:
                        yield f'data: ✗ {fname}: demasiado grande ({size_mb:.0f}MB)\n\n'
                        _mark(fname, 'ERR:TOO_LARGE')
                        continue
                except: pass

                tipo = _detect_file_type(fname)
                yield f'data: >> [{i+1}/{len(docs)}] {fname}...\n\n'
                # SSE keep-alive para evitar timeout de Render en procesados largos
                yield ': ping\n\n'

                try:
                    import subprocess as _sp

                    fl = fname.lower()
                    _ext_lower = os.path.splitext(fname)[1].lower()
                    _is_spreadsheet = _ext_lower in ('.xlsx', '.xls', '.csv')
                    _is_image = _ext_lower in ('.jpg', '.jpeg', '.png', '.webp', '.heic')

                    # Clasificación inteligente — keywords SOLO para hojas de cálculo.
                    # Las FOTOS van siempre al lector universal (Claude las clasifica:
                    # factura, extracto, ventas, comisiones OTA, inventario, mermas...).
                    is_ar = (not _is_image) and (tipo == 'AR' or (tipo == 'AR_o_AP' and any(
                        x in fl for x in ['booking','expedia','hotels','despegar','ota','comision','commission']
                    )))
                    # El nombre propone y las cabeceras confirman (ver _destino_capa1).
                    # Las fotos ya no las captura ninguna regla de nombre: un
                    # 'rooming_grupo.jpg' iba antes a una rama que no extraia nada
                    # en vez de llegar a Claude Vision.
                    _destino = _destino_capa1(fname, fpath)
                    is_drr_file = _destino == 'DRR'
                    is_bank     = _destino == 'BANCO'
                    is_fb       = _destino == 'F&B'
                    is_inventory= _destino == 'INVENTARIO'
                    is_merma    = _destino == 'MERMAS'
                    is_rooming  = _destino == 'ROOMING'

                    if is_drr_file:
                        yield f'data: >> Procesando DRR {fname} (puede tardar)...\n\n'
                        _m_drr, _marca_drr = _procesar_drr(fpath, fname)
                        yield f'data: {_m_drr}\n\n'
                        _mark(fname, _marca_drr)
                        continue
                    if is_bank:
                        ext = os.path.splitext(fname)[1].lower()
                        try:
                            import pandas as _pdb
                            if ext == '.csv':
                                _df_bank = _pdb.read_csv(fpath)
                            elif ext in ('.xlsx', '.xls'):
                                _df_bank = _pdb.read_excel(fpath)
                            else:
                                import shutil as _sh3b
                                _sh3b.copy2(fpath, os.path.join(_ddir(), 'extracto_banco_upload' + ext))
                                yield f'data: ✓ Banco {fname}: archivo copiado (formato {ext})\n\n'
                                _mark(fname, 'BANK_OK')
                                continue
                            # Cabeceras a canonicas (Fecha/Concepto/Importe/Saldo, date/amount...):
                            # sin esto un extracto con mayusculas se guardaba con SUS columnas y
                            # la conciliacion y el cuadre lo veian como filas vacias.
                            _df_bank = _normalize_cols(_df_bank, _BANK_COL_MAP)
                            _estampar_hotel_banco(_df_bank)   # modo por_hotel: marca las filas nuevas
                            # Integrar con extracto existente
                            banco_path = os.path.join(_ddir(), 'extracto_banco.xlsx')
                            if os.path.exists(banco_path):
                                _df_exist = _pdb.read_excel(banco_path)
                                _df_bank = _pdb.concat([_df_exist, _df_bank], ignore_index=True)
                                _df_bank.drop_duplicates(keep='last', inplace=True)
                            _df_bank.to_excel(banco_path, index=False)
                            yield f'data: ✓ Banco {fname}: {len(_df_bank)} movimientos integrados\n\n'
                        except Exception as _eb:
                            import shutil as _sh3c
                            _sh3c.copy2(fpath, os.path.join(_ddir(), 'extracto_banco_upload' + ext))
                            yield f'data: ✓ Banco {fname}: archivo cargado (revisar formato)\n\n'
                        _mark(fname, 'BANK_OK')
                        continue
                    if is_fb:
                        ext = os.path.splitext(fname)[1].lower()
                        try:
                            import pandas as _pdf
                            if ext == '.csv':
                                _df_fb = _pdf.read_csv(fpath)
                            elif ext in ('.xlsx', '.xls'):
                                _df_fb = _pdf.read_excel(fpath)
                            else:
                                import shutil as _sh4b
                                _sh4b.copy2(fpath, os.path.join(_ddir(), 'ventas_fb_upload' + ext))
                                yield f'data: ✓ F&B {fname}: archivo copiado\n\n'
                                _mark(fname, 'FB_OK')
                                continue
                            
                            # Las columnas, con el mismo mapa que las otras
                            # dos puertas: sin normalizar, el panel no encuentra
                            # ni `nombre_plato` ni `unidades_vendidas`.
                            _df_fb = _normalize_cols(_df_fb, _VEN_COL_MAP)
                            _df_fb, _n_fb = _guardar_fb_del_hotel(_df_fb, 'ventas_fb_diarias.xlsx')
                            yield (f'data: ✓ F&B {fname}: {_n_fb} registros integrados '
                                   f'· {len(_df_fb)} en el hotel\n\n')
                        except Exception as _efb:
                            import shutil as _sh4c
                            _sh4c.copy2(fpath, os.path.join(_ddir(), 'ventas_fb_upload' + ext))
                            yield f'data: ✓ F&B {fname}: archivo cargado (revisar columnas)\n\n'
                        _mark(fname, 'FB_OK')
                        continue
                    if is_inventory:
                        ext = os.path.splitext(fname)[1].lower()
                        try:
                            import pandas as _pdi
                            _df_i = _pdi.read_csv(fpath) if ext == '.csv' else _pdi.read_excel(fpath)
                            _df_i = _normalize_cols(_df_i, _INV_COL_MAP)
                            _df_i, _n_i = _guardar_fb_del_hotel(_df_i, 'inventario.xlsx')
                            yield (f'data: ✓ Inventario {fname}: {_n_i} items integrados '
                                   f'· {len(_df_i)} en el hotel\n\n')
                        except Exception:
                            import shutil as _sh5
                            _sh5.copy2(fpath, os.path.join(_ddir(), 'inventario_upload' + ext))
                            yield f'data: ✓ Inventario {fname}: archivo cargado (revisar columnas)\n\n'
                        _mark(fname, 'INV_OK')
                        continue
                    if is_merma:
                        ext = os.path.splitext(fname)[1].lower()
                        try:
                            import pandas as _pdm
                            _df_m = _pdm.read_csv(fpath) if ext == '.csv' else _pdm.read_excel(fpath)
                            _df_m = _normalize_cols(_df_m, _MER_COL_MAP)
                            _df_m, _n_m = _guardar_fb_del_hotel(_df_m, 'mermas.xlsx')
                            yield (f'data: ✓ Mermas {fname}: {_n_m} registros integrados '
                                   f'· {len(_df_m)} en el hotel\n\n')
                        except Exception:
                            import shutil as _sh5m
                            _sh5m.copy2(fpath, os.path.join(_ddir(), 'mermas_upload' + ext))
                            yield f'data: ✓ Mermas {fname}: archivo cargado (revisar columnas)\n\n'
                        _mark(fname, 'INV_OK')
                        continue
                    if is_rooming:
                        # Este camino NO extrae nada: solo reconoce el nombre del
                        # fichero y sigue. Antes marcaba 'ROOMING' igual que el
                        # camino de la IA (que si guarda datos), asi que en el
                        # historial parecian lo mismo.
                        yield (f'data: ℹ {fname}: parece un rooming por el nombre — '
                               f'no se ha extraido ningun dato\n\n')
                        _mark(fname, 'ROOMING_NO_LEIDO')
                        continue

                    # Hojas de calculo que no encajaron por nombre en banco/fb/drr/
                    # inventario. NO pueden ir a pdfplumber (crashea con 'No /Root
                    # object'), pero antes de rendirse se las lee y las clasifica la
                    # IA: es la misma oportunidad que tiene cualquier PDF. Sin esto,
                    # TODA hoja sin keyword reconocible acababa omitida sin abrirse.
                    _ext_check = os.path.splitext(fname)[1].lower()
                    if _ext_check in ('.xlsx', '.xls', '.xlsm', '.csv'):
                        _txt_hoja = _hoja_a_texto(fpath)
                        _reg_hoja = None
                        if _txt_hoja:
                            yield f'data: >> {fname}: sin keyword reconocible — lo lee la IA...\n\n'
                            yield ': ping\n\n'
                            try:
                                from lector_facturas_ap import extraer_con_claude as _eclaude
                                _reg_hoja = _eclaude(_txt_hoja, fname)
                            except Exception as _eia:
                                yield f'data: ⚠ {fname}: la IA no pudo leerlo — {str(_eia)[:60]}\n\n'
                                _reg_hoja = None
                        _tipo_hoja = _tipo_documento(_reg_hoja)
                        if _tipo_hoja and not _reg_hoja.get('_skip'):
                            # FACTURA no viene con tipo_documento en el JSON: se
                            # normaliza aqui para que el enrutado la vea como un
                            # tipo mas y no se caiga al "sin clasificar".
                            _reg_hoja['tipo_documento'] = _tipo_hoja
                            _msg, _marca, _flags = _enrutar_tipo_doc(_reg_hoja, fname, fpath)
                            yield f'data: {_msg}\n\n'
                            _mark(fname, _marca)
                            if _flags.get('has_ar'):
                                has_ar = True
                            if _flags.get('has_ap'):
                                has_ap = True
                                ap_extra += max(0, int(_flags.get('ap_n', 1)) - 1)
                            if _flags.get('albaran'):
                                has_albaran = True
                        else:
                            yield f'data: ⚠ {fname}: hoja de cálculo sin clasificar — revisar manualmente\n\n'
                            _mark(fname, 'SKIP')
                        continue

                    if is_ar:
                        # AR: usar subprocess (lector_ota.py tiene su propia lógica)
                        cmd = ['python3', 'lector_ota.py', '--file', fpath]
                        # env=_env_tenant(): sin esto el subproceso escribia en
                        # el arbol del tenant `default` fuera quien fuera el
                        # usuario, y ahora ademas necesita saber el hotel.
                        r = _sp.run(cmd, capture_output=True, text=True, cwd=BASE_DIR,
                                    timeout=60, env=_env_tenant())
                        # lector_ota distingue: 0=OK, 2=guardado incompleto,
                        # 3=leido pero sin datos OTA (no guarda nada). Antes
                        # cualquier returncode 0 se cantaba como "✓ OK" aunque
                        # la fila fuera entera NO_ENCONTRADO.
                        _faltan = ''
                        _hotel_doc = ''
                        for _ln in (r.stdout or '').splitlines():
                            if _ln.startswith('FALTAN:'):
                                _faltan = _ln.split(':', 1)[1].strip()
                            elif _ln.startswith('HOTEL_DOC:'):
                                _hotel_doc = _ln.split(':', 1)[1].strip()
                        # El aviso de "te has equivocado de hotel" solo existia
                        # en la rama de escaneo, y este es el camino que se usa
                        # de verdad: subes la liquidacion de un hotel teniendo
                        # elegido otro y se guardaba callando. No corrige nada
                        # —el hotel lo decide la sesion, no el papel—, avisa.
                        _aviso_hotel = _aviso_otro_hotel(_hotel_doc)
                        if r.returncode == 0:
                            yield f'data: ✓ AR {fname}: OK\n\n'
                            _mark(fname, 'AR_OK')
                            has_ar = True
                        elif r.returncode == 2:
                            _det = f' (faltan: {_faltan})' if _faltan else ''
                            yield f'data: ⚠ AR {fname}: guardado con campos incompletos{_det} — revisar manualmente\n\n'
                            _mark(fname, 'AR_PARCIAL')
                            has_ar = True
                        elif r.returncode == 3:
                            yield f'data: ⚠ {fname}: no se pudo extraer ningún dato de factura OTA — revisar manualmente\n\n'
                            _mark(fname, 'SKIP')
                        else:
                            msg = r.stderr[:80] or r.stdout[:80] or 'error'
                            yield f'data: ✗ AR {fname}: {msg}\n\n'
                            _mark(fname, f'ERR:{msg[:30]}')
                        # DESPUES de toda la cadena, no en medio. Metido entre
                        # el `elif 2` y el `elif 3` partia la cadena en dos: el
                        # `elif 3` y el `else` pasaban a colgar de este `if`,
                        # asi que un fichero correcto SIN aviso caia en el
                        # `else` y se cantaba como error, con un "✗" y un
                        # trozo de la salida del lector como mensaje.
                        #
                        # Va en su propia linea y empezando por ⚠ para que el
                        # log lo pinte de aviso: pegado al final del "✓ ...:
                        # OK" salia en verde, al final de una linea larga, en
                        # un panel que va scrolleando — o sea, invisible.
                        if _aviso_hotel:
                            yield f'data: ⚠ {fname}: {_aviso_hotel}\n\n'
                    else:
                        # AP: import directo (más rápido, sin cargar Python de nuevo)
                        try:
                            from lector_facturas_ap import procesar_factura_ap, cargar_proveedores, guardar_excel, SALIDA_DIR as _AP_DIR
                            _provs = cargar_proveedores()
                            reg = procesar_factura_ap(fpath, _provs)
                            if reg is None or (isinstance(reg, dict) and reg.get('_skip')):
                                motivo = reg.get('_motivo', 'documento no procesable') if isinstance(reg, dict) else 'documento no procesable'
                                yield f'data: ⚠ {fname}: {motivo}\n\n'
                                _mark(fname, 'SKIP')
                            elif isinstance(reg, dict) and reg.get('tipo_documento'):
                                # Claude clasificó el documento como otro tipo — enrutar
                                # (árbol de enrutado extraído a _enrutar_tipo_doc)
                                _msg, _marca, _flags = _enrutar_tipo_doc(reg, fname, fpath)
                                yield f'data: {_msg}\n\n'
                                _mark(fname, _marca)
                                # TODAS las banderas que devuelve el enrutador,
                                # no solo has_ar.
                                #
                                # Este es el camino de los PDF y las fotos, o sea
                                # el mas usado, y aqui se perdian `albaran`,
                                # `has_ap` y `ap_n`. Consecuencia medida: un
                                # albaran en PDF se guardaba perfectamente y
                                # despues el cruce NO se relanzaba, porque
                                # has_albaran seguia en False — asi que una
                                # entrega nueva no volvia a evaluar la factura
                                # que ayer no cuadraba, que es justo lo que la
                                # fase 3b·2 existe para conseguir.
                                #
                                # El camino de hojas de calculo (~linea 2075) SI
                                # las leia, y de ahi venia lo desconcertante: el
                                # mismo albaran funcionaba en Excel y no en PDF.
                                # `ap_n` tambien se perdia, y por eso el resumen
                                # del lote contaba menos facturas de las que
                                # entraron cuando un documento traia varias.
                                if _flags.get('has_ar'):
                                    has_ar = True
                                if _flags.get('has_ap'):
                                    has_ap = True
                                    ap_extra += max(0, int(_flags.get('ap_n', 1)) - 1)
                                if _flags.get('albaran'):
                                    has_albaran = True
                            elif reg and not reg.get('error') and not _ap_tiene_datos(reg):
                                # La extraccion no saco NI proveedor NI numero NI
                                # importe: guardar la fila solo ensucia el Excel y
                                # cantar "✓ AP" es mentir sobre lo que ha pasado.
                                yield f'data: ⚠ {fname}: no se pudo extraer ningún dato de la factura — revisar manualmente\n\n'
                                _mark(fname, 'SKIP')
                            elif reg and not reg.get('error'):
                                # Mismo guardado que la hoja de calculo y que la
                                # foto: un solo sitio. _filas_limpias despliega
                                # las varias facturas si el documento traia mas
                                # de una, en vez de quedarse con la primera.
                                from lector_facturas_ap import _filas_limpias as _fl_ap
                                _filas = _fl_ap(reg)
                                _n_ap = _guardar_factura_ap(_filas)
                                yield f'data: ✓ AP {fname}: {_resumen_factura_ap(_filas)}\n\n'
                                _mark(fname, 'AP_OK')
                                has_ap = True
                                ap_extra += max(0, _n_ap - 1)
                                
                                # 3-WAY MATCHING: buscar BEO/contrato del mismo evento/cliente
                                try:
                                    ref_path = os.path.join(_ddir(), 'eventos_referencia.json')
                                    if os.path.exists(ref_path):
                                        refs = json.load(open(ref_path))
                                        proveedor = (reg.get('nombre_proveedor') or '').lower()
                                        concepto = (reg.get('descripcion_concepto') or '').lower()
                                        total_factura = reg.get('total_factura', 0)
                                        
                                        match_found = None
                                        for ref in refs:
                                            cliente = (ref.get('cliente') or '').lower()
                                            evento = (ref.get('evento') or '').lower()
                                            # Match por cliente o evento mencionado en la factura
                                            if cliente and (cliente in proveedor or cliente in concepto or proveedor in cliente):
                                                match_found = ref
                                                break
                                            if evento and (evento in concepto or any(w in concepto for w in evento.split()[:3] if len(w)>3)):
                                                match_found = ref
                                                break
                                        
                                        if match_found:
                                            # OJO: no llamar 'docs' a esto — 'docs' es la
                                            # lista de archivos del bucle y se pisaba, dejando
                                            # el contador de progreso en [3/2].
                                            docs_ev = match_found.get('documentos', {})
                                            docs_list = list(docs_ev.keys())
                                            # Comparar totales
                                            discrepancias = []
                                            for doc_tipo, doc_data in docs_ev.items():
                                                doc_total = doc_data.get('total', 0)
                                                if doc_total and isinstance(total_factura, (int, float)) and total_factura > 0:
                                                    diff = abs(total_factura - doc_total)
                                                    diff_pct = diff / doc_total * 100 if doc_total > 0 else 0
                                                    if diff_pct > 5:  # Más de 5% de diferencia
                                                        discrepancias.append(f'{doc_tipo} dice €{doc_total:,.2f} vs factura €{total_factura:,.2f} ({diff_pct:.0f}% diff)')
                                            
                                            evento_nombre = match_found.get('evento', '—')
                                            if discrepancias:
                                                yield f'data: ⚠ MATCHING {evento_nombre}: {" | ".join(discrepancias)}\n\n'
                                            else:
                                                yield f'data: ✓ MATCHING {evento_nombre}: factura cuadra con {", ".join(docs_list)}\n\n'
                                except Exception:
                                    pass  # Matching es best-effort
                            else:
                                err = reg.get('error','error desconocido') if reg else 'sin resultado'
                                yield f'data: ✗ AP {fname}: {err[:80]}\n\n'
                                _mark(fname, f'ERR:{err[:30]}')
                        except Exception as _eap:
                            yield f'data: ✗ AP {fname}: {str(_eap)[:80]}\n\n'
                            _mark(fname, f'ERR:{str(_eap)[:30]}')

                except subprocess.TimeoutExpired:
                    yield f'data: ✗ {fname}: TIMEOUT (60s)\n\n'
                    _mark(fname, 'ERR:TIMEOUT')
                except Exception as e2:
                    yield f'data: ✗ {fname}: {str(e2)[:80]}\n\n'
                    _mark(fname, f'CRASH:{str(e2)[:30]}')

            # ── Resultado del contrato de grupo (si las fotos lo eran) -> AR Real ──
            if _contrato_res is not None:
                if _contrato_res.get('ok'):
                    _eur = f"{_contrato_res.get('total_receivable', 0):,.2f}"
                    _com = f"{_contrato_res.get('comision_total', 0):,.2f}"
                    _di = ' · certificado DI pendiente' if _contrato_res.get('requiere_certificado_di') else ''
                    yield f"data: ✓ Contrato {_contrato_res.get('contrato','')} · {_contrato_res.get('cliente','')} · {_eur}€ (comisión {_com}€){_di}\n\n"
                    # marca propia: este contrato SI acaba en AR Real, a diferencia
                    # del CONTRATO de evento, que solo se guarda para el cruce.
                    for _a in imgs: _mark(_a, 'AR_REAL_OK')
                    has_ar_real = True
                elif _contrato_res.get('needs_review'):
                    yield f"data: ⚠ Contrato de grupo: no se pudo leer — {str(_contrato_res.get('error',''))[:80]}\n\n"
                    for _a in imgs: _mark(_a, 'SKIP')
                else:
                    yield f"data: ✗ Contrato de grupo: {str(_contrato_res.get('error','error'))[:80]}\n\n"
                    for _a in imgs: _mark(_a, 'ERR:CONTRATO')

            # ── Lo que queda por cerrar ──
            # El cruce y el asignador ya NO van aqui. Iban al final de CADA
            # lote y se quedaban a medias: el frontend parte los archivos en
            # lotes de 4 y cierra el EventSource a los 60 s, asi que los dos
            # subprocesos del cierre no llegaban a arrancar en el ultimo lote.
            # Ahora el lote solo DICE lo que hay pendiente, y el frontend llama
            # UNA vez a /api/cerrar_pipeline_stream cuando ha acabado todo.
            _pend = []
            if has_ap: _pend.append('ap')
            if has_albaran: _pend.append('albaran')
            if has_ar: _pend.append('ar')
            if _pend:
                yield f'data: CIERRE_PENDIENTE:{",".join(_pend)}\n\n'

            # ── Resumen de procesado ──
            ap_n = sum(1 for v in lote.values() if v.get('resultado') == 'AP_OK') + ap_extra
            ar_n = sum(1 for v in lote.values() if v.get('resultado') == 'AR_OK')
            drr_n = sum(1 for v in lote.values() if v.get('resultado') == 'DRR_OK')
            bank_n = sum(1 for v in lote.values() if v.get('resultado') == 'BANK_OK')
            fb_n = sum(1 for v in lote.values() if v.get('resultado') == 'FB_OK')
            inv_n = sum(1 for v in lote.values() if v.get('resultado') == 'INV_OK')
            rooming_n = sum(1 for v in lote.values() if v.get('resultado') == 'ROOMING')
            alb_n = sum(1 for v in lote.values() if v.get('resultado') == 'ALBARAN_OK')
            po_n = sum(1 for v in lote.values() if v.get('resultado') == 'PO_OK')
            skip_n = sum(1 for v in lote.values() if 'SKIP' in str(v.get('resultado','')))
            err_n = sum(1 for v in lote.values() if 'ERR' in str(v.get('resultado','')) or 'CRASH' in str(v.get('resultado','')))
            parts = []
            if ap_n: parts.append(f'{ap_n} facturas AP')
            if ar_n: parts.append(f'{ar_n} informes OTA')
            parc_n = sum(1 for v in lote.values() if v.get('resultado') == 'AR_PARCIAL')
            if parc_n: parts.append(f'{parc_n} OTA incompletos')
            pact_n = sum(1 for v in lote.values() if v.get('resultado') == 'CONTRATO_OTA_OK')
            if pact_n: parts.append(f'{pact_n} contratos OTA')
            if drr_n: parts.append(f'{drr_n} DRR')
            if bank_n: parts.append(f'{bank_n} banco')
            if fb_n: parts.append(f'{fb_n} F&B')
            if inv_n: parts.append(f'{inv_n} inventario/mermas')
            if alb_n: parts.append(f'{alb_n} albaranes')
            if po_n: parts.append(f'{po_n} órdenes de compra')
            rooming_nl_n = sum(1 for v in lote.values() if v.get('resultado') == 'ROOMING_NO_LEIDO')
            if rooming_n: parts.append(f'{rooming_n} rooming')
            if rooming_nl_n: parts.append(f'{rooming_nl_n} rooming sin leer')
            beo_n = sum(1 for v in lote.values() if v.get('resultado') in ('BEO_OK','TM_OK','CONTRATO_OK'))
            if beo_n: parts.append(f'{beo_n} docs evento')
            resumen = ' · '.join(parts) if parts else 'sin documentos procesables'
            yield f'data: \n\n'
            yield f'data: ✅ {resumen}'
            if skip_n: yield f' · {skip_n} omitidos'
            if err_n: yield f' · {err_n} errores'
            yield f'\n\n'
            # Indicar qué tabs consultar
            tabs_updated = []
            if ap_n: tabs_updated.append('AP — Proveedores')
            if ar_n: tabs_updated.append('AR — OTAs')
            if bank_n: tabs_updated.append('Banco')
            if fb_n or inv_n: tabs_updated.append('F&B Cost')
            if has_ar_real: tabs_updated.append('AR Real')
            if tabs_updated:
                yield f'data: 📍 Consulta: {", ".join(tabs_updated)}\n\n'
            # Rooming y documentos de evento se guardan, pero NO hay pantalla
            # donde consultarlos. Antes salian en "Consulta:" mandando a
            # pestañas que no existen; ahora se dice lo que hay.
            if alb_n:
                yield f'data: 📋 {alb_n} albarán(es) guardado(s) — en AP → Albaranes, con sus líneas y su factura\n\n'
            _sin_pantalla = rooming_n + beo_n + po_n
            if _sin_pantalla:
                yield (f'data: ℹ {_sin_pantalla} documento(s) guardado(s) para cruces internos — '
                       f'todavia no hay pantalla donde consultarlos\n\n')
            yield 'data: PIPELINE_COMPLETO\n\n'
        except Exception as e:
            yield f'data: ERROR: {str(e)[:200]}\n\n'
            yield 'data: PIPELINE_CON_ERRORES\n\n'
        finally:
            _pipeline_running = False

    return Response(stream_with_context(generar()), mimetype='text/event-stream',
                    headers={'Cache-Control':'no-cache','X-Accel-Buffering':'no','Connection':'keep-alive'})



# ══ El paso de cierre del pipeline ═══════════════════════════════════════
# Vivia al final de CADA lote, y de ahi salian dos agujeros:
#
#   1. el frontend parte los archivos en lotes de 4 y cierra el EventSource a
#      los 60 s (`setTimeout`). El cruce y el asignador son dos subprocesos con
#      180 s de margen cada uno: en el ultimo lote la conexion moria antes de
#      que arrancaran, y el generador se quedaba cortado ahi.
#   2. la FOTO no pasa por el lote —va a `/api/scan_documento`, un POST por
#      imagen— y ahi no habia paso de cierre NINGUNO. Una foto de un albaran se
#      guardaba y no relanzaba el cruce; y peor, una foto de una FACTURA se
#      guardaba y no llegaba nunca a Aprobaciones AP, porque nadie le asignaba
#      cuenta ni asiento. Se veia el ✓ verde y el documento no existia para
#      quien tiene que aprobarlo.
#
# Ahora es UN paso con su propio endpoint, que el frontend llama UNA vez
# cuando ha terminado todo: vengan los documentos por lote, por foto o
# mezclados. Y al ser su propia conexion, tiene su propio reloj.
_PASOS_CIERRE = ('ap', 'albaran', 'ar')


def _generar_cierre(pasos):
    """Las lineas SSE del cierre. `pasos` es un conjunto de `_PASOS_CIERRE`.

    NADA de Oracle aqui, igual que antes: esto genera informes, no contabiliza.
    """
    global _pipeline_running
    with _pipeline_lock:
        if _pipeline_running:
            yield 'data: ℹ Ya hay un proceso activo — espera\n\n'
            yield 'data: CIERRE_CON_ERRORES\n\n'
            return
        _pipeline_running = True
    try:
        if not pasos:
            yield 'data: CIERRE_COMPLETO\n\n'
            return
        if 'ar' in pasos:
            yield 'data: >> Verificando comisiones OTA...\n\n'
            try:
                import subprocess as _sp2
                _sp2.run(['python3','verificador_comisiones.py'], cwd=BASE_DIR, timeout=30, capture_output=True, env=_env_tenant())
                _sp2.run(['python3','detector_doble_imposicion.py'], cwd=BASE_DIR, timeout=30, capture_output=True, env=_env_tenant())
                yield 'data: ✓ Verificación completada\n\n'
            except: pass

        if 'ap' in pasos or 'albaran' in pasos:
            # ORDEN: el cruce va ANTES que el asignador porque el asignador
            # une su informe al generar facturas_contabilizadas. Al reves,
            # el estado llegaria al panel un lote tarde.
            # Tambien corre si SOLO han entrado albaranes: una entrega nueva
            # puede completar una factura que ayer no cuadraba.
            yield 'data: >> Cruzando facturas con albaranes...\n\n'
            yield ': ping\n\n'
            try:
                import subprocess as _sp4
                _r_alb = _sp4.run(['python3', 'matching_ap_albaran.py'], cwd=BASE_DIR,
                                  timeout=180, capture_output=True, text=True,
                                  env=_env_tenant())
                # el modulo lo dice el mismo: contar lineas de su consola
                # se comia tambien la fila del resumen y salian 2 donde habia 1
                _n_fac = _n_alb = 0
                for _ln in (_r_alb.stdout or '').splitlines():
                    if _ln.startswith('INCIDENCIAS:'):
                        try:
                            _a, _b = _ln.split(':', 1)[1].strip().split('|')
                            _n_fac, _n_alb = int(_a), int(_b)
                        except Exception:
                            pass
                if _r_alb.returncode == 0:
                    _partes = []
                    if _n_fac:
                        _partes.append(f'{_n_fac} factura(s) sin entrega que las respalde '
                                       f'o con diferencia de importe')
                    if _n_alb:
                        _partes.append(f'{_n_alb} albarán(es) entregado(s) sin facturar')
                    if _partes:
                        yield f'data: ⚠ Cruce con albaranes: {" · ".join(_partes)}\n\n'
                    else:
                        yield 'data: ✓ Cruce con albaranes: sin incidencias\n\n'
                else:
                    _e_alb = (_r_alb.stderr or _r_alb.stdout or 'error').strip().splitlines()
                    _e_alb = (_e_alb[-1] if _e_alb else 'error')[:90]
                    yield f'data: ⚠ No se ha podido cruzar con los albaranes — {_e_alb}\n\n'
            except Exception as _ea2:
                yield f'data: ⚠ No se ha podido cruzar con los albaranes — {str(_ea2)[:80]}\n\n'

        if 'ap' in pasos or 'albaran' in pasos:
            # Las facturas ya estan guardadas; ahora se les pone cuenta y
            # asiento. Es lo que alimenta /aprobaciones-ap, que hasta ahora
            # solo se llenaba pulsando el boton del pipeline a mano.
            #
            # NO se lanzan matching_ap_otras ni matching_ap_fb: revientan
            # con KeyError 'proveedor' y necesitan el concepto de PO, que
            # va en su fase. Y NADA de Oracle: aqui solo se genera el
            # informe, no se contabiliza nada en el libro mayor.
            yield 'data: >> Asignando cuentas contables...\n\n'
            yield ': ping\n\n'
            try:
                import subprocess as _sp3
                _r_asig = _sp3.run(['python3', 'asignador_cuentas.py'], cwd=BASE_DIR,
                                   timeout=180, capture_output=True, text=True,
                                   env=_env_tenant())
                if _r_asig.returncode == 0:
                    yield 'data: ✓ Cuentas y asientos asignados — pendientes de aprobar en Aprobaciones AP\n\n'
                else:
                    # Honestidad: las facturas SI estan guardadas. Lo que ha
                    # fallado es la contabilizacion, y hay que decirlo.
                    _e_asig = (_r_asig.stderr or _r_asig.stdout or 'error').strip().splitlines()
                    _e_asig = (_e_asig[-1] if _e_asig else 'error')[:90]
                    yield f'data: ⚠ Facturas guardadas, pero no se han podido asignar las cuentas — {_e_asig}\n\n'
            except Exception as _ea:
                yield f'data: ⚠ Facturas guardadas, pero no se han podido asignar las cuentas — {str(_ea)[:80]}\n\n'

        yield 'data: CIERRE_COMPLETO\n\n'
    except Exception as _ec:
        yield f'data: ⚠ El cierre del pipeline ha fallado — {str(_ec)[:120]}\n\n'
        yield 'data: CIERRE_CON_ERRORES\n\n'
    finally:
        _pipeline_running = False


@app.route('/api/cerrar_pipeline_stream')
@login_required
def api_cerrar_pipeline_stream():
    """El cruce y el asignador, una vez, cuando ya esta todo guardado.

    Se llama con `?pasos=ap,albaran,ar` — lo que el frontend haya visto entrar.
    Un paso que no se pide no corre: no tiene sentido cruzar albaranes cuando
    solo han entrado extractos de banco.
    """
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
    pedidos = (request.args.get('pasos') or '').split(',')
    pasos = {p.strip().lower() for p in pedidos if p.strip().lower() in _PASOS_CIERRE}
    return Response(stream_with_context(_generar_cierre(pasos)),
                    mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no',
                             'Connection': 'keep-alive'})


@app.route('/demo')
def demo_view():
    """Demo pública sin login — entra como admin directamente."""
    from flask_login import login_user
    from auth import login as auth_login
    user = auth_login('admin', 'admin123')
    if user:
        login_user(user)
    return redirect('/')



@app.route('/api/scan_documento', methods=['POST'])
@login_required
def api_scan_documento():
    """Procesa una imagen de documento físico con Claude Vision."""
    import base64
    # La foto del movil es otra puerta a lo mismo: tambien estampa hotel.
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
    if 'image' not in request.files:
        return jsonify({'ok': False, 'error': 'No se recibió imagen'}), 400
    
    img_file = request.files['image']
    img_data = img_file.read()
    img_b64 = base64.b64encode(img_data).decode()
    
    # Determinar media type
    fname = img_file.filename or 'scan.jpg'
    ext = fname.rsplit('.', 1)[-1].lower() if '.' in fname else 'jpg'
    media_map = {'jpg':'image/jpeg','jpeg':'image/jpeg','png':'image/png','heic':'image/heic','webp':'image/webp'}
    media_type = media_map.get(ext, 'image/jpeg')
    
    try:
        import anthropic
        client = anthropic.Anthropic()
        
        # Prompt COMPARTIDO con el lector de documentos: una sola fuente de
        # verdad. Antes habia aqui una copia con solo 4 de los 12 esquemas.
        from lector_facturas_ap import prompt_foto as _prompt_foto
        prompt = _prompt_foto(fname)

        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{"role":"user","content":[
                {"type":"image","source":{"type":"base64","media_type":media_type,"data":img_b64}},
                {"type":"text","text":prompt}
            ]}]
        )
        raw = resp.content[0].text.strip()
        
        # Extraer JSON
        import re as _re
        raw = _re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
        first = raw.find('{')
        last = raw.rfind('}')
        if first >= 0 and last > first:
            raw = raw[first:last+1]
        
        datos = json.loads(raw)
        tipo = _tipo_documento(datos) or 'OTRO'
        
        # Guardar en historial
        from datetime import datetime as _dt2
        log = _load_proc_log()
        log[fname] = _entrada_proc(f'{tipo}_OK')                # M8
        _save_proc_log(log)
        
        # Integrar datos según tipo — MISMO flujo que Procesar Archivos
        items_count = 0
        mensaje = ''
        # None = "este tipo no dice nada al respecto" -> el frontend mantiene su
        # ✓ de siempre. Solo el albaran la pone hoy; los demas tipos podrian
        # adoptarla despues sin mover nada de lo que ya funciona.
        guardado = None
        
        if tipo == 'FACTURA' and datos.get('es_factura'):
            # Mismo normalizador y mismo guardado que el PDF y la hoja de
            # calculo. Esto era una TERCERA copia: no autocalculaba base ni IVA
            # cuando faltaban, y guardaba con guardar_excel(), que SOBRESCRIBE
            # el fichero del dia y se llevaba por delante las facturas que el
            # lote hubiera guardado antes.
            from lector_facturas_ap import facturas_de_respuesta, cargar_proveedores
            _filas = [f for f in facturas_de_respuesta(datos, fname, cargar_proveedores())
                      if _ap_tiene_datos(f)]
            if _filas:
                _guardar_factura_ap(_filas)
                # items_count solo cuenta cuando hay MAS de una: con una sola,
                # la pantalla decia 0 y no se toca lo que hoy funciona.
                items_count = len(_filas) if len(_filas) > 1 else 0
                mensaje = _resumen_factura_ap(_filas)
                guardado = True
            else:
                mensaje = 'factura detectada, pero no se pudo extraer ningún dato — revisar manualmente'
                guardado = False
            
        elif tipo in ('BEO','TM','CONTRATO'):
            ref_path = os.path.join(_ddir(), 'eventos_referencia.json')
            refs = json.load(open(ref_path)) if os.path.exists(ref_path) else []
            evento = datos.get('evento', fname)
            evento_key = evento.lower().strip()[:50]
            found = False
            for ref in refs:
                if ref.get('evento_key') == evento_key:
                    ref['documentos'][tipo] = {'archivo': fname, 'total': datos.get('total_estimado', datos.get('importe_total',0)),
                        'items': datos.get('items',[]), 'fecha': date.today().isoformat()}
                    found = True
                    break
            if not found:
                refs.append({'evento': evento, 'evento_key': evento_key,
                    'cliente': datos.get('cliente',''), 'documentos': {
                        tipo: {'archivo': fname, 'total': datos.get('total_estimado', datos.get('importe_total',0)),
                            'items': datos.get('items',[]), 'fecha': date.today().isoformat()}}})
            json.dump(refs, open(ref_path,'w'), indent=2, ensure_ascii=False)
            items_count = len(datos.get('items',[]))
            n_docs = len([r for r in refs if r.get('evento_key')==evento_key][0].get('documentos',{}))
            mensaje = f'{evento} ({datos.get("cliente","")}) — {items_count} items · {n_docs} docs del evento'
            # el documento se registra igual, pero sin items ni nombre de evento
            # no se ha leido nada aprovechable: no puede salir en verde
            guardado = bool(items_count) or evento_key != fname.lower().strip()[:50]
            
        elif tipo == 'EXTRACTO_BANCO':
            movimientos = datos.get('movimientos', [])
            if movimientos:
                _df_movs = _normalize_cols(pd.DataFrame(movimientos), _BANK_COL_MAP)
                _estampar_hotel_banco(_df_movs)   # modo por_hotel: marca las filas nuevas
                banco_path = os.path.join(_ddir(), 'extracto_banco.xlsx')
                if os.path.exists(banco_path):
                    _df_old = pd.read_excel(banco_path)
                    _df_movs = pd.concat([_df_old, _df_movs], ignore_index=True)
                _df_movs.to_excel(banco_path, index=False)
                items_count = len(movimientos)
                mensaje = f'{items_count} movimientos bancarios integrados'
                guardado = True
            else:
                mensaje = 'Extracto detectado (sin movimientos extraíbles)'
                guardado = False
                
        elif tipo == 'VENTAS_POS':
            platos = datos.get('platos', [])
            total = datos.get('total_ventas', 0)
            if platos:
                _df_v = _normalize_cols(pd.DataFrame(platos), _VEN_COL_MAP)
                fecha = datos.get('fecha', date.today().isoformat())
                if 'fecha' not in _df_v.columns:
                    _df_v['fecha'] = fecha
                _df_v, _ = _guardar_fb_del_hotel(_df_v, 'ventas_fb_diarias.xlsx')
                items_count = len(platos)
                mensaje = f'{items_count} platos, €{total} integrados'
                guardado = True
            else:
                mensaje = f'Ventas detectadas — €{total}'
                guardado = False
                
        elif tipo == 'INVENTARIO':
            inv_items = datos.get('items', datos.get('productos', []))
            if inv_items:
                _df_inv = _normalize_cols(pd.DataFrame(inv_items), _INV_COL_MAP)
                _df_inv, _ = _guardar_fb_del_hotel(_df_inv, 'inventario.xlsx')
                items_count = len(inv_items)
                nombres = [str(i.get('ingrediente', '?'))[:15] for i in inv_items[:4]]
                mensaje = f'{items_count} productos ({", ".join(nombres)}...) integrados'
                guardado = True
            else:
                mensaje = 'Inventario detectado (sin items extraíbles)'
                guardado = False
                
        elif tipo == 'MERMAS':
            merma_items = datos.get('items', datos.get('mermas', []))
            if merma_items:
                _df_m = _normalize_cols(pd.DataFrame(merma_items), _MER_COL_MAP)
                _df_m, _ = _guardar_fb_del_hotel(_df_m, 'mermas.xlsx')
                items_count = len(merma_items)
                total_merma = sum(float(m.get('coste_merma', m.get('coste', 0)) or 0) for m in merma_items)
                mensaje = f'{items_count} registros — €{total_merma:.2f} integrados'
                guardado = True
            else:
                mensaje = 'Mermas detectadas (sin registros extraíbles)'
                guardado = False
                
        elif tipo == 'ROOMING':
            grupo = datos.get('grupo', '—')
            habs = datos.get('num_habitaciones', '?')
            checkin = datos.get('checkin', '')
            checkout = datos.get('checkout', '')
            rooming_path = os.path.join(_ddir(), 'rooming_grupos.json')
            rooming_data = json.load(open(rooming_path)) if os.path.exists(rooming_path) else []
            rooming_data.append({'archivo': fname, 'grupo': grupo, 'habitaciones': habs,
                'checkin': checkin, 'checkout': checkout, 'fecha_procesado': date.today().isoformat()})
            json.dump(rooming_data, open(rooming_path, 'w'), indent=2, ensure_ascii=False)
            mensaje = f'{grupo} — {habs} hab. ({checkin}→{checkout})'
            # el json se escribe igual, pero un rooming del que no se ha sacado
            # ni el grupo ni el numero de habitaciones no es un rooming leido
            guardado = bool(str(grupo).strip() not in ('', '—')) or str(habs).strip() not in ('', '?')
            
        elif tipo in ('COMISIONES_OTA', 'CONTRATO_OTA', 'ALBARAN', 'ORDEN_COMPRA'):
            # Reutiliza el enrutado del pipeline: asi una foto de una factura de
            # comisiones, de un contrato con tarifas pactadas o de un ALBARAN
            # aterriza en el mismo sitio que su equivalente en PDF (y GUARDA,
            # que antes esta rama solo componia un mensaje).
            # El albaran acaba en _guardar_albaran por dentro de
            # _enrutar_tipo_doc: ni una linea de guardado duplicada aqui. Un
            # albaran en papel fotografiado con el movil es el caso MAS normal
            # de todos, y hasta ahora caia al else generico y no se guardaba.
            _m, _mk, _fl = _enrutar_tipo_doc(datos, fname)
            log[fname] = _entrada_proc(_mk)                     # M8
            _save_proc_log(log)
            mensaje = _m.split(': ', 1)[-1] if ': ' in _m else _m
            items_count = len(datos.get('facturas', datos.get('tarifas',
                                        datos.get('lineas', []))))
            # TODOS los tipos que pasan por el enrutador comun heredan su
            # veredicto: si volvio SKIP no se guardo nada, y la rama de OTA
            # cantaba verde igual sobre un SKIP.
            if True:
                # Misma honestidad que en el camino de archivos: un albaran del
                # que no se ha podido leer una sola linea NO puede cantar
                # "✓ Albaran". El frontend solo miraba 'ok', asi que pintaba
                # verde igualmente; ahora se le dice si se guardo algo.
                # Igual con el pedido: una foto de una orden de compra sin
                # proveedor ni importe no puede salir en verde.
                guardado = (_mk != 'SKIP')
            
        else:
            # tipo sin rama propia (OTRO y los que no tienen destino todavia):
            # no se guarda nada en ningun sitio, asi que no puede salir verde
            desc = datos.get('descripcion', tipo)
            mensaje = desc
            guardado = False
        
        # Lo que esta foto deja pendiente de cerrar. Es la mitad del bug que
        # no se arreglaba con banderas: este endpoint guarda el documento y
        # ahi se acababa todo. Una foto de una factura no llegaba a
        # Aprobaciones AP, y una foto de un albaran no relanzaba el cruce.
        # El frontend junta estos pasos de todas las fotos y llama UNA vez a
        # /api/cerrar_pipeline_stream.
        _cierre = []
        if guardado is not False:
            if tipo == 'FACTURA':
                _cierre.append('ap')
            elif tipo == 'ALBARAN':
                _cierre.append('albaran')
            elif tipo in ('COMISIONES_OTA', 'CONTRATO_OTA'):
                _cierre.append('ar')
        return jsonify({'ok': True, 'tipo': tipo, 'mensaje': mensaje,
                        'items': str(items_count) if items_count else None,
                        'guardado': guardado, 'cierre': _cierre, 'datos': datos})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500


@app.route('/api/eventos_referencia')
@login_required
def api_eventos_referencia():
    """Devuelve los eventos con sus documentos de referencia (BEO, TM, contrato)."""
    ref_path = os.path.join(_ddir(), 'eventos_referencia.json')
    if not os.path.exists(ref_path):
        return jsonify({'ok': True, 'eventos': []})
    try:
        refs = json.load(open(ref_path))
        return jsonify({'ok': True, 'eventos': refs})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/test_clasificador')
@login_required
def api_test_clasificador():
    """Copia los 8 PDFs de test al directorio de upload y devuelve la lista."""
    import shutil as _sht
    test_dir = os.path.join(BASE_DIR, 'facturas-entrada', 'test_clasificador')
    upload_dir = os.path.join(BASE_DIR, 'facturas-entrada')
    archivos = []
    for pdf in sorted(glob.glob(os.path.join(test_dir, '*.pdf'))):
        fname = os.path.basename(pdf)
        dest = os.path.join(upload_dir, fname)
        _sht.copy2(pdf, dest)
        archivos.append(fname)
    return jsonify({'ok': True, 'archivos': archivos})

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
        "reports":   len(_glob.glob(os.path.join(_rdir(), "*.xlsx"))),
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
    """Aprueba en lote las facturas AP cuyo cruce cuadra.

    BOMBA 1 — ANTES buscaba una columna `aprobacion` en los informes
    `matching_*.xlsx` que NINGUN modulo escribe: `aprobadas` valia siempre 0
    y el navegador enseñaba "N aprobadas" habiendo aprobado cero.

    AHORA escribe en `aprobaciones_ap.xlsx` por `registrar_acciones`, el
    mismo registro (y la misma forma de fila) que la pantalla "Facturas por
    aprobar" — que es lo que lee Oracle para decidir que contabiliza. El
    navegador PROPONE claves; aqui se decide: solo se aprueba lo que (1) esta
    en el panel del hotel activo, (2) tiene un estado de cruce correcto
    (`_ESTADOS_OK`, la misma lista que la pantalla de aprobar) y (3) no tiene
    ya una decision. Se devuelven las cifras reales, sin inventar nada.
    """
    from app_aprobacion_ap import (_ESTADOS_OK, clave_factura, _acciones_por_clave,
                                   registrar_acciones, decidir_accion, _firma1_por_clave,
                                   umbral_doble_firma)
    data = request.get_json(force=True, silent=True) or {}
    pedidas = {safe_str(x) for x in (data.get("facturas") or []) if safe_str(x)}
    if not pedidas:
        return jsonify({"ok": False, "error": "No se especificaron facturas"}), 400
    try:
        df = cargar_datos_ap()                 # lo que el usuario esta viendo
        decididas = _acciones_por_clave()      # la ultima decision de cada clave
        try:
            import censo_hoteles as _ch
            _hid = _ch.activo() or ""
        except Exception:
            _hid = ""
        from datetime import datetime as _dt
        # `session["username"]` no existe (nadie lo escribe): el que aprueba es
        # el usuario logueado de flask-login, igual que en el resto de la app.
        usuario = getattr(current_user, "username", None) or "sistema"
        ahora = _dt.now().strftime("%d/%m/%Y %H:%M:%S")
        firma1 = _firma1_por_clave()           # doble firma: quien puso la primera
        filas, ya_decididas, no_cuadran, vistas = [], 0, 0, set()
        primera_firma, esperan_segunda = 0, 0
        if not df.empty:
            for _, r in df.iterrows():
                clave = clave_factura(r)
                if clave not in pedidas or clave in vistas:
                    continue
                vistas.add(clave)
                est = (safe_str(r.get("estado_matching")) or safe_str(r.get("estado"))).upper()
                if est not in _ESTADOS_OK:
                    no_cuadran += 1
                    continue
                if decididas.get(clave) in ("APROBADA", "RECHAZADA"):
                    ya_decididas += 1
                    continue
                # Misma regla que "Facturas por aprobar": por encima del umbral
                # la primera firma es FIRMA_1 (Oracle no la ve) y la segunda,
                # de OTRA persona, es la APROBADA.
                accion_real, info = decidir_accion(clave, r.get("total_factura"), "APROBADA",
                                                   usuario, firma1)
                if accion_real is None:
                    esperan_segunda += 1
                    continue
                if accion_real == "FIRMA_1":
                    primera_firma += 1
                num = safe_str(r.get("numero_factura"))
                filas.append({
                    "fecha_hora":     ahora,
                    # Oracle lee ESTA columna: sin numero se escribe la clave y
                    # Oracle no encontrara correspondencia — falla en cerrado.
                    "numero_factura": num or clave,
                    "clave_factura":  clave,
                    "accion":         accion_real,
                    "comentario":     f"Aprobacion en lote desde el panel de AP: cruce correcto ({est})",
                    "departamento":   safe_str(r.get("departamento")) or "AP",
                    "aprobador":      usuario,
                    "hotel_id":       _hid,
                })
        registrar_acciones(filas)
        aprobadas = sum(1 for f in filas if f["accion"] == "APROBADA")
        no_encontradas = len(pedidas - vistas)
        _audit("AP_LOTE_APROBADO",
               f"{aprobadas} facturas aprobadas, {primera_firma} con primera firma "
               f"({ya_decididas} ya decididas, {esperan_segunda} esperan otra persona, "
               f"{no_cuadran} sin cruce correcto, {no_encontradas} no encontradas)",
               usuario)
        return jsonify({"ok": True, "aprobadas": aprobadas, "primera_firma": primera_firma,
                        "esperan_segunda": esperan_segunda, "umbral": umbral_doble_firma(),
                        "ya_decididas": ya_decididas, "no_cuadran": no_cuadran,
                        "no_encontradas": no_encontradas,
                        "claves": [f["clave_factura"] for f in filas if f["accion"] == "APROBADA"]})
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
    # Del hotel activo: si no, con un hotel sin DRR salia "ya hay DRR" y el
    # panel de al lado, vacio.
    has_drr = bool(drr_del_hotel())
    has_ar  = bool(_g.glob(os.path.join(_rdir(), 'doble_imposicion_*.xlsx')) or
                   _g.glob(os.path.join(_rdir(), 'verificacion_*.xlsx')))
    has_ap  = bool(_g.glob(os.path.join(_rdir(), 'matching_*.xlsx')))
    has_cfg = os.path.exists(os.path.join(_ddir(), 'hotel_config.json'))
    
    # Oracle mode
    oracle_mode = 'real' if os.environ.get('ORACLE_BASE_URL') else 'simulation'
    
    # SMTP configured
    smtp_ok = bool(os.environ.get('SMTP_USER') and os.environ.get('SMTP_PASSWORD'))
    
    
    health['components'] = {
        'drr':      {'ok': has_drr,   'msg': 'DRR procesado' if has_drr else 'Sin DRR'},
        'ar':       {'ok': has_ar,    'msg': 'AR datos OK' if has_ar else 'Sin datos AR'},
        'ap':       {'ok': has_ap,    'msg': 'AP datos OK' if has_ap else 'Sin datos AP'},
        'config':   {'ok': has_cfg,   'msg': 'Hotel configurado' if has_cfg else 'Sin configuración'},
        'oracle':   {'ok': True,      'msg': f'Oracle {oracle_mode}', 'mode': oracle_mode},
        'smtp':     {'ok': smtp_ok,   'msg': 'SMTP configurado' if smtp_ok else 'SMTP no configurado'},

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
        ap_hits = _g.glob(os.path.join(_rdir(), "matching_*.xlsx"))
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

ENTRADA_DIR_LEGACY   = os.path.join(BASE_DIR, 'facturas-entrada')
PROCESADAS_DIR_LEGACY2 = os.path.join(BASE_DIR, 'facturas-procesadas')
PROC_LOG_PATH  = os.path.join(_ddir(), 'archivos_procesados.json')
os.makedirs(_edir(),   exist_ok=True)
os.makedirs(_pdir(), exist_ok=True)

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


def _entrada_proc(resultado):
    """Una entrada del historial de procesados, con SU hotel (M8).

    El historial era otra puerta que no filtraba: enseñaba los archivos de
    todos los hoteles mezclados. Se estampa aqui, en el UNICO sitio que
    construye la entrada, para que las cuatro puertas que escriben en el log
    (el lote, el escaneo por foto, el camino no-stream y el reproceso) no
    puedan volver a olvidarse.
    """
    from datetime import datetime as _dtp
    try:
        _h = censo_hoteles.para_guardar()
    except Exception:
        _h = ''
    return {'fecha': _dtp.now().strftime('%Y-%m-%d %H:%M'),
            'resultado': resultado, 'hotel_id': _h}

def _detect_file_type(filename):
    """Detect what section a file belongs to."""
    name = filename.lower()
    # DRR: solo si tiene keyword explícito de DRR (no cualquier .xlsm)
    if any(x in name for x in ['drr', 'daily revenue', 'daily_revenue', 'revenue report']):
        return 'DRR'
    if any(ota in name for ota in ['booking', 'expedia', 'hotelbeds', 'hotusa', 'ota', 'comision', 'commission']):
        return 'AR'
    if name.endswith('.pdf'):
        return 'AR_o_AP'
    if name.endswith(('.xlsx', '.xls', '.csv', '.xlsm')):
        return 'AR_o_AP'  # Podría ser extracto, ventas, inventario...
    return 'AP'

@app.route('/api/historial_procesado')
@login_required
def api_historial_procesado():
    """Devuelve historial de archivos procesados con tipo y resultado."""
    log = _load_proc_log()
    # M8 · igualdad estricta, como el cruce factura<->albaran: con un hotel
    # elegido se ven SOLO sus archivos; lo que no lleva hotel (entradas de
    # antes de este cambio) se ve en la vista de grupo. El vacio NO es
    # comodin. En vista de grupo —o con 0/1 hoteles— se ve todo, que es
    # exactamente lo de siempre.
    try:
        _hact = censo_hoteles.activo()
    except Exception:
        _hact = ''
    items = []
    for fname, info in sorted(log.items(), key=lambda x: x[1].get('fecha',''), reverse=True):
        if _hact and str(info.get('hotel_id', '') or '') != _hact:
            continue
        resultado = info.get('resultado', '—')
        # Determinar el tab que se actualizó
        tab = '—'
        if resultado in ('AP_OK',): tab = 'AP — Proveedores'
        elif resultado in ('AR_OK',): tab = 'AR — OTAs'
        elif resultado in ('AR_PARCIAL',): tab = 'AR — OTAs (incompleto)'
        elif resultado in ('CONTRATO_OTA_OK',): tab = 'AR — OTAs (tarifas pactadas)'
        elif resultado in ('DRR_OK',): tab = 'DRR'
        elif resultado in ('DRR_RECIBIDO',): tab = 'DRR (recibido, sin procesar)'
        elif resultado in ('BANK_OK',): tab = 'Banco'
        elif resultado in ('FB_OK',): tab = 'F&B Cost'
        elif resultado in ('INV_OK',): tab = 'F&B Cost (Inventario/Mermas)'
        elif resultado in ('AR_REAL_OK',): tab = 'AR Real'
        # Estos se GUARDAN pero no hay pantalla donde verlos. Decir el nombre de
        # una pestaña que no existe es mandar al usuario a buscar algo que no
        # esta: se dice lo que hay.
        elif resultado in ('ALBARAN_OK',): tab = 'Albarán (sin pantalla)'
        elif resultado in ('PO_OK',): tab = 'Orden de compra (sin pantalla)'
        elif resultado in ('ROOMING',): tab = 'Rooming (sin pantalla)'
        elif resultado in ('ROOMING_NO_LEIDO',): tab = 'Rooming (no leído)'
        elif resultado in ('BEO_OK',): tab = 'Evento BEO (sin pantalla)'
        elif resultado in ('TM_OK',): tab = 'Evento TM (sin pantalla)'
        elif resultado in ('CONTRATO_OK',): tab = 'Evento contrato (sin pantalla)'
        elif 'SKIP' in resultado: tab = 'Omitido'
        elif 'ERR' in resultado or 'CRASH' in resultado: tab = 'Error'
        
        icono = '⚠' if resultado in ('AR_PARCIAL', 'DRR_RECIBIDO') else (
                '✓' if 'OK' in resultado else ('⚠' if 'SKIP' in resultado else ('ℹ' if resultado.startswith('ROOMING') else '✗')))
        items.append({
            'archivo': fname,
            'fecha': info.get('fecha', '—'),
            'resultado': resultado,
            'tab': tab,
            'icono': icono,
        })
    return jsonify({'ok': True, 'items': items[:50]})  # Últimos 50


@app.route('/api/archivos_estado', methods=['GET'])
@login_required
def api_archivos_estado():
    """List files in facturas-entrada with processed status."""
    log = _load_proc_log()
    files = []
    if os.path.exists(_edir()):
        for fname in sorted(os.listdir(_edir())):
            fpath = os.path.join(_edir(), fname)
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
        fpath = os.path.join(_edir(), fname)
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
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
    data = request.get_json(force=True, silent=True) or {}
    solo_nuevos = data.get('solo_nuevos', True)  # default: skip already processed
    tipos = data.get('tipos', ['AR', 'AP', 'DRR', 'AR_o_AP'])  # which types to process
    archivos_seleccionados = data.get('archivos', [])  # specific filenames to process
    
    if _pipeline_running:
        return jsonify({'ok': False, 'error': 'Ya hay un proceso en ejecución'}), 409
    
    log = _load_proc_log()
    from datetime import datetime as _dt
    
    def _mark_processed(fname, resultado='OK'):
        log[fname] = _entrada_proc(resultado)                   # M8
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
                candidatos = sorted(os.listdir(_edir())) if os.path.exists(_edir()) else []
            
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
                fpath = os.path.join(_edir(), fname)
                if not os.path.exists(fpath):
                    yield f'data: ✗ {fname}: archivo no encontrado\n\n'
                    continue
                
                yield f'data: >> Procesando {fname} ({tipo})...\n\n'
                
                try:
                    if tipo == 'DRR':
                        _m_drr, _marca_drr = _procesar_drr(fpath, fname)
                        yield f'data: {_m_drr}\n\n'
                        _mark_processed(fname, _marca_drr)
                        has_drr = (_marca_drr == 'DRR_OK')
                    
                    elif tipo in ('AR', 'AR_o_AP', 'AP'):
                        # Run OTA reader for AR, AP reader for others
                        if tipo == 'AR' or (tipo == 'AR_o_AP' and any(x in fname.lower() for x in ['booking','expedia','ota'])):
                            import subprocess
                            result = subprocess.run(['python3', 'lector_ota.py', '--file', fpath], 
                                capture_output=True, text=True, cwd=BASE_DIR, timeout=120, env=_env_tenant())
                            if result.returncode == 0:
                                yield f'data: ✓ AR OTA {fname}: procesado\n\n'
                                _mark_processed(fname, 'AR_OK')
                                has_ar = True
                            else:
                                yield f'data: ✗ AR {fname}: {result.stderr[:100]}\n\n'
                                _mark_processed(fname, f'ERROR: {result.stderr[:50]}')
                        else:
                            result = subprocess.run(['python3', 'lector_facturas_ap.py', '--file', fpath],
                                capture_output=True, text=True, cwd=BASE_DIR, timeout=120, env=_env_tenant())
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
                        capture_output=True, text=True, cwd=BASE_DIR, timeout=60, env=_env_tenant())
                    yield f'data: ✓ Verificación comisiones completada\n\n'
                    result2 = subprocess.run(['python3', 'detector_doble_imposicion.py'],
                        capture_output=True, text=True, cwd=BASE_DIR, timeout=60, env=_env_tenant())
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
    # Igual que el lote: antes de abrir el stream.
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
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
FACTURAS_AP_DIR_LEGACY      = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_AP_DIR_LEGACY  = os.path.join(BASE_DIR, "aprobaciones")

def _solo_hotel_activo(df):
    """Deja solo las filas del hotel elegido, cruzando por `hotel_id`.

    Lo usan AP (fase 2) y AR (fase 3). No cruza por NOMBRE: los nombres se
    editan, llevan acentos y se parecen entre si — con "Hotel Sol" y "Hotel Sol
    Mar" en el mismo grupo, el primero se llevaba las filas del segundo.

    Falla en CERRADO: si no hay columna de hotel, con un hotel elegido no se
    devuelve nada. Devolver todo seria repetir el fallo que estamos quitando —
    un filtro que parece filtrar y no filtra.

    El cuerpo vive ahora en `almacen_datos.solo_del_hotel_activo`, porque hay
    mas paneles que necesitan lo mismo (reclamaciones, aprobar AR,
    notificaciones, los emails) y tener el criterio repetido en cada uno es
    justo como se acaba con cuatro respuestas distintas a "que datos son de
    este hotel". Aqui se queda el nombre, que es el que usan AP y AR.
    """
    from almacen_datos import solo_del_hotel_activo as _solo
    return _solo(df)


def cargar_datos_ap_sin_filtrar():
    """Lo mismo que `cargar_datos_ap()` pero SIN acotar al hotel elegido.

    Mismo corte que en AR y por lo mismo: el agregador del grupo necesita las
    filas de todos los hoteles, y las necesita YA enriquecidas con las
    aprobaciones, porque si no sus contadores no cuadrarian con los del panel.
    """
    from almacen_datos import facturas_ap as _facturas_ap
    df = _facturas_ap(_pdir(), _rdir())
    if df is None or df.empty:
        return pd.DataFrame()

    # Merge con aprobaciones AP
    apro_path = os.path.join(_adir(), "aprobaciones_ap.xlsx")
    if os.path.exists(apro_path):
        try:
            df_apro = pd.read_excel(apro_path)
            if not df_apro.empty and "numero_factura" in df_apro.columns:
                ultimas = df_apro.sort_values("fecha_hora").groupby("numero_factura").last().reset_index()
                df = df.merge(ultimas[["numero_factura","accion","comentario"]], on="numero_factura", how="left")
        except Exception:
            pass
    return df


def cargar_datos_ap():
    """Carga facturas AP de TODOS los dias, ya consolidadas.

    La lectura y el deduplicado viven en almacen_datos: ese es el UNICO sitio
    que habra que tocar cuando migremos a persistencia con almacen por hotel.
    Aqui solo queda el enriquecimiento (aprobaciones, hotel activo).
    """
    return _solo_hotel_activo(cargar_datos_ap_sin_filtrar())


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
        # MATCH_ALBARAN_OK y DIFERENCIA_IMPORTE vienen del cruce con albaranes:
        # sin esto los tiles se quedaban a cero aunque el cruce hubiera corrido.
        matches        = int((estados.isin(["MATCH_CORRECTO","MATCH_3WAY_OK",
                                            "MATCH_ALBARAN_OK"])).sum())
        discrepancias  = int((estados.isin(["DISCREPANCIA_PO","DIFERENCIA_IMPORTE"])).sum())
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


def safe_str(val):
    """Texto de una celda, tratando el NaN como vacio.

    `str(float('nan'))` es 'nan', que es VERDADERO, asi que el patron habitual
    `str(x).strip() or "—"` nunca cae en el fallback y pinta el NaN en crudo.
    Es la misma trampa de pandas 3 del banco y de F&B, aqui en la pantalla:
    cuando el asignador regenera el informe contable, lo que era '' vuelve de
    Excel como NaN. Reproducido en produccion. `app_aprobacion_ap.safe_str`
    ya hacia esto mismo; el panel no lo tenia.
    """
    if val is None:
        return ''
    s = str(val).strip()
    # NO_ENCONTRADO se deja pasar a proposito: dice algo (la IA no lo encontro)
    # y hoy se ve asi. Cambiarlo seria mover otra cosa en el mismo paso.
    return '' if s.lower() in ('', 'nan', 'none', '<na>', 'nat') else s


def _cuenta_str(val):
    """El codigo contable sin el '.0' que le pega el viaje por Excel.

    cuenta_contable vuelve como float64 ('629.0'). Un numero de cuenta con
    decimal no es un numero de cuenta.
    """
    s = safe_str(val)
    if not s:
        return ''
    try:
        f = float(s)
    except (TypeError, ValueError):
        return s                      # 'REVISAR_MANUAL' y demas: tal cual
    return str(int(f)) if f == int(f) else s


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
        # OJO con el orden: r.get('estado_matching') devuelve el NaN, no el
        # segundo argumento, asi que el fallback tiene que ir por safe_str.
        # Una factura que no ha pasado por el cruce NO esta "sin PO": esta
        # PENDIENTE de cruzar, y decir otra cosa seria mentir sobre lo hecho.
        est = safe_str(r.get("estado_matching")) or safe_str(r.get("estado"))
        rows.append({
            "numero_factura":    safe_str(r.get("numero_factura")) or "N/D",
            # La identidad con la que se aprueba: el numero, o el fichero si
            # no lo hay (misma regla que app_aprobacion_ap.clave_factura).
            "clave":             safe_str(r.get("numero_factura")) or safe_str(r.get("archivo")),
            "proveedor":         safe_str(r.get("nombre_proveedor")) or "Desconocido",
            "tipo":              safe_str(r.get("tipo_proveedor")).upper() or "OTRAS",
            "total":             total,
            "cuenta_contable":   _cuenta_str(r.get("cuenta_contable")) or "—",
            "estado":            est.upper() or "PENDIENTE",
            "accion":            safe_str(r.get("accion")).upper(),
            "detalle_alerta":    safe_str(r.get("detalle_alerta")),
            "duplicados":        int(pd.to_numeric(r.get("duplicados"), errors="coerce") or 0) if r.get("duplicados") is not None else 0,
            "duplicado_de":      safe_str(r.get("duplicado_de")),
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
    _bloqueo = _falta_hotel()
    if _bloqueo:
        return _bloqueo
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
            "El asistente IA necesita una **API key de Anthropic** para responder.\n\n"
            "Mientras tanto, puedes consultar todos los datos directamente en las pestañas del dashboard: "
            "**AR**, **AP**, **DRR** y **Banco**.\n\n"
            "Para activarme, añade `ANTHROPIC_API_KEY` en las variables de entorno de Render."}), 200

    contexto = _cargar_contexto_chat()

    system_prompt = f"""Eres Yve, copiloto financiero de Yve.01 integrado en el dashboard del hotel.
Tienes acceso COMPLETO y en tiempo real a todos los módulos: AR (comisiones OTA), AP (facturas proveedores con 3-way matching), DRR (Revenue Report), Banco (conciliación), F&B Cost y Multi-Hotel.

Tu misión: dar respuestas ACCIONABLES. No solo describir el estado — decir QUÉ hacer a continuación.
Ejemplo bueno: "Tienes 3 discrepancias con Expedia por €847. Puedes reclamarlas desde AR → botón Reclamar, o te genero el email ahora."
Ejemplo malo: "Hay 3 discrepancias con Expedia."

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
            model="claude-haiku-4-5-20251001",
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
DRR_UPLOAD_DIR_LEGACY       = os.path.join(BASE_DIR, "facturas-entrada")

def drr_del_hotel(reportes_dir=None, hotel=None):
    """El ultimo informe de DRR del hotel elegido, o None.

    UNICO sitio que decide que DRR le toca a quien. Lo usan el panel, el aviso
    de "hay DRR", las notificaciones y el informe en PDF: en AR el fallo fue
    justamente tener el criterio repartido y filtrar solo un camino.

    El hotel viaja en el NOMBRE del fichero, no en una columna, porque el DRR
    es un informe por subida y todo el mundo coge "el ultimo": con una columna,
    el ultimo seria el del hotel que subio mas tarde.

    Sin hotel elegido devuelve el mas reciente de todos, que es el
    comportamiento de siempre y el que mantiene intacto el caso de 0 hoteles.
    """
    import censo_hoteles as _censo
    rdir = reportes_dir or _rdir()
    hits = glob.glob(os.path.join(rdir, "drr_procesado_*.xlsx"))
    if not hits:
        return None
    hid = hotel if hotel is not None else _censo.activo()
    if hid:
        hits = [p for p in hits
                if _censo.fichero_es_de(os.path.basename(p), hid)]
        if not hits:
            return None
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return hits[0]


def _cargar_drr_procesado():
    """El DRR del hotel activo. El criterio vive en `drr_del_hotel`."""
    return drr_del_hotel()

def num_drr(s):
    """El numero que hay dentro de un valor del DRR ya formateado, o None.

    Tolera las DOS formas en que este mismo proyecto escribe los importes, que
    no son la misma:

        '€16,360'      lo que produce `_fmt` cuando la celda trae un numero
        '16,360 EUR'   lo que escribe `lector_drr` en el fichero procesado

    El parser de antes solo quitaba '€', '%' y comas, asi que con la SEGUNDA
    forma —la de los ficheros reales— devolvia None siempre. Consecuencia: en
    la cadena del GOP, `rev_val` salia None con cualquier DRR de verdad y la
    rama que deriva el GOP del porcentaje del propio hotel no llegaba a
    ejecutarse nunca. Se veia como "no hay datos" cuando si los habia.

    Se descubrio montando la fase E, al ir a ponderar: los numeros del grupo
    salian todos a cero. Un fallo que solo aparece con el formato real y no con
    el de los tests es exactamente el que hay que dejar cerrado con una
    funcion sola y compartida.
    """
    if s is None:
        return None
    t = str(s).strip()
    if t in ("", "N/D", "nan", "None", "NaT"):
        return None
    # Fuera moneda, porcentaje y separador de millares. El productor es codigo
    # nuestro y siempre usa la coma como millar, asi que no hay ambiguedad.
    t = t.replace("€", "").replace("%", "").replace(",", "")
    t = t.replace("EUR", "").replace("eur", "").strip()
    try:
        return float(t)
    except ValueError:
        return None


def _leer_drr_stats(ruta):
    """Lee el Excel procesado del DRR y devuelve stats para el frontend."""
    def _fmt(v, is_pct=False, is_eur=False, dec=0, es_conteo=False):
        """El valor tal y como se ve en la tarjeta.

        `dec` son los decimales. NO es un adorno: el ADR, el RevPAR y el Spend
        PAR son importes POR UNIDAD, y ahi el decimal es informacion. El panel
        los formateaba con `:,.0f` como cualquier total, asi que un RevPAR de
        83,70 EUR se enseñaba como 84 EUR. Y el redondeo no se quedaba en la
        pantalla: `agregador_grupo` vuelve a leer ESTA cadena para ponderar el
        RevPAR del grupo, o sea que el numero del grupo se calculaba con 79 en
        vez de 79,20.

        `es_conteo` es para las habitaciones ocupadas: un recuento no lleva
        decimales ni simbolo de moneda. Salia "7,200.00", que no es un numero de
        habitaciones que nadie escriba.

        Y ya no hay atajo "si trae coma esta formateado". Ese atajo era el
        tercer sintoma del mismo enredo: `lector_drr` escribe los importes como
        "40,130 EUR", con coma, asi que se devolvian TAL CUAL y las tarjetas
        salian en dos formatos distintos a la vez —unas "€135" y otras
        "40,130 EUR"—. Ahora se parsea siempre con `num_drr`, que entiende las
        dos formas, y se formatea en un solo sitio.
        """
        if v is None: return "N/D"
        s = str(v).strip()
        if s in ("", "nan", "None", "N/D", "NaT"): return "N/D"
        f = num_drr(s)
        if f is None:
            return s          # texto de verdad ("REVISAR", "n/a"...): tal cual
        if is_pct:
            pct = f * 100 if abs(f) <= 1 else f
            return f"{pct:.1f}%"
        if es_conteo:
            return f"{f:,.0f}"
        if is_eur:
            return f"€{f:,.{dec}f}"
        return f"{f:,.{dec}f}" if dec else s

    # El parser vive fuera (`num_drr`) para que el agregador del grupo use
    # EXACTAMENTE el mismo. Con dos copias, la del grupo entendia un formato y
    # la del panel otro, y los numeros dejaban de cuadrar por una razon que no
    # tiene nada que ver con los hoteles.
    _num = num_drr

    try:
        # Hoja Resumen — métricas KPI
        df_res = pd.read_excel(ruta, sheet_name="Resumen", header=None)
        metricas = {}
        KEYS = ["Total Revenue", "Occupancy %", "ADR", "Revenue PAR", "GOP", "GOP %",
                "Rooms Revenue", "F&B Revenue Total", "Rooms Occupied", "Spend PAR"]
        PCT_KEYS = {"Occupancy %", "GOP %"}
        EUR_KEYS = {"Total Revenue", "GOP", "Rooms Revenue", "F&B Revenue Total", "ADR", "Revenue PAR", "Spend PAR"}
        # Importes POR UNIDAD: el decimal es informacion, no adorno. Un RevPAR
        # de 83,70 no es 84, y la diferencia se multiplica por las habitaciones
        # disponibles del mes.
        EUR_2DEC = {"ADR", "Revenue PAR", "Spend PAR"}
        # Recuentos: sin decimales y sin moneda.
        CONTEO_KEYS = {"Rooms Occupied"}
        for _, row in df_res.iterrows():
            name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if name in KEYS:
                _kw = dict(is_pct=name in PCT_KEYS, is_eur=name in EUR_KEYS,
                           dec=2 if name in EUR_2DEC else 0,
                           es_conteo=name in CONTEO_KEYS)
                metricas[name] = {
                    "today":    _fmt(row.iloc[1] if pd.notna(row.iloc[1]) else None, **_kw),
                    "mtd":      _fmt(row.iloc[2] if pd.notna(row.iloc[2]) else None, **_kw),
                    "forecast": _fmt(row.iloc[3] if pd.notna(row.iloc[3]) else None, **_kw),
                    "budget":   _fmt(row.iloc[4] if len(row) > 4 and pd.notna(row.iloc[4]) else None, **_kw),
                }
        # ── FASE D · De donde sale el GOP, dicho ────────────────────────────
        #
        # Antes habia TRES formas de rellenar un GOP que el DRR no traia, y las
        # tres terminaban poniendo un " ~" al final del texto. O sea: el
        # marcador existia pero no distinguia. Un GOP calculado con el
        # porcentaje real del hotel y un GOP inventado con la media del sector
        # llegaban a la pantalla exactamente iguales.
        #
        # Ahora cada periodo lleva su PROCEDENCIA, en un campo aparte del valor:
        #
        #   medido    el DRR trae el numero. Se enseña y se agrega.
        #   derivado  aritmetica sobre datos DEL PROPIO HOTEL (sus ingresos por
        #             su GOP%, o por el GOP% de su presupuesto). Se enseña
        #             marcado, y el grupo dice cuantos son derivados.
        #   inventado no se enseña y NO entra en ninguna suma.
        #
        # Y la rama del 22% se ha ido. No se marca: se borra. `ingresos × 0,22`
        # no dice nada del hotel que no dijeran ya los ingresos — es un numero
        # decorativo con pinta de medido, y un `~` es una nota al pie que no
        # sobrevive a un copiar y pegar a un correo o a un consejo. Un hueco es
        # informacion ("el DRR no trae el GOP, pidelo"); un numero falso, no.
        #
        # `inventado` se queda en el vocabulario aunque hoy no lo produzca
        # nadie: si mañana alguien añade otra estimacion de la nada, la regla de
        # no enseñarla ya esta puesta y escrita, no hay que acordarse de ella.
        procedencia = {}

        for period in ("today", "mtd", "forecast"):
            gop_val  = metricas.get("GOP",   {}).get(period, "N/D")
            gpct_val = metricas.get("GOP %", {}).get(period, "N/D")
            rev_val  = _num(metricas.get("Total Revenue", {}).get(period, "N/D"))

            if gop_val != "N/D":
                procedencia[period] = "medido"
                # Falta solo el porcentaje: se saca del euro medido y los
                # ingresos medidos. Sigue siendo del hotel.
                if gpct_val == "N/D" and rev_val and rev_val > 0:
                    g = _num(gop_val)
                    if g:
                        metricas.setdefault("GOP %", {})[period] = f"{g/rev_val*100:.1f}%"
                continue

            if not rev_val:
                procedencia[period] = "sin_datos"
                continue

            # 1) Su propio GOP% del periodo.
            pct = _num(gpct_val) if gpct_val != "N/D" else 0
            origen = "su GOP% del periodo"
            # 2) Si no, el GOP% de su presupuesto.
            if not pct:
                pct = _num(metricas.get("GOP %", {}).get("budget", "N/D"))
                origen = "el GOP% de su presupuesto"
            # 3) Si no, el que se deduce del presupuesto en euros.
            if not pct:
                bgt_eur = _num(metricas.get("GOP", {}).get("budget", "N/D"))
                bgt_rev = _num(metricas.get("Total Revenue", {}).get("budget", "N/D"))
                if bgt_eur and bgt_rev and bgt_rev > 0:
                    pct = bgt_eur / bgt_rev * 100
                    origen = "su presupuesto"

            if not pct:
                # Aqui es donde estaba el 22%. Ahora se queda en N/D.
                procedencia[period] = "sin_datos"
                continue

            p = pct / 100 if pct > 1 else pct
            metricas.setdefault("GOP", {})[period]   = f"€{rev_val*p:,.0f} ~"
            metricas.setdefault("GOP %", {})[period] = f"{p*100:.1f}% ~"
            procedencia[period] = "derivado"
            procedencia[period + "_origen"] = origen

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
            # FASE D: de donde sale el GOP de cada periodo — medido, derivado,
            # inventado o sin_datos. Viaja al lado del valor, no dentro: un
            # marcador metido en el texto ("22.0% ~") se pierde en cuanto
            # alguien copia la cifra.
            "gop_procedencia": procedencia,
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
        # Las secciones las nombra `lector_drr`, no este fichero. Aqui habia la
        # cadena "INCOME" escrita a mano, y los DRR reales traen "REVENUE": el
        # filtro casaba cero filas y el grafico salia vacio en silencio. Con la
        # lista compartida, las dos partes no pueden volver a discrepar.
        from lector_drr import SECCIONES_INGRESO, SECCIONES_GASTO
        _secc = df["Sección"].astype(str).str.strip().str.upper()
        income = df[_secc.isin(SECCIONES_INGRESO)].copy()
        income["Total"] = pd.to_numeric(income["Total"], errors="coerce").abs()
        expenses = df[_secc.isin(SECCIONES_GASTO)].copy()
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
        # MM-DD, PARSEANDO la fecha, no cortando la cadena por los ultimos 5
        # caracteres. `lector_drr` escribe la fecha en dos formatos segun lo que
        # traiga el .xlsm —ISO cuando la celda es una fecha de verdad, 'Jul-01'
        # cuando es texto— y con el corte a ciegas el eje del grafico salia con
        # etiquetas como 'ul-01'. Si no se puede parsear se usa la fecha tal cual,
        # que dice algo, en vez de un trozo de cadena que no dice nada.
        def _etiqueta(d):
            f = str(fechas_map.get(d, "") or "")
            try:
                return pd.to_datetime(f, errors="raise").strftime("%m-%d")
            except Exception:
                return f or str(d)
        labels = [_etiqueta(d) for d in dias]
        # Una ausencia tambien tiene que decir de donde viene. Si no hay dias,
        # el panel necesita saber si es que el DRR no trae hojas de dia o es que
        # las trae y ninguna cuadra con las secciones de ingreso — que es
        # exactamente el bug que estuvo tapado aqui. Sin esto, "vacio" y "roto"
        # se ven igual.
        _motivo = ""
        if not dias:
            _secciones = sorted(set(_secc.dropna()) - {"NAN", ""})
            if df.empty:
                _motivo = "el DRR procesado no trae ninguna hoja de dia"
            else:
                _motivo = ("el DRR trae dias pero ninguna fila de ingresos: "
                           f"secciones encontradas {_secciones or 'ninguna'}")
        return jsonify({
            "dias":     dias,
            "labels":   labels,
            "fechas":   [fechas_map.get(d, str(d)) for d in dias],
            "revenue":  [round(float(daily_rev.get(d, 0)), 0) for d in dias],
            "expenses": [round(float(daily_exp.get(d, 0)), 0) for d in dias],
            "oob":      [d in oob_dias for d in dias],
            # Antes contaba los descuadres del fichero ENTERO, independiente del
            # filtro, asi que decia a la vez "cero dias" y "un dia descuadrado".
            # Un contador tiene que contar lo que se esta mostrando.
            "oob_count": len([d for d in dias if d in oob_dias]),
            "motivo":   _motivo,
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

    os.makedirs(_edir(), exist_ok=True)
    save_path = os.path.join(_edir(), f.filename)
    f.save(save_path)

    # Ejecutar lector_drr.py
    script = os.path.join(BASE_DIR, "lector_drr.py")
    if not os.path.exists(script):
        return jsonify({"ok": False, "error": "lector_drr.py no encontrado"}), 500
    try:
        # env=_env_tenant(): sin esto el informe de un cliente se escribe en el
        # arbol del tenant base y _cargar_drr_procesado() (que lee el del
        # tenant de la sesion) no lo encuentra nunca.
        res = subprocess.run(
            [sys.executable, script, save_path],
            capture_output=True, text=True, timeout=120, cwd=BASE_DIR,
            env=_env_tenant()
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
    """Ejecuta la conciliación bancaria automática (extracto ↔ facturas AP/AR)."""
    try:
        from conciliacion_bancaria import cargar_extracto, cargar_facturas, conciliar, generar_reporte
        extracto = cargar_extracto()
        if extracto is None or extracto.empty:
            return jsonify({"ok": False,
                            "error": "No hay extracto bancario. Sube uno con ⚡ Procesar Archivos o 📸."}), 404
        facturas = cargar_facturas()
        resultado = conciliar(extracto, facturas)
        ruta = generar_reporte(resultado)
        conciliados = int((resultado["estado"] == "CONCILIADO").sum())
        diferencias = int((resultado["estado"] == "DIFERENCIA").sum())
        pendientes  = int((resultado["estado"] == "PENDIENTE").sum())
        _audit("CONCILIACION_RUN", f"{conciliados} conciliados, {diferencias} diferencias, {pendientes} pendientes")
        return jsonify({"ok": True, "total": int(len(resultado)), "conciliados": conciliados,
                        "diferencias": diferencias, "pendientes": pendientes,
                        "facturas_disponibles": len(facturas),
                        "archivo": os.path.basename(ruta)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:300]}), 500

def stats_banco(df):
    """KPIs de banco a partir de los movimientos ya juntados por el almacen.

    Funcion PURA: entra un df, sale un dict. Ni lee ficheros ni mira la sesion.
    La sacamos del endpoint para que el agregador del grupo (fase A) cuente lo
    mismo que el panel en vez de tener su propio criterio de "conciliado".

    Ojo con una cosa al usarla desde el grupo: el banco NO se parte por hotel.
    El extracto es de la cuenta de la sociedad, y `movimientos_banco()` es la
    unica funcion del almacen sin argumento `hotel` justamente por eso.
    Repartirlo entre hoteles seria inventar.
    """
    movs = df.to_dict("records")

    def _estado(m):
        e = str(m.get("estado", "") or "").strip().upper()
        return e if e else "PENDIENTE"

    total = len(movs)
    conc = sum(1 for m in movs if _estado(m) == "CONCILIADO")
    pend = sum(1 for m in movs if _estado(m) == "PENDIENTE")
    diff = sum(1 for m in movs if _estado(m) == "DIFERENCIA")

    importes = [safe_float(m.get("importe", 0)) for m in movs]
    cargos = abs(sum(x for x in importes if x < 0))
    abonos = sum(x for x in importes if x > 0)
    imp_pend = sum(abs(safe_float(m.get("importe", 0)))
                   for m in movs if _estado(m) == "PENDIENTE")

    # Alertas: pendientes de mas de 7 dias
    alertas = []
    from datetime import datetime as _dtb
    hoy = _dtb.now()
    for m in movs:
        if _estado(m) != "PENDIENTE":
            continue
        try:
            f = pd.to_datetime(m.get("fecha"), dayfirst=True)
            dias = (hoy - f).days
            if dias > 7:
                alertas.append({"concepto": str(m.get("concepto", ""))[:50],
                                "importe": safe_float(m.get("importe", 0)),
                                "dias": dias})
        except Exception:
            pass
    alertas.sort(key=lambda a: a["dias"], reverse=True)

    return {"total": total, "conciliados": conc, "pendientes": pend,
            "diferencias": diff, "importe_pendiente": round(imp_pend, 2),
            "cargos": round(cargos, 2), "abonos": round(abonos, 2),
            "alertas": alertas[:10]}


@app.route("/api/config_banco", methods=["GET", "POST"])
@login_required
def api_config_banco():
    """Cómo funciona el banco de la empresa: 'grupo' o 'por_hotel'.

    Se elige UNA vez (modal la primera vez que se abre Banco) y gobierna cómo se
    muestra/filtra. Vive en el servidor (`config_banco.json` del tenant), no en el
    navegador: así funciona igual desde el móvil y el PC. Sobrevive a los deploys
    si se siembra `YVE_BANCO_MODO` en Render (como `YVE_HOTELES_SEED`).
    """
    import config_banco as _cfg
    if request.method == "POST":
        data = request.get_json(force=True) or {}
        m = _cfg.elegir(data.get("modo"))
        if not m:
            return jsonify({"ok": False, "error": "modo inválido (grupo|por_hotel)"}), 400
        _audit("BANCO_CONFIG", f"modo={m}")
        return jsonify({"ok": True, "modo": m})
    return jsonify({"ok": True, "modo": _cfg.modo(), "elegido": _cfg.elegido()})


@app.route("/api/stats_banco")
def api_stats_banco():
    """Resumen de conciliacion bancaria para el dashboard.

    El extracto manda: total y movimientos salen SIEMPRE de extracto_banco.xlsx,
    asi que un movimiento recien subido se ve en el acto. El estado conciliado
    se recupera del ultimo informe (incluidas las asignaciones manuales) y lo
    que no este en el informe cuenta como pendiente. Quien junta los dos
    ficheros es almacen_datos.movimientos_banco(), no este endpoint.

    Segun la config del banco (config_banco):
      - 'grupo' (o sin elegir): se muestra junto, como siempre.
      - 'por_hotel': con un hotel activo, se filtra por `hotel_id`. Lo que no
        lleva hotel (extractos viejos, o subidos sin hotel) NO se esconde: se
        cuenta aparte como `sin_asignar`.
    """
    try:
        import almacen_datos as _alm
        import config_banco as _cfg
        import censo_hoteles as _censo
        _modo = _cfg.modo()
        df, info = _alm.movimientos_banco(reportes_dir=_rdir())

        if df is None or df.empty:
            return jsonify(None)

        sin_asignar = 0
        hotel = ""
        if _modo == "por_hotel":
            hotel = _censo.activo()
            if hotel:
                if "hotel_id" in df.columns:
                    # Al leer el Excel, las celdas vacías vuelven como NaN (float),
                    # no como '' — por eso fillna('') ANTES de comparar: si no, un
                    # movimiento sin hotel se contaría como 0 "sin asignar" y se
                    # perdería justo la visibilidad honesta que prometemos.
                    _h = df["hotel_id"].fillna("").astype(str).str.strip()
                    _sin = _h.isin(["", "nan", "None", "NaN"])
                    sin_asignar = int(_sin.sum())
                    df = df[_h == hotel]
                else:
                    # extracto sin columna de hotel: nada asignado a este hotel
                    sin_asignar = len(df)
                    df = df.iloc[0:0]

        return jsonify(dict(stats_banco(df),
                            modo=_modo or "grupo",
                            hotel=hotel,
                            sin_asignar=sin_asignar,
                            sin_conciliar=info.get("informe") is None,
                            archivo=info.get("informe"),
                            extracto=info.get("extracto")))
    except Exception as e:
        return jsonify({"error": str(e)}), 500



@app.route("/api/smtp_status")
def api_smtp_status():
    """Comprueba si el email está configurado (Resend o SMTP)."""
    resend_key = os.environ.get("RESEND_API_KEY", "")
    smtp_user  = os.environ.get("SMTP_USER", "")
    smtp_pass  = os.environ.get("SMTP_PASSWORD", "")

    # Prioridad 1: Brevo
    brevo_key = os.environ.get("BREVO_API_KEY", "")
    if brevo_key:
        return jsonify({
            "configured": True, "method": "brevo",
            "user": smtp_user or "Brevo API",
            "ok": True,
            "msg": f"Brevo configurado — listo para enviar"
        })

    # Prioridad 2: Resend
    if resend_key:
        return jsonify({
            "configured": True, "method": "resend",
            "user": smtp_user or "Resend API",
            "ok": True,
            "msg": "Resend configurado (requiere dominio verificado para enviar a cualquier email)"
        })

    # Fallback: SMTP (bloqueado en Render free tier)
    if smtp_user and smtp_pass:
        return jsonify({
            "configured": True, "method": "smtp",
            "user": smtp_user,
            "ok": False,
            "msg": "SMTP configurado pero Render free tier bloquea conexiones SMTP. Añade RESEND_API_KEY."
        })

    return jsonify({
        "configured": False, "ok": False, "user": "",
        "msg": "Sin configuración de email. Añade RESEND_API_KEY en Render."
    })

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
            "message": ("✓ Enviado por: " + ", ".join(k for k,v in resultados.items() if v)) if any(resultados.values()) else ("⚠ Error al enviar por: " + ", ".join(resultados.keys()) + ". Comprueba la configuración de Resend.") if resultados else "No se pudo enviar. Verifica que el email de destino está configurado y guardado."
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/hotel_config")
def api_hotel_config():
    """Devuelve configuración del hotel (nombre, email, etc.)."""
    path = os.path.join(_ddir(), "hotel_config.json")
    try:
        data = json.load(open(path)) if os.path.exists(path) else {}
        return jsonify(data)
    except Exception:
        return jsonify({})

@app.route("/api/notif_config", methods=["GET"])
def api_notif_config_get():
    """Devuelve la configuración de notificaciones."""
    path = os.path.join(_ddir(), "notif_config.json")
    default = {
        "canales": {"email": False, "whatsapp": False, "slack": False, "push": False},
        "email": "", "whatsapp": "", "slack_webhook": "",
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
    path = os.path.join(_ddir(), "notif_config.json")
    data = request.get_json(silent=True) or {}
    try:
        with open(path, "w") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500



def _leer_csv_flexible(f):
    """Lee un CSV sin dar por hecho el separador ni el decimal.

    Se prueban los cuatro separadores reales por orden de cuantas columnas
    sacan: gana el que mas produce, porque un separador equivocado deja el
    fichero en una sola columna. Y si los numeros vienen a la española
    (1.234,56) se convierten al vuelo — la coma decimal es lo normal en un TPV
    de aqui, y una columna de cantidades que llega como texto vale lo mismo que
    no llegar.
    """
    import pandas as pd, io
    crudo = f.read()
    if isinstance(crudo, str):
        crudo = crudo.encode('utf-8')
    mejor = None
    for sep in (';', ',', '\t', '|'):
        for cod in ('utf-8-sig', 'latin-1'):
            try:
                d = pd.read_csv(io.BytesIO(crudo), sep=sep, encoding=cod)
            except Exception:
                continue
            if mejor is None or len(d.columns) > len(mejor.columns):
                mejor = d
            break
    if mejor is None:
        raise ValueError('no he podido leer el CSV con ningun separador conocido')
    # numeros a la española: solo donde la columna ha llegado como texto.
    #
    # La comprobacion es "NO es numerica", no "es object": en pandas 3 una
    # columna de texto tiene dtype `str`, no `object`, asi que preguntar por
    # object no encontraba ninguna y los importes se quedaban en '7.595,00'
    # tal cual. Reventaba luego, al dividir para sacar el precio unitario:
    # "unsupported operand type(s) for /: 'str' and 'int'".
    for c in mejor.columns:
        if pd.api.types.is_numeric_dtype(mejor[c]):
            continue
        muestra = mejor[c].dropna().astype(str).head(20)
        if len(muestra) and all(_re_num_es.match(v.strip()) for v in muestra):
            mejor[c] = [safe_float(v) for v in mejor[c]]
    return mejor


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
            # Un TPV español exporta con punto y coma, que es lo normal cuando la
            # coma es el separador decimal. Con `read_csv(f)` a secas el fichero
            # entero entraba como UNA sola columna y la subida se rechazaba por
            # columnas que si estaban.
            df_new = _leer_csv_flexible(f)
        elif fname.endswith(('.xlsx', '.xls')):
            df_new = pd.read_excel(f)
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Usa .xlsx o .csv"}), 400

        # El MISMO mapa que usan los otros dos caminos de ventas (el lote y la
        # foto), en vez de una tercera cadena de `elif` propia. Aqui vivia el
        # bug: comparaba por SUBCADENA con `'id' in cl`, e "id" esta dentro de
        # `cantidad`, de `unidad` y hasta de `unidades_vendidas`. La columna de
        # cantidades acababa renombrada a `id_receta`, `unidades_vendidas` se
        # rellenaba con 1 unos renglones mas abajo, y el food cost salia 0,06%
        # en vez de 16,25%. Con cabeceras reales de TPV
        # (Fecha·Plato·Categoria·Cantidad·Importe) era peor: `Importe` no casaba
        # con nada y la subida se rechazaba por "falta total_venta".
        df_new = _normalize_cols(df_new, _VEN_COL_MAP)

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

        # El hotel y el deduplicado, en el mismo sitio que las otras tres
        # puertas. La lista de columnas se queda —aqui SI interesa recortar a
        # las del esquema— pero ojo: es una lista blanca, y una lista blanca se
        # come en silencio lo que no este en ella. Es exactamente como
        # `guardar_excel` de lector_ota se comio el hotel en la fase 3 y nos
        # costo una verificacion entera. Por eso `hotel_id` lo pone
        # `_guardar_fb_del_hotel` DESPUES del recorte, no antes.
        df_new = df_new[['fecha', 'id_receta', 'nombre_plato', 'categoria',
                         'unidades_vendidas', 'precio_unitario', 'total_venta']]
        df_combined, _ = _guardar_fb_del_hotel(df_new, 'ventas_fb_diarias.xlsx')

        return jsonify({
            "ok": True,
            "filas_importadas": len(df_new),
            "total_filas": len(df_combined),
            "fechas": df_new['fecha'].dropna().unique().tolist()[:5],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# Como se llama cada cosa en el Excel de un hotel. El orden importa: se mira
# de la mas especifica a la mas general, porque "precio" a secas puede ser el
# PVP del plato o el coste del ingrediente.
_REC_COLS = [
    ('id_receta',      ('id_receta', 'id receta', 'idreceta', 'codigo', 'cod', 'ref')),
    ('nombre',         ('receta', 'plato', 'nombre_plato', 'nombre receta', 'nombre', 'elaboracion')),
    ('categoria',      ('categoria', 'familia', 'tipo', 'grupo', 'seccion')),
    ('precio_venta',   ('precio_venta', 'pvp', 'precio venta', 'precio_pvp', 'venta')),
    ('ingrediente',    ('ingrediente', 'producto', 'articulo', 'materia_prima', 'materia prima', 'item')),
    ('cantidad',       ('cantidad', 'cant', 'qty', 'peso', 'dosis', 'racion')),
    ('unidad',         ('unidad', 'ud', 'medida', 'um')),
    ('coste_unitario', ('coste_unitario', 'coste unitario', 'coste', 'precio_coste',
                        'precio coste', 'coste_kg', 'importe_unitario')),
]


def _rec_norm_cab(col):
    """Cabecera -> nombre canonico, o None si no se reconoce."""
    import unicodedata as _u
    s = _u.normalize('NFKD', str(col))
    s = ''.join(c for c in s if not _u.combining(c)).strip().lower()
    for ch in ('.', '-', '/', '\\'):
        s = s.replace(ch, ' ')
    s = ' '.join(s.split())
    # Las columnas del formato INTERNO se respetan tal cual y salen antes de
    # cualquier alias. Si no, 'ingredientes_json' cae en la busqueda laxa por
    # 'ingrediente' —es subcadena suya— y el fichero exportado por Yve se leia
    # como si fuera el del hotelero (la trampa de siempre con los nombres que
    # son sufijo de otro).
    if s.replace(' ', '_') in ('ingredientes_json', 'id_receta', 'precio_venta'):
        return s.replace(' ', '_')
    for destino, alias in _REC_COLS:
        if s == destino or s in alias:
            return destino
    for destino, alias in _REC_COLS:
        if any(a in s for a in alias):
            return destino
    return None


@app.route("/fb/api/upload_recetas", methods=["POST"])
@login_required
def api_upload_recetas():
    """Sube el escandallo del cliente a recetas.xlsx.

    DOS formatos, detectados solos:
      1. UNA FILA POR INGREDIENTE — el que tiene de verdad un hotelero:
         receta | ingrediente | cantidad | unidad | coste | PVP | categoria
         Se agrupa por receta y se construye el `ingredientes_json` por dentro.
      2. El formato INTERNO (una fila por receta, con `ingredientes_json`),
         para poder reimportar lo que exporte Yve.

    Se REEMPLAZA POR id_receta: una receta corregida pisa a la vieja, y las que
    no vengan en el fichero se conservan. Asi se puede subir un fichero con dos
    platos para arreglarlos sin perder la carta entera.
    """
    import pandas as pd, json as _json
    from tab_fb_dashboard import (_clave_plato, _invalidate as _fb_inv,
                                  _num_fb, _txt_ing)

    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "No se recibió archivo"}), 400
    fname = (f.filename or '').lower()
    try:
        if fname.endswith('.csv'):
            df = pd.read_csv(f, sep=None, engine='python')
        elif fname.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(f)
        else:
            return jsonify({"ok": False, "error": "Formato no soportado. Usa .xlsx o .csv"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"No se ha podido leer el archivo: {str(e)[:120]}"}), 400

    df = df.rename(columns={c: (_rec_norm_cab(c) or c) for c in df.columns})
    filas = df.to_dict('records')
    avisos = []

    # ── formato interno: una fila por receta, con el JSON ya hecho ─────────
    if 'ingredientes_json' in df.columns:
        nuevas = []
        for r in filas:
            nombre = str(r.get('nombre') or '').strip()
            if not nombre:
                continue
            ings = r.get('ingredientes_json')
            if not isinstance(ings, str) or not ings.strip():
                ings = '[]'
            try:
                _json.loads(ings)
            except Exception:
                avisos.append(f"«{nombre}»: ingredientes_json ilegible, se ha dejado vacío")
                ings = '[]'
            nuevas.append({
                'id_receta': str(r.get('id_receta') or '').strip() or _clave_plato(nombre),
                'nombre': nombre,
                'categoria': str(r.get('categoria') or '').strip() or 'General',
                'precio_venta': _num_fb(r.get('precio_venta'), 0.0),
                'ingredientes_json': ings,
            })
        n_ing = sum(len(_json.loads(x['ingredientes_json'])) for x in nuevas)
        formato = 'interno (una fila por receta)'
    else:
        # ── formato del hotelero: una fila por ingrediente ─────────────────
        if 'ingrediente' not in df.columns or 'nombre' not in df.columns:
            return jsonify({"ok": False, "error":
                "No encuentro las columnas mínimas. El archivo debe tener una fila por "
                "ingrediente con al menos: receta (o plato) e ingrediente. Y a poder ser "
                "cantidad, unidad, coste y precio de venta."}), 400
        grupos, orden = {}, []
        for r in filas:
            nombre = str(r.get('nombre') or '').strip()
            if not nombre:
                continue
            # la clave de agrupacion es el id si lo traen, y si no el nombre
            # normalizado: tiene que ser ESTABLE entre subidas para que
            # "reemplazar por id_receta" funcione al volver a subir el fichero.
            rid = str(r.get('id_receta') or '').strip() or _clave_plato(nombre)
            if rid not in grupos:
                grupos[rid] = {'nombre': nombre, 'categoria': '', 'precio_venta': 0.0, 'ings': []}
                orden.append(rid)
            g = grupos[rid]
            if not g['categoria']:
                g['categoria'] = str(r.get('categoria') or '').strip()
            if not g['precio_venta']:
                g['precio_venta'] = _num_fb(r.get('precio_venta'), 0.0)
            ing_nom = str(r.get('ingrediente') or '').strip()
            cant = _num_fb(r.get('cantidad'), 0.0)
            if not ing_nom or cant <= 0:
                continue        # una linea sin ingrediente o sin cantidad no aporta
            linea = {'ingrediente': ing_nom, 'cantidad': cant}
            _u = str(r.get('unidad') or '').strip()
            if _u:
                linea['unidad'] = _u
            _c = _num_fb(r.get('coste_unitario'), None)
            if _c is not None and _c > 0:
                linea['coste_unitario'] = _c
            g['ings'].append(linea)

        nuevas, sin_ing = [], []
        for rid in orden:
            g = grupos[rid]
            if not g['ings']:
                # Una receta sin una sola linea util costaria 0 y saldria con
                # food cost 0%: eso es peor que no tenerla. No se guarda y se
                # dice cual.
                sin_ing.append(g['nombre'])
                continue
            nuevas.append({
                'id_receta': rid,
                'nombre': g['nombre'],
                'categoria': g['categoria'] or 'General',
                'precio_venta': g['precio_venta'],
                'ingredientes_json': _json.dumps(g['ings'], ensure_ascii=False),
            })
        if sin_ing:
            avisos.append(f"{len(sin_ing)} receta(s) sin ningún ingrediente con cantidad, "
                          f"no se han guardado: {', '.join(sin_ing[:4])}"
                          + ("..." if len(sin_ing) > 4 else ""))
        n_ing = sum(len(g['ings']) for g in grupos.values())
        formato = 'una fila por ingrediente'

    if not nuevas:
        return jsonify({"ok": False, "error":
            "No se ha podido leer ninguna receta del archivo. " + (avisos[0] if avisos else "")}), 400

    # ── avisar de lo que costara de menos ─────────────────────────────────
    # Un ingrediente que no esta en el inventario Y no trae coste propio vale 0:
    # la receta sale mas barata de lo que es y el food cost mas bajo. Callarselo
    # seria dar por bueno un numero que no lo es.
    try:
        _inv_path = os.path.join(_ddir(), 'inventario.xlsx')
        _inv = pd.read_excel(_inv_path) if os.path.exists(_inv_path) else pd.DataFrame()
        # _txt_ing y no una normalizacion propia: es LA MISMA clave con la que
        # `_calc_recipe_costs` cruza contra el inventario. Con dos criterios
        # distintos, el aviso cantaria "Café molido no esta en tu inventario"
        # mientras el calculo ya lo esta cobrando contra "Cafe molido".
        _conocidos = {k for k in (_txt_ing(x) for x in _inv.get('ingrediente', [])) if k}
        _huerfanos = set()
        for x in nuevas:
            for ing in _json.loads(x['ingredientes_json']):
                if (_txt_ing(ing.get('ingrediente')) not in _conocidos
                        and not ing.get('coste_unitario')):
                    _huerfanos.add(str(ing.get('ingrediente', ''))[:30])
        if _huerfanos:
            avisos.append(f"{len(_huerfanos)} ingrediente(s) no están en tu inventario y no "
                          f"traen coste, así que cuentan 0 €: "
                          f"{', '.join(sorted(_huerfanos)[:4])}"
                          + ("..." if len(_huerfanos) > 4 else ""))
    except Exception:
        pass

    # ── guardar: REEMPLAZAR por id_receta, conservar el resto ─────────────
    ruta = os.path.join(_ddir(), 'recetas.xlsx')
    try:
        df_old = pd.read_excel(ruta) if os.path.exists(ruta) else pd.DataFrame()
    except Exception:
        df_old = pd.DataFrame()
    ids_nuevos = {x['id_receta'] for x in nuevas}
    if not df_old.empty and 'id_receta' in df_old.columns:
        df_old = df_old[df_old['id_receta'].map(lambda v: str(v).strip() not in ids_nuevos)]
    _COLS = ['id_receta', 'nombre', 'categoria', 'precio_venta', 'ingredientes_json']
    df_fin = pd.concat([df_old, pd.DataFrame(nuevas)], ignore_index=True)
    for c in _COLS:
        if c not in df_fin.columns:
            df_fin[c] = ''
    df_fin[_COLS].to_excel(ruta, index=False)

    if 'recetas.xlsx' in _EXCEL_CACHE:
        del _EXCEL_CACHE['recetas.xlsx']
    try:
        _fb_inv()
    except Exception:
        pass

    return jsonify({"ok": True, "formato": formato,
                    "recetas_importadas": len(nuevas),
                    "ingredientes": n_ing,
                    "total_recetas": len(df_fin),
                    "avisos": avisos})


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
                _CAL_CACHE.clear()
    except Exception:
        pass
    return jsonify({"ok": True, "message": "Cache limpiado"})


@app.route("/api/notificaciones")
def api_notificaciones():
    """Devuelve historial de notificaciones."""
    hist_path = os.path.join(_ddir(), "notificaciones_historial.json")
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
        from notificaciones import enviar_pendientes
        alertas = enviar_pendientes()
        return jsonify({"ok": True, "enviadas": len(alertas)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:200]}), 500


# ── Chat AI — Yve Copilot ──────────────────────────────────────────────

def _hotel_name():
    cfg_path = os.path.join(_ddir(), "hotel_config.json")
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
    tag = name if name else ""
    configured = "true" if name else "false"
    user_name = current_user.nombre
    user_rol  = current_user.rol
    out = HTML.replace("__HOTEL_TAG__", tag).replace("__CONFIGURED__", configured)
    admin_display = "inline" if user_rol in ("admin", "financial_controller") else "none"
    out = out.replace("__USER_NAME__", user_name).replace("__USER_ROL__", user_rol)
    out = out.replace("__ADMIN_DISPLAY__", admin_display)
    # Sello de los estaticos: sin el, el navegador sirve de su cache un i18n
    # viejo y una traduccion arreglada tarda dias en llegar (medido).
    out = out.replace("__ASSETS_V__", SELLO_ESTATICOS)
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
<meta name="theme-color" content="#0f172a">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Yve.01">
<link rel="apple-touch-icon" href="/static/icons/apple-touch-icon.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/icons/favicon-32.png">
<link rel="icon" type="image/png" sizes="16x16" href="/static/icons/favicon-16.png">
<link rel="mask-icon" href="/static/icons/favicon.svg" color="#3b82f6">
<link rel="icon" type="image/svg+xml" href="/static/icons/favicon.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Dashboard</title>
<script async src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
/* Light mode */
body.light-mode{--bg:#f8fafc;--s1:#fff;--s2:#e2e8f0;--s3:#cbd5e1;--tx:var(--bg);--mut:#475569;--dim:#64748b;--acc:var(--acc-dark);--acc2:var(--acc-dark);--grn:#16a34a;--red:#dc2626;--ora:#ea580c;--pur:#7c3aed}
body.light-mode .nav{background:rgba(248,250,252,.9);border-bottom-color:#e2e8f0}
body.light-mode .tab-btn{color:#475569}
body.light-mode .tab-btn.active{color:var(--acc-dark)}
/* ── Skeleton loading ─────────────────────────────── */
@keyframes confettiFall{0%{transform:translateY(0) translateX(0) rotate(0);opacity:1}100%{transform:translateY(110vh) translateX(8vw) rotate(900deg);opacity:0}}
@keyframes tourRing{0%{width:20px;height:20px;opacity:.9}100%{width:120vmax;height:120vmax;opacity:0}}
@keyframes shimmer{0%{background-position:-400px 0}100%{background-position:400px 0}}
.skeleton{background:linear-gradient(90deg,var(--s1) 25%,var(--s2) 50%,var(--s1) 75%);
  background-size:800px 100%;animation:shimmer 1.4s infinite;border-radius:6px;
  color:transparent!important;pointer-events:none}
.skeleton *{visibility:hidden}
/* ─────────────────────────────────────────────────── */
.show-mobile{display:none!important}
:root{
  --tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;--yel:#eab308;--pur:#8b5cf6;
}
*{box-sizing:border-box;margin:0;padding:0}
/* Reservar SIEMPRE el hueco de la barra de desplazamiento. Sin esto, un
   apartado que no llena la ventana se queda sin barra, la pagina es 10 px
   mas ancha y toda la interfaz se corre a la derecha; al volver a uno
   largo, salta. Medido: nav 501 con barra / 511 sin ella. */
html{overflow-x:hidden;scrollbar-gutter:stable}
body{
  overflow-x:hidden;background:var(--bg);color:var(--tx);
  font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  min-height:100vh;line-height:1.5;position:relative;
}
/* Gradient glow igual que el login — sutil, solo ambiente */
body::before{
  content:'';position:fixed;inset:0;z-index:0;pointer-events:none;
  background:
    radial-gradient(900px 500px at 90% -5%,rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1),transparent 60%),
    radial-gradient(700px 400px at -5% 105%,rgba(139,92,246,.08),transparent 55%)
}
.main{position:relative;z-index:1}

/* ── NAV ── */
.nav{
  background:rgba(var(--bg-r,15),var(--bg-g,23),var(--bg-b,42),.92);
  backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  border-bottom:1px solid var(--s2);
  padding:env(safe-area-inset-top) 24px 0 24px;height:calc(60px + env(safe-area-inset-top));box-sizing:border-box;
  display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200
}
.logo{display:flex;align-items:baseline;gap:10px;flex-shrink:0}
.logo-name{font-family:'Space Grotesk','Inter',sans-serif;font-size:20px;font-weight:700;color:#fff;letter-spacing:-0.3px}
.logo-tag{font-size:11px;color:var(--mut);font-weight:400;white-space:nowrap}
.logo-dot{width:10px;height:10px;border-radius:50%;background:var(--acc,#3b82f6);flex-shrink:0;box-shadow:0 0 6px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.6),0 0 14px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.35),0 0 28px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)}
.logo-dot-one{color:var(--acc2)}
.logo-mark{display:none}
.nav-mid{flex:1}
.nav-right{display:flex;align-items:center;gap:10px;flex-shrink:0}
.pill{font-size:11px;color:var(--mut);background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.07);padding:4px 12px;border-radius:20px;white-space:nowrap;letter-spacing:.2px}
.btn-ref{background:none;border:1px solid var(--s2);color:var(--mut);padding:6px 12px;border-radius:8px;font-size:12px;cursor:pointer;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;white-space:nowrap}
.btn-ref:hover{border-color:var(--acc);color:var(--acc2)}
.btn-run{background:linear-gradient(135deg,rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.18),rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10));color:var(--acc2);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4);padding:9px 18px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:8px;white-space:nowrap;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s}
.btn-run:hover{background:linear-gradient(135deg,rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.30),rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.18));border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.7)}
.btn-run:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
@media(max-width:640px){
  .logo-tag,.pill{display:none}
  .nav{padding:env(safe-area-inset-top) 10px 0 10px;gap:6px;height:calc(50px + env(safe-area-inset-top));box-sizing:border-box}
  .logo-name{font-size:16px}
  .btn-run{padding:8px 12px;font-size:12px}
  .btn-ref{padding:4px 8px;font-size:11px}
  .nav-right{gap:6px}
  #daily-alerts-panel{display:none!important}
  #tour-box{max-width:280px!important;padding:14px!important;font-size:12px!important}
  #tour-box h3{font-size:14px!important}
  #tour-congrats>div{padding:20px!important;max-width:300px!important}
  #tour-congrats>div>div:first-child{font-size:36px!important}
  #mobile-kpi-bar{display:none!important}
  #exec-summary{display:none!important}
  /* Stats cards: 2 columnas en móvil */
  .stats,.metrics,.kpi-row{grid-template-columns:repeat(2,1fr)!important;gap:8px!important}
  /* Sub-tabs F&B: scroll horizontal */
  .fb-sub{font-size:11px!important;padding:6px 10px!important}
  /* Tabs principales: scroll horizontal */
  .tabs{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;flex-wrap:nowrap!important}
  .tabs::-webkit-scrollbar{display:none}
  .tab{font-size:11px;padding:8px 10px;white-space:nowrap;flex-shrink:0}
  /* Main padding reducido */
  .main{padding:12px!important}
  /* Tablas: scroll horizontal */
  .table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  table{min-width:auto!important;font-size:11px}
  /* Gráficos: ancho completo */
  .mid,.chart-grid{grid-template-columns:1fr!important}
  /* Modal: ocupa más pantalla */
  .modal{width:calc(100% - 16px)!important;max-width:none!important;margin:8px!important;padding:16px!important}
  .fb-kpi-grid{grid-template-columns:repeat(2,1fr)!important;gap:8px!important}
  .fb-chart-grid{grid-template-columns:1fr!important}
  .fb-kpi-card{padding:12px 10px!important}
  .fb-kpi-val{font-size:20px!important}
  .fb-kpi-lbl{font-size:9px!important}
  h2{font-size:15px!important}
}

/* ── MAIN ── */
.main{padding:24px;max-width:1440px;margin:0 auto}
@media(max-width:640px){.main{padding:14px}}

/* ── ALERT ── */
.alert{display:none;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.25);border-radius:12px;padding:12px 18px;font-size:13px;color:#fca5a5;margin-bottom:20px;align-items:center;gap:10px}
.alert.on{display:flex}

/* ── STATS ── */
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:22px}
@media(max-width:1200px){.stats{grid-template-columns:repeat(3,1fr)}}


@media(max-width:900px){
  #ar-real-grid{grid-template-columns:1fr}
  .metrics{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:768px){
  /* Nav */
  .nav{padding:env(safe-area-inset-top) 10px 0 10px;gap:4px;height:calc(52px + env(safe-area-inset-top));box-sizing:border-box}
  .logo-name{font-size:15px}
  .logo-tag{display:none}
  /* M10 · La nav derecha medía 409 px en una pantalla de 370: el selector de
     hotel (151) + Procesar (141) + los dos botones. Con `flex-shrink:0` no
     cabia y arrastraba TODA la pagina hacia la derecha —162 px medidos—, que
     es el "se descoloca" al cambiar de apartado. Ahora encoge y, si aun asi
     no cabe, se desliza DENTRO de la nav en vez de mover la pagina. */
  /* M10 (revisado) · La barra NO se desliza. El `overflow-x:auto` que se puso
     aqui hacia dos cosas mal: la barra se movia de lado con el dedo, y ademas
     RECORTABA el menu de la rueda —que vive dentro de esta caja—, asi que al
     pulsarla no se veia nada. Medido: el menu iba de x=250 a x=468 y la barra
     acababa en x=348.
     Ahora cabe todo sin deslizar, porque el boton de Yve se ha ido a la
     burbuja flotante y el selector de hotel es mas estrecho. */
  .nav-right{gap:4px;flex-shrink:1;min-width:0}
  /* El boton de procesar, solo el rayo. Medido a 370 px: los hijos de la barra
     pedian 277 px (selector 92 + procesar 137 + rueda 40 + huecos) en una caja
     de 225, asi que la rueda acababa en x=400 — fuera de la pantalla y sin
     poder pulsarla. Quitando el texto sobra sitio para las tres. */
  #run-lbl{font-size:0;letter-spacing:0}
  #run-lbl::after{content:'⚡';font-size:15px;letter-spacing:normal}
  .btn-run{padding:7px 11px}
  /* `!important` a proposito: el selector lleva `max-width:190px` EN LINEA
     (esta pintado con estilo inline), y sin esto la regla no gana nunca —
     medido en el navegador: seguia en 151 px. */
  #hotel-activo-sel{max-width:92px!important;font-size:10.5px;padding:3px 8px}
  /* El desplegable, anclado a la PANTALLA y no al boton. Da igual donde acabe
     la barra: siempre cabe entero. */
  .menu{position:fixed;top:calc(52px + env(safe-area-inset-top));right:8px;left:auto;
        min-width:0;width:min(268px,calc(100vw - 16px));max-height:calc(100vh - 74px)}
  /* Red de seguridad: que ningun elemento pueda volver a desplazar la pagina
     de lado. La causa se arregla arriba; esto es para que no vuelva. */
  html,body{max-width:100%;overflow-x:hidden}
  /* M9 · El historial de procesados, en tarjetas: una tabla de 4 columnas no
     cabe en un movil y se leia a trozos. */
  .hist-t thead{display:none}
  .hist-t,.hist-t tbody,.hist-t tr,.hist-t td{display:block;width:auto}
  .hist-t{min-width:0!important}
  .hist-t tr{border:1px solid var(--s2)!important;border-radius:10px;margin-bottom:8px;padding:8px 10px;position:relative}
  .hist-t td{padding:2px 0!important;max-width:none!important;white-space:normal!important;overflow:visible!important;word-break:break-word}
  .hist-t td[data-r]::before{content:attr(data-r) ': ';color:var(--dim);font-size:10px;text-transform:uppercase;letter-spacing:.4px}
  .hist-t td:first-child{position:absolute;right:8px;top:8px;padding:0!important}
  .btn-ref{font-size:11px;padding:5px 8px}
  .btn-run{font-size:12px;padding:7px 12px}
  #btn-install-pwa{display:none}
  .hide-mobile{display:none!important}
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
  .show-mobile{display:inline-flex!important}
  
  #cal-grid{grid-template-columns:1fr!important}
  /* Multi-hotel table */
  #mh-kpis .sc-lbl{font-size:9px}
  /* Status bar */
  .status-bar{font-size:10px;padding:5px 10px}
  /* Demo banner */
  #demo-banner{font-size:10px;padding:4px 8px}
  /* Back to top */
  #back-top{display:none!important}
  #notif-canales{grid-template-columns:repeat(2,1fr)!important;gap:10px!important}
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
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s}
.sc:hover{border-color:var(--s3);transform:translateY(-1px)}
.sc.hl{border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4);background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.05)}
.sc-lbl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px}
.sc-val{font-size:28px;font-weight:800;line-height:1;letter-spacing:-1px}
.sc-sub{font-size:11px;color:var(--dim);margin-top:6px}
.sc.c-blu .sc-val{color:var(--acc2)}
.sc.c-grn .sc-val{color:var(--grn)}
.sc.c-red .sc-val{color:var(--red)}
.sc.c-ora .sc-val{color:var(--ora)}
.sc.c-yel .sc-val{color:var(--yel)}
.sc.c-pur .sc-val{color:var(--pur)}
/* Modo acento total: todos los contenedores responden al acento */
body.acentuar-todo .sc,
body.acentuar-todo .card,
body.acentuar-todo .drr-mc,
body.acentuar-todo .fb-kpi-card{border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)!important;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.06)!important}
body.acentuar-todo .sc .sc-val{color:var(--acc2)!important}
body.acentuar-todo .fb-kpi-val{color:var(--acc2)!important}
body.acentuar-todo .card-title,
body.acentuar-todo .fb-kpi-lbl{color:var(--acc2)!important;opacity:.8}
/* Modo OFF: neutralizar el .hl del primer card para que sea igual al resto */
body:not(.acentuar-todo) .sc.hl{border-color:var(--s2)!important;background:var(--s1)!important}
.sc.upd-green,.card.upd-green,.fb-kpi-card.upd-green{border-color:rgba(34,197,94,.6)!important;background:rgba(34,197,94,.08)!important}
.sc.upd-green .sc-val,.fb-kpi-card.upd-green .fb-kpi-val{color:#22c55e!important}
body.acentuar-todo .sc.upd-green,body.acentuar-todo .card.upd-green,body.acentuar-todo .fb-kpi-card.upd-green{border-color:rgba(34,197,94,.75)!important;background:rgba(34,197,94,.1)!important}

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
.b-cok{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.12);color:var(--acc2);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2)}
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
.l-ok{color:#4ade80}@keyframes pulse-green{0%,100%{box-shadow:0 0 4px rgba(34,197,94,.4)}50%{box-shadow:0 0 12px rgba(34,197,94,.8)}}.l-err{color:#f87171}.l-info{color:var(--acc2);font-weight:700}.l-warn{color:#facc15}.l-dim{color:#475569}
.modal-f{margin-top:16px;display:flex;justify-content:flex-end;gap:10px}
.btn-cl{background:var(--s2);color:var(--tx);border:none;padding:9px 20px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s}
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
.tab{padding:10px 18px;background:none;border:none;color:var(--mut);cursor:pointer;font-size:.85rem;font-weight:600;border-bottom:2px solid transparent;transition:background-color .18s,border-color .18s,color .18s,box-shadow .18s,transform .18s,opacity .18s;white-space:nowrap}
.tab:hover{color:var(--tx)}
.tab.active{color:var(--acc2);border-bottom-color:var(--acc)}
.panel{display:none}.panel.active{display:block;animation:fadeIn .18s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:none}}
/* ── AP Cards ─────────────────────────────────────────── */
.ap-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.75rem;font-weight:700;letter-spacing:.04em}
.ap-badge.fb{background:rgba(139,92,246,.2);color:#c4b5fd}
.ap-badge.otras{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2);color:var(--acc3)}
.ap-badge.ok{background:rgba(34,197,94,.2);color:#86efac}
.ap-badge.disc{background:rgba(239,68,68,.2);color:#fca5a5}
.ap-badge.alerta{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15);color:var(--acc3)}
.ap-badge.sinpo{background:rgba(234,179,8,.2);color:#fde047}
.ap-badge.manual{background:rgba(249,115,22,.2);color:#fed7aa}
.alerta-box{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.3);border-radius:6px;padding:8px 12px;margin-top:8px;font-size:.8rem;color:var(--acc3)}
.disc-box{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);border-radius:6px;padding:8px 12px;margin-top:8px;font-size:.8rem;color:#fca5a5}

/* ── Chat AI — Yve Copilot ─────────────────────────────── */
/* ── Chat / Ask Yve ─────────────────────────────────── */
#chat-fab{
  position:fixed;bottom:24px;right:24px;z-index:1000;
  display:flex;align-items:center;gap:10px;
  background:linear-gradient(135deg,#7c3aed,var(--acc));
  color:#fff;border:none;border-radius:50px;
  padding:13px 20px 13px 16px;cursor:pointer;
  font-size:.9rem;font-weight:700;
  box-shadow:0 4px 20px rgba(124,58,237,.5);transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s;
}
#chat-fab:hover{transform:translateY(-2px);box-shadow:0 8px 32px rgba(124,58,237,.6)}
#chat-fab .fab-dot{
  width:9px;height:9px;border-radius:50%;
  background:#22c55e;box-shadow:0 0 6px #22c55e;
  animation:pulse-dot 2s infinite;flex-shrink:0;
}
@keyframes pulse-dot{0%,100%{opacity:1}50%{opacity:.4}}
/* La burbuja de Yve tambien en el movil, como en el PC: asi no ocupa sitio en
   la barra de arriba, que es lo que hacia que no cupiera todo. Un poco mas
   pequeña y pegada a la esquina para no taparle nada al contenido. */
@media(max-width:768px){
  #chat-fab{bottom:16px;right:12px;padding:10px 14px;font-size:12px}
  #chat-fab span{display:none}
  #chat-fab::after{content:'Yve';font-weight:800}
}

#chat-panel{
  position:fixed;top:0;right:0;width:420px;height:100%;
  background:var(--bg);border-left:1px solid var(--s1);
  z-index:1200;display:flex;flex-direction:column;
  transform:translateX(100%);
  transition:transform .3s cubic-bezier(.4,0,.2,1);
  box-shadow:-8px 0 40px rgba(0,0,0,.5);
}
#chat-panel.open{transform:translateX(0)}
@media(max-width:768px){
  #chat-panel{
    width:100%;height:100%;
    top:0;left:0;right:0;bottom:0;
    border-left:none;
    transform:translateY(100%);
  }
  #chat-panel.open{transform:translateY(0)}
}

#chat-header{
  padding:16px 18px;border-bottom:1px solid var(--s1);
  display:flex;align-items:center;justify-content:space-between;
  flex-shrink:0;background:var(--bg);
}
.chat-title{display:flex;align-items:center;gap:12px}
.chat-title span{font-size:1.8rem;line-height:1}
.chat-title h3{font-size:1rem;font-weight:700;color:#f1f5f9;margin:0}
.chat-title p{font-size:.75rem;color:#64748b;margin:0;margin-top:2px}
#chat-close{
  background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);
  color:#94a3b8;width:34px;height:34px;border-radius:50%;
  cursor:pointer;font-size:1.1rem;display:flex;align-items:center;justify-content:center;
  transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;flex-shrink:0;
}
#chat-close:hover{background:rgba(239,68,68,.15);border-color:rgba(239,68,68,.3);color:#f87171}

#chat-msgs{
  flex:1;overflow-y:auto;padding:16px;
  display:flex;flex-direction:column;gap:12px;
  scroll-behavior:smooth;
}
#chat-msgs::-webkit-scrollbar{width:4px}
#chat-msgs::-webkit-scrollbar-thumb{background:var(--s2);border-radius:2px}

.msg-user{align-self:flex-end;max-width:80%;background:var(--acc);color:#fff;border-radius:16px 16px 4px 16px;padding:10px 14px;font-size:.88rem;line-height:1.5}
.msg-bot{align-self:flex-start;max-width:88%;background:var(--s1);color:#e2e8f0;border-radius:16px 16px 16px 4px;padding:12px 14px;font-size:.88rem;line-height:1.6}
.msg-bot .msg-label{font-size:.7rem;font-weight:700;color:var(--acc2);margin-bottom:6px;display:block;letter-spacing:.04em}
.msg-typing{align-self:flex-start;background:var(--s1);border-radius:16px;padding:12px 16px;display:flex;gap:5px}
.dot-pulse{width:7px;height:7px;border-radius:50%;background:var(--acc2);animation:dotPulse 1.2s infinite}
.dot-pulse:nth-child(2){animation-delay:.2s}
.dot-pulse:nth-child(3){animation-delay:.4s}
@keyframes dotPulse{0%,80%,100%{opacity:.2;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}

#chat-suggestions{
  padding:0 16px 12px;display:flex;flex-wrap:wrap;gap:7px;flex-shrink:0;
}
.sug{
  background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2);
  color:var(--acc3);border-radius:20px;padding:7px 13px;font-size:.78rem;
  cursor:pointer;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;white-space:nowrap;font-weight:500;
}
.sug:hover{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.2);border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)}

#chat-input-row{
  padding:12px 14px;border-top:1px solid var(--s1);
  display:flex;gap:10px;align-items:flex-end;flex-shrink:0;background:var(--bg);
  padding-bottom:max(12px, env(safe-area-inset-bottom));
}
#chat-input{
  flex:1;background:var(--s1);border:1px solid var(--s2);color:#f1f5f9;
  border-radius:20px;padding:11px 16px;font-size:.9rem;outline:none;
  resize:none;font-family:inherit;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;max-height:120px;
  line-height:1.5;
}
#chat-input:focus{border-color:var(--acc2);box-shadow:0 0 0 2px rgba(var(--acc-r,96),var(--acc-g,165),var(--acc-b,250),.12)}
#chat-input::placeholder{color:#475569}
#chat-send{
  background:linear-gradient(135deg,#7c3aed,var(--acc));border:none;
  color:#fff;border-radius:50%;width:42px;height:42px;cursor:pointer;
  font-size:1.1rem;flex-shrink:0;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;
  display:flex;align-items:center;justify-content:center;
}
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
.drr-upload label{margin:0;padding:10px 18px;background:linear-gradient(135deg,var(--acc),var(--acc-dark));color:#fff;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;white-space:nowrap}
.drr-upload label:hover{box-shadow:0 0 20px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)}
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

/* ── DRR rediseño (fase presentación) — reorganización por grupos ───────
   Solo maquetación y tamaños. El acento de las burbujas .drr-mc lo pone la
   regla global `body.acentuar-todo` (ahi se anadio .drr-mc), igual que .card:
   asi todas las burbujas del panel responden al color personalizado. */
.rd-actions{display:flex;justify-content:flex-end;gap:8px;align-items:center;margin-bottom:20px}
.rd-group{margin-bottom:18px}
.rd-glabel{font-size:10px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:var(--mut);margin:0 2px 10px}
.rd-prov{font-size:15px;color:var(--grn);font-weight:600;display:flex;align-items:center;gap:8px;line-height:1.5}
.rd-grid{display:grid;gap:12px}
.rd-grid.rate{grid-template-columns:1.3fr 1fr 1fr 1fr}
.rd-grid.rev{grid-template-columns:1.3fr 1fr 2fr}
.rd-grid.gop{grid-template-columns:1.3fr 1fr 2fr}
@media(max-width:900px){.rd-grid.rate,.rd-grid.rev,.rd-grid.gop{grid-template-columns:1fr 1fr}}
@media(max-width:560px){.rd-grid.rate,.rd-grid.rev,.rd-grid.gop{grid-template-columns:1fr}}
.rd-tile .rd-hero{font-size:21px;font-weight:800;letter-spacing:-.02em;line-height:1.15;color:var(--tx)}
.rd-tile.hero .rd-hero{font-size:26px;color:var(--acc2)}
.rd-tile.hero.grn .rd-hero{color:var(--grn)}
.rd-hero .per{font-size:10px;font-weight:600;color:var(--dim);margin-left:6px;letter-spacing:.4px}
.rd-mini{display:flex;gap:16px;margin-top:10px;padding-top:9px;border-top:1px solid var(--s2);flex-wrap:wrap}
.rd-mini>div{font-size:10.5px;color:var(--dim)}
.rd-mini b{display:block;font-size:12px;color:var(--tx);font-weight:700;margin-top:1px}
.rd-bud{display:flex;flex-direction:column;justify-content:center}
.rd-daywrap{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:10px}
.rd-daycard{background:var(--bg);border:1px solid var(--s2);border-radius:12px;padding:12px 14px}
.rd-daycard.oob{border-color:rgba(239,68,68,.4);background:rgba(239,68,68,.05)}
.rd-dc-top{font-size:11px;color:var(--mut);margin-bottom:6px}
.rd-dc-amt{font-size:18px;font-weight:800;color:var(--tx);letter-spacing:-.01em}
.rd-dc-sub{font-size:10px;color:var(--dim);margin-top:1px}
.rd-dc-st{font-size:11.5px;font-weight:700;margin-top:8px}
.rd-dc-st.ok{color:var(--grn)} .rd-dc-st.oob{color:var(--red)}


/* ── GUIDED TOUR ─────────────────────────────────── */
#tour-overlay{
  position:fixed;inset:0;z-index:9000;pointer-events:none;
  background:transparent;transition:background .35s;
}
#tour-overlay.active{pointer-events:all;background:rgba(4,9,20,.72)}

@keyframes spotGlow{
  0%,100%{border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.55);box-shadow:0 0 0 9999px rgba(4,9,20,.72),0 0 0 4px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.12)}
  50%{border-color:rgba(96,165,250,.9);box-shadow:0 0 0 9999px rgba(4,9,20,.72),0 0 0 6px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.22)}
}

#tour-card.entering{animation:cardEnter .28s cubic-bezier(.2,.8,.2,1)}
@keyframes cardEnter{from{opacity:0;transform:translateY(8px) scale(.97)}to{opacity:1;transform:none}}

.tour-content-wrap
#tour-card h3{font-size:15px;font-weight:800;margin-bottom:8px;color:var(--tx);letter-spacing:-.2px}
#tour-card p{font-size:13.5px;color:var(--mut);line-height:1.65;margin-bottom:0}









.tour-btn-skip:hover{color:var(--mut)}

.tour-btn-prev:hover{background:var(--s2);color:var(--tx)}

.tour-btn-next:hover{box-shadow:0 0 22px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.6);transform:translateY(-1px)}
.tour-target{outline:2px solid var(--acc)!important;outline-offset:4px!important;border-radius:10px!important;animation:tourPulse 2s ease-in-out infinite!important;position:relative!important;z-index:9001!important}
@keyframes tourPulse{0%,100%{outline-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.6)}50%{outline-color:rgba(96,165,250,1)}}
@media(max-width:600px){
  
}

/* ── F&B Sub-tabs ─────────────────────────────────── */
.fb-sub{background:none;border:none;color:var(--mut);padding:7px 14px;border-radius:7px;cursor:pointer;font-size:13px;font-weight:500;font-family:inherit;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s;white-space:nowrap}
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
  .nav{height:calc(54px + env(safe-area-inset-top));padding:env(safe-area-inset-top) 12px 0 12px;gap:8px;box-sizing:border-box}
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

  
  #cal-hoteles{grid-template-columns:1fr!important}
  

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
.menu{display:none;position:absolute;top:46px;right:0;background:var(--s1);border:1px solid var(--s2);border-radius:11px;padding:7px;z-index:1000;min-width:218px;box-shadow:0 12px 40px rgba(0,0,0,.45);max-height:calc(100vh - 60px);overflow-y:auto}
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
[data-tip]::after{content:attr(data-tip);position:absolute;bottom:calc(100% + 8px);left:50%;transform:translateX(-50%);background:var(--s1);color:#f1f5f9;padding:7px 12px;border-radius:8px;font-size:11px;line-height:1.4;max-width:240px;white-space:normal;text-align:center;border:1px solid var(--s2);pointer-events:none;opacity:0;transition:opacity .15s;z-index:5000;box-shadow:0 4px 12px rgba(0,0,0,.3)}
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
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&display=swap" rel="stylesheet">
<style id="yve-icon-css">/*yve-icons-v1*/
svg.yvi{width:1em;height:1em;vertical-align:-0.125em;flex-shrink:0;display:inline-block}
</style>
</head>
<body>
<!-- ── Pantalla de inicio (splash) al abrir la app — no saltable, precarga recursos ── -->
<style>
#yve-splash{position:fixed;inset:0;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;
  background:linear-gradient(180deg,#101a2e 0%,#0c1424 55%,#090e1a 100%);padding:24px;
  transition:opacity .55s ease,visibility .55s ease}
#yve-splash.hide{opacity:0;visibility:hidden;pointer-events:none}
#yve-splash .sp-logo{width:110px;height:110px;border-radius:27px;box-shadow:0 22px 60px rgba(0,0,0,.55);animation:spPop .6s cubic-bezier(.2,.8,.2,1)}
#yve-splash .sp-brand{font-family:'Space Grotesk','Inter',sans-serif;margin-top:24px;font-size:31px;font-weight:700;letter-spacing:-.4px;color:#fff;animation:spFade .6s ease .12s both}
#yve-splash .sp-brand span{color:var(--acc2,#60a5fa)}
#yve-splash .sp-sub{margin-top:9px;font-size:13px;color:#94a3b8;animation:spFade .6s ease .22s both}
#yve-splash .sp-loader{margin-top:30px;width:32px;height:32px;border-radius:50%;border:3px solid rgba(148,163,184,.22);border-top-color:var(--acc,#3b82f6);animation:spSpin .8s linear infinite}
@keyframes spPop{from{opacity:0;transform:scale(.82)}to{opacity:1;transform:scale(1)}}
@keyframes spFade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
@keyframes spSpin{to{transform:rotate(360deg)}}
/* ── Arreglos responsive móvil ── */
@media(max-width:480px){
  .stats,#stats-ap-grid,#ar-real-stats{grid-template-columns:repeat(2,1fr)!important;gap:8px!important}
}
@media(max-width:768px){
  .panel > div[style*="display:flex"]{flex-wrap:wrap!important}
  .panel > div[style*="flex-end"]{justify-content:flex-start!important}
  #ar-real-grid{grid-template-columns:1fr!important}
  #panel-ar_real{overflow-x:hidden}
  div:has(> .fb-sub){overflow:visible!important;flex-wrap:nowrap!important}
  .fb-sub{flex:1 1 0!important;min-width:0;font-size:12.5px!important;padding:10px 4px!important;text-align:center;white-space:normal!important;line-height:1.15}
}
</style>
<div id="yve-splash" role="status" aria-label="Cargando Yve.01">
  <img class="sp-logo" src="/static/icons/yve-logo-192.png" alt="Yve.01">
  <div class="sp-brand">Yve<span>.01</span></div>
  <div class="sp-sub">Automatización financiera para hoteles</div>
  <div class="sp-loader"></div>
</div>
<script>
(function(){
  var sp=document.getElementById('yve-splash'); if(!sp) return;
  var shown=false;
  try{ shown=sessionStorage.getItem('yve_splash_shown')==='1'; }catch(e){}
  if(shown){ if(sp.parentNode) sp.parentNode.removeChild(sp); return; }
  try{ sessionStorage.setItem('yve_splash_shown','1'); }catch(e){}
  // Precargar la traducción del idioma guardado mientras se ve el splash
  try{ var lang=localStorage.getItem('yve_lang'); if(lang && lang!=='es'){ fetch('/static/i18n/'+lang+'.json?v=__ASSETS_V__').catch(function(){}); } }catch(e){}
  var start=Date.now(), MIN=1500, MAX=6000, done=false;
  // El servidor manda el HTML SIEMPRE en español (no sabe tu idioma), asi que
  // al entrar con otro idioma se veia español antes de traducir. El splash ya
  // dura 1,5 s: se aprovecha para traducir DEBAJO y no se suelta hasta que la
  // primera pasada ha terminado.
  function ocultar(){
    if(done) return; done=true;
    try{ if(window._pintarYa) window._pintarYa(document.body); }catch(e){}
    sp.classList.add('hide'); setTimeout(function(){ if(sp.parentNode) sp.parentNode.removeChild(sp); }, 600); }
  function tryHide(){ var el=Date.now()-start; if(el>=MIN) ocultar(); else setTimeout(ocultar, MIN-el); }
  if(document.readyState==='complete') tryHide(); else window.addEventListener('load', tryHide);
  setTimeout(ocultar, MAX);
})();
</script>

<nav class="nav" id="app-header">
  <div class="logo">
    <div class="logo-dot"></div>
    <span class="logo-name">Yve<span style="color:var(--acc2)">.01</span></span>
    <span class="logo-tag">__HOTEL_TAG__</span>
  <span style="font-size:9px;color:#334155;margin-left:4px;font-weight:500">v1.5</span>
  </div>
  <div class="nav-mid"></div>
  <div id="demo-banner" style="display:none;position:fixed;top:0;left:0;right:0;z-index:8000;background:linear-gradient(90deg,#f59e0b,#d97706);color:#000;text-align:center;padding:6px 16px;font-size:13px;font-weight:700;letter-spacing:.3px">
    🎭 MODO DEMO · <span style="font-weight:400">Datos de ejemplo para demostración</span>
    <button onclick="toggleDemoMode()" style="margin-left:16px;background:rgba(0,0,0,.2);border:none;padding:3px 10px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:700">✕ Salir</button>
  </div>
  <div class="nav-right">
    <!-- DESKTOP: fecha + instalar + tema + atajos + usuario -->
    <span class="pill hide-mobile" id="date-pill">—</span>
    <button id="btn-install-pwa" onclick="if(_deferredInstall){_deferredInstall.prompt();}" style="display:none;background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.2);color:#22c55e;padding:4px 10px;border-radius:8px;font-size:11px;cursor:pointer">📲 Instalar</button>
    
    <select id="hotel-activo-sel" onchange="seleccionarHotelActivo(this.value)" style="display:none;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);color:var(--acc2);padding:4px 10px;border-radius:20px;font-size:11px;cursor:pointer;max-width:190px;outline:none" title="Hotel activo"></select>
    <span class="pill hide-mobile" style="color:var(--acc2)">👤 __USER_NAME__</span>

    <!-- El acceso a Yve vive en la burbuja flotante (#chat-fab), igual que en
         el PC. Estaba aqui duplicado solo para el movil y era justo lo que
         hacia que la barra no cupiera. -->





    <button class="btn-run" id="btn-run" onclick="openUploadModal()">
      <div class="spin" id="spin"></div>
      <span id="run-lbl" data-i18n="nav.procesar">⚡ Procesar Archivos</span>
    </button>

    <div class="dropdown">
      <button class="btn-ref" onclick="toggleMenu('main-menu')" title="Ajustes, idioma y administración" style="font-size:16px;line-height:1;padding:5px 11px">⚙️</button>
      <div id="main-menu" class="menu">
        <div class="menu-head" data-i18n="menu.reportes">Reportes</div>
        <a href="/api/reportes/diario" class="menu-item">📄 Diario</a>
        <a href="/api/reportes/semanal" class="menu-item">📊 Semanal</a>
        <a href="/api/reportes/mensual" class="menu-item">📈 Mensual</a>
        <a href="/api/reportes/ejecutivo.pdf" class="menu-item">🎯 Ejecutivo PDF</a>
        <a href="/api/reportes/consolidado.xlsx" class="menu-item">📊 Consolidado Excel</a>
        <div class="menu-sep"></div>
        <!-- Language switcher -->
        <div class="menu-head" data-i18n="nav.idioma" style="margin-top:0">Idioma</div>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:3px;padding:0 4px;margin-bottom:3px">
          <button class="lang-btn menu-item" data-lang="es" onclick="cambiarIdioma('es');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Español">ES</button>
          <button class="lang-btn menu-item" data-lang="en" onclick="cambiarIdioma('en');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="English">EN</button>
          <button class="lang-btn menu-item" data-lang="fr" onclick="cambiarIdioma('fr');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Français">FR</button>
          <button class="lang-btn menu-item" data-lang="de" onclick="cambiarIdioma('de');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Deutsch">DE</button>
        </div>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:3px;padding:0 4px;margin-bottom:8px">
          <button class="lang-btn menu-item" data-lang="it" onclick="cambiarIdioma('it');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Italiano">IT</button>
          <button class="lang-btn menu-item" data-lang="pt" onclick="cambiarIdioma('pt');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Português">PT</button>
          <button class="lang-btn menu-item" data-lang="ca" onclick="cambiarIdioma('ca');document.getElementById('main-menu').classList.remove('open')" style="text-align:center;padding:6px 2px;font-size:11px;font-weight:800;line-height:1;justify-content:center;display:flex;border-radius:7px;letter-spacing:.5px" title="Català">CAT</button>
          <button class="lang-btn menu-item" data-lang="es" style="visibility:hidden;pointer-events:none"></button>
        </div>
        <div class="menu-sep"></div>
        <div class="menu-head" data-i18n="menu.presentacion">Presentación</div>
        <button class="menu-item" data-i18n="nav.tour" onclick="startTour();document.getElementById('main-menu').classList.remove('open')">🎯 Tour guiado</button>
        <button class="menu-item" onclick="mostrarHistorialProcesado();document.getElementById('main-menu').classList.remove('open')">📋 Historial de procesado</button>
        <button class="menu-item" id="btn-demo" onclick="toggleDemoMode()"><span data-i18n="nav.demo">🎭 Demo Mode</span></button>
        <div class="menu-sep"></div>
        <div class="menu-head" data-i18n="menu.cambiarRol">Cambiar rol</div>
        <button class="menu-item" id="rol-btn">👤 Admin</button>
        <div id="rol-menu" class="rol-sub" style="display:none">
          <button class="menu-item" onclick="cambiarRol('admin')" data-i18n="rol.admin">🔑 Administrador</button>
          <button class="menu-item" onclick="cambiarRol('financial_controller')" data-i18n="rol.fc">💰 Controller Financiero</button>
          <button class="menu-item" onclick="cambiarRol('income_auditor')">📊 Income Auditor</button>
          <button class="menu-item" onclick="cambiarRol('fb_manager')">🍽️ Jefe F&B</button>
          <button class="menu-item" onclick="cambiarRol('jefe_otras')">🛠️ Jefe Servicios</button>
        </div>
        <div class="menu-sep"></div>
        <button class="menu-item" onclick="loadAll();document.getElementById('main-menu').classList.remove('open')">↻ Actualizar datos</button>
        <a href="/admin/" class="menu-item" style="display:__ADMIN_DISPLAY__" data-i18n="menu.admin">👥 Administración</a>
        <!-- Colores personalizados -->
        <button class="menu-item" onclick="_openColorPicker();document.getElementById('main-menu').classList.remove('open')" id="btn-color-picker">🎨 Personalizar colores</button>
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
  <div id="top-bar" style="position:fixed;top:0;left:0;height:2px;background:linear-gradient(90deg,var(--acc,#3b82f6),#a78bfa);z-index:9999;transition:width .3s ease;width:0"></div>

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
  <div id="exec-summary" style="display:none;padding:8px 16px;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.05);border-bottom:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1);font-size:12px;color:var(--mut);text-align:center">
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
    <button class="tab" onclick="switchTab('multi_hotel',this)" id="tab-multi-hotel" data-i18n="tab.multihotel">🏨 Multi-Hotel</button>
    <button class="tab" onclick="switchTab('cierre',this)" id="tab-cierre" data-i18n="tab.cierre">🧾 Cierre</button>
  </div>

  <div id="panel-ar" class="panel active">
    
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
        <div class="empty"><div class="ei">📂</div><p>Sin datos.<br>Pulsa ⚡ Procesar Archivos.</p></div>
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
          <tr><td colspan="11" class="empty"><p>Sin datos. Pulsa ⚡ Procesar Archivos para empezar.</p></td></tr>
        </tbody>
      </table>
    </div>
  </div>

    <!-- ── Reclamaciones OTA (loop de reclamación automática) ── -->
    <div id="ar-recl-section" style="margin-top:24px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <div style="display:flex;align-items:center;gap:8px">
          <div style="font-size:13px;font-weight:700;letter-spacing:.2px">&#128257; Reclamaciones OTA pendientes de aprobar</div>
          <span data-tip="Yve detecta comisiones cobradas por encima del contrato, redacta el email con IA y lo deja listo para que lo apruebes y env&iacute;es. Nada se env&iacute;a sin tu OK." style="cursor:help;color:var(--dim);font-size:12px">&#9432;</span>
        </div>
        <div id="ar-recl-resumen" style="font-size:12px;color:var(--mut)"></div>
      </div>
      <div id="ar-recl-list" style="display:flex;flex-direction:column;gap:12px">
        <div class="empty card" style="padding:20px;text-align:center;color:var(--dim);font-size:12px;border-style:dashed;border-radius:12px">
          Cuando Yve detecte comisiones cobradas por encima del contrato, aparecer&aacute;n aqu&iacute; para reclamar.
        </div>
      </div>
    </div>

  </div><!-- /panel-ar -->

  <!-- PANEL AP -->
  <div id="panel-ap" class="panel">
  <div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:14px"><a href="/api/exportar/ap" style="background:var(--acc);color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px" data-i18n="btn.downloadExcel" data-i18n="btn.downloadExcel">⬇️ Descargar Excel</a><a href="/aprobaciones-ap/" class="btn-ref" style="text-decoration:none" title="Abrir panel de aprobaciones AP" data-i18n="btn.aprobarAP">📲 Aprobar facturas AP</a>
        <button class="btn-ref" onclick="aprobarMatchOK()" style="font-size:12px" title="Aprueba automáticamente todas las facturas con 3-way match correcto">✅ Aprobar Match OK</button>
        <button class="btn-ref" id="btnOracle" onclick="procesarOracle()" style="font-size:12px" title="Genera los asientos de las facturas aprobadas">🔮 Contabilizar en Oracle</button><span id="oracle-modo-chip" style="display:none;font-size:11px;padding:3px 8px;border-radius:999px;background:rgba(245,158,11,.15);color:#f59e0b;border:1px solid rgba(245,158,11,.4);align-self:center"></span><a href="/api/oracle/export_excel" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="oracle.exportGl" title="GL_INTERFACE con los asientos que ha producido el pipeline (simulación o real)">⬇️ Asientos GL</a>
        <button class="btn-ref" onclick="filtrarAPPorEstado()" id="btn-filter-ap" style="font-size:12px">🔍 Filtrar</button>
        <select id="ap-estado-filter" onchange="filtrarAPPorEstado(this.value)" style="background:var(--s1);border:1px solid var(--s2);color:var(--tx);padding:6px 10px;border-radius:8px;font-size:12px;cursor:pointer">
          <option value="" data-i18n-opt="lbl.todos">Todos</option>
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

    <!-- Provisiones de cierre (Ola A): solo lectura, sale de lo que ya hay -->
    <div class="card" id="card-provisiones" style="margin-top:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <div class="card-title" style="margin:0" data-i18n="prov.titulo">Provisiones de cierre</div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <input type="month" id="prov-mes" onchange="loadProvisiones()" style="background:var(--s1);border:1px solid var(--s2);color:var(--tx);padding:6px 10px;border-radius:8px;font-size:12px">
          <a id="prov-descarga" href="/api/exportar/provisiones" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="prov.descargar">⬇️ Excel del cierre</a>
        </div>
      </div>
      <div id="prov-body" style="font-size:13px;color:var(--mut)"><div class="empty"><p data-i18n="prov.cargando">Calculando provisiones…</p></div></div>
    </div>

    <!-- Aging AP (Ola A): a quien debemos y desde cuando. Solo lectura. -->
    <div class="card" id="card-aging-ap" style="margin-top:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <div class="card-title" style="margin:0" data-i18n="aging.titulo">Antigüedad de la deuda (aging AP)</div>
        <a href="/api/exportar/aging_ap" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="aging.descargar">⬇️ Excel del aging</a>
      </div>
      <div id="aging-tramos" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:8px;margin-bottom:12px"></div>
      <div id="aging-body" style="font-size:13px;color:var(--mut)"><div class="empty"><p data-i18n="aging.cargando">Calculando antigüedad…</p></div></div>
    </div>

    <!-- Reclamar al proveedor (Ola A): factura rectificativa o abono. Nada sale sin "Aprobar y enviar". -->
    <div class="card" id="card-albaranes" style="margin-top:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <div class="card-title" style="margin:0" data-i18n="alb.titulo">Albaranes de entrega</div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <span id="alb-resumen" style="font-size:12px;color:var(--mut)"></span>
          <a href="/api/exportar/albaranes" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="alb.excel">⬇️ Excel</a>
        </div>
      </div>
      <div style="font-size:12px;color:var(--mut);margin-bottom:10px" data-i18n="alb.ayuda">Cada entrega con sus líneas y la factura con la que ha cruzado. Pulsa una fila para ver las líneas.</div>
      <div id="alb-list"></div>
    </div>
    <div class="card" id="card-recl-ap" style="margin-top:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <div class="card-title" style="margin:0" data-i18n="reclap.titulo">Reclamar al proveedor (rectificativa / abono)</div>
        <span id="ap-recl-resumen" style="font-size:12px;color:var(--mut)"></span>
      </div>
      <div id="ap-recl-list" style="display:flex;flex-direction:column;gap:12px"><div class="empty"><p data-i18n="reclap.cargando">Buscando facturas que reclamar…</p></div></div>
    </div>
  </div><!-- /panel-ap -->

  <!-- PANEL DRR -->
  <div id="panel-drr" class="panel">
    <!-- Acciones (sin titulo ni chips: decision de diseno). El estado va oculto,
         que lo leen la subida y el onboarding; el OOB va dentro del panel. -->
    <div class="rd-actions">
      <label for="drr-file-input" class="btn-run" style="cursor:pointer;font-size:13px;margin:0" data-i18n="btn.uploadDrr">📂 Subir DRR</label>
      <input type="file" id="drr-file-input" accept=".xlsm,.xlsx" style="display:none" onchange="uploadDRR(this)">
      <a href="/api/exportar/drr" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="btn.downloadExcel">⬇️ Excel</a>
      <a href="/api/exportar/drr/pdf" class="btn-ref" style="text-decoration:none;font-size:12px">📄 PDF</a>
    </div>
    <span class="drr-status" id="drr-status" style="display:none" data-i18n="drr.sinArchivo">Sin archivo cargado</span>
    <span id="drr-oob-badge" style="display:none"></span>
    <!-- Cuerpo: lo pinta renderDRR. #drr-metrics vive aqui (ancla del tour) y lo
         reconstruye renderDRR con los tres grupos. Al arrancar, la zona de subida. -->
    <div id="drr-body">
      <div class="drr-metrics" id="drr-metrics" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px">
        <div class="empty" id="drr-drop-zone"
      style="border:2px dashed var(--s3);border-radius:12px;padding:32px;cursor:pointer;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s"
      ondragover="event.preventDefault();this.style.borderColor='var(--acc)';this.style.background='rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.05)'"
      ondragleave="this.style.borderColor='var(--s3)';this.style.background=''"
      ondrop="event.preventDefault();this.style.borderColor='var(--s3)';this.style.background='';uploadDRR({files:event.dataTransfer.files})"
      onclick="document.getElementById('drr-file-input').click()">
    <div class="ei">📊</div>
    <p style="margin-bottom:6px">Arrastra tu DRR aquí o</p>
    <p style="font-size:12px;color:var(--acc2);font-weight:600" data-i18n="drr.hazClic">haz clic para seleccionar (.xlsm/.xlsx)</p>
  </div>
      </div>
    </div>
  </div><!-- /panel-drr -->

  <!-- PANEL BANCO -->
  <div id="panel-banco" class="panel">
    <!-- Modo del banco (grupo / por hotel), elegido por el usuario. La etiqueta
         y el enlace de "cambiar" los rellena loadBanco cuando ya hay elección. -->
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:14px">
      <span id="banco-modo-chip" style="display:none;font-size:11.5px;font-weight:600;padding:5px 12px;border-radius:999px;background:var(--s1);border:1px solid var(--s2);color:var(--mut)"></span>
      <a id="banco-modo-cambiar" href="javascript:void(0)" onclick="abrirModoBanco()" style="display:none;font-size:11.5px;color:var(--acc2);text-decoration:none">⚙ Cambiar cómo funciona el banco</a>
    </div>
    <div id="banco-progress-bar" style="display:none;margin-bottom:14px"></div>
    <div class="stats lite-visible" id="banco-stats">
      <div class="sc hl c-blu"><div class="sc-lbl" data-i18n="sc.movimientos">Movimientos</div><div class="sc-val" id="bk-total" data-tip="Movimientos totales en el extracto">—</div><div class="sc-sub" data-i18n="sc.delExtracto">del extracto</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-i18n="sc.conciliados">Conciliados</div><div class="sc-val" id="bk-conc" data-tip="Cruzados con factura en el sistema">—</div><div class="sc-sub" data-i18n="sc.conFactura">con factura</div></div>
      <div class="sc c-ora"><div class="sc-lbl" data-i18n="sc.pendientes">Pendientes</div><div class="sc-val" id="bk-pend" data-tip="Sin factura asociada — pendientes">—</div><div class="sc-sub" id="bk-imp-pend">—</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-i18n="sc.diferencias">Diferencias</div><div class="sc-val" id="bk-diff">—</div><div class="sc-sub" data-i18n="sc.importeNoCuadra">importe no cuadra</div></div>
    </div>
    <div class="card">
      <div class="card-title" data-i18n="card.alertasBanco">Alertas Bancarias</div>
      <div id="bk-alertas"><div class="empty"><p>—</p></div></div>
    </div>
    <div style="margin-top:16px;display:flex;flex-wrap:wrap;gap:8px;align-items:center">
      <a href="/api/exportar/banco" class="btn-ref" style="text-decoration:none;font-size:12px">⬇️ Excel</a>
        <a href="/api/exportar/asientos" class="btn-ref" style="text-decoration:none;font-size:12px;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15);border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4);color:var(--acc2)" title="Exportar Libro Diario para A3, Sage, Holded...">📒 Libro Diario</a>
      <button class="btn-run" onclick="runConciliacion()" style="font-size:12px">⚡ Conciliar</button>
      <a href="/conciliacion/" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="btn.verConciliacion">🏦 Ver conciliación</a>
    </div>

    <!-- Modal de primera vez: ¿cómo funciona el banco de esta empresa? Resuelve
         el nº10 convirtiéndolo en una elección (grupo vs por hotel), no un bug.
         La elección se guarda en el servidor (config_banco.json) y gobierna todo. -->
    <div id="modal-banco-config" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:9000;align-items:center;justify-content:center;padding:20px">
      <div class="card" style="max-width:600px;width:100%">
        <div style="font-size:18px;font-weight:800;margin-bottom:6px">🏦 ¿Cómo funciona el banco de tu empresa?</div>
        <div style="font-size:13px;color:var(--mut);margin-bottom:20px;line-height:1.6">Un extracto bancario no dice a qué hotel es cada movimiento, así que lo eliges tú una vez. Se guarda en tu empresa (funciona igual desde el móvil y el PC) y puedes cambiarlo luego.</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <button type="button" onclick="elegirModoBanco('grupo')" style="text-align:left;background:var(--s1);border:1px solid var(--s2);border-radius:12px;padding:16px;cursor:pointer;color:var(--tx);transition:border-color .15s,background-color .15s" onmouseover="this.style.borderColor='var(--acc)'" onmouseout="this.style.borderColor='var(--s2)'">
            <div style="font-size:15px;font-weight:700;margin-bottom:6px">🏛️ Una cuenta del grupo</div>
            <div style="font-size:12px;color:var(--mut);line-height:1.5">Una sola cuenta para toda la empresa. El banco se muestra junto, igual en todos los hoteles.</div>
          </button>
          <button type="button" onclick="elegirModoBanco('por_hotel')" style="text-align:left;background:var(--s1);border:1px solid var(--s2);border-radius:12px;padding:16px;cursor:pointer;color:var(--tx);transition:border-color .15s,background-color .15s" onmouseover="this.style.borderColor='var(--acc)'" onmouseout="this.style.borderColor='var(--s2)'">
            <div style="font-size:15px;font-weight:700;margin-bottom:6px">🏨 Cada hotel su cuenta</div>
            <div style="font-size:12px;color:var(--mut);line-height:1.5">Cada hotel tiene su cuenta. El banco se separa por hotel; subes cada extracto dentro de su hotel.</div>
          </button>
        </div>
        <div style="text-align:right;margin-top:16px"><a href="javascript:void(0)" onclick="cerrarModoBanco()" id="banco-modal-cancelar" style="display:none;font-size:12px;color:var(--mut);text-decoration:none">Cancelar</a></div>
      </div>
    </div>
  </div><!-- /panel-banco -->

  <!-- PANEL NOTIFICACIONES -->
  <div id="panel-notif" class="panel">
    <!-- Banner de estado SMTP -->
    <div id="notif-smtp-banner" style="margin-bottom:16px;display:none"></div>
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

  <!-- PANEL CIERRE DE MES (Ola B). Solo lectura: asientos del mes + reconciliacion. -->
  <div id="panel-cierre" class="panel">
    <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <label for="cierre-mes" style="font-size:12px;color:var(--mut)" data-i18n="cierre.mes">Mes</label>
        <input type="month" id="cierre-mes" onchange="loadCierre(true)" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:6px 8px;border-radius:8px;font-size:12px">
        <span id="cierre-hotel" style="font-size:12px;color:var(--dim)"></span>
      </div>
      <a id="cierre-excel" href="/api/exportar/cierre" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="cierre.descargar">⬇️ Excel del cierre</a>
    </div>
    <div class="stats" id="cierre-stats" style="margin-bottom:14px">
      <div class="stat"><div class="stat-label" data-i18n="cierre.kAsientos">Asientos</div><div class="stat-value" id="cierre-k-asientos">—</div></div>
      <div class="stat"><div class="stat-label" data-i18n="cierre.kDebe">Debe</div><div class="stat-value" id="cierre-k-debe">—</div></div>
      <div class="stat"><div class="stat-label" data-i18n="cierre.kHaber">Haber</div><div class="stat-value" id="cierre-k-haber">—</div></div>
      <div class="stat"><div class="stat-label" data-i18n="cierre.kCuadre">Cuadre</div><div class="stat-value" id="cierre-k-cuadre">—</div></div>
    </div>
    <div class="card" id="card-cierre-paquete" style="margin-bottom:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="paq.titulo">Archivo de fin de mes para la central</div>
        <div style="display:flex;align-items:center;gap:8px"><span id="paq-estado" style="font-size:12px;font-weight:700"></span>
        <a id="paq-excel" href="/api/exportar/cierre_paquete" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="paq.descargar">📦 Descargar paquete</a></div>
      </div>
      <div id="paq-resultado" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-bottom:12px"></div>
      <div id="paq-body" style="font-size:12px;color:var(--mut)"><div class="empty"><p data-i18n="cierre.cargando">Montando el mes…</p></div></div>
    </div>
    <div class="card" id="card-cierre-recon" style="margin-bottom:22px">
      <div class="card-title" data-i18n="cierre.recon">Reconciliación de cuentas</div>
      <div id="cierre-recon-body" style="font-size:13px;color:var(--mut)"><div class="empty"><p data-i18n="cierre.cargando">Montando el mes…</p></div></div>
    </div>
    <div class="card" id="card-cierre-banco" style="margin-bottom:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="cbanco.titulo">Cuadre de banco por pestañas</div>
        <div style="display:flex;align-items:center;gap:10px"><span id="cbanco-saldo" style="font-size:12px;color:var(--mut)"></span>
        <a id="cbanco-excel" href="/api/exportar/cuadre_banco" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="cbanco.descargar">⬇️ Excel</a></div>
      </div>
      <div id="cbanco-pestanas" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-bottom:12px"></div>
      <div id="cbanco-body" style="font-size:12px;color:var(--mut);overflow-x:auto"><div class="empty"><p data-i18n="cierre.cargando">Montando el mes…</p></div></div>
    </div>
    <div class="card" id="card-cierre-inv" style="margin-bottom:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="inv.titulo">Inventarios de cierre</div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <a id="inv-hoja" href="/api/inventarios/hoja_recuento" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="inv.hoja">📋 Hoja de recuento</a>
          <label for="inv-file" class="btn-ref" style="cursor:pointer;font-size:12px;margin:0" data-i18n="inv.subir">📤 Subir recuento</label><input type="file" id="inv-file" accept=".xlsx,.xls,.csv" style="display:none" onchange="_invSubir(this)">
          <a id="inv-excel" href="/api/exportar/inventarios" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="inv.descargar">⬇️ Excel</a>
        </div>
      </div>
      <div id="inv-resumen" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-bottom:12px"></div>
      <div id="inv-body" style="font-size:12px;color:var(--mut);overflow-x:auto"><div class="empty"><p data-i18n="cierre.cargando">Montando el mes…</p></div></div>
    </div>
    <div class="card" id="card-cierre-fiscal" style="margin-bottom:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="fis.titulo">Fiscal: IVA 303, 349 y SII</div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap"><span id="fis-estado" style="font-size:12px;font-weight:700"></span>
          <a id="fis-excel" href="/api/exportar/fiscal" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="fis.descargar">⬇️ Excel</a>
        </div>
      </div>
      <div style="font-size:12px;color:var(--mut);margin-bottom:10px" data-i18n="fis.ayuda">Preparado a partir de los mismos datos que los asientos del mes. Nada se envía a Hacienda: el envío del SII exige certificado digital y lo hace la gestoría.</div>
      <div id="fis-resumen" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:12px"></div>
      <div id="fis-body"></div>
    </div>
    <div class="card" id="card-cierre-inm" style="margin-bottom:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="inm.titulo">Inmovilizado y amortizaciones</div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <button class="btn-ref" style="font-size:12px" onclick="_inmForm()" data-i18n="inm.alta">➕ Dar de alta</button>
          <a id="inm-excel" href="/api/exportar/inmovilizado" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="inm.descargar">⬇️ Excel</a>
        </div>
      </div>
      <div id="inm-form" style="display:none;margin-bottom:12px;padding:10px;border:1px dashed var(--s2);border-radius:10px">
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:8px">
          <input id="inm-desc" placeholder="Descripción" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px">
          <select id="inm-cat" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px"></select>
          <input id="inm-fecha" type="date" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px">
          <input id="inm-coste" type="number" step="0.01" placeholder="Coste (sin IVA)" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px">
          <input id="inm-vida" type="number" step="0.5" placeholder="Vida útil (años)" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px">
          <input id="inm-doc" placeholder="Nº factura" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:7px;border-radius:8px;font-size:12px">
        </div>
        <div style="margin-top:8px;display:flex;gap:8px"><button class="btn-run" style="font-size:12px" onclick="_inmGuardar()" data-i18n="inm.guardar">Guardar</button><button class="btn-ref" style="font-size:12px" onclick="document.getElementById('inm-form').style.display='none'" data-i18n="inm.cancelar">Cancelar</button></div>
      </div>
      <div id="inm-resumen" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-bottom:12px"></div>
      <div id="inm-body" style="font-size:12px;color:var(--mut);overflow-x:auto"><div class="empty"><p data-i18n="cierre.cargando">Montando el mes…</p></div></div>
    </div>
    <div class="card" id="card-cierre-mayor" style="margin-bottom:22px">
      <div class="card-title" data-i18n="cierre.mayor">Mayor del mes (por cuenta)</div>
      <div id="cierre-mayor-body" style="font-size:13px;color:var(--mut)"></div>
    </div>
    <div class="card" id="card-cierre-diario">
      <div class="card-title" data-i18n="cierre.diario">Libro Diario del mes</div>
      <div id="cierre-avisos" style="font-size:12px;color:#f59e0b;margin-bottom:8px"></div>
      <div id="cierre-diario-body" style="font-size:12px;color:var(--mut);overflow-x:auto"></div>
    </div>
  </div><!-- /panel-cierre -->

  <!-- PANEL F&B -->

  <div id="panel-fb" class="panel">
    <!-- F&B Sub-tabs -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:10px">
      <div id="fb-subtabs" style="display:flex;gap:4px;background:var(--s1);border-radius:10px;padding:4px;border:1px solid var(--s2);overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none">
        <button class="fb-sub active" onclick="fbSub('resumen',this)" data-i18n="fb.resumen">📊 Resumen</button>
        <button class="fb-sub" onclick="fbSub('inventario',this)" data-i18n="fb.inventario">📦 Inventario</button>
        <button class="fb-sub" onclick="fbSub('mermas',this)" data-i18n="fb.mermas">⚠️ Mermas</button>
        <button class="fb-sub" onclick="fbSub('recetas',this)" data-i18n="fb.recetas">📋 Recetas</button>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <label for="fb-upload-input" class="btn-ref" style="cursor:pointer;font-size:12px" data-i18n="btn.importarPos">📤 Importar ventas POS</label>
        <input type="file" id="fb-upload-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="fbUploadPOS(this)">
        <label for="fb-rec-input" class="btn-ref" style="cursor:pointer;font-size:12px" data-i18n="btn.importarRecetario">📋 Importar recetario</label>
        <input type="file" id="fb-rec-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="fbUploadRecetas(this)">
        <a href="/api/exportar/fb/pdf" class="btn-ref" style="text-decoration:none;font-size:12px">📄 PDF</a>
      </div>
    </div>
    <div id="fb-upload-msg" style="font-size:12px;margin-bottom:12px;min-height:16px"></div>
    <div id="fb-resumen" class="lite-visible"><div class="empty"><p>Cargando...</p></div></div>
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
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:16px" id="ar-real-stats" class="lite-visible">
      <div class="sc hl c-ora"><div class="sc-lbl" data-tip="Facturas emitidas aún no cobradas">PENDIENTE COBRO</div><div class="sc-val" id="arp-pendiente">—</div></div>
      <div class="sc c-red"><div class="sc-lbl" data-tip="Facturas con más de 60 días sin cobrar">VENCIDO >60d</div><div class="sc-val" id="arp-vencido">—</div></div>
      <div class="sc c-grn"><div class="sc-lbl" data-tip="Cobrado este mes">COBRADO MES</div><div class="sc-val" id="arp-cobrado">—</div></div>
      <!-- Cuenta las altas de clientes_credito.xlsx, no los clientes que salen
           en las facturas: por eso la etiqueta dice de credito y no activos. -->
      <div class="sc"><div class="sc-lbl">CLIENTES DE CRÉDITO</div><div class="sc-val" id="arp-nclientes">—</div></div>
    </div>

    <!-- Two-column layout: clients + invoices -->
    <div style="display:grid;grid-template-columns:1fr 1.6fr;gap:16px" id="ar-real-grid">

      <!-- Client list -->
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;gap:8px">
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)">Clientes</div>
          <button onclick="abrirNuevoCliente()" class="btn-ref" data-i18n="ar.btnNuevoCliente" style="font-size:11px;padding:6px 10px;min-height:36px">➕ Nuevo cliente</button>
        </div>
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
          <table style="width:100%;min-width:520px;border-collapse:collapse;font-size:12px">
            <thead><tr style="border-bottom:2px solid var(--s2)">
              <th style="text-align:left;padding:8px;color:var(--mut);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px">Nº / Cliente</th>
              <th style="text-align:right;padding:8px;color:var(--mut);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px">Importe</th>
              <th style="text-align:center;padding:8px;color:var(--mut);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px">Días</th>
              <th style="text-align:left;padding:8px;color:var(--mut);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px">Estado</th>
              <th style="padding:8px;color:var(--mut);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px">Acciones</th>
            </tr></thead>
            <tbody id="ar-facturas-tbody">
            <tr><td colspan="8" class="empty" style="padding:32px;text-align:center;color:var(--dim)">
              <div style="font-size:24px;margin-bottom:8px">📋</div>
              <div style="font-weight:600;margin-bottom:4px">Sin facturas AR todavía</div>
              <div style="font-size:12px">Usa <b>Nueva factura</b> para emitir a clientes corporativos, grupos o agencias.</div>
            </td></tr>
          </tbody>
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

    <!-- Direct bill: la factura a credito contra el bono de la agencia (Ola A). Solo lectura. -->
    <div id="ar-bonos-section" class="card" style="margin-top:22px">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:10px">
        <div class="card-title" style="margin:0" data-i18n="bonos.titulo">Direct bill: factura vs bono de agencia</div>
        <div style="display:flex;align-items:center;gap:10px"><span id="ar-bonos-resumen" style="font-size:12px;color:var(--mut)"></span>
        <a href="/api/exportar/bonos" class="btn-ref" style="text-decoration:none;font-size:12px" data-i18n="bonos.descargar">⬇️ Excel</a></div>
      </div>
      <div id="ar-bonos-list" style="display:flex;flex-direction:column;gap:8px"><div class="empty"><p data-i18n="bonos.cargando">Cotejando bonos…</p></div></div>
    </div>

    <!-- BEOs generados automáticamente desde contratos -->
    <div id="ar-beos-section" style="margin-top:22px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)">BEOs desde contratos <span id="ar-beos-count" style="color:var(--dim)"></span></div>
        <span data-tip="Yve crea el BEO (partidas e importes) desde el contrato de grupo y coteja la factura contra él." style="cursor:help;color:var(--dim);font-size:12px">&#9432;</span>
      </div>
      <div id="ar-beos-list" style="display:flex;flex-direction:column;gap:10px">
        <div class="empty card" style="padding:20px;text-align:center;color:var(--dim);font-size:12px;border-style:dashed;border-radius:12px">
          Procesa un contrato de grupo en <b>Procesar Archivos</b> y aquí ver&aacute;s su BEO con el cotejo de la factura.
        </div>
      </div>
    </div>

    <div id="ar-real-status" style="display:none"></div>
  </div><!-- /panel-ar_real -->



  <div id="panel-multi_hotel" class="panel" style="overflow-x:hidden">

    <!-- Header -->
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;flex-wrap:wrap;gap:12px">
      <div>
        <h2 style="font-size:20px;font-weight:800;margin:0">🌍 Multi-Hotel Dashboard</h2>
        <div style="font-size:12px;color:var(--mut);margin-top:3px">Vista consolidada del grupo</div>
      </div>
      <!-- Month selector + perspective toggle + export -->
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <select id="mh-mes-select" onchange="_mh_loaded=false;loadMultiHotel()"
          style="background:var(--s1);border:1px solid var(--s2);color:var(--txt);padding:7px 10px;border-radius:9px;font-size:12px;cursor:pointer;outline:none">
          <option value="">Mes actual</option>
        </select>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
      </div>
        <!-- Perspectiva toggle -->
        <div style="display:inline-flex;background:var(--s1);border:1px solid var(--s2);border-radius:9px;padding:3px;gap:2px">
          <button id="mh-view-cards" onclick="setMHView('cards')" style="background:var(--acc2);color:#fff;border:none;padding:6px 13px;border-radius:6px;font-size:12px;cursor:pointer;font-weight:500;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s">📊 Resumen</button>
          <button id="mh-view-ranking" onclick="setMHView('ranking')" style="background:transparent;color:var(--mut);border:none;padding:6px 13px;border-radius:6px;font-size:12px;cursor:pointer;font-weight:500;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s">🏆 Ranking</button>
        </div>
        <a href="/api/exportar/multihotel" class="btn-ref" style="text-decoration:none;font-size:12px">⬇️ Excel</a>
      </div>
    </div>

    <!-- ═══ VISTA RESUMEN (cards + gráficos) ═══ -->
    <div id="mh-view-resumen">
      <!-- KPI Cards 2x2 -->
      <div id="mh-kpis" class="lite-visible" style="margin-bottom:20px"></div>

      <!-- Smart Insights row -->
      <div id="mh-insights" style="margin-bottom:20px"></div>

      <!-- Trend charts 2 col -->
      <div id="mh-trend-row" style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">
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
      <div id="mh-hotel-cards" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(290px,100%),1fr));gap:16px;overflow:hidden"></div>
    </div>

    <!-- ═══ VISTA RANKING (clásica: status + top performers + tabla) ═══ -->
    <div id="mh-view-clasica" style="display:none">
      <!-- Status cards -->
      <div id="mh-status" style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px"></div>

      <!-- Top performers + Alertas -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">
        <div class="card">
          <div class="card-title">🏆 Top Performers (RevPAR)</div>
          <div id="mh-rankings"></div>
        </div>
        <div class="card">
          <div class="card-title">⚠️ Alertas activas</div>
          <div id="mh-alertas"></div>
        </div>
      </div>

      <!-- Full table -->
      <div class="card">
        <div class="card-title">Todos los hoteles</div>
        <div class="tbl-wrap"><table style="width:100%"><thead><tr>
          <th>Hotel</th><th>Categoría</th><th style="text-align:right">Hab.</th>
          <th style="text-align:right">Ocup.</th><th style="text-align:right">ADR</th>
          <th style="text-align:right">RevPAR</th><th style="text-align:right">Revenue</th>
          <th style="text-align:right">GOP%</th><th style="text-align:center">Estado</th>
        </tr></thead><tbody id="mh-tbody-full"></tbody></table></div>
      </div>
    </div>

  </div><!-- /panel-multi_hotel -->

</div><!-- /main -->

<!-- MODAL PIPELINE -->
<!-- ── File Upload Modal ─────────────────────────────────────────────── -->
<div id="demo-setup-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9100;align-items:center;justify-content:center">
  <div style="background:var(--s1);border:1px solid rgba(245,158,11,.35);border-radius:20px;padding:26px;width:min(520px,95vw);max-height:85vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <h2 style="font-size:17px;font-weight:800;margin:0">🎭 Modo Demo personalizado</h2>
      <button onclick="document.getElementById('demo-setup-modal').style.display='none'" style="background:none;border:none;color:var(--mut);font-size:22px;cursor:pointer">×</button>
    </div>
    <div style="font-size:12.5px;color:var(--mut);line-height:1.6;margin-bottom:14px">Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.</div>
    <div style="font-size:11px;font-weight:700;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Una línea por cadena · formato "Cadena: Hotel 1, Hotel 2"</div>
    <textarea id="demo-setup-input" rows="4" placeholder="Hotel Miramar
Cadena Sol: Hotel Sol Mar, Hotel Sol Playa
Gestoría Nord: Hotel Pirineus, Hotel Vall" style="width:100%;background:var(--bg);border:1px solid var(--s2);color:var(--tx);border-radius:10px;padding:12px;font-size:13px;font-family:inherit;resize:vertical;outline:none"></textarea>
    <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:14px">
      <button onclick="document.getElementById('demo-setup-modal').style.display='none'" style="background:var(--s2);border:1px solid var(--s3);color:var(--tx);padding:9px 18px;border-radius:9px;font-size:13px;cursor:pointer">Cancelar</button>
      <button id="btn-demo-generar" onclick="generarDemo()" style="background:linear-gradient(135deg,#f59e0b,#d97706);border:none;color:#000;padding:9px 20px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer">🎭 Generar demo</button>
    </div>
  </div>
</div>
<div id="upload-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9000;align-items:center;justify-content:center">
  <div style="background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:28px;width:min(600px,95vw);max-height:85vh;overflow-y:auto;position:relative">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:22px">
      <div>
        <h2 style="font-size:18px;font-weight:800;margin:0">⚡ Procesar Archivos</h2>
        <div style="font-size:12px;color:var(--mut);margin-top:4px" data-i18n="upload.subtitulo">Facturas · Extractos bancarios · Ventas POS · Inventario · Mermas · Comisiones OTA · Rooming — clasificación automática por IA</div>
      </div>
      <button onclick="closeUploadModal()" style="background:none;border:none;color:var(--mut);font-size:24px;cursor:pointer">×</button>
    </div>
    <!-- Hotel de destino. Solo se ve con 2+ hoteles: con 0 o 1 no hay nada que
         elegir y meter una pregunta ahi seria fricción por nada. -->
    <div id="upload-hotel-row" style="display:none;margin-bottom:16px;padding:12px 14px;border-radius:12px;border:1px solid var(--s3);background:var(--bg)">
      <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <div style="font-size:12px;font-weight:700;color:var(--tx)">🏨 <span data-i18n="upload.hotelDestino">Hotel al que pertenece</span></div>
        <select id="upload-hotel-sel" onchange="_elegirHotelSubida(this.value)"
                style="flex:1;min-width:180px;background:var(--s1);border:1px solid var(--s3);color:var(--tx);padding:7px 10px;border-radius:9px;font-size:13px;cursor:pointer;outline:none"></select>
      </div>
      <div id="upload-hotel-aviso" style="display:none;font-size:11px;color:var(--ora);margin-top:8px;line-height:1.45"></div>
    </div>
    <div id="upload-drop-zone"
         onclick="document.getElementById('upload-file-input').click()"
         ondragover="event.preventDefault();this.style.borderColor='var(--acc)';this.style.background='rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.08)'"
         ondragleave="this.style.borderColor='var(--s3)';this.style.background=''"
         ondrop="handleUploadDrop(event)"
         style="border:2px dashed var(--s3);border-radius:14px;padding:32px;text-align:center;cursor:pointer;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s;margin-bottom:16px">
      <div style="font-size:36px;margin-bottom:10px">📂</div>
      <div style="font-size:15px;font-weight:600;color:var(--tx);margin-bottom:6px">Arrastra archivos aquí o haz clic</div>
      <div style="font-size:12px;color:var(--dim);margin-bottom:14px" data-i18n="upload.tiposZona">PDF · Excel · CSV · Fotos — clasificación automática</div>
      <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
        <button onclick="event.stopPropagation();document.getElementById('upload-file-input').click()"
                data-i18n="upload.selArchivos"
                style="background:var(--acc);border:none;color:#fff;padding:8px 18px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer">📄 Seleccionar archivos</button>
        <button class="show-mobile" onclick="event.stopPropagation();document.getElementById('upload-photo-input').click()"
                data-i18n="upload.selFotos"
                style="background:var(--s2);border:1px solid var(--s3);color:var(--tx);padding:8px 18px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer">📸 Hacer o elegir fotos</button>
        <button class="hide-mobile" onclick="event.stopPropagation();document.getElementById('upload-folder-input').click()"
                data-i18n="upload.selCarpeta"
                style="background:var(--s2);border:1px solid var(--s3);color:var(--tx);padding:8px 18px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer">📁 Seleccionar carpeta</button>
      </div>
    </div>
    <!-- DOS puertas, a proposito. Antes habia una sola con un `accept` que
         mezclaba documentos e `image/*`: en el movil eso hace que el sistema
         abra la GALERIA, asi que todo lo que se elegia acababa siendo una foto
         —y por eso salia "Foto" en todo y se ofrecia unir donde no tocaba. -->
    <input id="upload-file-input" type="file" multiple accept=".pdf,.xlsm,.xlsx,.xls,.csv,.jpg,.jpeg,.png,.webp,.heic" style="display:none" onchange="handleUploadFiles(this.files, this)">
    <input id="upload-photo-input" type="file" multiple accept="image/*" style="display:none" onchange="handleUploadFiles(this.files, this)">
    <input id="upload-folder-input" type="file" multiple webkitdirectory style="display:none" onchange="handleUploadFiles(this.files, this)">
    <!-- Nada se descarta en silencio. Va FUERA de #upload-file-list a
         proposito: si se descartan TODOS los ficheros la lista esta vacia y
         escondida, y es justo cuando mas falta hace el aviso. -->
    <div id="upload-aviso-descartes" style="display:none;margin-bottom:14px;padding:11px 13px;border-radius:11px;
         border:1px solid rgba(245,158,11,.35);background:rgba(245,158,11,.08);font-size:12px;line-height:1.5;color:var(--tx)"></div>
    <!-- Already uploaded files on server -->
    <div id="server-files-section" style="display:none;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">EN SERVIDOR (facturas-entrada)</div>
      <div id="server-files-list" style="max-height:120px;overflow-y:auto;display:flex;flex-direction:column;gap:5px"></div>
    </div>
    
    <div id="upload-file-list" style="display:none;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:var(--dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">ARCHIVOS SELECCIONADOS</div>
      <div id="upload-files-container" style="max-height:220px;overflow-y:auto;display:flex;flex-direction:column;gap:6px"></div>
      <!-- Agrupar fotos: se pinta solo cuando hay 2+ fotos que unir. Va FUERA
           del contenedor con scroll a proposito, para que el boton no se
           escape hacia arriba mientras se marcan fotos en una lista larga. -->
      <div id="upload-unir-bar" style="display:none;margin-top:10px"></div>
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
                style="background:var(--acc);border:none;color:#fff;padding:10px 22px;border-radius:10px;font-size:14px;font-weight:700;cursor:not-allowed;opacity:.4;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s">
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
    <div class="modal-f" style="display:flex;gap:10px;justify-content:flex-end">
      <button class="btn-cl" id="btn-cl" onclick="closeModalAndRefresh()" disabled>Cerrar</button>
      <button class="btn-cl" id="btn-retry" onclick="retryLastBatch()" style="display:none;background:#1db954;color:#fff;border:none">🔄 Reintentar</button>
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
    <button class="sug" onclick="askSug(this)">🌅 Briefing de hoy</button>
    <button class="sug" onclick="askSug(this)">⚠️ ¿Qué discrepancias tengo abiertas?</button>
    <button class="sug" onclick="askSug(this)">💰 ¿Cuánto puedo reclamar este mes?</button>
    <button class="sug" onclick="askSug(this)">📋 ¿Qué necesita mi firma hoy?</button>
  </div>
  <div id="chat-input-row">
    <textarea id="chat-input" rows="1" placeholder="Escribe aquí…"
      onkeydown="chatKeydown(event)" oninput="autoResize(this)"></textarea>
    <button id="chat-send" onclick="sendChat()">➤</button>
  </div>
</div>

<!-- GUIDED TOUR (tour-box created dynamically by startTour) -->
<div id="tour-overlay" style="display:none;position:fixed;inset:0;z-index:9900;pointer-events:none"></div>

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
    // Sin estos dos, el estado se pintaba con su nombre en crudo
    // ('SIN_TARIFA_HOTEL') o se confundia con una discrepancia reclamable.
    COBRO_POR_DEBAJO: ['b-unk', '↓ Cobrado por debajo'],
    SIN_TARIFA_HOTEL: ['b-unk', '? Sin tarifa del hotel'],
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
  if (!a || a === '') return '<span class="badge b-pen">· ' + (t('lbl.pendiente', 'Pendiente')) + '</span>';
  if (a === 'APROBADA')  return '<span class="badge b-apr">✓ ' + (t('lbl.aprobado', 'Aprobada')) + '</span>';
  if (a === 'RECHAZADA') return '<span class="badge b-rec">✗ ' + (t('lbl.rechazado', 'Rechazada')) + '</span>';
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
      // notif AR "todo en orden" quitada
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
  document.getElementById('status-txt').textContent = t('status.actualizando', 'Actualizando...');
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
        parts.push(stats.discrepancias + ' ' + (t('alert.discrepancias', 'discrepancia(s) reclamables')));
      if (stats.di_pendientes > 0)
        parts.push(stats.di_pendientes + ' ' + (t('alert.sinDI', 'factura(s) sin cert. DI')));
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
    // clock handled by _startClock()
      hoy.toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'});

    document.getElementById('status-txt').textContent =
      'Actualizado' + ' · ' + (stats.total || 0) + ' facturas cargadas';
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
    else { sumEl.style.display = 'none'; }
  }
  if (topBar) { topBar.style.width = '100%'; setTimeout(() => { topBar.style.opacity = '0'; setTimeout(() => { topBar.style.width = '0'; topBar.style.opacity = '1'; }, 300); }, 400); }

  // Recargar SIEMPRE todos los tabs con datos
  try {
    // AP — siempre recargar
    if (typeof cargarStatsAP === 'function') cargarStatsAP();
    if (typeof cargarFacturasAP === 'function') cargarFacturasAP();
    if (typeof cargarReclamacionesOTA === 'function') cargarReclamacionesOTA();
    // Tab activo extra
    var activePanel = document.querySelector('.panel.active');
    if (activePanel) {
      var pid = activePanel.id || '';
      if (pid === 'panel-drr' && typeof cargarDRR === 'function') cargarDRR();
      if (pid === 'panel-banco' && typeof cargarBanco === 'function') cargarBanco();
      // F&B no estaba en esta lista, y desde la fase 4b sus datos son POR
      // HOTEL: al cambiar de hotel con la pestaña de F&B abierta se quedaban
      // los numeros del hotel anterior. Mismo fallo que el de las tarjetas de
      // reclamacion, en otro panel. `_refrescarFB` marca los cuatro subtabs
      // como no cargados y repinta el que se este mirando.
      if (pid === 'panel-fb' && typeof _refrescarFB === 'function') _refrescarFB();
      if (pid === 'panel-multi_hotel') {
        window._mhGrupo = null;
        window._mhGrupoLabel = null;
        window._mhGrupoSub   = null;
        _mh_loaded = false;
        if (typeof loadMultiHotel === 'function') loadMultiHotel();
      }
    }
  } catch(e2) { console.warn('Error recargando tabs:', e2); }

  } catch(e) {
    console.error('Error en loadAll:', e);
    document.getElementById('status-txt').textContent = t('status.error', 'Error al cargar datos');
  }
  // Re-apply string map after all data has rendered
  if (_i18nLang && _i18nLang !== 'es') {
    setTimeout(function() { _applyStrMap(_i18nLang); }, 500);
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
  if (typeof Chart === 'undefined') { console.warn('Chart.js aún no cargado'); return; }
  // Mismo fallo que en las reclamaciones: sin datos se salia sin tocar nada y
  // el grafico del hotel ANTERIOR se quedaba pintado. Al cambiar a un hotel
  // sin facturas seguias viendo la barra de Booking del otro.
  if (!ch || !ch.labels || !ch.labels.length) {
    if (otaChart) { otaChart.destroy(); otaChart = null; }
    return;
  }
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
  document.getElementById('tbl-count').textContent = rows.length ? rows.length + ' ' + (t('lbl.registros', 'registros')) : '';
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="12" style="padding:32px;text-align:center"><div style="font-size:32px;margin-bottom:8px">📦</div><div style="font-weight:600;color:var(--mut);margin-bottom:4px">Sin facturas AP</div><div style="font-size:12px;color:var(--dim)">Pulsa ⚡ Procesar Archivos</div></td></tr>';
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
    el.innerHTML = '<div class="empty"><div class="ei">📂</div><p>Sin datos.<br>Pulsa ⚡ Procesar Archivos.</p></div>';
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
    // Las dos que antes no se contaban en ninguna linea: una factura sin
    // tarifa del hotel, o cobrada por debajo, no salia en el resumen.
    { dot:'o', n: c.COBRO_POR_DEBAJO     || 0, txt: 'cobradas por debajo de lo pactado', key:'res.porDebajo' },
    { dot:'m', n: (c.SIN_TARIFA_HOTEL||0) + (c.OTA_DESCONOCIDA||0), txt: 'sin tarifa pactada con la que comparar', key:'res.sinTarifa' },
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
  title.textContent = _tSSE('Pipeline AR — Procesando...');
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

    p.textContent = _tSSE(txt);
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      // Han entrado documentos: lo precargado ya no vale. Sin esto, "no
      // repintar al volver" dejaria Banco, DRR o Multi-Hotel con los numeros
      // de antes de procesar — un parpadeo cosmetico cambiado por datos
      // viejos, que es peor.
      try { if (typeof _invalidarPaneles === 'function') _invalidarPaneles(); } catch(e){}
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = _tSSE(ok ? 'Pipeline completado con éxito' : 'Pipeline finalizado con errores');
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = _tSSE('⚡ Procesar Archivos');
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
    lbl.textContent = _tSSE('⚡ Procesar Archivos');
    btnCl.disabled = false;
    icon.textContent = '⚠️';
    title.textContent = _tSSE('Error de conexión');

    const actionsDiv = document.createElement('div');
    actionsDiv.style.cssText = 'display:flex;gap:12px;margin-top:16px;justify-content:flex-end';
    actionsDiv.innerHTML =
      '<button onclick="closeModal()" style="background:transparent;border:1px solid #444;color:#aaa;padding:10px 16px;border-radius:8px;cursor:pointer;font-size:13px">Cerrar</button>' +
      '<button onclick="closeModal();setTimeout(runPipeline,300)" style="background:#1db954;border:none;color:#fff;padding:10px 16px;border-radius:8px;cursor:pointer;font-weight:600;font-size:13px">🔄 Reintentar</button>';
    log.appendChild(actionsDiv);
    btnCl.disabled = false;
  };
}

function closeModalAndRefresh() {
  closeModal();
  setTimeout(function() {
    loadAll();
    // Detectar dónde hay nuevos datos y cambiar al tab correcto
    fetch('/api/stats_ap').then(r=>r.json()).then(ap=>{
      fetch('/api/stats').then(r=>r.json()).then(ar=>{
        var apCount = ap.total || 0;
        var arCount = ar.total || 0;
        // Si hay más datos en AP que en AR, cambiar al tab AP
        if (apCount > arCount && apCount > 0) {
          var apTab = document.getElementById('tab-ap') || document.querySelector('[onclick*="ap_proveedores"]') || document.querySelector('[onclick*="switchTab(\'ap\'"]');
          if (apTab) {
            apTab.click();
            // notif AP quitada
          }
        } else if (arCount > 0) {
          // notif AR quitada
        }
      }).catch(()=>{});
    }).catch(()=>{});
  }, 300);
}

var _lastBatchFiles = [];
function retryLastBatch() {
  var retryBtn = document.getElementById('btn-retry');
  if (retryBtn) retryBtn.style.display = 'none';
  _procesarSiguiente && _procesarSiguiente();
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
  // Si todos los valores son 0, dibujar una línea tenue y salir
  if (data.every(function(v){ return v === 0; })) {
    ctx.beginPath();
    ctx.moveTo(pad, H - pad - 1);
    ctx.lineTo(W - pad, H - pad - 1);
    ctx.strokeStyle = color;
    ctx.globalAlpha = 0.15;
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.globalAlpha = 1;
    return;
  }
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
  {valId:'s-tot',  color:'#60a5fa', baseColor:'#60a5fa', accent:true},
  {valId:'s-imp',  color:'#60a5fa', baseColor:'#60a5fa', accent:true},
  {valId:'s-ok',   color:'#22c55e', baseColor:'#22c55e'},
  {valId:'s-disc', color:'#ef4444', baseColor:'#ef4444'},
  {valId:'s-di',   color:'#f97316', baseColor:'#f97316'},
  {valId:'s-pend', color:'#8b5cf6', baseColor:'#8b5cf6'},
];
const AP_SPARKS = [
  {valId:'ap-total',    color:'#60a5fa', baseColor:'#60a5fa', accent:true},
  {valId:'ap-importe',  color:'#60a5fa', baseColor:'#60a5fa', accent:true},
  {valId:'ap-matches',  color:'#22c55e', baseColor:'#22c55e'},
  {valId:'ap-disc',     color:'#ef4444', baseColor:'#ef4444'},
  {valId:'ap-sinpo',    color:'#f97316', baseColor:'#f97316'},
  {valId:'ap-aprobadas',color:'#8b5cf6', baseColor:'#8b5cf6'},
];

// ══════════════════════════════════════════════════════════════
// DRR REVENUE CHART
// ══════════════════════════════════════════════════════════════
var _drrChart = null;

async function renderDRRChart() {
  try {
    // Lee el dato que ya cargó renderDRR; si se llama suelto (tras subir un DRR),
    // lo pide. El guard mira la LONGITUD de `dias`: un array vacío es truthy en
    // JS, así que `!d.dias` no abortaba nunca — de ahí el "gráfico vacío".
    let d = window._drrChartData;
    if (!d) { const r = await fetch('/api/drr_daily_chart'); d = await r.json(); }
    const card = document.getElementById('drr-chart-card');
    const vacio = document.getElementById('drr-chart-vacio');
    const canvas = document.getElementById('drr-revenue-chart');

    // Sin DRR cargado no hay nada que contar: la tarjeta se queda escondida y
    // el bloque de "sube un DRR" de arriba ya lo dice.
    if (!d || d.error) { if (card) card.style.display = 'none'; return; }

    // Y aqui el segundo agujero: el guard era `!d.dias`, y en JavaScript un
    // array VACIO es truthy, asi que no abortaba nunca. Enseñaba la tarjeta y
    // dibujaba un grafico de cero barras — el "sale vacio" que costo una
    // prueba de integracion entera. Ahora, si no hay dias, se dice POR QUE.
    if (!Array.isArray(d.dias) || d.dias.length === 0) {
      if (card) card.style.display = 'block';
      if (canvas) canvas.style.display = 'none';
      if (vacio) {
        vacio.style.display = 'block';
        vacio.textContent = _tSSE('No hay revenue por día que mostrar')
          + (d.motivo ? ' — ' + d.motivo : '')
          + '. El resto de la pestaña sigue siendo válido.';
      }
      if (_drrChart) { _drrChart.destroy(); _drrChart = null; }
      return;
    }
    if (card) card.style.display = 'block';
    if (vacio) vacio.style.display = 'none';
    if (canvas) canvas.style.display = '';
    if (!canvas || !window.Chart) return;
    if (_drrChart) { _drrChart.destroy(); _drrChart = null; }

    // Grafico de AREA (evolucion por dia). El eje se ajusta al RANGO del revenue
    // para que se note el sube y baja aunque la diferencia de euros sea poca; el
    // dia descuadrado va en rojo con ⚠, y cada punto lleva su importe encima.
    const ctx = canvas.getContext('2d');
    const gv = v => getComputedStyle(document.body).getPropertyValue(v).trim();
    const acc = gv('--acc') || '#3b82f6', tx = gv('--tx') || '#f1f5f9', red = gv('--red') || '#ef4444', mut = gv('--mut') || '#94a3b8', s2v = gv('--s2') || '#334155', bg = gv('--bg') || '#0f172a';
    const _rgba = (c, a) => { c = (c || '').trim(); if (c[0] === '#') { let h = c.slice(1); if (h.length === 3) h = h.split('').map(x => x + x).join(''); const n = parseInt(h, 16); return 'rgba(' + ((n >> 16) & 255) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + a + ')'; } const m = c.match(/(\d+)[,\s]+(\d+)[,\s]+(\d+)/); return m ? 'rgba(' + m[1] + ',' + m[2] + ',' + m[3] + ',' + a + ')' : c; };
    const _eur = n => '€' + Math.round(n).toLocaleString('en-US');
    const rev = d.revenue, minR = Math.min.apply(null, rev), maxR = Math.max.apply(null, rev);
    const _pad = Math.max((maxR - minR) * 0.5, maxR * 0.01) || 1000;
    const yMin = Math.max(0, Math.floor((minR - _pad) / 100) * 100), yMax = Math.ceil((maxR + _pad) / 100) * 100;
    const grad = ctx.createLinearGradient(0, 0, 0, 215);
    grad.addColorStop(0, _rgba(acc, .16)); grad.addColorStop(1, _rgba(acc, 0));
    const N = rev.length;
    const etiquetas = { id: 'drrRevLab', afterDatasetsDraw(c) {
      const cx = c.ctx, meta = c.getDatasetMeta(0);
      cx.save(); cx.font = '700 11px Inter,sans-serif';
      meta.data.forEach((pt, i) => {
        cx.fillStyle = d.oob[i] ? red : tx;
        cx.textAlign = i === 0 ? 'left' : (i === N - 1 ? 'right' : 'center');
        const dx = i === 0 ? -2 : (i === N - 1 ? 2 : 0);
        cx.fillText(_eur(rev[i]) + (d.oob[i] ? ' ⚠' : ''), pt.x + dx, pt.y - 14);
      });
      cx.restore();
    } };
    _drrChart = new Chart(canvas, {
      type: 'line',
      data: { labels: d.labels || d.dias.map(String), datasets: [{
        label: 'Revenue', data: rev, borderColor: acc, backgroundColor: grad, fill: true, tension: .35, borderWidth: 2.5,
        pointRadius: d.oob.map(o => o ? 6 : 5), pointHoverRadius: 7,
        pointBackgroundColor: d.oob.map(o => o ? red : acc), pointBorderColor: bg, pointBorderWidth: 2
      }] },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 26, left: 6, right: 16 } },
        plugins: { legend: { display: false }, tooltip: { callbacks: {
          title: items => 'Día ' + d.dias[items[0].dataIndex] + (d.oob[items[0].dataIndex] ? ' ⚠ OOB' : ''),
          label: item => 'Revenue: ' + _eur(item.raw)
        } } },
        scales: {
          x: { grid: { display: false }, ticks: { color: mut, font: { size: 11 }, padding: 6 } },
          y: { min: yMin, max: yMax, grid: { color: _rgba(s2v, .45), drawBorder: false }, ticks: { color: mut, font: { size: 10 }, maxTicksLimit: 5, callback: v => '€' + Math.round(v / 1000) + 'K' } }
        }
      },
      plugins: [etiquetas]
    });
  } catch(e) { console.warn('DRR chart error:', e); }
}

// FASE C: `renderMHMap` (el mapa de puntos por ciudad) se va con
// `openHotelDetail`. No lo llamaba nadie desde hacia tiempo y pintaba sobre
// `#mh-dots`, que ya no existe en el panel.



// ══════════════════════════════════════════════════════════════
// GUIDED TOUR — Demo interactiva paso a paso
// ══════════════════════════════════════════════════════════════

// ── Tour stubs (old system removed) ─────────────────────────────────
function tourStart() { startTour(); }   // alias — button in nav calls this
function tourPrev()  { prevTourStep(); }
function tourNext()  { nextTourStep(); }
function tourEnd()   { endTour(); }
function tourGo(n)   { startTour(); }
// ─────────────────────────────────────────────────────────────────────



// ── Inline action confirmation (replaces confirm() dialogs) ─────────
function _dismissConfirm() { var c=document.getElementById('yve-confirm'); if(c) c.remove(); }
function showConfirmAction(title, subtitle, btnLabel, onConfirm) {
  var existing = document.getElementById('yve-confirm');
  if (existing) existing.remove();
  var el = document.createElement('div');
  el.id = 'yve-confirm';
  el.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);' +
    'background:#1e293b;border:1px solid rgba(245,158,11,.4);border-radius:14px;' +
    'padding:14px 18px;z-index:9500;display:flex;align-items:center;gap:14px;' +
    'box-shadow:0 8px 32px rgba(0,0,0,.5);max-width:420px;width:calc(100% - 32px);' +
    'animation:slideUp .2s ease';
  el.innerHTML =
    '<div style="flex:1"><div style="font-weight:600;color:#f1f5f9;font-size:13px">' + title + '</div>' +
    '<div style="font-size:11px;color:#94a3b8;margin-top:2px">' + subtitle + '</div></div>' +
    '<button onclick="_dismissConfirm()" ' +
      'style="background:transparent;border:1px solid #334155;color:#64748b;padding:6px 12px;border-radius:8px;font-size:12px;cursor:pointer;flex-shrink:0">Cancelar</button>' +
    '<button id="yve-confirm-yes" ' +
      'style="background:#f59e0b;border:none;color:#fff;padding:6px 14px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;flex-shrink:0">' + btnLabel + '</button>';
  document.body.appendChild(el);
  document.getElementById('yve-confirm-yes').onclick = function() {
    el.remove();
    onConfirm();
  };
  setTimeout(function() { if (el.parentNode) el.remove(); }, 8000);
}

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
    '/api/ar_real/data',
    '/api/stats_banco',
  ];
  preloads.forEach(url => fetch(url).catch(() => {}));
}, 3000);

const _i18nCache = {};
const _i18nOriginal = {}; // textos ES originales — para restaurar al volver a español
var _i18nStrMap = {
  en: {
    "🎭 Modo Demo personalizado": "🎭 Custom Demo Mode",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Type your hotel, your chain or several chains and I'll generate realistic sample data with those names: invoices, bank, F&B and Multi-Hotel. Perfect for showing the product to clients and accounting firms.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "One line per chain · format \"Chain: Hotel 1, Hotel 2\"",
    "🎭 Generar demo": "🎭 Generate demo",
    "Cancelar": "Cancel",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "The system is ready to automate your hotel's finances.",
    "El primer paso: procesa las facturas OTA del mes en": "First step: process this month's OTA invoices in",
    "👋 Bienvenido a Yve.01": "👋 Welcome to Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "The hotel finance system that automates AR, AP, DRR and reporting. This tour walks you through every module from left to right — 3 minutes and you'll master it all. Drag me out of the way: I'll dock wherever you drop me.",
    "📥 AR — Comisiones OTA": "📥 AR — OTA Commissions",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Automatically verifies Booking.com and Expedia commissions. Processed invoices, total amount, claimable discrepancies and pending DI certificates. The red number is euros you can recover.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "For every supplier invoice, Yve cross-checks 3 documents: invoice, purchase order (PO) and delivery note. Everything matches → automatic Match OK. Any difference → alert plus an AI-generated email to the supplier.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Drag your .xlsm file here. Yve extracts RevPAR, ADR, GOP%, occupancy and the 7,000+ Trial Balance lines in seconds. It detects Out of Balance automatically and alerts you instantly.",
    "🏦 Banco — Conciliación": "🏦 Bank — Reconciliation",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Automatically matches the bank statement against supplier invoices. It flags unmatched transactions, amount differences and duplicate payments. From 8 hours down to 2 minutes.",
    "🔔 Notificaciones": "🔔 Notifications",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Set up automatic alerts by email or Telegram: OTA discrepancies, unsigned invoices, DRR Out of Balance or low F&B stock. Yve warns you proactively.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Calculates real vs theoretical Food Cost by category. It connects POS data, recipes and inventory. Detects waste, spots the dishes with the best margin and optimizes restaurant performance.",
    "🏢 AR Real — Grupos Corporativos": "🏢 Real AR — Corporate Groups",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Full corporate client management: issue invoices, track aging (0-30 / 31-60 / +90 days), collect with one click and send automatic email reminders.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hotel — Group View",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "For the group's Financial Controller: consolidated KPIs, per-hotel performance ranking, 6-month trend and centralized alerts. One screen, the whole group.",
    "¡Ya conoces Yve.01!": "You now know Yve.01!",
    "Empezar con AR →": "Start with AR →",
    "Sin alertas bancarias pendientes.": "No pending bank alerts.",
    "Sin alertas bancarias.": "No bank alerts.",
    "● Activo": "● Active",
    "○ Inactivo": "○ Inactive",
    "Email de notificaciones": "Notification email",
    "No hay facturas con este filtro": "No invoices with this filter",
    "No hay hoteles en el grupo": "No hotels in the group",
    "Mes actual": "Current month",
    "🏆 Top Performers (RevPAR)": "🏆 Top Performers (RevPAR)",
    "Todos los hoteles": "All hotels",
    "Hab.": "Rooms",
    "Ocup.": "Occ.",
    "Categoría": "Category",
    "Estado": "Status",
    "📄 Diario": "📄 Daily",
    "📊 Semanal": "📊 Weekly",
    "📈 Mensual": "📈 Monthly",
    "🎯 Ejecutivo PDF": "🎯 Executive PDF",
    "📊 Consolidado Excel": "📊 Consolidated Excel",
    "📋 Historial de procesado": "📋 Processing history",
    "↻ Actualizar datos": "↻ Refresh data",
    "🎨 Personalizar colores": "🎨 Customize colors",
    "🌅 Briefing de hoy": "🌅 Today's briefing",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ What open discrepancies do I have?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 How much can I claim this month?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 What needs my signature today?",
    "Escribe aquí…": "Type here…",
    "Conciliado:": "Matched:",
    "FC% promedio": "Average FC%",
    "media ponderada": "weighted average",
    "Alertas FC alto": "High FC alerts",
    "media del menú": "menu average",
    "Mejor margen": "Best margin",
    "menor FC%": "lowest FC%",
    "FC% Medio": "Avg FC%",
    "revisar urgente": "review urgently",
    "Receta": "Recipe",
    "PVP": "Price",
    "Coste": "Cost",
    "Margen": "Margin",
    "Crítico": "Critical",
    "Bajo": "Low",
    "📸 Escanear Documento": "📸 Scan Document",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Take a photo of the physical document (invoice, BEO, contract, statement...) and Yve will read it with AI.",
    "📸 Cámara": "📸 Camera",
    "🖼️ Galería": "🖼️ Gallery",
    "⚡ Procesar documento": "⚡ Process document",
    "📸 Escanear": "📸 Scan",
    "Cerrar": "Close",
    "📸 Escanear más": "📸 Scan more",
    '% Com.': '% Comm.',
    '0 ya procesados (se saltarán)': '0 already processed (will skip)',
    '2 minutos para conocer todo Yve': '2 minutes to discover Yve',
    'AP firma': 'AP signature',
    'AP pendientes': 'Pending AP',
    'AR Dashboard - Facturas Pendientes': 'AR Dashboard — Pending Invoices',
    'AR pend': 'AR pending',
    'ARCHIVOS SELECCIONADOS': 'FILES SELECTED',
    'Abrir chat Yve': 'Open Yve chat',
    'Acceso en tiempo real a los datos del hotel': 'Real-time hotel data access',
    'Acciones': 'Actions',
    'Actualizar datos': 'Refresh data',
    'Ahora no': 'Not now',
    'Alerta': 'Alert',
    'Alertas': 'Alerts',
    'Alertas Bancarias': 'Banking Alerts',
    'Alertas DRR': 'DRR Alerts',
    'Alto FC': 'High FC',
    'Antigüedad de saldo': 'Receivables aging',
    'Aprobación': 'Approval',
    'Aprobadas': 'Approved',
    'Archivo': 'File',
    'Arrastra archivos aquí o haz clic': 'Drag files here or click',
    'Arrastra tu DRR aquí o': 'Drag your DRR here or',
    'Asunto': 'Subject',
    'Bajo': 'Low',
    'Búsqueda global': 'Global search',
    'CATEGORÍA CON MÁS MERMA': 'CATEGORY WITH MOST WASTE',
    'CLIENTES DE CRÉDITO': 'CREDIT CLIENTS',
    'COBRADO MES': 'COLLECTED MONTH',
    'COSTE TOTAL MERMAS': 'TOTAL WASTE COST',
    'Calcular': 'Calculate',
    'Cambiar de pestaña': 'Switch tab',
    'Cambiar rol': 'Change role',
    'Canales de notificación': 'Notification channels',
    'Cancelar': 'Cancel',
    'Cantidad': 'Quantity',
    'Cargando datos...': 'Loading data...',
    'Cargando...': 'Loading...',
    'Categoría': 'Category',
    'Causa': 'Cause',
    'Cerrar': 'Close',
    'Cerrar modales': 'Close modals',
    'Certif. DI pendiente': 'DI cert. pending',
    'Cliente': 'Client',
    'Clientes': 'Clients',
    'Clientes Directos': 'Direct Clients',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Credit clients · Corporate billing · Collections control',
    'Cobradas': 'Collected',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Complete these steps to get the most out of Yve.01',
    'Con avisos': 'With warnings',
    'Concepto': 'Concept',
    'Conciliado:': 'Reconciled:',
    'Conciliados': 'Reconciled',
    'Configura dónde recibes alertas de Yve': 'Configure where you receive Yve alerts',
    'Correctas': 'Correct',
    'Coste': 'Cost',
    'Crítico': 'Critical',
    'Críticos': 'Critical',
    'Cuenta': 'Account',
    'Destinatario': 'Recipient',
    'Detalle factura': 'Invoice detail',
    'Diferencias': 'Differences',
    'Discrepancia': 'Discrepancy',
    'Discrepancias': 'Discrepancies',
    'Días': 'Days',
    'EN SERVIDOR (facturas-entrada)': 'ON SERVER (invoices-inbox)',
    'EUR procesados': 'EUR processed',
    'Empezar →': 'Start →',
    'Entrada': 'Entry',
    'Error inventario': 'Inventory error',
    'Error mermas': 'Waste error',
    'Error recetas': 'Recipe error',
    'Estado': 'Status',
    'Eventos que disparan alerta': 'Alert trigger events',
    'Explorar sin tour': 'Explore without tour',
    'F&B + OTRAS': 'F&B + OTHER',
    'Factura': 'Invoice',
    'Facturas': 'Invoices',
    'Fecha': 'Date',
    'Ficha de Recetas con Coste Teórico': 'Recipe Cards with Theoretical Cost',
    'Food Cost % — Teórico vs Real': 'Food Cost % — Theoretical vs Actual',
    'Food Cost por Categoría': 'Food Cost by Category',
    'GOP% MEDIO': 'AVG GOP%',
    'Generar email': 'Generate email',
    'Grupo': 'Group',
    'Grupos Corporativos': 'Corporate Groups',
    'Guardar': 'Save',
    'Habitaciones': 'Rooms',
    'Historial de Mermas · Total:': 'Waste History · Total:',
    'Hoteles OK': 'Hotels OK',
    'Importe': 'Amount',
    'KPIs Operativos': 'Operational KPIs',
    'MEJOR GOP%': 'BEST GOP%',
    'Margen': 'Margin',
    'Mermas por Causa': 'Waste by Cause',
    'Mostrar atajos': 'Show shortcuts',
    'Nueva factura': 'New invoice',
    'Ocupacion': 'Occupancy',
    'Ocupación': 'Occupancy',
    'PVP': 'RRP',
    'Pendiente': 'Pending',
    'Performance Financiero': 'Financial Performance',
    'Proveedor': 'Supplier',
    'REGISTROS': 'RECORDS',
    'Registrar Merma': 'Register Waste',
    'Revenue hoy vs Budget:': 'Today\'s revenue vs Budget:',
    'Seleccionar cliente...': 'Select client...',
    'Siguiente →': 'Next →',
    'Sin alertas activas': 'No active alerts',
    'Sin datos': 'No data',
    'Sin facturas AP': 'No AP invoices',
    'Stock de Ingredientes': 'Ingredient Stock',
    'Tipo': 'Type',
    'Yve — Copiloto Financiero': 'Yve — Financial Copilot',
    '¿Hacemos el tour?': 'Take the tour?',
    '← Atrás': '← Back',
    '✓ Finalizar': '✓ Finish',
  
    '⌨ Atajos': '⌨ Shortcuts',
    '📊 Vista lite': '📊 Lite view',
    '📊 Vista completa': '📊 Full view',
    '📱 Vista lite': '📱 Lite view',
    '⚡ Procesar Archivos': '⚡ Process Files',
    '⚡ Procesar archivos nuevos': '⚡ Process new files',
    '▶ Procesar pendientes del servidor': '▶ Process server pending',
    '✅ Aprobar Match OK': '✅ Approve Match OK',
    '📲 Aprobar facturas AR': '📲 Approve AR invoices',
    '⚡ Conciliar': '⚡ Reconcile',
    '🔍 Filtrar': '🔍 Filter',
    'Sin facturas AR todavía': 'No AR invoices yet',
    'Sin datos.': 'No data.',
    'Pendiente emitir': 'Pending issue',
    'Pendientes': 'Pending',
    'Todas las facturas del ciclo': 'All cycle invoices',
    'Pulsa ⚡ Procesar Archivos': 'Press ⚡ Process AP Invoices',
    'Pulsa ⚡ Procesar Archivos.': 'Press ⚡ Process Files.',
    '📊 Resumen': '📊 Summary',
    'Vista consolidada del grupo': 'Group consolidated view',
    '⚠️ Alertas activas': '⚠️ Active alerts',
    '⌨ Atajos de teclado': '⌨ Keyboard shortcuts',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Global search Ctrl+K · ⌨ Shortcuts 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Lite view on all panels',
    'PDF (facturas) · XLSM (DRR)': 'PDF (invoices) · XLSM (DRR)',
    'Actualizado': 'Updated',
    'facturas cargadas': 'invoices loaded',
    'Todos': 'All',
    'Sin archivo cargado': 'No file loaded',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'click to select (.xlsm/.xlsx)',
    'Revenue Diario': 'Daily Revenue',
    'Trial Balance — Estado Diario': 'Trial Balance — Daily Status',
    'Sin alertas.': 'No alerts.',
    'Historial de Notificaciones': 'Notification History',
    'Sin notificaciones.': 'No notifications.',
    'Historial de notificaciones enviadas': 'Sent notifications history',
    'Ingrediente': 'Ingredient',
    'Receta': 'Recipe',
    '▶ Ejecutar Análisis': '▶ Run Analysis',
    'Sin datos F&B.': 'No F&B data.',
    'Reportes': 'Reports',
    'Administrador': 'Administrator',
    '👤 Administrador': '👤 Administrator',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 OTA claims awaiting approval",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "When Yve finds commissions charged above the contract, they will show up here to claim.",
    "✍️ Redactar con IA": "✍️ Draft with AI",
    "✅ Aprobar y enviar": "✅ Approve and send",
    "🔄 Regenerar": "🔄 Regenerate",
    "🗑 Descartar": "🗑 Discard",
    "Enviar a": "Send to",
    "Descartada": "Discarded",
    "✓ Enviada": "✓ Sent",
    "email de la OTA": "OTA email",
  },
  ca: {
    "🎭 Modo Demo personalizado": "🎭 Mode Demo personalitzat",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Escriu el teu hotel, la teva cadena o diverses cadenes i genero dades d'exemple realistes amb aquests noms: factures, banc, F&B i Multi-Hotel. Ideal per ensenyar el producte a clients i gestories.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "Una línia per cadena · format \"Cadena: Hotel 1, Hotel 2\"",
    "🎭 Generar demo": "🎭 Generar demo",
    "Cancelar": "Cancel·lar",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "El sistema està a punt per automatitzar les finances del teu hotel.",
    "El primer paso: procesa las facturas OTA del mes en": "El primer pas: processa les factures OTA del mes a",
    "👋 Bienvenido a Yve.01": "👋 Benvingut a Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "El sistema de finances hoteleres que automatitza AR, AP, DRR i reporting. Aquest tour et porta per cada mòdul d'esquerra a dreta — 3 minuts i ja ho domines tot. Arrossega'm si et faig nosa: m'acoblo sol on em deixis anar.",
    "📥 AR — Comisiones OTA": "📥 AR — Comissions OTA",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Verifica automàticament les comissions de Booking.com i Expedia. Factures processades, import total, discrepàncies reclamables i certificats DI pendents. El número vermell són euros que pots recuperar.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "Per a cada factura de proveïdor, Yve creua 3 documents: factura, comanda (PO) i albarà. Si tot quadra → Match OK automàtic. Si hi ha diferència → alerta i email al proveïdor generat amb IA.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Arrossega el teu fitxer .xlsm aquí. Yve extreu RevPAR, ADR, GOP%, ocupació i les més de 7.000 línies del Trial Balance en segons. Detecta Out of Balance automàticament i t'avisa a l'instant.",
    "🏦 Banco — Conciliación": "🏦 Banc — Conciliació",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Creua automàticament l'extracte bancari amb les factures de proveïdors. Identifica moviments no conciliats, diferències d'import i pagaments duplicats. De 8 hores a 2 minuts.",
    "🔔 Notificaciones": "🔔 Notificacions",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Configura alertes automàtiques per email o Telegram: discrepàncies OTA, factures sense signar, Out of Balance al DRR o estoc baix a F&B. Yve t'avisa proactivament.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Calcula el Food Cost real vs teòric per categoria. Connecta les dades del POS, receptes i inventari. Detecta minves, identifica quins plats tenen millor marge i optimitza el rendiment del restaurant.",
    "🏢 AR Real — Grupos Corporativos": "🏢 AR Real — Grups Corporatius",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Gestió completa de clients corporatius: emet factures, controla l'aging (0-30 / 31-60 / +90 dies), cobra amb un clic i envia recordatoris automàtics per email.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hotel — Vista de Grup",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "Per al Financial Controller del grup: KPIs consolidats, rànquing de rendiment per hotel, tendència de 6 mesos i alertes centralitzades. Una pantalla, tot el grup.",
    "¡Ya conoces Yve.01!": "Ja coneixes Yve.01!",
    "Empezar con AR →": "Començar amb AR →",
    "Sin alertas bancarias pendientes.": "Sense alertes bancàries pendents.",
    "Sin alertas bancarias.": "Sense alertes bancàries.",
    "● Activo": "● Actiu",
    "○ Inactivo": "○ Inactiu",
    "Email de notificaciones": "Email de notificacions",
    "No hay facturas con este filtro": "No hi ha factures amb aquest filtre",
    "No hay hoteles en el grupo": "No hi ha hotels al grup",
    "Mes actual": "Mes actual",
    "🏆 Top Performers (RevPAR)": "🏆 Millors hotels (RevPAR)",
    "Todos los hoteles": "Tots els hotels",
    "Hab.": "Hab.",
    "Ocup.": "Ocup.",
    "Categoría": "Categoria",
    "Estado": "Estat",
    "📄 Diario": "📄 Diari",
    "📊 Semanal": "📊 Setmanal",
    "📈 Mensual": "📈 Mensual",
    "🎯 Ejecutivo PDF": "🎯 Executiu PDF",
    "📊 Consolidado Excel": "📊 Consolidat Excel",
    "📋 Historial de procesado": "📋 Historial de processament",
    "↻ Actualizar datos": "↻ Actualitzar dades",
    "🎨 Personalizar colores": "🎨 Personalitzar colors",
    "🌅 Briefing de hoy": "🌅 Briefing d'avui",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ Quines discrepàncies tinc obertes?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 Quant puc reclamar aquest mes?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 Què necessita la meva signatura avui?",
    "Escribe aquí…": "Escriu aquí…",
    "Conciliado:": "Conciliat:",
    "FC% promedio": "FC% mitjà",
    "media ponderada": "mitjana ponderada",
    "Alertas FC alto": "Alertes FC alt",
    "media del menú": "mitjana de la carta",
    "Mejor margen": "Millor marge",
    "menor FC%": "menor FC%",
    "FC% Medio": "FC% mitjà",
    "revisar urgente": "revisar urgent",
    "Receta": "Recepta",
    "PVP": "PVP",
    "Coste": "Cost",
    "Margen": "Marge",
    "Crítico": "Crític",
    "Bajo": "Baix",
    "📸 Escanear Documento": "📸 Escanejar Document",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Fes una foto al document físic (factura, BEO, contracte, extracte...) i Yve el llegirà amb IA.",
    "📸 Cámara": "📸 Càmera",
    "🖼️ Galería": "🖼️ Galeria",
    "⚡ Procesar documento": "⚡ Processar document",
    "📸 Escanear": "📸 Escanejar",
    "Cerrar": "Tancar",
    "📸 Escanear más": "📸 Escanejar més",
    '0 ya procesados (se saltarán)': '0 ja processats (s\'ometran)',
    '2 minutos para conocer todo Yve': '2 minuts per conèixer Yve',
    'AP firma': 'Firma AP',
    'AP pendientes': 'AP pendents',
    'AR Dashboard - Facturas Pendientes': 'Tauler AR — Factures Pendents',
    'AR pend': 'AR pendent',
    'ARCHIVOS SELECCIONADOS': 'ARXIUS SELECCIONATS',
    'Abrir chat Yve': 'Obrir xat Yve',
    'Acceso en tiempo real a los datos del hotel': 'Accés temps real a les dades',
    'Acciones': 'Accions',
    'Actualizar datos': 'Actualitzar dades',
    'Ahora no': 'Ara no',
    'Alertas': 'Alertes',
    'Alertas Bancarias': 'Alertes Bancàries',
    'Alertas DRR': 'Alertes DRR',
    'Alto FC': 'FC alt',
    'Antigüedad de saldo': 'Antiguitat saldo',
    'Aprobación': 'Aprovació',
    'Aprobadas': 'Aprovades',
    'Archivo': 'Arxiu',
    'Arrastra archivos aquí o haz clic': 'Arrossega arxius aquí o clica',
    'Arrastra tu DRR aquí o': 'Arrossega el DRR aquí o',
    'Asunto': 'Assumpte',
    'Bajo': 'Baix',
    'Budget': 'Pressupost',
    'Búsqueda global': 'Cerca global',
    'CATEGORÍA CON MÁS MERMA': 'CATEGORIA AMB MÉS MALBÉ',
    'CLIENTES DE CRÉDITO': 'CLIENTS DE CRÈDIT',
    'COBRADO MES': 'COBRAT MES',
    'COSTE TOTAL MERMAS': 'COST TOTAL MALBÉ',
    'Cambiar de pestaña': 'Canviar pestanya',
    'Cambiar rol': 'Canviar rol',
    'Canales de notificación': 'Canals notificació',
    'Cancelar': 'Cancel·lar',
    'Cantidad': 'Quantitat',
    'Cargando datos...': 'Carregant dades...',
    'Cargando...': 'Carregant...',
    'Categoría': 'Categoria',
    'Cerrar': 'Tancar',
    'Cerrar modales': 'Tancar modals',
    'Certif. DI pendiente': 'Certif. DI pendent',
    'Cliente': 'Client',
    'Clientes': 'Clients',
    'Clientes Directos': 'Clients Directes',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Clients crèdit · Facturació corporativa · Control cobraments',
    'Cobradas': 'Cobrades',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Completa aquests passos per aprofitar Yve.01',
    'Con avisos': 'Amb avisos',
    'Concepto': 'Concepte',
    'Conciliado:': 'Conciliat:',
    'Conciliados': 'Conciliats',
    'Configura dónde recibes alertas de Yve': 'Configura on reps les alertes',
    'Correctas': 'Correctes',
    'Coste': 'Cost',
    'Crítico': 'Crític',
    'Críticos': 'Crítics',
    'Cuenta': 'Compte',
    'Destinatario': 'Destinatari',
    'Detalle factura': 'Detall factura',
    'Diferencias': 'Diferències',
    'Discrepancia': 'Discrepància',
    'Discrepancias': 'Discrepàncies',
    'Días': 'Dies',
    'EN SERVIDOR (facturas-entrada)': 'AL SERVIDOR (factures-entrada)',
    'EUR procesados': 'EUR processats',
    'Empezar →': 'Començar →',
    'Error inventario': 'Error inventari',
    'Error mermas': 'Error malbé',
    'Error recetas': 'Error receptes',
    'Estado': 'Estat',
    'Eventos que disparan alerta': 'Esdeveniments d\'alerta',
    'Explorar sin tour': 'Explorar sense tour',
    'F&B + OTRAS': 'F&B + ALTRES',
    'F&B Cost %': 'Cost F&B %',
    'F&B Cost Control': 'Control Cost F&B',
    'Facturas': 'Factures',
    'Fecha': 'Data',
    'Ficha de Recetas con Coste Teórico': 'Fitxa Receptes amb Cost Teòric',
    'Food Cost % — Teórico vs Real': 'Food Cost % — Teòric vs Real',
    'Food Cost por Categoría': 'Food Cost per Categoria',
    'Forecast': 'Previsió',
    'GOP% MEDIO': 'GOP% MITJÀ',
    'Grupo': 'Grup',
    'Grupos Corporativos': 'Grups Corporatius',
    'Guardar': 'Desar',
    'Habitaciones': 'Habitacions',
    'Historial de Mermas · Total:': 'Historial Malbés · Total:',
    'Hoteles OK': 'Hotels OK',
    'Importe': 'Import',
    'KPIs Operativos': 'KPIs Operatius',
    'MEJOR GOP%': 'MILLOR GOP%',
    'Margen': 'Marge',
    'Mermas por Causa': 'Malbé per Causa',
    'Mostrar atajos': 'Mostrar dreceres',
    'Nueva factura': 'Nova factura',
    'Ocupacion': 'Ocupació',
    'Ocupación': 'Ocupació',
    'Pendiente': 'Pendent',
    'Performance Financiero': 'Rendiment Financer',
    'Proveedor': 'Proveïdor',
    'REGISTROS': 'REGISTRES',
    'REVENUE': 'INGRESSOS',
    'Registrar Merma': 'Registrar Malbé',
    'Revenue MTD': 'Ingressos MTD',
    'Revenue hoy vs Budget:': 'Ingressos avui vs Budget:',
    'Seleccionar cliente...': 'Seleccionar client...',
    'Siguiente →': 'Següent →',
    'Sin alertas activas': 'Sense alertes actives',
    'Sin datos': 'Sense dades',
    'Sin facturas AP': 'Sense factures AP',
    'Stock de Ingredientes': 'Estoc d\'Ingredients',
    'Tipo': 'Tipus',
    'Yve — Copiloto Financiero': 'Yve — Copilot Financer',
    '¿Hacemos el tour?': 'Fem el tour?',
    '← Atrás': '← Enrere',
    '✓ Finalizar': '✓ Finalitzar',
  
    '⌨ Atajos': '⌨ Dreceres',
    '📊 Vista lite': '📊 Vista reduïda',
    '📱 Vista lite': '📱 Vista reduïda',
    '⚡ Procesar Archivos': '⚡ Processar Arxius',
    '⚡ Procesar archivos nuevos': '⚡ Processar arxius nous',
    '▶ Procesar pendientes del servidor': '▶ Processar pendents servidor',
    '✅ Aprobar Match OK': '✅ Aprovar Match OK',
    '📲 Aprobar facturas AR': '📲 Aprovar factures AR',
    'Sin facturas AR todavía': 'Sense factures AR encara',
    'Sin datos.': 'Sense dades.',
    'Pendiente emitir': 'Pendent d\'emetre',
    'Pendientes': 'Pendents',
    'Todas las facturas del ciclo': 'Totes les factures del cicle',
    'Pulsa ⚡ Procesar Archivos': 'Prem ⚡ Processar Factures AP',
    'Pulsa ⚡ Procesar Archivos.': 'Prem ⚡ Processar Arxius.',
    '📊 Resumen': '📊 Resum',
    'Vista consolidada del grupo': 'Vista consolidada del grup',
    '⚠️ Alertas activas': '⚠️ Alertes actives',
    '⌨ Atajos de teclado': '⌨ Dreceres de teclat',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Cerca global Ctrl+K · ⌨ Dreceres 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Vista reduïda a tots els panells',
    'PDF (facturas) · XLSM (DRR)': 'PDF (factures) · XLSM (DRR)',
    'Actualizado': 'Actualitzat',
    'facturas cargadas': 'factures carregades',
    'Todos': 'Tots',
    'Sin archivo cargado': 'Sense arxiu carregat',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'fes clic per seleccionar (.xlsm/.xlsx)',
    'Revenue Diario': 'Ingressos Diaris',
    'Trial Balance — Estado Diario': 'Balanç — Estat Diari',
    'Sin alertas.': 'Sense alertes.',
    'Historial de Notificaciones': 'Historial Notificacions',
    'Sin notificaciones.': 'Sense notificacions.',
    'Historial de notificaciones enviadas': 'Historial notificacions enviades',
    'Ingrediente': 'Ingredient',
    'Receta': 'Recepta',
    '▶ Ejecutar Análisis': '▶ Executar Anàlisi',
    'Sin datos F&B.': 'Sense dades F&B.',
    'Reportes': 'Informes',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 Reclamacions OTA pendents d'aprovar",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "Quan Yve detecti comissions cobrades per sobre del contracte, apareixeran aquí per reclamar.",
    "✍️ Redactar con IA": "✍️ Redactar amb IA",
    "✅ Aprobar y enviar": "✅ Aprovar i enviar",
    "🔄 Regenerar": "🔄 Regenerar",
    "🗑 Descartar": "🗑 Descartar",
    "Enviar a": "Enviar a",
    "Descartada": "Descartada",
    "✓ Enviada": "✓ Enviada",
    "email de la OTA": "correu de l'OTA",
  },
  fr: {
    "🎭 Modo Demo personalizado": "🎭 Mode Démo personnalisé",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Saisissez votre hôtel, votre chaîne ou plusieurs chaînes et je génère des données d'exemple réalistes avec ces noms : factures, banque, F&B et Multi-Hôtel. Idéal pour présenter le produit aux clients et cabinets comptables.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "Une ligne par chaîne · format « Chaîne : Hôtel 1, Hôtel 2 »",
    "🎭 Generar demo": "🎭 Générer la démo",
    "Cancelar": "Annuler",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "Le système est prêt à automatiser les finances de votre hôtel.",
    "El primer paso: procesa las facturas OTA del mes en": "Première étape : traitez les factures OTA du mois dans",
    "👋 Bienvenido a Yve.01": "👋 Bienvenue sur Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "Le système de finances hôtelières qui automatise AR, AP, DRR et reporting. Ce tour vous guide module par module — 3 minutes et vous maîtrisez tout. Déplacez-moi si je gêne : je m'ancre là où vous me lâchez.",
    "📥 AR — Comisiones OTA": "📥 AR — Commissions OTA",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Vérifie automatiquement les commissions de Booking.com et Expedia. Factures traitées, montant total, écarts réclamables et certificats DI en attente. Le chiffre rouge, ce sont des euros à récupérer.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "Pour chaque facture fournisseur, Yve croise 3 documents : facture, bon de commande (PO) et bon de livraison. Tout correspond → Match OK automatique. Un écart → alerte et email au fournisseur généré par IA.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Glissez votre fichier .xlsm ici. Yve extrait RevPAR, ADR, GOP%, occupation et les 7 000+ lignes du Trial Balance en quelques secondes. Il détecte l'Out of Balance automatiquement et vous alerte aussitôt.",
    "🏦 Banco — Conciliación": "🏦 Banque — Rapprochement",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Rapproche automatiquement le relevé bancaire des factures fournisseurs. Il identifie les mouvements non rapprochés, les écarts de montant et les paiements en double. De 8 heures à 2 minutes.",
    "🔔 Notificaciones": "🔔 Notifications",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Configurez des alertes automatiques par email ou Telegram : écarts OTA, factures non signées, Out of Balance du DRR ou stock F&B bas. Yve vous prévient proactivement.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Calcule le Food Cost réel vs théorique par catégorie. Il connecte les données POS, recettes et inventaire. Détecte les pertes, identifie les plats les plus rentables et optimise la performance du restaurant.",
    "🏢 AR Real — Grupos Corporativos": "🏢 AR Réel — Groupes Corporate",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Gestion complète des clients corporate : émettez des factures, suivez l'aging (0-30 / 31-60 / +90 jours), encaissez en un clic et envoyez des rappels automatiques par email.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hôtel — Vue Groupe",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "Pour le Financial Controller du groupe : KPIs consolidés, classement de performance par hôtel, tendance sur 6 mois et alertes centralisées. Un écran, tout le groupe.",
    "¡Ya conoces Yve.01!": "Vous connaissez Yve.01 !",
    "Empezar con AR →": "Commencer avec AR →",
    "Sin alertas bancarias pendientes.": "Aucune alerte bancaire en attente.",
    "Sin alertas bancarias.": "Aucune alerte bancaire.",
    "● Activo": "● Actif",
    "○ Inactivo": "○ Inactif",
    "Email de notificaciones": "Email de notifications",
    "No hay facturas con este filtro": "Aucune facture avec ce filtre",
    "No hay hoteles en el grupo": "Aucun hôtel dans le groupe",
    "Mes actual": "Mois en cours",
    "🏆 Top Performers (RevPAR)": "🏆 Meilleurs hôtels (RevPAR)",
    "Todos los hoteles": "Tous les hôtels",
    "Hab.": "Ch.",
    "Ocup.": "Occ.",
    "Categoría": "Catégorie",
    "Estado": "Statut",
    "📄 Diario": "📄 Quotidien",
    "📊 Semanal": "📊 Hebdomadaire",
    "📈 Mensual": "📈 Mensuel",
    "🎯 Ejecutivo PDF": "🎯 PDF exécutif",
    "📊 Consolidado Excel": "📊 Excel consolidé",
    "📋 Historial de procesado": "📋 Historique de traitement",
    "↻ Actualizar datos": "↻ Actualiser les données",
    "🎨 Personalizar colores": "🎨 Personnaliser les couleurs",
    "🌅 Briefing de hoy": "🌅 Briefing du jour",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ Quels écarts sont ouverts ?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 Combien puis-je réclamer ce mois-ci ?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 Qu'attend ma signature aujourd'hui ?",
    "Escribe aquí…": "Écris ici…",
    "Conciliado:": "Rapproché :",
    "FC% promedio": "FC% moyen",
    "media ponderada": "moyenne pondérée",
    "Alertas FC alto": "Alertes FC élevé",
    "media del menú": "moyenne du menu",
    "Mejor margen": "Meilleure marge",
    "menor FC%": "FC% le plus bas",
    "FC% Medio": "FC% moyen",
    "revisar urgente": "à vérifier d'urgence",
    "Receta": "Recette",
    "PVP": "Prix",
    "Coste": "Coût",
    "Margen": "Marge",
    "Crítico": "Critique",
    "Bajo": "Bas",
    "📸 Escanear Documento": "📸 Scanner un document",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Prenez une photo du document physique (facture, BEO, contrat, relevé...) et Yve le lira avec l'IA.",
    "📸 Cámara": "📸 Caméra",
    "🖼️ Galería": "🖼️ Galerie",
    "⚡ Procesar documento": "⚡ Traiter le document",
    "📸 Escanear": "📸 Scanner",
    "Cerrar": "Fermer",
    "📸 Escanear más": "📸 Scanner plus",
    '0 ya procesados (se saltarán)': '0 déjà traités (seront ignorés)',
    '2 minutos para conocer todo Yve': '2 minutes pour découvrir Yve',
    'AP firma': 'Signature AP',
    'AP pendientes': 'AP en attente',
    'AR Dashboard - Facturas Pendientes': 'Tableau AR — Factures en attente',
    'AR pend': 'AR en attente',
    'ARCHIVOS SELECCIONADOS': 'FICHIERS SÉLECTIONNÉS',
    'Abrir chat Yve': 'Ouvrir chat Yve',
    'Acceso en tiempo real a los datos del hotel': 'Accès temps réel données hôtel',
    'Acciones': 'Actions',
    'Actualizar datos': 'Actualiser données',
    'Ahora no': 'Pas maintenant',
    'Alerta': 'Alerte',
    'Alertas': 'Alertes',
    'Alertas Bancarias': 'Alertes Bancaires',
    'Alertas DRR': 'Alertes DRR',
    'Alto FC': 'FC élevé',
    'Antigüedad de saldo': 'Ancienneté solde',
    'Aprobación': 'Approbation',
    'Aprobadas': 'Approuvées',
    'Archivo': 'Fichier',
    'Arrastra archivos aquí o haz clic': 'Glissez les fichiers ici ou cliquez',
    'Arrastra tu DRR aquí o': 'Glissez votre DRR ici ou',
    'Asunto': 'Sujet',
    'Bajo': 'Faible',
    'Búsqueda global': 'Recherche globale',
    'CATEGORÍA CON MÁS MERMA': 'CATÉGORIE AVEC PLUS DE PERTES',
    'CLIENTES DE CRÉDITO': 'CLIENTS À CRÉDIT',
    'COBRADO MES': 'ENCAISSÉ CE MOIS',
    'COSTE TOTAL MERMAS': 'COÛT TOTAL PERTES',
    'Calcular': 'Calculer',
    'Cambiar de pestaña': 'Changer d\'onglet',
    'Cambiar rol': 'Changer rôle',
    'Canales de notificación': 'Canaux notification',
    'Cancelar': 'Annuler',
    'Cantidad': 'Quantité',
    'Cargando datos...': 'Chargement données...',
    'Cargando...': 'Chargement...',
    'Categoría': 'Catégorie',
    'Causa': 'Cause',
    'Cerrar': 'Fermer',
    'Cerrar modales': 'Fermer modales',
    'Certif. DI pendiente': 'Cert. DI en attente',
    'Cliente': 'Client',
    'Clientes': 'Clients',
    'Clientes Directos': 'Clients Directs',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Clients crédit · Facturation corporative · Contrôle encaissements',
    'Cobradas': 'Encaissées',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Complétez ces étapes pour profiter au maximum de Yve.01',
    'Con avisos': 'Avec avertissements',
    'Concepto': 'Concept',
    'Conciliado:': 'Rapproché:',
    'Conciliados': 'Rapprochés',
    'Configura dónde recibes alertas de Yve': 'Configurez où recevoir les alertes',
    'Correctas': 'Correctes',
    'Coste': 'Coût',
    'Crítico': 'Critique',
    'Críticos': 'Critiques',
    'Cuenta': 'Compte',
    'Destinatario': 'Destinataire',
    'Detalle factura': 'Détail facture',
    'Diferencias': 'Différences',
    'Discrepancia': 'Écart',
    'Discrepancias': 'Écarts',
    'Días': 'Jours',
    'EN SERVIDOR (facturas-entrada)': 'SUR SERVEUR (factures-entrée)',
    'EUR procesados': 'EUR traités',
    'Empezar →': 'Commencer →',
    'Entrada': 'Entrée',
    'Error inventario': 'Erreur inventaire',
    'Error mermas': 'Erreur pertes',
    'Error recetas': 'Erreur recettes',
    'Estado': 'Statut',
    'Eventos que disparan alerta': 'Événements déclencheurs',
    'Explorar sin tour': 'Explorer sans visite',
    'F&B + OTRAS': 'F&B + AUTRES',
    'F&B Cost %': 'Coût F&B %',
    'F&B Cost Control': 'Contrôle Coût F&B',
    'Factura': 'Facture',
    'Facturas': 'Factures',
    'Fecha': 'Date',
    'Ficha de Recetas con Coste Teórico': 'Fiche Recettes avec Coût Théorique',
    'Food Cost % — Teórico vs Real': 'Coût Alim. % — Théorique vs Réel',
    'Food Cost por Categoría': 'Coût Alimentaire par Catégorie',
    'Forecast': 'Prévision',
    'GOP% MEDIO': 'GOP% MOYEN',
    'Generar email': 'Générer email',
    'Grupo': 'Groupe',
    'Grupos Corporativos': 'Groupes Corporatifs',
    'Guardar': 'Enregistrer',
    'Habitaciones': 'Chambres',
    'Historial de Mermas · Total:': 'Historique Pertes · Total:',
    'Hoteles OK': 'Hôtels OK',
    'Importe': 'Montant',
    'KPIs Operativos': 'KPIs Opérationnels',
    'MEJOR GOP%': 'MEILLEUR GOP%',
    'Margen': 'Marge',
    'Mermas por Causa': 'Pertes par Cause',
    'Mostrar atajos': 'Afficher raccourcis',
    'Nueva factura': 'Nouvelle facture',
    'Ocupacion': 'Occupation',
    'Ocupación': 'Occupation',
    'PVP': 'PVR',
    'Pendiente': 'En attente',
    'Performance Financiero': 'Performance Financière',
    'Proveedor': 'Fournisseur',
    'REGISTROS': 'ENREGISTREMENTS',
    'REVENUE': 'REVENUS',
    'Registrar Merma': 'Enregistrer Perte',
    'Revenue MTD': 'Revenus MTD',
    'Revenue hoy vs Budget:': 'Revenus aujourd\'hui vs Budget:',
    'Seleccionar cliente...': 'Sélectionner client...',
    'Siguiente →': 'Suivant →',
    'Sin alertas activas': 'Aucune alerte active',
    'Sin datos': 'Pas de données',
    'Sin facturas AP': 'Aucune facture AP',
    'Stock de Ingredientes': 'Stock d\'Ingrédients',
    'Tipo': 'Type',
    'Yve — Copiloto Financiero': 'Yve — Copilote Financier',
    '¿Hacemos el tour?': 'Faire la visite?',
    '← Atrás': '← Retour',
    '✓ Finalizar': '✓ Terminer',
  
    '⌨ Atajos': '⌨ Raccourcis',
    '📊 Vista lite': '📊 Vue réduite',
    '📊 Vista completa': '📊 Vue complète',
    '📱 Vista lite': '📱 Vue réduite',
    '⚡ Procesar Archivos': '⚡ Traiter Fichiers',
    '⚡ Procesar archivos nuevos': '⚡ Traiter nouveaux fichiers',
    '▶ Procesar pendientes del servidor': '▶ Traiter en attente serveur',
    '✅ Aprobar Match OK': '✅ Approuver Match OK',
    '📲 Aprobar facturas AR': '📲 Approuver factures AR',
    '⚡ Conciliar': '⚡ Rapprocher',
    '🔍 Filtrar': '🔍 Filtrer',
    'Sin facturas AR todavía': 'Aucune facture AR encore',
    'Sin datos.': 'Pas de données.',
    'Pendiente emitir': 'En attente d\'émission',
    'Pendientes': 'En attente',
    'Todas las facturas del ciclo': 'Toutes les factures du cycle',
    'Pulsa ⚡ Procesar Archivos': 'Appuyez ⚡ Traiter Factures AP',
    'Pulsa ⚡ Procesar Archivos.': 'Appuyez ⚡ Traiter Fichiers.',
    '📊 Resumen': '📊 Résumé',
    'Vista consolidada del grupo': 'Vue consolidée du groupe',
    '⚠️ Alertas activas': '⚠️ Alertes actives',
    '⌨ Atajos de teclado': '⌨ Raccourcis clavier',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Recherche globale Ctrl+K · ⌨ Raccourcis 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Vue réduite sur tous les panneaux',
    'PDF (facturas) · XLSM (DRR)': 'PDF (factures) · XLSM (DRR)',
    'Actualizado': 'Mis à jour',
    'facturas cargadas': 'factures chargées',
    'Todos': 'Tous',
    'Sin archivo cargado': 'Aucun fichier chargé',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'cliquez pour sélectionner (.xlsm/.xlsx)',
    'Revenue Diario': 'Revenus Journaliers',
    'Trial Balance — Estado Diario': 'Balance — État Journalier',
    'Sin alertas.': 'Aucune alerte.',
    'Historial de Notificaciones': 'Historique Notifications',
    'Sin notificaciones.': 'Aucune notification.',
    'Historial de notificaciones enviadas': 'Historique notifications envoyées',
    'Ingrediente': 'Ingrédient',
    'Receta': 'Recette',
    '▶ Ejecutar Análisis': '▶ Lancer Analyse',
    'Sin datos F&B.': 'Pas de données F&B.',
    'Reportes': 'Rapports',
    'Administrador': 'Administrateur',
    '👤 Administrador': '👤 Administrateur',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 Réclamations OTA en attente d'approbation",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "Lorsque Yve détectera des commissions facturées au-dessus du contrat, elles apparaîtront ici pour réclamation.",
    "✍️ Redactar con IA": "✍️ Rédiger avec l'IA",
    "✅ Aprobar y enviar": "✅ Approuver et envoyer",
    "🔄 Regenerar": "🔄 Régénérer",
    "🗑 Descartar": "🗑 Écarter",
    "Enviar a": "Envoyer à",
    "Descartada": "Écartée",
    "✓ Enviada": "✓ Envoyée",
    "email de la OTA": "e-mail de l'OTA",
  },
  de: {
    "🎭 Modo Demo personalizado": "🎭 Individueller Demo-Modus",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Gib dein Hotel, deine Kette oder mehrere Ketten ein und ich erzeuge realistische Beispieldaten mit diesen Namen: Rechnungen, Bank, F&B und Multi-Hotel. Ideal, um das Produkt Kunden und Steuerbüros zu zeigen.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "Eine Zeile pro Kette · Format \"Kette: Hotel 1, Hotel 2\"",
    "🎭 Generar demo": "🎭 Demo erzeugen",
    "Cancelar": "Abbrechen",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "Das System ist bereit, die Finanzen deines Hotels zu automatisieren.",
    "El primer paso: procesa las facturas OTA del mes en": "Erster Schritt: Verarbeite die OTA-Rechnungen des Monats in",
    "👋 Bienvenido a Yve.01": "👋 Willkommen bei Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "Das Hotelfinanz-System, das AR, AP, DRR und Reporting automatisiert. Diese Tour führt dich Modul für Modul von links nach rechts — 3 Minuten und du beherrschst alles. Zieh mich beiseite: Ich docke dort an, wo du mich loslässt.",
    "📥 AR — Comisiones OTA": "📥 AR — OTA-Provisionen",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Prüft automatisch die Provisionen von Booking.com und Expedia. Verarbeitete Rechnungen, Gesamtbetrag, reklamierbare Abweichungen und offene DI-Zertifikate. Die rote Zahl sind Euros, die du zurückholen kannst.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "Für jede Lieferantenrechnung gleicht Yve 3 Dokumente ab: Rechnung, Bestellung (PO) und Lieferschein. Stimmt alles → automatisches Match OK. Bei Differenz → Warnung und KI-generierte E-Mail an den Lieferanten.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Zieh deine .xlsm-Datei hierher. Yve extrahiert RevPAR, ADR, GOP%, Auslastung und die 7.000+ Zeilen der Trial Balance in Sekunden. Out of Balance wird automatisch erkannt und sofort gemeldet.",
    "🏦 Banco — Conciliación": "🏦 Bank — Abstimmung",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Gleicht den Kontoauszug automatisch mit den Lieferantenrechnungen ab. Erkennt nicht abgestimmte Bewegungen, Betragsdifferenzen und Doppelzahlungen. Von 8 Stunden auf 2 Minuten.",
    "🔔 Notificaciones": "🔔 Benachrichtigungen",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Richte automatische Alerts per E-Mail oder Telegram ein: OTA-Abweichungen, unsignierte Rechnungen, Out of Balance im DRR oder niedriger F&B-Bestand. Yve warnt dich proaktiv.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Berechnet den realen vs. theoretischen Food Cost pro Kategorie. Verbindet POS-Daten, Rezepte und Inventar. Erkennt Schwund, findet die margenstärksten Gerichte und optimiert die Restaurant-Performance.",
    "🏢 AR Real — Grupos Corporativos": "🏢 AR Real — Firmenkunden",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Komplette Verwaltung von Firmenkunden: Rechnungen ausstellen, Aging überwachen (0-30 / 31-60 / +90 Tage), mit einem Klick einziehen und automatische E-Mail-Erinnerungen senden.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hotel — Gruppenansicht",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "Für den Financial Controller der Gruppe: konsolidierte KPIs, Performance-Ranking pro Hotel, 6-Monats-Trend und zentrale Alerts. Ein Bildschirm, die ganze Gruppe.",
    "¡Ya conoces Yve.01!": "Du kennst jetzt Yve.01!",
    "Empezar con AR →": "Mit AR starten →",
    "Sin alertas bancarias pendientes.": "Keine offenen Bankwarnungen.",
    "Sin alertas bancarias.": "Keine Bankwarnungen.",
    "● Activo": "● Aktiv",
    "○ Inactivo": "○ Inaktiv",
    "Email de notificaciones": "Benachrichtigungs-E-Mail",
    "No hay facturas con este filtro": "Keine Rechnungen mit diesem Filter",
    "No hay hoteles en el grupo": "Keine Hotels in der Gruppe",
    "Mes actual": "Aktueller Monat",
    "🏆 Top Performers (RevPAR)": "🏆 Top-Hotels (RevPAR)",
    "Todos los hoteles": "Alle Hotels",
    "Hab.": "Zim.",
    "Ocup.": "Ausl.",
    "Categoría": "Kategorie",
    "Estado": "Status",
    "📄 Diario": "📄 Täglich",
    "📊 Semanal": "📊 Wöchentlich",
    "📈 Mensual": "📈 Monatlich",
    "🎯 Ejecutivo PDF": "🎯 Executive-PDF",
    "📊 Consolidado Excel": "📊 Konsolidiertes Excel",
    "📋 Historial de procesado": "📋 Verarbeitungsverlauf",
    "↻ Actualizar datos": "↻ Daten aktualisieren",
    "🎨 Personalizar colores": "🎨 Farben anpassen",
    "🌅 Briefing de hoy": "🌅 Briefing von heute",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ Welche offenen Abweichungen habe ich?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 Wie viel kann ich diesen Monat zurückfordern?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 Was braucht heute meine Unterschrift?",
    "Escribe aquí…": "Hier schreiben…",
    "Conciliado:": "Abgestimmt:",
    "FC% promedio": "Ø WEK %",
    "media ponderada": "gewichteter Durchschnitt",
    "Alertas FC alto": "Warnungen hoher WEK",
    "media del menú": "Menü-Durchschnitt",
    "Mejor margen": "Beste Marge",
    "menor FC%": "niedrigster WEK %",
    "FC% Medio": "Ø WEK %",
    "revisar urgente": "dringend prüfen",
    "Receta": "Rezept",
    "PVP": "VK-Preis",
    "Coste": "Kosten",
    "Margen": "Marge",
    "Crítico": "Kritisch",
    "Bajo": "Niedrig",
    "📸 Escanear Documento": "📸 Dokument scannen",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Fotografiere das physische Dokument (Rechnung, BEO, Vertrag, Kontoauszug...) und Yve liest es mit KI.",
    "📸 Cámara": "📸 Kamera",
    "🖼️ Galería": "🖼️ Galerie",
    "⚡ Procesar documento": "⚡ Dokument verarbeiten",
    "📸 Escanear": "📸 Scannen",
    "Cerrar": "Schließen",
    "📸 Escanear más": "📸 Mehr scannen",
    '% Com.': '% Prov.',
    '0 ya procesados (se saltarán)': '0 bereits verarbeitet (werden übersprungen)',
    '2 minutos para conocer todo Yve': '2 Minuten um Yve zu entdecken',
    'AP firma': 'AP-Unterschrift',
    'AP pendientes': 'Ausstehende AP',
    'AR Dashboard - Facturas Pendientes': 'AR-Dashboard — Ausstehende Rechnungen',
    'AR pend': 'Ausstehende AR',
    'ARCHIVOS SELECCIONADOS': 'AUSGEWÄHLTE DATEIEN',
    'Abrir chat Yve': 'Yve Chat öffnen',
    'Acceso en tiempo real a los datos del hotel': 'Echtzeitzugriff Hoteldaten',
    'Acciones': 'Aktionen',
    'Actualizar datos': 'Daten aktualisieren',
    'Ahora no': 'Nicht jetzt',
    'Alerta': 'Warnung',
    'Alertas': 'Warnungen',
    'Alertas Bancarias': 'Bank-Warnungen',
    'Alertas DRR': 'DRR-Warnungen',
    'Alto FC': 'Hohe LK',
    'Antigüedad de saldo': 'Forderungsalterung',
    'Aprobación': 'Genehmigung',
    'Aprobadas': 'Genehmigt',
    'Archivo': 'Datei',
    'Arrastra archivos aquí o haz clic': 'Dateien hier ablegen oder klicken',
    'Arrastra tu DRR aquí o': 'DRR hier ablegen oder',
    'Asunto': 'Betreff',
    'Bajo': 'Niedrig',
    'Búsqueda global': 'Globale Suche',
    'CATEGORÍA CON MÁS MERMA': 'KATEGORIE MIT DEN MEISTEN VERLUSTEN',
    'CLIENTES DE CRÉDITO': 'KREDITKUNDEN',
    'COBRADO MES': 'EINGEZOGEN MONAT',
    'COSTE TOTAL MERMAS': 'GESAMTKOSTEN VERLUSTE',
    'Calcular': 'Berechnen',
    'Cambiar de pestaña': 'Tab wechseln',
    'Cambiar rol': 'Rolle wechseln',
    'Canales de notificación': 'Benachrichtigungskanäle',
    'Cancelar': 'Abbrechen',
    'Cantidad': 'Menge',
    'Cargando datos...': 'Daten laden...',
    'Cargando...': 'Laden...',
    'Categoría': 'Kategorie',
    'Causa': 'Ursache',
    'Cerrar': 'Schließen',
    'Cerrar modales': 'Dialoge schließen',
    'Certif. DI pendiente': 'DI-Zert. ausstehend',
    'Cliente': 'Kunde',
    'Clientes': 'Kunden',
    'Clientes Directos': 'Direktkunden',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Kreditkunden · Unternehmensabrechnung · Inkassokontrolle',
    'Cobradas': 'Eingezogen',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Schließe diese Schritte ab, um Yve.01 optimal zu nutzen',
    'Con avisos': 'Mit Warnungen',
    'Concepto': 'Konzept',
    'Conciliado:': 'Abgestimmt:',
    'Conciliados': 'Abgestimmt',
    'Configura dónde recibes alertas de Yve': 'Konfiguriere Benachrichtigungsorte',
    'Correctas': 'Korrekt',
    'Coste': 'Kosten',
    'Crítico': 'Kritisch',
    'Críticos': 'Kritisch',
    'Cuenta': 'Konto',
    'Destinatario': 'Empfänger',
    'Detalle factura': 'Rechnungsdetail',
    'Diferencias': 'Unterschiede',
    'Discrepancia': 'Abweichung',
    'Discrepancias': 'Abweichungen',
    'Días': 'Tage',
    'EN SERVIDOR (facturas-entrada)': 'AUF SERVER (Rechnungen-Eingang)',
    'EUR procesados': 'EUR verarbeitet',
    'Empezar →': 'Starten →',
    'Entrada': 'Eingang',
    'Error inventario': 'Inventarfehler',
    'Error mermas': 'Verlustfehler',
    'Error recetas': 'Rezeptfehler',
    'Estado': 'Status',
    'Eventos que disparan alerta': 'Auslöseende Ereignisse',
    'Explorar sin tour': 'Ohne Tour erkunden',
    'F&B + OTRAS': 'F&B + ANDERE',
    'F&B Cost %': 'F&B Kosten %',
    'F&B Cost Control': 'F&B Kostenkontrolle',
    'FC%': 'LK%',
    'Factura': 'Rechnung',
    'Facturas': 'Rechnungen',
    'Fecha': 'Datum',
    'Ficha de Recetas con Coste Teórico': 'Rezeptkarten mit Theoretischen Kosten',
    'Food Cost % — Teórico vs Real': 'Lebensmittelkosten % — Soll vs Ist',
    'Food Cost por Categoría': 'Lebensmittelkosten nach Kategorie',
    'Forecast': 'Prognose',
    'GOP% MEDIO': 'DURCHSCHN. GOP%',
    'Generar email': 'E-Mail generieren',
    'Grupo': 'Gruppe',
    'Grupos Corporativos': 'Unternehmensgruppen',
    'Guardar': 'Speichern',
    'Habitaciones': 'Zimmer',
    'Historial de Mermas · Total:': 'Verlaufsprotokoll · Gesamt:',
    'Hoteles OK': 'Hotels OK',
    'Importe': 'Betrag',
    'KPIs Operativos': 'Operative KPIs',
    'MEJOR GOP%': 'BESTES GOP%',
    'Margen': 'Marge',
    'Mermas por Causa': 'Verluste nach Ursache',
    'Mostrar atajos': 'Tastenkürzel zeigen',
    'Nueva factura': 'Neue Rechnung',
    'Ocupacion': 'Auslastung',
    'Ocupación': 'Auslastung',
    'PVP': 'UVP',
    'Pendiente': 'Ausstehend',
    'Performance Financiero': 'Finanzielle Leistung',
    'Proveedor': 'Lieferant',
    'REGISTROS': 'EINTRÄGE',
    'REVENUE': 'UMSATZ',
    'Registrar Merma': 'Verlust erfassen',
    'Revenue MTD': 'Umsatz MTD',
    'Revenue hoy vs Budget:': 'Heutiger Umsatz vs Budget:',
    'Seleccionar cliente...': 'Kunde auswählen...',
    'Siguiente →': 'Weiter →',
    'Sin alertas activas': 'Keine aktiven Warnungen',
    'Sin datos': 'Keine Daten',
    'Sin facturas AP': 'Keine AP-Rechnungen',
    'Stock': 'Bestand',
    'Stock de Ingredientes': 'Zutatenbestand',
    'Tipo': 'Typ',
    'Total': 'Gesamt',
    'Yve — Copiloto Financiero': 'Yve — Finanz-Copilot',
    '¿Hacemos el tour?': 'Tour machen?',
    '← Atrás': '← Zurück',
    '✓ Finalizar': '✓ Fertig',
  
    '⌨ Atajos': '⌨ Tastenkürzel',
    '📊 Vista lite': '📊 Kompaktansicht',
    '📊 Vista completa': '📊 Vollansicht',
    '📱 Vista lite': '📱 Kompaktansicht',
    '⚡ Procesar Archivos': '⚡ Dateien verarbeiten',
    '⚡ Procesar archivos nuevos': '⚡ Neue Dateien verarbeiten',
    '▶ Procesar pendientes del servidor': '▶ Server-Ausstehende verarbeiten',
    '✅ Aprobar Match OK': '✅ Match OK genehmigen',
    '📲 Aprobar facturas AR': '📲 AR genehmigen',
    '⚡ Conciliar': '⚡ Abstimmen',
    '🔍 Filtrar': '🔍 Filtern',
    'Sin facturas AR todavía': 'Noch keine AR-Rechnungen',
    'Sin datos.': 'Keine Daten.',
    'Pendiente emitir': 'Ausstehend',
    'Pendientes': 'Ausstehend',
    'Todas las facturas del ciclo': 'Alle Rechnungen des Zyklus',
    'Pulsa ⚡ Procesar Archivos': '⚡ AP-Rechnungen verarbeiten',
    'Pulsa ⚡ Procesar Archivos.': '⚡ Dateien verarbeiten.',
    '📊 Resumen': '📊 Übersicht',
    'Vista consolidada del grupo': 'Konsolidierte Gruppenansicht',
    '⚠️ Alertas activas': '⚠️ Aktive Warnungen',
    '⌨ Atajos de teclado': '⌨ Tastenkürzel',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Globale Suche Ctrl+K · ⌨ Tastenkürzel 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Kompaktansicht auf allen Panels',
    'PDF (facturas) · XLSM (DRR)': 'PDF (Rechnungen) · XLSM (DRR)',
    'Actualizado': 'Aktualisiert',
    'facturas cargadas': 'Rechnungen geladen',
    'Todos': 'Alle',
    'Sin archivo cargado': 'Keine Datei geladen',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'klicken zum Auswählen (.xlsm/.xlsx)',
    'Revenue Diario': 'Tagesumsatz',
    'Trial Balance — Estado Diario': 'Probebilanz — Tagesstatus',
    'Sin alertas.': 'Keine Warnungen.',
    'Historial de Notificaciones': 'Benachrichtigungsverlauf',
    'Sin notificaciones.': 'Keine Benachrichtigungen.',
    'Historial de notificaciones enviadas': 'Gesendete Benachrichtigungen',
    'Ingrediente': 'Zutat',
    'Receta': 'Rezept',
    '▶ Ejecutar Análisis': '▶ Analyse ausführen',
    'Sin datos F&B.': 'Keine F&B-Daten.',
    'Reportes': 'Berichte',
    'Administrador': 'Administrator',
    '👤 Administrador': '👤 Administrator',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 OTA-Reklamationen zur Freigabe",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "Sobald Yve über dem Vertrag berechnete Provisionen erkennt, erscheinen sie hier zur Reklamation.",
    "✍️ Redactar con IA": "✍️ Mit KI verfassen",
    "✅ Aprobar y enviar": "✅ Freigeben und senden",
    "🔄 Regenerar": "🔄 Neu erzeugen",
    "🗑 Descartar": "🗑 Verwerfen",
    "Enviar a": "Senden an",
    "Descartada": "Verworfen",
    "✓ Enviada": "✓ Gesendet",
    "email de la OTA": "E-Mail der OTA",
  },
  it: {
    "🎭 Modo Demo personalizado": "🎭 Modalità Demo personalizzata",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Scrivi il tuo hotel, la tua catena o più catene e genero dati di esempio realistici con quei nomi: fatture, banca, F&B e Multi-Hotel. Ideale per mostrare il prodotto a clienti e studi contabili.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "Una riga per catena · formato \"Catena: Hotel 1, Hotel 2\"",
    "🎭 Generar demo": "🎭 Genera demo",
    "Cancelar": "Annulla",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "Il sistema è pronto ad automatizzare le finanze del tuo hotel.",
    "El primer paso: procesa las facturas OTA del mes en": "Primo passo: elabora le fatture OTA del mese in",
    "👋 Bienvenido a Yve.01": "👋 Benvenuto in Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "Il sistema di finanza alberghiera che automatizza AR, AP, DRR e reporting. Questo tour ti guida modulo per modulo — 3 minuti e padroneggi tutto. Trascinami se ti intralcio: mi aggancio dove mi lasci.",
    "📥 AR — Comisiones OTA": "📥 AR — Commissioni OTA",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Verifica automaticamente le commissioni di Booking.com ed Expedia. Fatture elaborate, importo totale, discrepanze reclamabili e certificati DI in sospeso. Il numero rosso sono euro che puoi recuperare.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "Per ogni fattura fornitore, Yve incrocia 3 documenti: fattura, ordine (PO) e bolla di consegna. Se tutto quadra → Match OK automatico. Se c'è differenza → avviso ed email al fornitore generata con IA.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Trascina qui il tuo file .xlsm. Yve estrae RevPAR, ADR, GOP%, occupazione e le oltre 7.000 righe del Trial Balance in pochi secondi. Rileva l'Out of Balance automaticamente e ti avvisa subito.",
    "🏦 Banco — Conciliación": "🏦 Banca — Riconciliazione",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Incrocia automaticamente l'estratto conto con le fatture dei fornitori. Identifica movimenti non riconciliati, differenze di importo e pagamenti duplicati. Da 8 ore a 2 minuti.",
    "🔔 Notificaciones": "🔔 Notifiche",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Configura avvisi automatici via email o Telegram: discrepanze OTA, fatture non firmate, Out of Balance nel DRR o scorte F&B basse. Yve ti avvisa proattivamente.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Calcola il Food Cost reale vs teorico per categoria. Collega dati POS, ricette e inventario. Rileva sprechi, individua i piatti con il margine migliore e ottimizza le prestazioni del ristorante.",
    "🏢 AR Real — Grupos Corporativos": "🏢 AR Reale — Gruppi Corporate",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Gestione completa dei clienti corporate: emetti fatture, controlla l'aging (0-30 / 31-60 / +90 giorni), incassa con un clic e invia promemoria automatici via email.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hotel — Vista di Gruppo",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "Per il Financial Controller del gruppo: KPI consolidati, ranking di performance per hotel, trend a 6 mesi e avvisi centralizzati. Uno schermo, tutto il gruppo.",
    "¡Ya conoces Yve.01!": "Ora conosci Yve.01!",
    "Empezar con AR →": "Inizia con AR →",
    "Sin alertas bancarias pendientes.": "Nessun avviso bancario in sospeso.",
    "Sin alertas bancarias.": "Nessun avviso bancario.",
    "● Activo": "● Attivo",
    "○ Inactivo": "○ Inattivo",
    "Email de notificaciones": "Email di notifica",
    "No hay facturas con este filtro": "Nessuna fattura con questo filtro",
    "No hay hoteles en el grupo": "Nessun hotel nel gruppo",
    "Mes actual": "Mese corrente",
    "🏆 Top Performers (RevPAR)": "🏆 Migliori hotel (RevPAR)",
    "Todos los hoteles": "Tutti gli hotel",
    "Hab.": "Cam.",
    "Ocup.": "Occ.",
    "Categoría": "Categoria",
    "Estado": "Stato",
    "📄 Diario": "📄 Giornaliero",
    "📊 Semanal": "📊 Settimanale",
    "📈 Mensual": "📈 Mensile",
    "🎯 Ejecutivo PDF": "🎯 PDF esecutivo",
    "📊 Consolidado Excel": "📊 Excel consolidato",
    "📋 Historial de procesado": "📋 Cronologia elaborazioni",
    "↻ Actualizar datos": "↻ Aggiorna dati",
    "🎨 Personalizar colores": "🎨 Personalizza colori",
    "🌅 Briefing de hoy": "🌅 Briefing di oggi",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ Quali discrepanze ho aperte?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 Quanto posso reclamare questo mese?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 Cosa richiede la mia firma oggi?",
    "Escribe aquí…": "Scrivi qui…",
    "Conciliado:": "Riconciliato:",
    "FC% promedio": "FC% medio",
    "media ponderada": "media ponderata",
    "Alertas FC alto": "Avvisi FC alto",
    "media del menú": "media del menu",
    "Mejor margen": "Miglior margine",
    "menor FC%": "FC% più basso",
    "FC% Medio": "FC% medio",
    "revisar urgente": "verificare urgentemente",
    "Receta": "Ricetta",
    "PVP": "Prezzo",
    "Coste": "Costo",
    "Margen": "Margine",
    "Crítico": "Critico",
    "Bajo": "Basso",
    "📸 Escanear Documento": "📸 Scansiona documento",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Scatta una foto al documento fisico (fattura, BEO, contratto, estratto...) e Yve lo leggerà con l'IA.",
    "📸 Cámara": "📸 Fotocamera",
    "🖼️ Galería": "🖼️ Galleria",
    "⚡ Procesar documento": "⚡ Elabora documento",
    "📸 Escanear": "📸 Scansiona",
    "Cerrar": "Chiudi",
    "📸 Escanear más": "📸 Scansiona altro",
    '% Com.': '% Comm.',
    '0 ya procesados (se saltarán)': '0 già elaborati (verranno saltati)',
    '2 minutos para conocer todo Yve': '2 minuti per scoprire Yve',
    'AP firma': 'Firma AP',
    'AP pendientes': 'AP in attesa',
    'AR Dashboard - Facturas Pendientes': 'AR Dashboard — Fatture in attesa',
    'AR pend': 'AR in attesa',
    'ARCHIVOS SELECCIONADOS': 'FILE SELEZIONATI',
    'Abrir chat Yve': 'Apri chat Yve',
    'Acceso en tiempo real a los datos del hotel': 'Accesso in tempo reale ai dati',
    'Acciones': 'Azioni',
    'Actualizar datos': 'Aggiorna dati',
    'Ahora no': 'Non ora',
    'Alerta': 'Avviso',
    'Alertas': 'Avvisi',
    'Alertas Bancarias': 'Avvisi Bancari',
    'Alertas DRR': 'Avvisi DRR',
    'Alto FC': 'FC alto',
    'Antigüedad de saldo': 'Anzianità crediti',
    'Aprobación': 'Approvazione',
    'Aprobadas': 'Approvate',
    'Archivo': 'File',
    'Arrastra archivos aquí o haz clic': 'Trascina file qui o fai clic',
    'Arrastra tu DRR aquí o': 'Trascina qui il DRR o',
    'Asunto': 'Oggetto',
    'Bajo': 'Basso',
    'Búsqueda global': 'Ricerca globale',
    'CATEGORÍA CON MÁS MERMA': 'CATEGORIA CON PIÙ SPRECHI',
    'CLIENTES DE CRÉDITO': 'CLIENTI A CREDITO',
    'COBRADO MES': 'INCASSATO MESE',
    'COSTE TOTAL MERMAS': 'COSTO TOTALE SPRECHI',
    'Calcular': 'Calcola',
    'Cambiar de pestaña': 'Cambia scheda',
    'Cambiar rol': 'Cambia ruolo',
    'Canales de notificación': 'Canali notifica',
    'Cancelar': 'Annulla',
    'Cantidad': 'Quantità',
    'Cargando datos...': 'Caricamento dati...',
    'Cargando...': 'Caricamento...',
    'Categoría': 'Categoria',
    'Cerrar': 'Chiudi',
    'Cerrar modales': 'Chiudi modali',
    'Certif. DI pendiente': 'Cert. DI in attesa',
    'Clientes': 'Clienti',
    'Clientes Directos': 'Clienti Diretti',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Clienti credito · Fatturazione aziendale · Controllo incassi',
    'Cobradas': 'Incassate',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Completa questi passaggi per sfruttare al meglio Yve.01',
    'Con avisos': 'Con avvisi',
    'Concepto': 'Concetto',
    'Conciliado:': 'Riconciliato:',
    'Conciliados': 'Riconciliati',
    'Configura dónde recibes alertas de Yve': 'Configura dove ricevi gli avvisi',
    'Correctas': 'Corrette',
    'Coste': 'Costo',
    'Crítico': 'Critico',
    'Críticos': 'Critici',
    'Cuenta': 'Conto',
    'Detalle factura': 'Dettaglio fattura',
    'Diferencias': 'Differenze',
    'Discrepancia': 'Discrepanza',
    'Discrepancias': 'Discrepanze',
    'Días': 'Giorni',
    'EN SERVIDOR (facturas-entrada)': 'SUL SERVER (fatture-entrata)',
    'EUR procesados': 'EUR elaborati',
    'Empezar →': 'Inizia →',
    'Entrada': 'Ingresso',
    'Error inventario': 'Errore inventario',
    'Error mermas': 'Errore sprechi',
    'Error recetas': 'Errore ricette',
    'Estado': 'Stato',
    'Eventos que disparan alerta': 'Eventi scatenanti',
    'Explorar sin tour': 'Esplora senza tour',
    'F&B + OTRAS': 'F&B + ALTRO',
    'F&B Cost %': 'Costo F&B %',
    'F&B Cost Control': 'Controllo Costi F&B',
    'Factura': 'Fattura',
    'Facturas': 'Fatture',
    'Fecha': 'Data',
    'Ficha de Recetas con Coste Teórico': 'Schede Ricette con Costo Teorico',
    'Food Cost % — Teórico vs Real': 'Costo Cibo % — Teorico vs Reale',
    'Food Cost por Categoría': 'Costo Cibo per Categoria',
    'Forecast': 'Previsione',
    'Generar email': 'Genera email',
    'Grupo': 'Gruppo',
    'Grupos Corporativos': 'Gruppi Aziendali',
    'Guardar': 'Salva',
    'Habitaciones': 'Camere',
    'Historial de Mermas · Total:': 'Storico Sprechi · Totale:',
    'Hoteles OK': 'Hotel OK',
    'Importe': 'Importo',
    'KPIs Operativos': 'KPI Operativi',
    'MEJOR GOP%': 'MIGLIOR GOP%',
    'Margen': 'Margine',
    'Mermas por Causa': 'Sprechi per Causa',
    'Mostrar atajos': 'Mostra scorciatoie',
    'Nueva factura': 'Nuova fattura',
    'Ocupacion': 'Occupazione',
    'Ocupación': 'Occupazione',
    'Pendiente': 'In attesa',
    'Performance Financiero': 'Performance Finanziaria',
    'Proveedor': 'Fornitore',
    'REGISTROS': 'REGISTRI',
    'REVENUE': 'RICAVI',
    'Registrar Merma': 'Registra Sprechi',
    'Revenue MTD': 'Ricavi MTD',
    'Revenue hoy vs Budget:': 'Ricavi oggi vs Budget:',
    'Seleccionar cliente...': 'Seleziona cliente...',
    'Siguiente →': 'Avanti →',
    'Sin alertas activas': 'Nessun avviso attivo',
    'Sin datos': 'Nessun dato',
    'Sin facturas AP': 'Nessuna fattura AP',
    'Stock de Ingredientes': 'Scorte Ingredienti',
    'Total': 'Totale',
    'Yve — Copiloto Financiero': 'Yve — Copilota Finanziario',
    '¿Hacemos el tour?': 'Fare il tour?',
    '← Atrás': '← Indietro',
    '✓ Finalizar': '✓ Fine',
  
    '⌨ Atajos': '⌨ Scorciatoie',
    '📊 Vista lite': '📊 Vista compatta',
    '📱 Vista lite': '📱 Vista compatta',
    '⚡ Procesar Archivos': '⚡ Elabora File',
    '⚡ Procesar archivos nuevos': '⚡ Elabora nuovi file',
    '▶ Procesar pendientes del servidor': '▶ Elabora in attesa server',
    '✅ Aprobar Match OK': '✅ Approva Match OK',
    '📲 Aprobar facturas AR': '📲 Approva fatture AR',
    '⚡ Conciliar': '⚡ Riconcilia',
    '🔍 Filtrar': '🔍 Filtra',
    'Sin facturas AR todavía': 'Ancora nessuna fattura AR',
    'Sin datos.': 'Nessun dato.',
    'Pendiente emitir': 'In attesa emissione',
    'Pendientes': 'In attesa',
    'Todas las facturas del ciclo': 'Tutte le fatture del ciclo',
    'Pulsa ⚡ Procesar Archivos': 'Premi ⚡ Elabora Fatture AP',
    'Pulsa ⚡ Procesar Archivos.': 'Premi ⚡ Elabora File.',
    '📊 Resumen': '📊 Riepilogo',
    'Vista consolidada del grupo': 'Vista consolidata del gruppo',
    '⚠️ Alertas activas': '⚠️ Avvisi attivi',
    '⌨ Atajos de teclado': '⌨ Scorciatoie',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Ricerca globale Ctrl+K · ⌨ Scorciatoie 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Vista compatta su tutti i pannelli',
    'PDF (facturas) · XLSM (DRR)': 'PDF (fatture) · XLSM (DRR)',
    'Actualizado': 'Aggiornato',
    'facturas cargadas': 'fatture caricate',
    'Todos': 'Tutti',
    'Sin archivo cargado': 'Nessun file caricato',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'clicca per selezionare (.xlsm/.xlsx)',
    'Revenue Diario': 'Ricavi Giornalieri',
    'Trial Balance — Estado Diario': 'Bilancio — Stato Giornaliero',
    'Sin alertas.': 'Nessun avviso.',
    'Historial de Notificaciones': 'Storico Notifiche',
    'Sin notificaciones.': 'Nessuna notifica.',
    'Historial de notificaciones enviadas': 'Storico notifiche inviate',
    'Receta': 'Ricetta',
    '▶ Ejecutar Análisis': '▶ Esegui Analisi',
    'Sin datos F&B.': 'Nessun dato F&B.',
    'Reportes': 'Report',
    'Administrador': 'Amministratore',
    '👤 Administrador': '👤 Amministratore',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 Reclami OTA in attesa di approvazione",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "Quando Yve rileverà commissioni addebitate oltre il contratto, appariranno qui per il reclamo.",
    "✍️ Redactar con IA": "✍️ Redigi con l'IA",
    "✅ Aprobar y enviar": "✅ Approva e invia",
    "🔄 Regenerar": "🔄 Rigenera",
    "🗑 Descartar": "🗑 Scarta",
    "Enviar a": "Invia a",
    "Descartada": "Scartata",
    "✓ Enviada": "✓ Inviata",
    "email de la OTA": "email dell'OTA",
  },
  pt: {
    "🎭 Modo Demo personalizado": "🎭 Modo Demo personalizado",
    "Escribe tu hotel, tu cadena o varias cadenas y genero datos de ejemplo realistas con esos nombres: facturas, banco, F&B y Multi-Hotel. Ideal para enseñar el producto a clientes y gestorías.": "Escreva seu hotel, sua rede ou várias redes e eu gero dados de exemplo realistas com esses nomes: faturas, banco, F&B e Multi-Hotel. Ideal para mostrar o produto a clientes e escritórios contábeis.",
    "Una línea por cadena · formato \"Cadena: Hotel 1, Hotel 2\"": "Uma linha por rede · formato \"Rede: Hotel 1, Hotel 2\"",
    "🎭 Generar demo": "🎭 Gerar demo",
    "Cancelar": "Cancelar",
    "El sistema está listo para automatizar las finanzas de tu hotel.": "O sistema está pronto para automatizar as finanças do seu hotel.",
    "El primer paso: procesa las facturas OTA del mes en": "Primeiro passo: processe as faturas OTA do mês em",
    "👋 Bienvenido a Yve.01": "👋 Bem-vindo ao Yve.01",
    "El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.": "O sistema de finanças hoteleiras que automatiza AR, AP, DRR e reporting. Este tour percorre cada módulo da esquerda para a direita — 3 minutos e você domina tudo. Arraste-me se atrapalhar: eu me encaixo onde você me soltar.",
    "📥 AR — Comisiones OTA": "📥 AR — Comissões OTA",
    "Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.": "Verifica automaticamente as comissões de Booking.com e Expedia. Faturas processadas, valor total, discrepâncias reclamáveis e certificados DI pendentes. O número vermelho são euros que você pode recuperar.",
    "Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.": "Para cada fatura de fornecedor, o Yve cruza 3 documentos: fatura, pedido (PO) e guia de remessa. Tudo confere → Match OK automático. Diferença → alerta e email ao fornecedor gerado com IA.",
    "Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.": "Arraste seu arquivo .xlsm aqui. O Yve extrai RevPAR, ADR, GOP%, ocupação e as mais de 7.000 linhas do Trial Balance em segundos. Detecta Out of Balance automaticamente e avisa na hora.",
    "🏦 Banco — Conciliación": "🏦 Banco — Conciliação",
    "Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.": "Cruza automaticamente o extrato bancário com as faturas de fornecedores. Identifica movimentos não conciliados, diferenças de valor e pagamentos duplicados. De 8 horas para 2 minutos.",
    "🔔 Notificaciones": "🔔 Notificações",
    "Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.": "Configure alertas automáticos por email ou Telegram: discrepâncias OTA, faturas sem assinatura, Out of Balance no DRR ou estoque baixo no F&B. O Yve avisa você proativamente.",
    "Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.": "Calcula o Food Cost real vs teórico por categoria. Conecta dados do POS, receitas e inventário. Detecta perdas, identifica os pratos com melhor margem e otimiza o desempenho do restaurante.",
    "🏢 AR Real — Grupos Corporativos": "🏢 AR Real — Grupos Corporativos",
    "Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.": "Gestão completa de clientes corporativos: emita faturas, controle o aging (0-30 / 31-60 / +90 dias), cobre com um clique e envie lembretes automáticos por email.",
    "🌍 Multi-Hotel — Vista de Grupo": "🌍 Multi-Hotel — Visão de Grupo",
    "Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.": "Para o Financial Controller do grupo: KPIs consolidados, ranking de desempenho por hotel, tendência de 6 meses e alertas centralizados. Uma tela, todo o grupo.",
    "¡Ya conoces Yve.01!": "Você já conhece o Yve.01!",
    "Empezar con AR →": "Começar com AR →",
    "Sin alertas bancarias pendientes.": "Sem alertas bancários pendentes.",
    "Sin alertas bancarias.": "Sem alertas bancários.",
    "● Activo": "● Ativo",
    "○ Inactivo": "○ Inativo",
    "Email de notificaciones": "Email de notificações",
    "No hay facturas con este filtro": "Sem faturas com este filtro",
    "No hay hoteles en el grupo": "Sem hotéis no grupo",
    "Mes actual": "Mês atual",
    "🏆 Top Performers (RevPAR)": "🏆 Melhores hotéis (RevPAR)",
    "Todos los hoteles": "Todos os hotéis",
    "Hab.": "Qua.",
    "Ocup.": "Ocup.",
    "Categoría": "Categoria",
    "Estado": "Estado",
    "📄 Diario": "📄 Diário",
    "📊 Semanal": "📊 Semanal",
    "📈 Mensual": "📈 Mensal",
    "🎯 Ejecutivo PDF": "🎯 PDF executivo",
    "📊 Consolidado Excel": "📊 Excel consolidado",
    "📋 Historial de procesado": "📋 Histórico de processamento",
    "↻ Actualizar datos": "↻ Atualizar dados",
    "🎨 Personalizar colores": "🎨 Personalizar cores",
    "🌅 Briefing de hoy": "🌅 Briefing de hoje",
    "⚠️ ¿Qué discrepancias tengo abiertas?": "⚠️ Que discrepâncias tenho em aberto?",
    "💰 ¿Cuánto puedo reclamar este mes?": "💰 Quanto posso reclamar este mês?",
    "📋 ¿Qué necesita mi firma hoy?": "📋 O que precisa da minha assinatura hoje?",
    "Escribe aquí…": "Escreva aqui…",
    "Conciliado:": "Conciliado:",
    "FC% promedio": "FC% médio",
    "media ponderada": "média ponderada",
    "Alertas FC alto": "Alertas FC alto",
    "media del menú": "média do menu",
    "Mejor margen": "Melhor margem",
    "menor FC%": "menor FC%",
    "FC% Medio": "FC% médio",
    "revisar urgente": "revisar urgente",
    "Receta": "Receita",
    "PVP": "PVP",
    "Coste": "Custo",
    "Margen": "Margem",
    "Crítico": "Crítico",
    "Bajo": "Baixo",
    "📸 Escanear Documento": "📸 Escanear Documento",
    "Haz una foto al documento físico (factura, BEO, contrato, extracto...) y Yve lo leerá con IA.": "Tire uma foto do documento físico (fatura, BEO, contrato, extrato...) e a Yve o lerá com IA.",
    "📸 Cámara": "📸 Câmera",
    "🖼️ Galería": "🖼️ Galeria",
    "⚡ Procesar documento": "⚡ Processar documento",
    "📸 Escanear": "📸 Escanear",
    "Cerrar": "Fechar",
    "📸 Escanear más": "📸 Escanear mais",
    '0 ya procesados (se saltarán)': '0 já processados (serão ignorados)',
    '2 minutos para conocer todo Yve': '2 minutos para conhecer Yve',
    'AP firma': 'Assinatura AP',
    'AP pendientes': 'AP pendentes',
    'AR Dashboard - Facturas Pendientes': 'AR Dashboard — Faturas Pendentes',
    'AR pend': 'AR pendente',
    'ARCHIVOS SELECCIONADOS': 'ARQUIVOS SELECIONADOS',
    'Acceso en tiempo real a los datos del hotel': 'Acesso em tempo real aos dados',
    'Acciones': 'Ações',
    'Actualizar datos': 'Atualizar dados',
    'Ahora no': 'Agora não',
    'Alertas Bancarias': 'Alertas Bancários',
    'Alto FC': 'FC alto',
    'Antigüedad de saldo': 'Envelhecimento',
    'Aprobación': 'Aprovação',
    'Aprobadas': 'Aprovadas',
    'Archivo': 'Arquivo',
    'Arrastra archivos aquí o haz clic': 'Arraste arquivos aqui ou clique',
    'Arrastra tu DRR aquí o': 'Arraste o DRR aqui ou',
    'Asunto': 'Assunto',
    'Bajo': 'Baixo',
    'Budget': 'Orçamento',
    'Búsqueda global': 'Pesquisa global',
    'CATEGORÍA CON MÁS MERMA': 'CATEGORIA COM MAIS PERDA',
    'CLIENTES DE CRÉDITO': 'CLIENTES DE CRÉDITO',
    'COBRADO MES': 'COBRADO MÊS',
    'COSTE TOTAL MERMAS': 'CUSTO TOTAL PERDAS',
    'Cambiar de pestaña': 'Mudar aba',
    'Cambiar rol': 'Mudar papel',
    'Canales de notificación': 'Canais notificação',
    'Cantidad': 'Quantidade',
    'Cargando datos...': 'Carregando dados...',
    'Cargando...': 'Carregando...',
    'Categoría': 'Categoria',
    'Cerrar': 'Fechar',
    'Cerrar modales': 'Fechar modais',
    'Certif. DI pendiente': 'Certif. DI pendente',
    'Clientes Directos': 'Clientes Diretos',
    'Clientes de crédito · Facturación corporativa · Control de cobros': 'Clientes crédito · Faturamento corporativo · Controle cobranças',
    'Completa estos pasos para sacar el máximo partido a Yve.01': 'Complete estas etapas para aproveitar ao máximo o Yve.01',
    'Con avisos': 'Com avisos',
    'Concepto': 'Conceito',
    'Configura dónde recibes alertas de Yve': 'Configure onde receber alertas',
    'Correctas': 'Corretas',
    'Coste': 'Custo',
    'Cuenta': 'Conta',
    'Destinatario': 'Destinatário',
    'Detalle factura': 'Detalhe fatura',
    'Diferencias': 'Diferenças',
    'Discrepancia': 'Discrepância',
    'Discrepancias': 'Discrepâncias',
    'Días': 'Dias',
    'EN SERVIDOR (facturas-entrada)': 'NO SERVIDOR (faturas-entrada)',
    'EUR procesados': 'EUR processados',
    'Empezar →': 'Começar →',
    'Error inventario': 'Erro inventário',
    'Error mermas': 'Erro perdas',
    'Error recetas': 'Erro receitas',
    'Eventos que disparan alerta': 'Eventos disparadores',
    'Explorar sin tour': 'Explorar sem tour',
    'F&B + OTRAS': 'F&B + OUTRAS',
    'F&B Cost %': 'Custo F&B %',
    'F&B Cost Control': 'Controle Custo F&B',
    'Factura': 'Fatura',
    'Facturas': 'Faturas',
    'Fecha': 'Data',
    'Ficha de Recetas con Coste Teórico': 'Fichas Receitas com Custo Teórico',
    'Food Cost % — Teórico vs Real': 'Custo Alim. % — Teórico vs Real',
    'Food Cost por Categoría': 'Custo Alimentar por Categoria',
    'Forecast': 'Previsão',
    'GOP% MEDIO': 'GOP% MÉDIO',
    'Generar email': 'Gerar email',
    'Guardar': 'Salvar',
    'Habitaciones': 'Quartos',
    'Historial de Mermas · Total:': 'Histórico Perdas · Total:',
    'Hoteles OK': 'Hotéis OK',
    'Importe': 'Valor',
    'KPIs Operativos': 'KPIs Operacionais',
    'MEJOR GOP%': 'MELHOR GOP%',
    'Margen': 'Margem',
    'Mermas por Causa': 'Perdas por Causa',
    'Mostrar atajos': 'Mostrar atalhos',
    'Nueva factura': 'Nova fatura',
    'Ocupacion': 'Ocupação',
    'Ocupación': 'Ocupação',
    'Pendiente': 'Pendente',
    'Performance Financiero': 'Desempenho Financeiro',
    'Proveedor': 'Fornecedor',
    'REVENUE': 'RECEITA',
    'Registrar Merma': 'Registrar Perda',
    'Revenue MTD': 'Receita MTD',
    'Revenue hoy vs Budget:': 'Receita hoje vs Orçamento:',
    'Seleccionar cliente...': 'Selecionar cliente...',
    'Siguiente →': 'Próximo →',
    'Sin alertas activas': 'Sem alertas ativos',
    'Sin datos': 'Sem dados',
    'Sin facturas AP': 'Sem faturas AP',
    'Stock': 'Estoque',
    'Stock de Ingredientes': 'Estoque de Ingredientes',
    'Yve — Copiloto Financiero': 'Yve — Copiloto Financeiro',
    '¿Hacemos el tour?': 'Fazer o tour?',
    '← Atrás': '← Voltar',
  
    '⌨ Atajos': '⌨ Atalhos',
    '📊 Vista lite': '📊 Vista compacta',
    '📱 Vista lite': '📱 Vista compacta',
    '⚡ Procesar Archivos': '⚡ Processar Ficheiros',
    '⚡ Procesar archivos nuevos': '⚡ Processar novos arquivos',
    '▶ Procesar pendientes del servidor': '▶ Processar pendentes servidor',
    '✅ Aprobar Match OK': '✅ Aprovar Match OK',
    '📲 Aprobar facturas AR': '📲 Aprovar faturas AR',
    'Sin facturas AR todavía': 'Sem faturas AR ainda',
    'Sin datos.': 'Sem dados.',
    'Pendiente emitir': 'Pendente emitir',
    'Pendientes': 'Pendentes',
    'Todas las facturas del ciclo': 'Todas as faturas do ciclo',
    'Pulsa ⚡ Procesar Archivos': 'Pressione ⚡ Processar Faturas AP',
    'Pulsa ⚡ Procesar Archivos.': 'Pressione ⚡ Processar Ficheiros.',
    '📊 Resumen': '📊 Resumo',
    'Vista consolidada del grupo': 'Vista consolidada do grupo',
    '⚠️ Alertas activas': '⚠️ Alertas ativos',
    '⌨ Atajos de teclado': '⌨ Atalhos de teclado',
    '🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9': '🔍 Pesquisa global Ctrl+K · ⌨ Atalhos 1-9',
    '📱 Vista lite en todos los paneles (F&B, Real AR, Multi)': '📱 Vista compacta em todos os painéis',
    'PDF (facturas) · XLSM (DRR)': 'PDF (faturas) · XLSM (DRR)',
    'Actualizado': 'Atualizado',
    'facturas cargadas': 'faturas carregadas',
    'Sin archivo cargado': 'Nenhum arquivo carregado',
    'haz clic para seleccionar (.xlsm/.xlsx)': 'clique para selecionar (.xlsm/.xlsx)',
    'Revenue Diario': 'Receita Diária',
    'Trial Balance — Estado Diario': 'Balancete — Estado Diário',
    'Sin alertas.': 'Sem alertas.',
    'Historial de Notificaciones': 'Histórico de Notificações',
    'Sin notificaciones.': 'Sem notificações.',
    'Historial de notificaciones enviadas': 'Histórico de notificações enviadas',
    'Receta': 'Receita',
    '▶ Ejecutar Análisis': '▶ Executar Análise',
    'Sin datos F&B.': 'Sem dados F&B.',
    'Reportes': 'Relatórios',
    "🔁 Reclamaciones OTA pendientes de aprobar": "🔁 Reclamações OTA pendentes de aprovação",
    "Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.": "Quando o Yve detetar comissões cobradas acima do contrato, aparecerão aqui para reclamar.",
    "✍️ Redactar con IA": "✍️ Redigir com IA",
    "✅ Aprobar y enviar": "✅ Aprovar e enviar",
    "🔄 Regenerar": "🔄 Regenerar",
    "🗑 Descartar": "🗑 Descartar",
    "Enviar a": "Enviar para",
    "Descartada": "Descartada",
    "✓ Enviada": "✓ Enviada",
    "email de la OTA": "email da OTA",
  },
};


// ── DOM text-node replacement engine ──────────────────────────────────────
// After applyI18n() translates data-i18n elements, this walks ALL text nodes
// and replaces known Spanish strings — covers JS-rendered content too.
// ── Traducir aunque el icono ya se haya comido el emoji ──────────────────
// Quita UN simbolo del principio (mas su selector de variante y los espacios
// que le sigan), que es exactamente lo que hace el iconizador.
var _RX_ICO = /^(?:[\uD800-\uDBFF][\uDC00-\uDFFF]|[\u2190-\u2BFF\u2600-\u27BF\u25A0-\u25FF])\uFE0F?\s*/;
function _sinIcono(s) { return String(s).replace(_RX_ICO, '').trim(); }

// El nodo de texto viene justo detras de un icono que puso el iconizador.
function _traeIconoDelante(n) {
  var p = n.previousSibling;
  return !!(p && p.nodeType === 1 && p.tagName && p.tagName.toLowerCase() === 'svg'
            && p.classList && p.classList.contains('yvi'));
}

// ── Indice inverso: de cualquier idioma de vuelta al español ────────────
// La tabla esta indexada por el español, asi que sin esto, pasar de ingles a
// catalan no encuentra nada y el texto se queda en ingles.
var _idxEsp = null;
function _aEspanol() {
  if (_idxEsp) return _idxEsp;
  var idx = {}, choque = {};
  for (var lg in _i18nStrMap) {
    var m = _i18nStrMap[lg];
    if (!m) continue;
    for (var k in m) {
      var v = m[k];
      if (!v || v === k) continue;
      if (idx[v] !== undefined && idx[v] !== k) { choque[v] = 1; continue; }
      idx[v] = k;
    }
  }
  // Si dos idiomas traducen cosas DISTINTAS con la misma palabra, no se puede
  // saber de cual venia: fuera del indice. Es mejor dejar el texto como esta
  // que cambiarlo por lo que no era.
  for (var c in choque) delete idx[c];
  _idxEsp = idx;
  return idx;
}
var _idxEspSin = null;
function _aEspanolSinIcono() {
  if (_idxEspSin) return _idxEspSin;
  var base = _aEspanol(), out = {};
  for (var v in base) {
    var s = _sinIcono(v);
    if (s && out[s] === undefined) out[s] = _sinIcono(base[v]);
  }
  _idxEspSin = out;
  return out;
}

// {español sin icono: traduccion sin icono}, una vez por idioma.
var _cacheSinIcono = {};
function _mapaSinIcono(lang, map) {
  if (_cacheSinIcono[lang]) return _cacheSinIcono[lang];
  var alt = {};
  for (var k in map) {
    var s = _sinIcono(k);
    if (s && s !== k && !alt[s]) alt[s] = _sinIcono(map[k]);
  }
  _cacheSinIcono[lang] = alt;
  return alt;
}

function _applyStrMap(lang, root) {
  if (!lang) return;
  // 'es' ya no se sale: hay que poder DESHACER una traduccion anterior. Sin
  // esto, volver a español dejaba el texto en el idioma en el que estuviera.
  var map = (lang === 'es') ? null : _i18nStrMap[lang];
  if (lang !== 'es' && !map) return;

  // `root` acota el recorrido al trozo recien pintado. Sin el, se recorren
  // TODOS los nodos de texto de la pagina cada vez, y con ocho paneles
  // poblados eso se hace dos veces por cada cambio (observer + segunda
  // pasada). Por defecto sigue siendo document.body: los 5 llamantes que ya
  // existian no cambian de comportamiento.
  var _raiz = root || document.body;
  if (!_raiz || !_raiz.nodeType) _raiz = document.body;

  // Walk all visible text nodes in the page
  var walker = document.createTreeWalker(
    _raiz,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode: function(node) {
        // Skip script, style, code nodes
        var p = node.parentNode;
        if (!p) return NodeFilter.FILTER_REJECT;
        var tag = p.tagName ? p.tagName.toLowerCase() : '';
        if (tag === 'script' || tag === 'style' || tag === 'code' || tag === 'pre') {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      }
    },
    false
  );

  var replacements = [];
  var n;
  while ((n = walker.nextNode())) {
    var text = n.textContent;
    if (!text || !text.trim()) continue;
    var trimmed = text.trim();
    // El texto puede estar en cualquier idioma: primero se lleva al español,
    // que es como esta indexada la tabla, y desde ahi al idioma que toca.
    var desdeOtro = _aEspanol()[trimmed];
    var clave = desdeOtro || trimmed;
    var destino = (lang === 'es') ? (desdeOtro || null) : map[clave];
    if (destino && destino !== trimmed) {
      // Preserve leading/trailing whitespace
      var leading  = text.match(/^\s*/)[0];
      var trailing = text.match(/\s*$/)[0];
      replacements.push([n, leading + destino + trailing]);
    } else if (_traeIconoDelante(n)) {
      // El iconizador ya se llevo el emoji de este nodo: la clave del mapa
      // sigue siendo "📊 Semanal" y aqui solo queda " Semanal", asi que la
      // comparacion de arriba falla y el texto se queda en español. Se busca
      // tambien por el texto SIN icono, y se escribe la traduccion SIN el
      // suyo: el icono ya esta al lado como hermano, repetirlo lo duplicaria.
      var desdeOtroSin = _aEspanolSinIcono()[trimmed];
      var claveSin = desdeOtroSin || trimmed;
      var v = (lang === 'es') ? (desdeOtroSin || null) : _mapaSinIcono(lang, map)[claveSin];
      if (v && v !== trimmed) {
        var l2 = text.match(/^\s*/)[0], t2 = text.match(/\s*$/)[0];
        replacements.push([n, l2 + v + t2]);
      }
    }
  }
  // Apply replacements (after walker is done, to avoid mutation issues)
  for (var i = 0; i < replacements.length; i++) {
    replacements[i][0].textContent = replacements[i][1];
  }
}

// ── Hook into applyI18n ───────────────────────────────────────────────────
function applyI18n(data) {
  _applyI18nBase(data);
  // After translating data-i18n elements, also walk text nodes
  if (_i18nLang && _i18nLang !== 'es') {
    // SINCRONO. Antes esperaba 120 ms "por si quedaba algun render pendiente",
    // y ese retraso es justo lo que se veia al cambiar de idioma.
    _applyStrMap(_i18nLang);
    if (typeof _applyPlaceholders === 'function') _applyPlaceholders(_i18nLang);
  } else {
    // Volver a español tambien es traducir: hay que DESHACER lo anterior.
    // _applyI18nBase solo restaura los elementos con data-i18n; el texto que
    // cambio la tabla de cadenas se quedaba en el idioma anterior (medido:
    // aleman -> español dejaba el panel de reclamaciones en aleman).
    _applyStrMap('es');
  }
  // Y los iconos AQUI MISMO. Traducir escribe textContent, o sea que borra los
  // SVG y devuelve el emoji del .json en su sitio. Confiar en el observador no
  // valia: la pagina muta ~11 veces por segundo y su espera no llegaba a saltar
  // (medido). Se veia en el menu: "Administracion" salia sin su icono.
  try { if (typeof iconizeIn === 'function') iconizeIn(document.body); } catch (e) {}
}

// Also expose so render functions can call it after innerHTML updates
function _i18nAfterRender() {
  if (_i18nLang && _i18nLang !== 'es') {
    _applyStrMap(_i18nLang);
  }
}

// ── Pintar YA: traducir e iconizar en el mismo instante en que se inyecta ──
// Antes esto lo hacian dos MutationObserver con 100/120 ms de espera y una
// segunda pasada a los 150. Se veia el español y los emojis en crudo durante
// un cuarto de segundo largo cada vez que se pintaba algo. Ahora se hace de
// forma SINCRONA sobre el trozo nuevo; los observers siguen ahi por si algo
// pinta por un camino que no pase por aqui.
function _pintarYa(root) {
  var r = (root && root.nodeType) ? root : document.body;
  try {
    if (_i18nLang && _i18nLang !== 'es') {
      _applyStrMap(_i18nLang, r);
      if (typeof _applyPlaceholders === 'function') _applyPlaceholders(_i18nLang);
    }
  } catch (e) {}
  try { if (typeof iconizeIn === 'function') iconizeIn(r); } catch (e) {}
}
window._pintarYa = _pintarYa;

// Que apartados ya estan poblados. Sin esto, `switchTab` vuelve a llamar al
// cargador CADA VEZ que entras y el panel se repinta desde cero — que es el
// mismo parpadeo, pero al volver.
var _panelCargado = {};
// Lista de apartados MIGRADOS al sistema nuevo (los que pintan ya traducido y
// no necesitan la red de seguridad). Se va llenando segun se migran; el script
// de cobertura lee esta lista, asi que la cuenta sale del codigo y no puede
// quedarse vieja.
var _PANELES_MIGRADOS = [];
window._PANELES_MIGRADOS = _PANELES_MIGRADOS;
// ──────────────────────────────────────────────────────────────────────────
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
    const r = await fetch('/static/i18n/' + lang + '.json?v=__ASSETS_V__');
    const data = await r.json();
    _i18nCache[lang] = data; _i18nData = data; _i18nLang = lang;
    applyI18n(data); localStorage.setItem('yve_lang', lang);
  } catch(e) { console.warn('i18n error:', e); }
}

function t(key, fb) {
  var v = _i18nData[key] || _i18nOriginal[key];
  if (v && v !== key) return v;
  if (fb !== undefined && fb !== null) return fb;
  return key;
}

// Referencia estable a la función de traducción + wrapper a prueba de "clobbering"
// (si algo global pisa 't', tt() sigue traduciendo con la función real capturada).
var _T_FN = t;
function tt(key, fb) {
  try {
    var fn = (typeof t === 'function') ? t : _T_FN;
    return (typeof fn === 'function') ? fn(key, fb) : (fb !== undefined && fb !== null ? fb : key);
  } catch (e) { return (fb !== undefined && fb !== null) ? fb : key; }
}


// ── Traducción de mensajes SSE / dinámicos por fragmentos ──────────────────
var _sseFrags = [
  "hoja de cálculo sin clasificar — revisar manualmente",
  "extracto detectado pero error al guardar",
  "rooming list detectado (ocupación)",
  "inventario detectado (sin items extraíbles)",
  "mermas detectadas (sin items extraíbles)",
  "Ya hay un proceso AP en ejecucion — espera",
  "Ya hay un proceso Oracle en ejecucion — espera",
  "Ya hay un proceso en ejecucion — espera",
  "Ya hay un proceso activo — espera",
  "Ya hay un proceso — espera",
  "No se especificaron archivos",
  "No hay archivos nuevos que procesar",
  "archivo cargado (revisar formato)",
  "archivo cargado (revisar columnas)",
  "archivo copiado (formato",
  "archivo copiado",
  "archivo no encontrado",
  "archivos ya procesados",
  "Verificación comisiones completada",
  "Verificando comisiones OTA...",
  "Verificación completada",
  "Análisis doble imposición completado",
  "Iniciando procesamiento batch",
  "copiado para procesamiento",
  "Procesado finalizado con avisos",
  "Procesado completado",
  "movimientos integrados",
  "registros integrados",
  "integrados por IA",
  "extraídos por IA",
  "detectadas por IA",
  "detectado por IA",
  "detectado pero error",
  "factura cuadra con",
  "ventas detectadas",
  "datos cargados",
  "inventario detectado",
  "mermas detectadas",
  "demasiado grande",
  "alerta(s) procesada(s)",
  "Batch completado",
  "ERROR CRÍTICO",
  "tardo demasiado",
  "Evento tiene",
  "Pipeline completado con éxito",
  "Pipeline finalizado con errores",
  "Pipeline AP completado",
  "Pipeline AP con errores",
  "Oracle: contabilización completada",
  "Oracle: pipeline con errores",
  "conexión con servidor perdida",
  "Error de conexión",
  "Timeout en lote",
  "Reconectando",
  "continuando",
  "Contabilizar en Oracle",
  "Contabilizando",
  "Procesar Archivos",
  "archivo(s) nuevos",
  "Saltando",
  "Procesando",
  ": procesado",
  ": cargado",
  "no encontrado",
  "archivo(s)",
  "Notificaciones:",
  "completado",
  "movimientos",
  "registros",
  "productos",
  "platos",
  "Consulta:",
  "comisiones",
  "ventas",
  "documento(s) guardado(s) para cruces internos — todavia no hay pantalla donde consultarlos",
  "factura(s) sin entrega que las respalde o con diferencia de importe",
  "Cuentas y asientos asignados — pendientes de aprobar en Aprobaciones AP",
  "parece una orden de compra, pero no se ha podido extraer ni proveedor ni importe",
  "albarán detectado, pero no se ha podido extraer ninguna línea",
  "parece un DRR — subelo desde el boton",
  "Cruce con albaranes: sin incidencias",
  "Cruzando facturas con albaranes",
  "recibido, no se ha podido procesar",
  "sin keyword reconocible — lo lee la IA",
  "Cruce con albaranes:",
  "Asignando cuentas contables",
  "órdenes de compra",
  "Orden de compra ",
  "albaranes",
  "Albarán ",
  "Rooming ",
  " línea(s)",
  "revisar manualmente",
  "ojo: el documento nombra otro hotel",
];
var _sseTrans = {
  en: [
  "unclassified spreadsheet — review manually",
  "bank statement detected but error saving",
  "rooming list detected (occupancy)",
  "inventory detected (no extractable items)",
  "waste detected (no extractable items)",
  "An AP process is already running — wait",
  "An Oracle process is already running — wait",
  "A process is already running — wait",
  "A process is already running — wait",
  "A process is already running — wait",
  "No files specified",
  "No new files to process",
  "file loaded (check format)",
  "file loaded (check columns)",
  "file copied (format",
  "file copied",
  "file not found",
  "files already processed",
  "Commission verification completed",
  "Verifying OTA commissions...",
  "Verification completed",
  "Double taxation analysis completed",
  "Starting batch processing",
  "copied for processing",
  "Processing finished with warnings",
  "Processing completed",
  "transactions integrated",
  "records integrated",
  "integrated by AI",
  "extracted by AI",
  "detected by AI",
  "detected by AI",
  "detected but error",
  "invoice matches",
  "sales detected",
  "data loaded",
  "inventory detected",
  "waste detected",
  "too large",
  "alert(s) processed",
  "Batch completed",
  "CRITICAL ERROR",
  "took too long",
  "Event has",
  "Pipeline completed successfully",
  "Pipeline finished with errors",
  "AP pipeline completed",
  "AP pipeline with errors",
  "Oracle: posting completed",
  "Oracle: pipeline with errors",
  "connection to server lost",
  "Connection error",
  "Timeout in batch",
  "Reconnecting",
  "continuing",
  "Post to Oracle",
  "Posting",
  "Process Files",
  "new file(s)",
  "Skipping",
  "Processing",
  ": processed",
  ": loaded",
  "not found",
  "file(s)",
  "Notifications:",
  "completed",
  "transactions",
  "records",
  "products",
  "dishes",
  "Query:",
  "commissions",
  "sales",
    "document(s) saved for internal cross-checks — there is no screen to view them yet",
    "invoice(s) with no delivery to back them or with an amount mismatch",
    "Accounts and journal entries assigned — awaiting approval in Invoices to approve",
    "looks like a purchase order, but neither supplier nor amount could be extracted",
    "delivery note detected, but no line could be extracted",
    "looks like a DRR — upload it from the button",
    "Delivery-note match: no issues",
    "Matching invoices against delivery notes",
    "received, could not be processed",
    "no recognisable keyword — the AI reads it",
    "Delivery-note match:",
    "Assigning ledger accounts",
    "purchase orders",
    "Purchase order ",
    "delivery notes",
    "Delivery note ",
    "Rooming ",
    " line(s)",
    "review manually",
    "careful: the document names a different hotel",
  ],
  ca: [
  "full de càlcul sense classificar — revisar manualment",
  "extracte detectat però error en desar",
  "rooming list detectada (ocupació)",
  "inventari detectat (sense ítems extraïbles)",
  "minves detectades (sense ítems extraïbles)",
  "Ja hi ha un procés AP en execució — espera",
  "Ja hi ha un procés Oracle en execució — espera",
  "Ja hi ha un procés en execució — espera",
  "Ja hi ha un procés actiu — espera",
  "Ja hi ha un procés — espera",
  "No s'han especificat fitxers",
  "No hi ha fitxers nous per processar",
  "fitxer carregat (revisar format)",
  "fitxer carregat (revisar columnes)",
  "fitxer copiat (format",
  "fitxer copiat",
  "fitxer no trobat",
  "fitxers ja processats",
  "Verificació de comissions completada",
  "Verificant comissions OTA...",
  "Verificació completada",
  "Anàlisi de doble imposició completada",
  "Iniciant processament batch",
  "copiat per processar",
  "Processament finalitzat amb avisos",
  "Processament completat",
  "moviments integrats",
  "registres integrats",
  "integrats per IA",
  "extrets per IA",
  "detectades per IA",
  "detectat per IA",
  "detectat però error",
  "factura quadra amb",
  "vendes detectades",
  "dades carregades",
  "inventari detectat",
  "minves detectades",
  "massa gran",
  "alerta/es processada/es",
  "Batch completat",
  "ERROR CRÍTIC",
  "ha trigat massa",
  "L'esdeveniment té",
  "Pipeline completat amb èxit",
  "Pipeline finalitzat amb errors",
  "Pipeline AP completat",
  "Pipeline AP amb errors",
  "Oracle: comptabilització completada",
  "Oracle: pipeline amb errors",
  "connexió amb el servidor perduda",
  "Error de connexió",
  "Timeout al lot",
  "Reconnectant",
  "continuant",
  "Comptabilitzar a Oracle",
  "Comptabilitzant",
  "Processar Fitxers",
  "fitxer(s) nous",
  "Saltant",
  "Processant",
  ": processat",
  ": carregat",
  "no trobat",
  "fitxer(s)",
  "Notificacions:",
  "completat",
  "moviments",
  "registres",
  "productes",
  "plats",
  "Consulta:",
  "comissions",
  "vendes",
    "document(s) desat(s) per a encreuaments interns — encara no hi ha cap pantalla per consultar-los",
    "factura/es sense lliurament que les avali o amb diferència d'import",
    "Comptes i assentaments assignats — pendents d'aprovar a Factures per aprovar",
    "sembla una ordre de compra, però no s'ha pogut extreure ni proveïdor ni import",
    "albarà detectat, però no s'ha pogut extreure cap línia",
    "sembla un DRR — puja'l des del botó",
    "Encreuament amb albarans: sense incidències",
    "Encreuant factures amb albarans",
    "rebut, no s'ha pogut processar",
    "sense cap paraula clau reconeixible — ho llegeix la IA",
    "Encreuament amb albarans:",
    "Assignant comptes comptables",
    "ordres de compra",
    "Ordre de compra ",
    "albarans",
    "Albarà ",
    "Rooming ",
    " línia/es",
    "revisar-ho manualment",
    "compte: el document anomena un altre hotel",
  ],
  fr: [
  "feuille de calcul non classée — vérifier manuellement",
  "relevé détecté mais erreur d'enregistrement",
  "rooming list détectée (occupation)",
  "inventaire détecté (aucun élément extractible)",
  "pertes détectées (aucun élément extractible)",
  "Un processus AP est déjà en cours — patientez",
  "Un processus Oracle est déjà en cours — patientez",
  "Un processus est déjà en cours — patientez",
  "Un processus est déjà actif — patientez",
  "Un processus est déjà en cours — patientez",
  "Aucun fichier spécifié",
  "Aucun nouveau fichier à traiter",
  "fichier chargé (vérifier le format)",
  "fichier chargé (vérifier les colonnes)",
  "fichier copié (format",
  "fichier copié",
  "fichier introuvable",
  "fichiers déjà traités",
  "Vérification des commissions terminée",
  "Vérification des commissions OTA...",
  "Vérification terminée",
  "Analyse double imposition terminée",
  "Démarrage du traitement batch",
  "copié pour traitement",
  "Traitement terminé avec avertissements",
  "Traitement terminé",
  "mouvements intégrés",
  "enregistrements intégrés",
  "intégrés par IA",
  "extraits par IA",
  "détectées par IA",
  "détecté par IA",
  "détecté mais erreur",
  "la facture correspond à",
  "ventes détectées",
  "données chargées",
  "inventaire détecté",
  "pertes détectées",
  "trop volumineux",
  "alerte(s) traitée(s)",
  "Batch terminé",
  "ERREUR CRITIQUE",
  "a pris trop de temps",
  "L'événement a",
  "Pipeline terminé avec succès",
  "Pipeline terminé avec des erreurs",
  "Pipeline AP terminé",
  "Pipeline AP avec erreurs",
  "Oracle : comptabilisation terminée",
  "Oracle : pipeline avec erreurs",
  "connexion au serveur perdue",
  "Erreur de connexion",
  "Timeout sur le lot",
  "Reconnexion",
  "poursuite",
  "Comptabiliser dans Oracle",
  "Comptabilisation",
  "Traiter les fichiers",
  "nouveau(x) fichier(s)",
  "Ignore",
  "Traitement",
  ": traité",
  ": chargé",
  "introuvable",
  "fichier(s)",
  "Notifications :",
  "terminé",
  "mouvements",
  "enregistrements",
  "produits",
  "plats",
  "Requête :",
  "commissions",
  "ventes",
    "document(s) enregistré(s) pour les rapprochements internes — aucun écran ne permet encore de les consulter",
    "facture(s) sans livraison correspondante ou avec un écart de montant",
    "Comptes et écritures affectés — en attente d'approbation dans Factures à approuver",
    "ressemble à un bon de commande, mais ni le fournisseur ni le montant n'ont pu être extraits",
    "bon de livraison détecté, mais aucune ligne n'a pu être extraite",
    "ressemble à un DRR — importez-le depuis le bouton",
    "Rapprochement avec les bons de livraison : aucun problème",
    "Rapprochement des factures avec les bons de livraison",
    "reçu, n'a pas pu être traité",
    "aucun mot-clé reconnaissable — l'IA le lit",
    "Rapprochement avec les bons de livraison :",
    "Affectation des comptes comptables",
    "bons de commande",
    "Bon de commande ",
    "bons de livraison",
    "Bon de livraison ",
    "Rooming ",
    " ligne(s)",
    "à vérifier manuellement",
    "attention : le document nomme un autre hôtel",
  ],
  de: [
  "Tabelle nicht klassifiziert — manuell prüfen",
  "Kontoauszug erkannt, aber Fehler beim Speichern",
  "Rooming-Liste erkannt (Belegung)",
  "Inventar erkannt (keine extrahierbaren Positionen)",
  "Schwund erkannt (keine extrahierbaren Positionen)",
  "Ein AP-Prozess läuft bereits — bitte warten",
  "Ein Oracle-Prozess läuft bereits — bitte warten",
  "Ein Prozess läuft bereits — bitte warten",
  "Ein Prozess läuft bereits — bitte warten",
  "Ein Prozess läuft bereits — bitte warten",
  "Keine Dateien angegeben",
  "Keine neuen Dateien zu verarbeiten",
  "Datei geladen (Format prüfen)",
  "Datei geladen (Spalten prüfen)",
  "Datei kopiert (Format",
  "Datei kopiert",
  "Datei nicht gefunden",
  "Dateien bereits verarbeitet",
  "Provisionsprüfung abgeschlossen",
  "OTA-Provisionen werden geprüft...",
  "Prüfung abgeschlossen",
  "Doppelbesteuerungsanalyse abgeschlossen",
  "Batch-Verarbeitung gestartet",
  "zur Verarbeitung kopiert",
  "Verarbeitung mit Warnungen beendet",
  "Verarbeitung abgeschlossen",
  "Bewegungen integriert",
  "Datensätze integriert",
  "per KI integriert",
  "per KI extrahiert",
  "per KI erkannt",
  "per KI erkannt",
  "erkannt, aber Fehler",
  "Rechnung stimmt überein mit",
  "Verkäufe erkannt",
  "Daten geladen",
  "Inventar erkannt",
  "Schwund erkannt",
  "zu groß",
  "Alarm(e) verarbeitet",
  "Batch abgeschlossen",
  "KRITISCHER FEHLER",
  "hat zu lange gedauert",
  "Event hat",
  "Pipeline erfolgreich abgeschlossen",
  "Pipeline mit Fehlern beendet",
  "AP-Pipeline abgeschlossen",
  "AP-Pipeline mit Fehlern",
  "Oracle: Buchung abgeschlossen",
  "Oracle: Pipeline mit Fehlern",
  "Verbindung zum Server verloren",
  "Verbindungsfehler",
  "Timeout bei Los",
  "Verbinde neu",
  "fahre fort",
  "In Oracle buchen",
  "Buche",
  "Dateien verarbeiten",
  "neue Datei(en)",
  "Überspringe",
  "Verarbeite",
  ": verarbeitet",
  ": geladen",
  "nicht gefunden",
  "Datei(en)",
  "Benachrichtigungen:",
  "abgeschlossen",
  "Bewegungen",
  "Datensätze",
  "Produkte",
  "Gerichte",
  "Abfrage:",
  "Provisionen",
  "Verkäufe",
    "Dokument(e) für interne Abgleiche gespeichert — es gibt noch keine Ansicht dafür",
    "Rechnung(en) ohne belegende Lieferung oder mit Betragsabweichung",
    "Konten und Buchungen zugeordnet — Freigabe steht aus unter Rechnungen zur Freigabe",
    "sieht nach einer Bestellung aus, aber weder Lieferant noch Betrag konnten gelesen werden",
    "Lieferschein erkannt, aber es konnte keine Position gelesen werden",
    "sieht nach einem DRR aus — lade ihn über die Schaltfläche hoch",
    "Abgleich mit Lieferscheinen: keine Auffälligkeiten",
    "Rechnungen werden mit Lieferscheinen abgeglichen",
    "empfangen, konnte nicht verarbeitet werden",
    "kein erkennbares Schlüsselwort — die KI liest es",
    "Abgleich mit Lieferscheinen:",
    "Sachkonten werden zugeordnet",
    "Bestellungen",
    "Bestellung ",
    "Lieferscheine",
    "Lieferschein ",
    "Rooming ",
    " Position(en)",
    "manuell prüfen",
    "Achtung: das Dokument nennt ein anderes Hotel",
  ],
  it: [
  "foglio di calcolo non classificato — verificare manualmente",
  "estratto conto rilevato ma errore nel salvataggio",
  "rooming list rilevata (occupazione)",
  "inventario rilevato (nessun elemento estraibile)",
  "sprechi rilevati (nessun elemento estraibile)",
  "Un processo AP è già in esecuzione — attendere",
  "Un processo Oracle è già in esecuzione — attendere",
  "Un processo è già in esecuzione — attendere",
  "Un processo è già attivo — attendere",
  "Un processo è già in esecuzione — attendere",
  "Nessun file specificato",
  "Nessun nuovo file da elaborare",
  "file caricato (verificare formato)",
  "file caricato (verificare colonne)",
  "file copiato (formato",
  "file copiato",
  "file non trovato",
  "file già elaborati",
  "Verifica commissioni completata",
  "Verifica commissioni OTA...",
  "Verifica completata",
  "Analisi doppia imposizione completata",
  "Avvio elaborazione batch",
  "copiato per l'elaborazione",
  "Elaborazione terminata con avvisi",
  "Elaborazione completata",
  "movimenti integrati",
  "record integrati",
  "integrati da IA",
  "estratti da IA",
  "rilevate da IA",
  "rilevato da IA",
  "rilevato ma errore",
  "la fattura corrisponde a",
  "vendite rilevate",
  "dati caricati",
  "inventario rilevato",
  "sprechi rilevati",
  "troppo grande",
  "avvisi elaborati",
  "Batch completato",
  "ERRORE CRITICO",
  "ha impiegato troppo tempo",
  "L'evento ha",
  "Pipeline completata con successo",
  "Pipeline terminata con errori",
  "Pipeline AP completata",
  "Pipeline AP con errori",
  "Oracle: contabilizzazione completata",
  "Oracle: pipeline con errori",
  "connessione al server persa",
  "Errore di connessione",
  "Timeout nel lotto",
  "Riconnessione",
  "continuo",
  "Contabilizza in Oracle",
  "Contabilizzazione",
  "Elabora file",
  "nuovo/i file",
  "Salto",
  "Elaborazione",
  ": elaborato",
  ": caricato",
  "non trovato",
  "file",
  "Notifiche:",
  "completato",
  "movimenti",
  "record",
  "prodotti",
  "piatti",
  "Query:",
  "commissioni",
  "vendite",
    "documento/i salvato/i per i riscontri interni — non c'è ancora una schermata per consultarli",
    "fattura/e senza consegna a supporto o con differenza di importo",
    "Conti e scritture assegnati — in attesa di approvazione in Fatture da approvare",
    "sembra un ordine d'acquisto, ma non è stato possibile estrarre né fornitore né importo",
    "documento di trasporto rilevato, ma non è stata estratta alcuna riga",
    "sembra un DRR — caricalo dal pulsante",
    "Riscontro con i documenti di trasporto: nessuna anomalia",
    "Riscontro delle fatture con i documenti di trasporto",
    "ricevuto, non è stato possibile elaborarlo",
    "nessuna parola chiave riconoscibile — lo legge l'IA",
    "Riscontro con i documenti di trasporto:",
    "Assegnazione dei conti contabili",
    "ordini d'acquisto",
    "Ordine d'acquisto ",
    "documenti di trasporto",
    "Documento di trasporto ",
    "Rooming ",
    " riga/righe",
    "da rivedere manualmente",
    "attenzione: il documento nomina un altro hotel",
  ],
  pt: [
  "planilha sem classificar — revisar manualmente",
  "extrato detectado mas erro ao salvar",
  "rooming list detectada (ocupação)",
  "inventário detectado (sem itens extraíveis)",
  "perdas detectadas (sem itens extraíveis)",
  "Já existe um processo AP em execução — aguarde",
  "Já existe um processo Oracle em execução — aguarde",
  "Já existe um processo em execução — aguarde",
  "Já existe um processo ativo — aguarde",
  "Já existe um processo — aguarde",
  "Nenhum arquivo especificado",
  "Nenhum arquivo novo para processar",
  "arquivo carregado (revisar formato)",
  "arquivo carregado (revisar colunas)",
  "arquivo copiado (formato",
  "arquivo copiado",
  "arquivo não encontrado",
  "arquivos já processados",
  "Verificação de comissões concluída",
  "Verificando comissões OTA...",
  "Verificação concluída",
  "Análise de dupla tributação concluída",
  "Iniciando processamento batch",
  "copiado para processamento",
  "Processamento finalizado com avisos",
  "Processamento concluído",
  "movimentos integrados",
  "registros integrados",
  "integrados por IA",
  "extraídos por IA",
  "detectadas por IA",
  "detectado por IA",
  "detectado mas erro",
  "fatura confere com",
  "vendas detectadas",
  "dados carregados",
  "inventário detectado",
  "perdas detectadas",
  "muito grande",
  "alerta(s) processado(s)",
  "Batch concluído",
  "ERRO CRÍTICO",
  "demorou demais",
  "Evento tem",
  "Pipeline concluído com sucesso",
  "Pipeline finalizado com erros",
  "Pipeline AP concluído",
  "Pipeline AP com erros",
  "Oracle: contabilização concluída",
  "Oracle: pipeline com erros",
  "conexão com o servidor perdida",
  "Erro de conexão",
  "Timeout no lote",
  "Reconectando",
  "continuando",
  "Contabilizar no Oracle",
  "Contabilizando",
  "Processar Arquivos",
  "arquivo(s) novos",
  "Pulando",
  "Processando",
  ": processado",
  ": carregado",
  "não encontrado",
  "arquivo(s)",
  "Notificações:",
  "concluído",
  "movimentos",
  "registros",
  "produtos",
  "pratos",
  "Consulta:",
  "comissões",
  "vendas",
    "documento(s) guardado(s) para cruzamentos internos — ainda não há ecrã para os consultar",
    "fatura(s) sem entrega que as suporte ou com diferença de valor",
    "Contas e lançamentos atribuídos — pendentes de aprovação em Faturas por aprovar",
    "parece uma ordem de compra, mas não foi possível extrair fornecedor nem valor",
    "guia de remessa detetada, mas não foi possível extrair nenhuma linha",
    "parece um DRR — carrega-o a partir do botão",
    "Cruzamento com guias de remessa: sem incidências",
    "A cruzar faturas com guias de remessa",
    "recebido, não foi possível processá-lo",
    "sem palavra-chave reconhecível — a IA lê-o",
    "Cruzamento com guias de remessa:",
    "A atribuir contas contabilísticas",
    "ordens de compra",
    "Ordem de compra ",
    "guias de remessa",
    "Guia de remessa ",
    "Rooming ",
    " linha(s)",
    "rever manualmente",
    "atenção: o documento nomeia outro hotel",
  ],
};
function _tSSE(txt) {
  if (!txt || !_i18nLang || _i18nLang === 'es') return txt;
  var tr = _sseTrans[_i18nLang];
  if (!tr) return txt;
  for (var i = 0; i < _sseFrags.length; i++) {
    if (txt.indexOf(_sseFrags[i]) !== -1) txt = txt.split(_sseFrags[i]).join(tr[i]);
  }
  return txt;
}

function _applyI18nBase(data) {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const k = el.getAttribute('data-i18n');
    if (data[k] !== undefined) el.textContent = data[k];
  });
}

function toggleLightMode() {
  var isLight = document.body.classList.toggle('light-mode');
  localStorage.setItem('yve_theme', isLight ? 'light' : 'dark');
  var btn = document.getElementById('btn-theme');
  if (btn) btn.textContent = isLight ? '🌙 Modo oscuro' : '☀️ Modo claro';
  var navBtn = document.getElementById('btn-theme-nav');
  if (navBtn) navBtn.textContent = isLight ? '🌙' : '☀️';
}
// ── Custom color picker ───────────────────────────────────────────────────────
var _customColors = {
  accent: localStorage.getItem('yve_accent') || '#3b82f6',
  bg:     localStorage.getItem('yve_bg')     || '#0f172a',
  hlAll:  localStorage.getItem('yve_hl_all') === '1',
};

function _applyCustomColors() {
  var r = document.documentElement;

  // helpers (definidos aquí para que acc también los use)
  function toHex(v) { v=Math.max(0,Math.min(255,Math.round(v))); return ('0'+v.toString(16)).slice(-2); }
  function lighter(val, amt) { return Math.min(255, val+amt); }

  // ── FONDO: solo cambia el fondo de página y el nav (--bg + componentes RGB)
  //    Las superficies (--s1/s2/s3) se calculan como offset fijo sobre el bg
  //    para mantener profundidad, pero NO pintan las burbujas del chat
  //    (las burbujas de usuario usan --acc, las de bot usan --s1 que se mantiene
  //    a ~+20/42/65 puntos del fondo, creando el contraste correcto)
  var bg = _customColors.bg;
  var bgHex = bg.replace('#','');
  var bgR = parseInt(bgHex.substr(0,2),16);
  var bgG = parseInt(bgHex.substr(2,2),16);
  var bgB = parseInt(bgHex.substr(4,2),16);
  r.style.setProperty('--bg',   bg);
  r.style.setProperty('--bg-r', String(bgR));
  r.style.setProperty('--bg-g', String(bgG));
  r.style.setProperty('--bg-b', String(bgB));
  // Superficies (--s1/s2/s3) NO se tocan — mantienen su valor CSS por defecto
  // así las tarjetas no cambian de color al cambiar el fondo

  // ── ACENTO: cambia burbujas de chat, botones, tabs, badges ──────────────
  var acc = _customColors.accent;
  var aHex = acc.replace('#','');
  var aR = parseInt(aHex.substr(0,2),16);
  var aG = parseInt(aHex.substr(2,2),16);
  var aB = parseInt(aHex.substr(4,2),16);
  // acc2 = 25% más claro, acc3 = 50% más claro, acc-dark = 20% más oscuro
  function blendW(v,t) { return Math.round(v+(255-v)*t); }
  function darken(v,t) { return Math.max(0, Math.round(v*(1-t))); }
  r.style.setProperty('--acc',      acc);
  r.style.setProperty('--acc2',     '#'+toHex(blendW(aR,.25))+toHex(blendW(aG,.25))+toHex(blendW(aB,.25)));
  r.style.setProperty('--acc3',     '#'+toHex(blendW(aR,.5))+toHex(blendW(aG,.5))+toHex(blendW(aB,.5)));
  r.style.setProperty('--acc-dark', '#'+toHex(darken(aR,.2))+toHex(darken(aG,.2))+toHex(darken(aB,.2)));
  // RGB components for rgba(var(--acc-r),...) in CSS
  r.style.setProperty('--acc-r', String(aR));
  r.style.setProperty('--acc-g', String(aG));
  r.style.setProperty('--acc-b', String(aB));

  // ── Sparklines: canvas no lee CSS vars — re-dibujar con nuevo color ─────
  var newAcc2 = '#'+toHex(blendW(aR,.25))+toHex(blendW(aG,.25))+toHex(blendW(aB,.25));
  [typeof AR_SPARKS!=='undefined' && AR_SPARKS,
   typeof AP_SPARKS!=='undefined' && AP_SPARKS].forEach(function(sparks){
    if (!sparks) return;
    sparks.forEach(function(s){
      // hlAll: todos usan acento; si no, solo los marcados; si se desactiva hlAll, restaurar color semántico
      if (_customColors.hlAll) s.color = newAcc2;
      else if (s.accent) s.color = newAcc2;
      else if (s.baseColor) s.color = s.baseColor;
    });
  });
  if (typeof injectSparklines === 'function') {
    if (typeof AR_SPARKS !== 'undefined') injectSparklines(AR_SPARKS);
    if (typeof AP_SPARKS !== 'undefined') injectSparklines(AP_SPARKS);
  }
  // ── Modo acentuar-todo: body class controla el CSS de los contenedores ──
  if (_customColors.hlAll) document.body.classList.add('acentuar-todo');
  else document.body.classList.remove('acentuar-todo');
  if (typeof _marcarStatsActualizadas === 'function') _marcarStatsActualizadas();
}

function _cpSwatch(id, c, cur) {
  var sel = cur === c;
  // Usamos data attributes en vez de onclick inline para evitar conflictos de comillas
  return '<div class="cp-swatch" data-cpid="' + id + '" data-cpc="' + c + '" ' +
    'style="width:24px;height:24px;border-radius:50%;background:' + c + ';cursor:pointer;flex-shrink:0;transition:background-color .12s,border-color .12s,color .12s,box-shadow .12s,transform .12s,opacity .12s;' +
    'box-shadow:0 0 0 ' + (sel ? '3px' : '0px') + ' #0f172a, 0 0 0 ' + (sel ? '5px' : '0px') + ' #fff' +
    (sel ? ', 0 0 8px 2px ' + c : '') + '"></div>';
}
function _cpSet(id, color) {
  var el = document.getElementById(id);
  if (el) el.value = color;
  // Actualizar aro blanco en todas las swatches de este input
  var s1 = getComputedStyle(document.documentElement).getPropertyValue('--s1').trim() || '#1e293b';
  document.querySelectorAll('[data-cpid="'+id+'"]').forEach(function(sw) {
    var isSel = sw.getAttribute('data-cpc') === color;
    sw.style.boxShadow = isSel
      ? '0 0 0 3px '+s1+', 0 0 0 5px #fff, 0 0 8px 2px '+color
      : '';
  });
  // Preview en tiempo real + actualizar label hex
  if (id === 'cp-accent') {
    _customColors.accent = color;
    var lbl = document.getElementById('cp-accent-label');
    if (lbl) lbl.textContent = color;
  }
  if (id === 'cp-bg') {
    _customColors.bg = color;
    var lbl = document.getElementById('cp-bg-label');
    if (lbl) lbl.textContent = color;
  }
  _applyCustomColors();
}
// Delegación de eventos para swatches del color picker (evita conflictos de comillas en onclick)
document.addEventListener('click', function(e) {
  var sw = e.target.closest('.cp-swatch');
  if (sw) _cpSet(sw.dataset.cpid, sw.dataset.cpc);
});

function _openColorPicker() {
  var existing = document.getElementById('color-picker-modal');
  if (existing) { existing.remove(); return; }

  // Guardar estado previo para poder cancelar
  var _prevColors = { accent: _customColors.accent, bg: _customColors.bg, hlAll: _customColors.hlAll };

  var accentSwatches = [
    '#3b82f6','#6366f1','#7c3aed','#a855f7','#8b5cf6',
    '#ec4899','#f43f5e','#ef4444','#f97316','#f59e0b',
    '#eab308','#22c55e','#10b981','#14b8a6','#06b6d4',
    '#0ea5e9','#3b82f6','#64748b','#e11d48','#7c3aed'
  ];
  // Deduplicar
  accentSwatches = accentSwatches.filter(function(c,i){ return accentSwatches.indexOf(c)===i; });

  var bgSwatches = [
    '#0f172a','#1a1a2e','#0d1117','#1e1e2e',
    '#111827','#0a0f1e','#13111c','#1c1917',
    '#1a1f2e','#0f2027','#1e1b2e','#0d0d0d'
  ];

  var modal = document.createElement('div');
  modal.id = 'color-picker-modal';
  modal.style.cssText = 'position:fixed;inset:0;z-index:9800;display:flex;align-items:center;' +
    'justify-content:center;background:rgba(0,0,0,.6);backdrop-filter:blur(2px)';
  modal.innerHTML =
    '<div style="background:var(--s1);border:1px solid var(--s2);border-radius:18px;padding:26px 24px 20px;width:340px;max-width:calc(100vw - 32px)">' +
      '<div style="font-size:16px;font-weight:700;margin-bottom:18px">🎨 Personalizar colores</div>' +

      '<div style="margin-bottom:16px">' +
        '<label style="font-size:11px;color:var(--mut);font-weight:600;display:block;margin-bottom:8px;text-transform:uppercase;letter-spacing:.06em">Color de acento</label>' +
        '<div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">' +
          '<input type="color" id="cp-accent" value="'+_customColors.accent+'" ' +
            'style="width:40px;height:40px;border:none;border-radius:10px;cursor:pointer;background:none;flex-shrink:0">' +
          '<span style="font-size:12px;color:var(--mut)" id="cp-accent-label">'+_customColors.accent+'</span>' +
        '</div>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap">' +
          accentSwatches.map(function(c){ return _cpSwatch('cp-accent',c,_customColors.accent); }).join('') +
        '</div>' +
      '</div>' +

      '<div style="margin-bottom:18px">' +
        '<label style="font-size:11px;color:var(--mut);font-weight:600;display:block;margin-bottom:8px;text-transform:uppercase;letter-spacing:.06em">Color de fondo</label>' +
        '<div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">' +
          '<input type="color" id="cp-bg" value="'+_customColors.bg+'" ' +
            'style="width:40px;height:40px;border:none;border-radius:10px;cursor:pointer;background:none;flex-shrink:0">' +
          '<span style="font-size:12px;color:var(--mut)" id="cp-bg-label">'+_customColors.bg+'</span>' +
        '</div>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap">' +
          bgSwatches.map(function(c){ return _cpSwatch('cp-bg',c,_customColors.bg); }).join('') +
        '</div>' +
      '</div>' +

      '<div style="margin-bottom:18px">' +
        '<label style="display:flex;align-items:center;gap:10px;cursor:pointer;user-select:none">' +
          '<div onclick="_toggleHlAll()" id="cp-hlall-track" style="position:relative;width:40px;height:22px;' +
            'background:'+ (_customColors.hlAll ? 'var(--acc)' : 'var(--s3)') +';' +
            'border-radius:11px;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s;flex-shrink:0">' +
            '<div id="cp-hlall-thumb" style="position:absolute;top:3px;left:'+ (_customColors.hlAll ? '21px' : '3px') +';' +
              'width:16px;height:16px;background:#fff;border-radius:50%;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s"></div>' +
          '</div>' +
          '<div>' +
            '<div style="font-size:12px;font-weight:600">Acento en todos los contenedores</div>' +
            '<div style="font-size:11px;color:var(--mut);margin-top:2px">Los 6 cards usan el color elegido</div>' +
          '</div>' +
        '</label>' +
      '</div>' +

      '<div style="display:flex;gap:8px">' +
        '<button onclick="_cancelColors()" style="flex:1;background:rgba(255,255,255,.06);border:1px solid var(--s2);' +
          'color:var(--mut);padding:9px;border-radius:10px;cursor:pointer;font-size:13px">Cancelar</button>' +
        '<button onclick="_resetColors()" style="flex:1;background:rgba(255,255,255,.06);border:1px solid var(--s2);' +
          'color:var(--mut);padding:9px;border-radius:10px;cursor:pointer;font-size:13px">Resetear</button>' +
        '<button onclick="_saveColors()" style="flex:1;background:var(--acc);border:none;color:#fff;' +
          'padding:9px;border-radius:10px;cursor:pointer;font-size:13px;font-weight:700">Aplicar</button>' +
      '</div>' +
    '</div>';

  modal.addEventListener('click', function(e){
    if (e.target === modal) _cancelColors(); // click fuera = cancelar
  });
  document.body.appendChild(modal);

  // ── Live preview: aplicar cambios en tiempo real ─────────────────────────
  function _previewColors() {
    var a = document.getElementById('cp-accent');
    var b = document.getElementById('cp-bg');
    if (a) { _customColors.accent = a.value; document.getElementById('cp-accent-label').textContent = a.value; }
    if (b) { _customColors.bg = b.value; document.getElementById('cp-bg-label').textContent = b.value; }
    _applyCustomColors();
  }
  var cpA = document.getElementById('cp-accent');
  var cpB = document.getElementById('cp-bg');
  if (cpA) cpA.addEventListener('input', _previewColors);
  if (cpB) cpB.addEventListener('input', _previewColors);
}

function _toggleHlAll() {
  _customColors.hlAll = !_customColors.hlAll;
  var track = document.getElementById('cp-hlall-track');
  var thumb = document.getElementById('cp-hlall-thumb');
  if (track) track.style.background = _customColors.hlAll ? 'var(--acc)' : 'var(--s3)';
  if (thumb) thumb.style.left = _customColors.hlAll ? '21px' : '3px';
  _applyCustomColors(); // preview en tiempo real
}

function _cancelColors() {
  // Restaurar estado previo al abrir el picker
  if (typeof _prevColors !== 'undefined') {
    _customColors.accent = _prevColors.accent;
    _customColors.bg     = _prevColors.bg;
    _customColors.hlAll  = _prevColors.hlAll;
    _applyCustomColors();
  }
  var m = document.getElementById('color-picker-modal');
  if (m) m.remove();
}

function _saveColors() {
  _customColors.accent = document.getElementById('cp-accent').value;
  _customColors.bg     = document.getElementById('cp-bg').value;
  localStorage.setItem('yve_accent', _customColors.accent);
  localStorage.setItem('yve_bg',     _customColors.bg);
  localStorage.setItem('yve_hl_all', _customColors.hlAll ? '1' : '0');
  _applyCustomColors();
  var m = document.getElementById('color-picker-modal');
  if (m) m.remove();
  showNotification('🎨 Colores guardados', 'success');
}

function _resetColors() {
  _customColors = { accent: '#3b82f6', bg: '#0f172a', hlAll: false };
  localStorage.removeItem('yve_accent');
  localStorage.removeItem('yve_bg');
  localStorage.removeItem('yve_hl_all');
  document.documentElement.removeAttribute('style');
  document.body.classList.remove('acentuar-todo');
  // Restaurar colores semánticos en sparklines
  [typeof AR_SPARKS!=='undefined' && AR_SPARKS, typeof AP_SPARKS!=='undefined' && AP_SPARKS].forEach(function(sparks){
    if (!sparks) return;
    sparks.forEach(function(s){ if (s.baseColor) s.color = s.baseColor; });
  });
  if (typeof injectSparklines === 'function') {
    if (typeof AR_SPARKS !== 'undefined') injectSparklines(AR_SPARKS);
    if (typeof AP_SPARKS !== 'undefined') injectSparklines(AP_SPARKS);
  }
  var m = document.getElementById('color-picker-modal');
  if (m) m.remove();
  showNotification('Colores restablecidos', 'info');
}
// ─────────────────────────────────────────────────────────────────────────────

// ── CSRF token (fetched after login, attached to all POST calls) ──────────
var _csrfToken = '';
(function(){
  fetch('/api/csrf_token')
    .then(function(r){ return r.ok ? r.json() : {token:''}; })
    .then(function(d){ _csrfToken = d.token || ''; })
    .catch(function(){});
})();
function _postJson(url, body) {
  return fetch(url, {
    method: 'POST',
    headers: {'Content-Type': 'application/json', 'X-CSRF-Token': _csrfToken},
    body: JSON.stringify(body || {})
  });
}

// ── Reclamaciones OTA (loop dato -> IA -> gate humano -> envío -> registro) ──
var _reclItems = [];
function _reclMoney(v){ return '€' + (Number(v)||0).toLocaleString('es-ES',{minimumFractionDigits:2}); }
function _reclEsc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
async function cargarReclamacionesOTA(){
  var wrap = document.getElementById('ar-recl-list');
  var resumen = document.getElementById('ar-recl-resumen');
  if (!wrap) return;
  try {
    var r = await fetch('/api/reclamaciones_ota/list');
    var d = await r.json();
    _reclItems = (d && d.items) || [];
    if (resumen) resumen.textContent = (d && d.n_pendientes) ? (d.n_pendientes + ' pendiente(s) · ' + _reclMoney(d.total_reclamable) + ' reclamable') : '';
    // Sin datos hay que BORRAR lo que hubiera, no marcharse.
    //
    // Antes esto era `if (!_reclItems.length) return;` con un comentario que
    // decia "deja el mensaje por defecto". El mensaje por defecto solo esta
    // ahi la primera vez: en cuanto se pintan tarjetas, el hueco ya no lo
    // tiene. Al cambiar de hotel el servidor devolvia 0 correctamente y esta
    // funcion se iba sin tocar nada, asi que las tarjetas del hotel ANTERIOR
    // se quedaban en pantalla. Desde fuera parecia que el filtro por hotel no
    // filtraba — y no habia forma de verlo comprobando la API, porque la API
    // contestaba bien; solo se veia mirando la pantalla sin recargar.
    if (!_reclItems.length) { wrap.innerHTML = _reclVacio(); return; }
    wrap.innerHTML = _reclItems.map(function(it,i){ return _reclCard(it,i); }).join('');
  } catch(e) {}
}
function _vacioCard(texto){
  return '<div class="empty card" style="padding:20px;text-align:center;color:var(--dim);' +
         'font-size:12px;border-style:dashed;border-radius:12px">' + texto + '</div>';
}
function _reclVacio(){
  return _vacioCard(t('recl.vacio', 'Cuando Yve detecte comisiones cobradas por encima del contrato, aparecerán aquí para reclamar.'));
}
function _reclCard(it, i){
  var badge, bg, col;
  if (it.estado==='ENVIADA'){ badge='✓ Enviada'; bg='rgba(34,197,94,.12)'; col='#22c55e'; }
  else if (it.estado==='DESCARTADA'){ badge='Descartada'; bg='rgba(148,163,184,.12)'; col='var(--mut)'; }
  else { badge='Pendiente'; bg='rgba(245,158,11,.12)'; col='#f59e0b'; }
  var head = '<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:8px">' +
    '<div style="min-width:0"><div style="font-weight:700;font-size:13px">' + _reclEsc(it.ota||'OTA') + ' · factura ' + _reclEsc(it.numero_factura||'') + '</div>' +
    '<div style="font-size:11px;color:var(--dim)">Cobrado ' + (it.comision_cobrada!=null?it.comision_cobrada+'%':'—') + ' vs contrato ' + (it.comision_contrato!=null?it.comision_contrato+'%':'—') +
      ' · a devolver <b style="color:#f87171">' + _reclMoney(it.importe_reclamable) + '</b></div></div>' +
    '<span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;background:'+bg+';color:'+col+'">'+badge+'</span></div>';
  if (it.estado==='ENVIADA'){
    return '<div class="card" style="padding:12px;border-radius:12px;opacity:.85">' + head +
      '<div style="font-size:12px;color:var(--mut)">Enviada a ' + _reclEsc(it.destinatario) + ' · ' + _reclEsc(it.fecha_enviada) + '</div></div>';
  }
  if (it.estado==='DESCARTADA'){
    return '<div class="card" style="padding:12px;border-radius:12px;opacity:.55">' + head + '</div>';
  }
  var body;
  if (!it.tiene_borrador){
    body = '<button onclick="_reclGenerar('+i+',this)" class="btn-run" style="font-size:12px">✍️ Redactar con IA</button>';
  } else {
    body = '<div style="display:flex;flex-direction:column;gap:8px">' +
      '<label style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">Enviar a</label>' +
      '<input id="recl-dest-'+i+'" value="'+_reclEsc(it.destinatario)+'" placeholder="email de la OTA" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px">' +
      '<input id="recl-asunto-'+i+'" value="'+_reclEsc(it.asunto)+'" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px;font-weight:600">' +
      '<textarea id="recl-cuerpo-'+i+'" rows="8" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px;font-family:inherit;line-height:1.5;resize:vertical">'+_reclEsc(it.cuerpo)+'</textarea>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap">' +
        '<button onclick="_reclEnviar('+i+',this)" class="btn-run" style="font-size:12px">✅ Aprobar y enviar</button>' +
        '<button onclick="_reclGenerar('+i+',this)" class="btn-ref" style="font-size:12px">🔄 Regenerar</button>' +
        '<button onclick="_reclDescartar('+i+')" class="btn-ref" style="font-size:12px">🗑 Descartar</button>' +
      '</div></div>';
  }
  return '<div class="card" style="padding:12px;border-radius:12px">' + head + body + '</div>';
}
async function _reclGenerar(i, btn){
  var it=_reclItems[i]; if(!it) return;
  if(btn){ btn.disabled=true; btn.textContent='✍️ Redactando…'; }
  try {
    var r = await _postJson('/api/reclamaciones_ota/generar', {id: it.id, idioma:'es'});
    var d = await r.json();
    if (d && d.ok){ it.asunto=d.asunto; it.cuerpo=d.cuerpo; it.tiene_borrador=true; cargarReclamacionesOTA(); }
    else { showNotification('✗ ' + ((d&&d.error)||'No se pudo redactar'), 'error'); if(btn){btn.disabled=false;btn.textContent='✍️ Redactar con IA';} }
  } catch(e){ showNotification('✗ '+e.message,'error'); if(btn){btn.disabled=false;btn.textContent='✍️ Redactar con IA';} }
}
async function _reclEnviar(i, btn){
  var it=_reclItems[i]; if(!it) return;
  var dest=(document.getElementById('recl-dest-'+i)||{}).value||'';
  var asunto=(document.getElementById('recl-asunto-'+i)||{}).value||'';
  var cuerpo=(document.getElementById('recl-cuerpo-'+i)||{}).value||'';
  if(!dest || dest.indexOf('@')<0){ showNotification('Introduce un email de destino válido','error'); return; }
  if(!confirm('¿Enviar la reclamación a '+dest+'?')) return;
  if(btn){ btn.disabled=true; btn.textContent='Enviando…'; }
  try {
    var r = await _postJson('/api/reclamaciones_ota/aprobar_enviar', {id: it.id, destinatario: dest, asunto: asunto, cuerpo: cuerpo});
    var d = await r.json();
    if (d && d.ok){ showNotification('✓ Reclamación enviada a '+dest,'success'); cargarReclamacionesOTA(); }
    else if (d && d.ya_enviada){ showNotification('ℹ '+d.error,'info'); cargarReclamacionesOTA(); }
    else if (d && d.sin_cifras){ showNotification('⚠ '+d.error,'error'); if(btn){btn.disabled=false;btn.textContent='✅ Aprobar y enviar';} }
    else { showNotification('✗ '+((d&&d.error)||'No se pudo enviar'),'error'); if(btn){btn.disabled=false;btn.textContent='✅ Aprobar y enviar';} }
  } catch(e){ showNotification('✗ '+e.message,'error'); if(btn){btn.disabled=false;btn.textContent='✅ Aprobar y enviar';} }
}
async function _reclDescartar(i){
  var it=_reclItems[i]; if(!it) return;
  if(!confirm('¿Descartar esta reclamación? No se enviará ningún email.')) return;
  try { await _postJson('/api/reclamaciones_ota/descartar', {id: it.id}); cargarReclamacionesOTA(); } catch(e){}
}

// ── Reclamar al PROVEEDOR (AP): rectificativa o abono. Mismo patron que la OTA. ──
var _reclApItems = [];
async function cargarAlbaranes(){
  var wrap = document.getElementById('alb-list'), res = document.getElementById('alb-resumen');
  if (!wrap) return;
  try {
    var r = await fetch('/api/albaranes'); var d = await r.json();
    if (!d || !d.ok) { wrap.innerHTML = '<div class="empty"><p>' + _cEsc((d && d.error) || 'Error') + '</p></div>'; return; }
    var s = d.resumen || {};
    res.textContent = s.n ? (s.n + ' ' + t('alb.kN','albaranes') + ' · ' + (s.facturados||0) + ' ' + t('alb.kFact','facturados') + ' · ' + (s.sin_facturar||0) + ' ' + t('alb.kSin','sin facturar') + ' · ' + eur(s.total)) : '';
    if (!s.n) { wrap.innerHTML = _vacioCard(t('alb.vacio','Todavía no hay albaranes. Súbelos con ⚡ Procesar archivos (PDF o foto) y aparecerán aquí con sus líneas.')); return; }
    var est = {ALBARAN_FACTURADO:['ok', t('alb.eFact','Facturado')], ALBARAN_SIN_FACTURAR:['sinpo', t('alb.eSin','Sin facturar')], SIN_CRUZAR:['', t('alb.eSinCruzar','Sin cruzar aún')]};
    var h = '<div class="tbl-wrap"><table class="tbl" style="width:100%;font-size:12px"><thead><tr><th>' + _cEsc(t('alb.cNum','Albarán')) + '</th><th>' + _cEsc(t('alb.cProv','Proveedor')) + '</th><th>' + _cEsc(t('alb.cFecha','Entrega')) + '</th><th style="text-align:right">' + _cEsc(t('alb.cTotal','Total sin IVA')) + '</th><th>' + _cEsc(t('alb.cEstado','Estado')) + '</th><th>' + _cEsc(t('alb.cFactura','Factura')) + '</th><th>' + _cEsc(t('alb.cLineas','Líneas')) + '</th></tr></thead><tbody>';
    d.albaranes.forEach(function(a, i){
      var e = est[a.estado] || ['', a.estado];
      h += '<tr style="cursor:pointer" onclick="var x=document.getElementById(\'alb-lin-' + i + '\');x.style.display=x.style.display===\'none\'?\'\':\'none\'">' +
        '<td><b>' + _cEsc(a.numero_albaran) + '</b>' + (a.referencia_pedido ? '<div style="font-size:11px;color:var(--mut)">' + _cEsc(t('alb.pedido','pedido')) + ' ' + _cEsc(a.referencia_pedido) + '</div>' : '') + '</td>' +
        '<td>' + _cEsc(a.proveedor) + '</td><td>' + _cEsc(a.fecha_entrega || '—') + '</td><td style="text-align:right">' + eur(a.total) + '</td>' +
        '<td><span class="ap-badge ' + e[0] + '">' + _cEsc(e[1]) + '</span></td>' +
        '<td>' + (a.numero_factura ? '<b>' + _cEsc(a.numero_factura) + '</b>' : '<span style="color:var(--mut)">—</span>') + (a.detalle ? '<div style="font-size:11px;color:var(--mut)">' + _cEsc(a.detalle) + '</div>' : '') + '</td>' +
        '<td>' + a.n_lineas + ' ▾</td></tr>';
      h += '<tr id="alb-lin-' + i + '" style="display:none"><td colspan="7" style="padding:4px 10px 10px 24px;background:rgba(127,127,127,.06)">';
      if (a.lineas.length) {
        h += '<table style="width:100%;font-size:12px"><thead><tr><th>#</th><th>' + _cEsc(t('alb.lDesc','Artículo')) + '</th><th style="text-align:right">' + _cEsc(t('alb.lCant','Cantidad')) + '</th><th style="text-align:right">' + _cEsc(t('alb.lPrecio','Precio')) + '</th><th style="text-align:right">' + _cEsc(t('alb.lImp','Importe')) + '</th></tr></thead><tbody>';
        a.lineas.forEach(function(l){ h += '<tr><td>' + l.n + '</td><td>' + _cEsc(l.descripcion) + '</td><td style="text-align:right">' + (l.cantidad==null?'—':l.cantidad + ' ' + _cEsc(l.unidad)) + '</td><td style="text-align:right">' + eur(l.precio_unitario) + '</td><td style="text-align:right">' + eur(l.importe) + '</td></tr>'; });
        h += '</tbody></table>';
      } else { h += '<span style="color:var(--mut);font-size:12px">' + _cEsc(t('alb.sinLineas','Sin líneas legibles en este albarán.')) + '</span>'; }
      h += (a.archivo ? '<div style="font-size:11px;color:var(--dim);margin-top:6px">' + _cEsc(a.archivo) + '</div>' : '') + '</td></tr>';
    });
    wrap.innerHTML = h + '</tbody></table></div>';
  } catch(e) { wrap.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
async function cargarReclamacionesAP(){
  var wrap = document.getElementById('ap-recl-list');
  var resumen = document.getElementById('ap-recl-resumen');
  if (!wrap) return;
  try {
    var r = await fetch('/api/reclamaciones_ap/list');
    var d = await r.json();
    _reclApItems = (d && d.items) || [];
    if (resumen) resumen.textContent = (d && d.n_pendientes)
      ? t('reclap.resumen', '{n} pendiente(s) · {total} en disputa').replace('{n}', d.n_pendientes).replace('{total}', _reclMoney(d.total_en_disputa))
      : '';
    if (!_reclApItems.length) { wrap.innerHTML = _vacioCard(t('reclap.vacio', 'Cuando una factura no cuadre con el albarán/pedido o se rechace, aparecerá aquí para pedir al proveedor la rectificativa o el abono.')); return; }
    wrap.innerHTML = _reclApItems.map(function(it,i){ return _reclApCard(it,i); }).join('');
  } catch(e) {}
}
function _reclApCard(it, i){
  var badge, bg, col;
  if (it.estado==='ENVIADA'){ badge=t('reclap.enviada','✓ Enviada'); bg='rgba(34,197,94,.12)'; col='#22c55e'; }
  else if (it.estado==='DESCARTADA'){ badge=t('reclap.descartada','Descartada'); bg='rgba(148,163,184,.12)'; col='var(--mut)'; }
  else { badge=t('reclap.pendiente','Pendiente'); bg='rgba(245,158,11,.12)'; col='#f59e0b'; }
  var tipo = it.tipo==='ABONO' ? t('reclap.tipoAbono','Pedir abono (rechazada)') : t('reclap.tipoCorreccion','Pedir rectificativa');
  var motivo = it.detalle || it.comentario || it.estado_matching || '';
  var head = '<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:8px">' +
    '<div style="min-width:0"><div style="font-weight:700;font-size:13px">' + _reclEsc(it.proveedor||'—') + ' · ' + t('reclap.factura','factura') + ' ' + _reclEsc(it.numero_factura||it.id) + (it.hotel?' · '+_reclEsc(it.hotel):'') + '</div>' +
    '<div style="font-size:11px;color:var(--dim)">' + _reclEsc(tipo) + ' · <b style="color:#f87171">' + _reclMoney(it.total_factura) + '</b>' + (motivo?' · '+_reclEsc(motivo):'') + '</div></div>' +
    '<span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;background:'+bg+';color:'+col+'">'+_reclEsc(badge)+'</span></div>';
  if (it.estado==='ENVIADA'){
    return '<div class="card" style="padding:12px;border-radius:12px;opacity:.85">' + head +
      '<div style="font-size:12px;color:var(--mut)">' + t('reclap.enviadaA','Enviada a') + ' ' + _reclEsc(it.destinatario) + ' · ' + _reclEsc(it.fecha_enviada) + '</div></div>';
  }
  if (it.estado==='DESCARTADA'){
    return '<div class="card" style="padding:12px;border-radius:12px;opacity:.55">' + head + '</div>';
  }
  var body;
  if (!it.tiene_borrador){
    body = '<button onclick="_reclApGenerar('+i+',this)" class="btn-run" style="font-size:12px">' + t('reclap.redactar','✍️ Redactar') + '</button>';
  } else {
    body = '<div style="display:flex;flex-direction:column;gap:8px">' +
      '<label style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + t('reclap.enviarA','Enviar a') + '</label>' +
      '<input id="reclap-dest-'+i+'" value="'+_reclEsc(it.destinatario)+'" placeholder="' + t('reclap.emailProveedor','email del proveedor') + '" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px">' +
      '<input id="reclap-asunto-'+i+'" value="'+_reclEsc(it.asunto)+'" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px;font-weight:600">' +
      '<textarea id="reclap-cuerpo-'+i+'" rows="9" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:8px;border-radius:8px;font-size:12px;font-family:inherit;line-height:1.5;resize:vertical">'+_reclEsc(it.cuerpo)+'</textarea>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap">' +
        '<button onclick="_reclApEnviar('+i+',this)" class="btn-run" style="font-size:12px">' + t('reclap.aprobarEnviar','✅ Aprobar y enviar') + '</button>' +
        '<button onclick="_reclApGenerar('+i+',this)" class="btn-ref" style="font-size:12px">' + t('reclap.regenerar','🔄 Regenerar') + '</button>' +
        '<button onclick="_reclApDescartar('+i+')" class="btn-ref" style="font-size:12px">' + t('reclap.descartar','🗑 Descartar') + '</button>' +
      '</div></div>';
  }
  return '<div class="card" style="padding:12px;border-radius:12px">' + head + body + '</div>';
}
async function _reclApGenerar(i, btn){
  var it=_reclApItems[i]; if(!it) return;
  var txt = btn ? btn.textContent : '';
  if(btn){ btn.disabled=true; btn.textContent=t('reclap.redactando','✍️ Redactando…'); }
  try {
    var r = await _postJson('/api/reclamaciones_ap/generar', {id: it.id, idioma: (typeof _i18nLang==='string' && _i18nLang==='en') ? 'en' : 'es'});
    var d = await r.json();
    if (d && d.ok){ cargarReclamacionesAP(); }
    else { showNotification('✗ ' + ((d&&d.error)||t('reclap.noRedacta','No se pudo redactar')), 'error'); if(btn){btn.disabled=false;btn.textContent=txt;} }
  } catch(e){ showNotification('✗ '+e.message,'error'); if(btn){btn.disabled=false;btn.textContent=txt;} }
}
async function _reclApEnviar(i, btn){
  var it=_reclApItems[i]; if(!it) return;
  var dest=(document.getElementById('reclap-dest-'+i)||{}).value||'';
  var asunto=(document.getElementById('reclap-asunto-'+i)||{}).value||'';
  var cuerpo=(document.getElementById('reclap-cuerpo-'+i)||{}).value||'';
  if(!dest || dest.indexOf('@')<0){ showNotification(t('reclap.emailInvalido','Introduce un email de destino válido'),'error'); return; }
  if(!confirm(t('reclap.confirmar','¿Enviar la reclamación a {dest}?').replace('{dest}', dest))) return;
  var txt = btn ? btn.textContent : '';
  if(btn){ btn.disabled=true; btn.textContent=t('reclap.enviando','Enviando…'); }
  try {
    var r = await _postJson('/api/reclamaciones_ap/aprobar_enviar', {id: it.id, destinatario: dest, asunto: asunto, cuerpo: cuerpo});
    var d = await r.json();
    if (d && d.ok){ showNotification(t('reclap.ok','✓ Reclamación enviada a {dest}').replace('{dest}', dest),'success'); cargarReclamacionesAP(); }
    else if (d && d.ya_enviada){ showNotification('ℹ '+d.error,'info'); cargarReclamacionesAP(); }
    else { showNotification('✗ '+((d&&d.error)||t('reclap.noEnvia','No se pudo enviar')),'error'); if(btn){btn.disabled=false;btn.textContent=txt;} }
  } catch(e){ showNotification('✗ '+e.message,'error'); if(btn){btn.disabled=false;btn.textContent=txt;} }
}
async function _reclApDescartar(i){
  var it=_reclApItems[i]; if(!it) return;
  if(!confirm(t('reclap.confirmarDescartar','¿Descartar esta reclamación? No se enviará ningún email.'))) return;
  try { await _postJson('/api/reclamaciones_ap/descartar', {id: it.id}); cargarReclamacionesAP(); } catch(e){}
}
// ─────────────────────────────────────────────────────────────────────────────

// ── Skeleton helpers ───────────────────────────────
function _skelOn(ids) {
  ids.forEach(function(id) {
    var el = document.getElementById(id);
    if (el) { el._origTxt = el.textContent; el.textContent = '——'; el.classList.add('skeleton'); }
  });
}
function _skelOff(ids) {
  ids.forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.classList.remove('skeleton');
  });
}
// ───────────────────────────────────────────────────

// Apply saved custom colors on load
_applyCustomColors();

// Apply saved theme on load
if (localStorage.getItem('yve_theme') === 'light') {
  document.body.classList.add('light-mode');
  var btn = document.getElementById('btn-theme');
  if (btn) btn.textContent = '🌙 Modo oscuro';
  var navBtn = document.getElementById('btn-theme-nav');
  if (navBtn) navBtn.textContent = '🌙';
}

async function cambiarIdioma(lang) {
  fetch('/api/set_lang/' + lang);   // fire-and-forget, no await
  await loadI18n(lang);
  document.querySelectorAll('.lang-btn').forEach(b => {
    var active = b.dataset.lang === lang;
    b.style.fontWeight = active ? '700' : '400';
    b.style.background = active ? 'rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)' : 'transparent';
    b.style.borderRadius = active ? '6px' : '';
    b.style.color = active ? 'var(--acc2)' : 'var(--tx)';
  });
  // Translate select option text (can't use data-i18n on options reliably)
  var apFilter = document.getElementById('ap-filter-estado');
  if (apFilter && _i18nStrMap[_i18nLang]) {
    var m = _i18nStrMap[_i18nLang];
    Array.from(apFilter.options).forEach(function(opt) {
      if (m[opt.textContent]) opt.textContent = m[opt.textContent];
    });
  }
}

loadAll();
setInterval(loadAll, 60000);

// ── Live clock ────────────────────────────────────────────────
var _CLOCK_LOCALES = { es:'es-ES', en:'en-GB', ca:'ca-ES', fr:'fr-FR', de:'de-DE', it:'it-IT', pt:'pt-PT' };
function _updateClock() {
  var el = document.getElementById('date-pill');
  if (!el) return;
  var now = new Date();
  var loc = _CLOCK_LOCALES[localStorage.getItem('yve_lang') || 'es'] || 'es-ES';
  var fecha;
  try { fecha = now.toLocaleDateString(loc, {weekday:'long', day:'numeric', month:'short'}); }
  catch(e) { fecha = now.toLocaleDateString('es-ES', {weekday:'long', day:'numeric', month:'short'}); }
  el.textContent = fecha + ' · ' + String(now.getHours()).padStart(2,'0') + ':' +
    String(now.getMinutes()).padStart(2,'0') + ':' + String(now.getSeconds()).padStart(2,'0');
}
_updateClock();
setInterval(_updateClock, 1000);
// ─────────────────────────────────────────────────────────────
// Init mobile lite mode
// Auto-apply saved language preference
(function() {
  var _savedLang = localStorage.getItem('yve_lang');
  if (_savedLang && _savedLang !== 'es') {
    loadI18n(_savedLang);
  }
})();
// ── Observador i18n: cualquier contenido nuevo del DOM se retraduce solo ──
var _i18nObsTimer = null, _i18nApplying = false;
function _applyPlaceholders(lang) {
  var map = _i18nStrMap[lang];
  if (!map) return;
  document.querySelectorAll('input[placeholder],textarea[placeholder]').forEach(function(el) {
    var p = el.getAttribute('placeholder');
    if (p && map[p] && map[p] !== p) el.setAttribute('placeholder', map[p]);
  });
}
(function() {
  var obs = new MutationObserver(function() {
    if (_i18nApplying) return;
    if (!_i18nLang || _i18nLang === 'es') return;
    clearTimeout(_i18nObsTimer);
    _i18nObsTimer = setTimeout(function() {
      _i18nApplying = true;
      try { _applyStrMap(_i18nLang); _applyPlaceholders(_i18nLang); }
      finally { setTimeout(function() {
        _i18nApplying = false;
        // segunda pasada: caza renders que llegaron durante la primera
        if (_i18nLang && _i18nLang !== 'es') { _applyStrMap(_i18nLang); _applyPlaceholders(_i18nLang); }
      }, 150); }
    }, 100);
  });
  obs.observe(document.body, {childList: true, subtree: true});
})();
// Changelog badge
if (localStorage.getItem('changelog_seen') !== '2026-06-v3') {
  const mb = document.getElementById('menu-badge');
  if (mb) mb.style.display = 'inline-block';
}
// Show keyboard hint on first visit
if (!localStorage.getItem('kbd_shown')) {
  localStorage.setItem('kbd_shown', '1');  // auto-set, no need to show hint
}
function _tourBannerSkip() {
  localStorage.setItem('tour_skipped', '1');
  var b = document.getElementById('tour-banner');
  if (b) b.remove();
}
function _tourBannerStart() {
  var b = document.getElementById('tour-banner');
  if (b) b.remove();
  startTour();
}
// Show tour offer on first login (only if not skipped and not done)
if (localStorage.getItem('tour_done') !== _TOUR_VER && !localStorage.getItem('tour_skipped')) {
  setTimeout(function() {
    var n = document.createElement('div');
    n.id = 'tour-banner';
    n.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);' +
      'background:var(--s1);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4);border-radius:16px;' +
      'padding:16px 20px;z-index:9000;display:flex;align-items:center;gap:16px;' +
      'box-shadow:0 8px 32px rgba(0,0,0,.5);max-width:420px;width:calc(100% - 32px)';
    n.innerHTML =
      '<div style="font-size:28px;flex-shrink:0">🗺️</div>' +
      '<div style="flex:1"><div style="font-weight:600;color:#f1f5f9;font-size:14px">¿Hacemos el tour?</div>' +
      '<div style="font-size:12px;color:#94a3b8;margin-top:3px">2 minutos para conocer todo Yve</div></div>' +
      '<div style="display:flex;gap:8px;flex-shrink:0">' +
        '<button onclick="_tourBannerSkip()" ' +
          'style="background:transparent;border:1px solid #334155;color:#64748b;padding:7px 13px;border-radius:8px;font-size:12px;cursor:pointer">Ahora no</button>' +
        '<button onclick="_tourBannerStart()" ' +
          'style="background:linear-gradient(135deg,var(--acc,#3b82f6),#7c3aed);border:none;color:#fff;padding:7px 13px;border-radius:8px;font-size:12px;cursor:pointer;font-weight:600">Empezar →</button>' +
      '</div>';
    document.body.appendChild(n);
    // Auto-dismiss after 12s
    setTimeout(function(){ if(n.parentNode) n.remove(); }, 12000);
  }, 2500);
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

// ── PWA: Service Worker, actualización, instalación y push ────────────────
var _deferredInstall = null;
var _swReg = null;

(function(){
  if (document.getElementById('yve-pwa-style')) return;
  var s = document.createElement('style'); s.id = 'yve-pwa-style';
  s.textContent = '@keyframes yveSlideUp{from{opacity:0;transform:translateX(-50%) translateY(20px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}';
  (document.head || document.documentElement).appendChild(s);
})();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', function() {
    navigator.serviceWorker.register('/sw.js', { scope: '/' }).then(function(reg) {
      _swReg = reg;
      if (reg.waiting && navigator.serviceWorker.controller) yveShowUpdateBanner(reg);
      reg.addEventListener('updatefound', function() {
        var nw = reg.installing; if (!nw) return;
        nw.addEventListener('statechange', function() {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) yveShowUpdateBanner(reg);
        });
      });
    }).catch(function(){});
    var _refreshing = false;
    navigator.serviceWorker.addEventListener('controllerchange', function() {
      if (_refreshing) return; _refreshing = true; window.location.reload();
    });
  });
}

// Prompt de instalación nativo (Android / escritorio Chrome / Edge)
window.addEventListener('beforeinstallprompt', function(e) {
  e.preventDefault(); _deferredInstall = e;
  var btn = document.getElementById('btn-install-pwa');
  if (btn) btn.style.display = 'inline-block';
  yvePwaMaybeShowInstall();
});
window.addEventListener('appinstalled', function() {
  _deferredInstall = null;
  try { localStorage.setItem('yve_pwa_installed', '1'); } catch(e){}
  var btn = document.getElementById('btn-install-pwa');
  if (btn) btn.style.display = 'none';
  yveRemoveBanner('yve-install-banner');
  if (typeof showNotification === 'function') showNotification('✓ Yve.01 instalada. Búscala en tu pantalla de inicio.', 'success');
});

function yveIsStandalone() {
  return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
}
function yveIsIOS() {
  return /iphone|ipad|ipod/i.test(navigator.userAgent) && !window.MSStream;
}

function yvePwaMaybeShowInstall() {
  try {
    if (yveIsStandalone()) return;
    if (localStorage.getItem('yve_pwa_install_dismissed') === '1') return;
    if (localStorage.getItem('yve_pwa_installed') === '1') return;
  } catch(e){}
  if (document.getElementById('yve-install-banner')) return;
  var ios = yveIsIOS();
  if (!_deferredInstall && !ios) return;
  var msg = ios
    ? '📲 Instala Yve: toca <b>Compartir</b> y luego <b>Añadir a pantalla de inicio</b>'
    : '📲 Instala Yve para acceso rápido';
  var btnHtml = ios ? '' :
    '<button onclick="yvePwaInstall()" style="background:#fff;color:#2563eb;border:none;padding:8px 16px;border-radius:9px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit;white-space:nowrap">Instalar</button>';
  yveShowBanner('yve-install-banner', msg, btnHtml, 'yveDismissInstall', '#3b82f6');
}
function yvePwaInstall() {
  if (!_deferredInstall) { yveRemoveBanner('yve-install-banner'); return; }
  _deferredInstall.prompt();
  _deferredInstall.userChoice.then(function(){ _deferredInstall = null; yveRemoveBanner('yve-install-banner'); });
}
function yveDismissInstall() {
  try { localStorage.setItem('yve_pwa_install_dismissed', '1'); } catch(e){}
  yveRemoveBanner('yve-install-banner');
}

function yveShowUpdateBanner(reg) {
  window._yveWaitingReg = reg;
  yveShowBanner('yve-update-banner', '🔄 Actualización disponible',
    '<button onclick="yveApplyUpdate()" style="background:#fff;color:#166534;border:none;padding:8px 16px;border-radius:9px;font-weight:700;font-size:13px;cursor:pointer;font-family:inherit;white-space:nowrap">Actualizar</button>',
    null, '#22c55e');
}
function yveApplyUpdate() {
  var reg = window._yveWaitingReg;
  yveRemoveBanner('yve-update-banner');
  if (reg && reg.waiting) reg.waiting.postMessage('SKIP_WAITING');
  else window.location.reload();
}

function yveShowBanner(id, htmlMsg, btnHtml, dismissFn, accent) {
  if (document.getElementById(id)) return;
  accent = accent || '#3b82f6';
  var bar = document.createElement('div');
  bar.id = id;
  bar.style.cssText = 'position:fixed;left:50%;transform:translateX(-50%);bottom:calc(16px + env(safe-area-inset-bottom));z-index:10000;display:flex;align-items:center;gap:14px;max-width:min(560px,calc(100% - 24px));width:max-content;background:linear-gradient(135deg,' + accent + ',#2563eb);color:#fff;padding:12px 14px 12px 18px;border-radius:14px;box-shadow:0 10px 40px rgba(0,0,0,.35);font-size:13px;font-weight:500;animation:yveSlideUp .3s ease';
  var txt = document.createElement('div'); txt.style.cssText = 'flex:1;line-height:1.4'; txt.innerHTML = htmlMsg;
  bar.appendChild(txt);
  if (btnHtml) { var wrap = document.createElement('div'); wrap.innerHTML = btnHtml; if (wrap.firstChild) bar.appendChild(wrap.firstChild); }
  var x = document.createElement('button');
  x.textContent = '✕'; x.setAttribute('aria-label', 'Cerrar');
  x.style.cssText = 'background:transparent;border:none;color:rgba(255,255,255,.85);font-size:15px;cursor:pointer;padding:4px 6px;font-family:inherit;line-height:1';
  x.onclick = function(){ if (dismissFn && window[dismissFn]) window[dismissFn](); else yveRemoveBanner(id); };
  bar.appendChild(x);
  document.body.appendChild(bar);
}
function yveRemoveBanner(id) { var b = document.getElementById(id); if (b) b.remove(); }

// ── Push (Web Push API) ───────────────────────────────────────────────────
function yveUrlB64ToUint8(base64) {
  var padding = '='.repeat((4 - base64.length % 4) % 4);
  var b64 = (base64 + padding).replace(/-/g, '+').replace(/_/g, '/');
  var raw = atob(b64); var arr = new Uint8Array(raw.length);
  for (var i = 0; i < raw.length; i++) arr[i] = raw.charCodeAt(i);
  return arr;
}
function yvePushSupported() {
  return ('serviceWorker' in navigator) && ('PushManager' in window) && ('Notification' in window);
}
async function yvePushSubscribe() {
  if (!yvePushSupported()) { showNotification('Tu navegador no soporta notificaciones push', 'error'); return false; }
  try {
    var perm = await Notification.requestPermission();
    if (perm !== 'granted') { showNotification('Permiso de notificaciones denegado', 'warning'); return false; }
    var reg = _swReg || await navigator.serviceWorker.ready;
    var r = await fetch('/api/push/public_key'); var j = await r.json();
    if (!j.publicKey) { showNotification('Push aún no está configurado en el servidor', 'warning'); return false; }
    var sub = await reg.pushManager.getSubscription();
    if (!sub) {
      sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: yveUrlB64ToUint8(j.publicKey) });
    }
    await _postJson('/api/push/subscribe', { subscription: sub });
    showNotification('✓ Notificaciones push activadas en este dispositivo', 'success');
    return true;
  } catch (e) { showNotification('No se pudo activar push: ' + (e && e.message ? e.message : e), 'error'); return false; }
}
async function yvePushUnsubscribe() {
  try {
    var reg = _swReg || await navigator.serviceWorker.ready;
    var sub = await reg.pushManager.getSubscription();
    if (sub) {
      await _postJson('/api/push/unsubscribe', { endpoint: sub.endpoint });
      await sub.unsubscribe();
    }
    return true;
  } catch (e) { return false; }
}
async function yvePushTest() {
  try {
    var r = await _postJson('/api/push/test', {});
    var d = await r.json();
    showNotification((d.ok ? '✓ ' : '⚠ ') + (d.message || d.error || 'Push'), d.ok ? 'success' : 'warning');
  } catch (e) { showNotification('✗ Error probando push', 'error'); }
}

// ── Toast notifications ──────────────────────────────────────────────────
var _toastTimeout;
function showNotification(msg, type = 'info') {
  let toast = document.getElementById('yve-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'yve-toast';
    toast.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(20px);opacity:0;max-width:420px;width:calc(100% - 40px);padding:12px 18px;border-radius:12px;font-size:13px;font-weight:500;z-index:9999;transition:background-color .25s ease,border-color .25s ease,color .25s ease,box-shadow .25s ease,transform .25s ease,opacity .25s ease;pointer-events:none;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.4)';
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
    // Paso 1 — siempre centrado, no highlight
    el: null, tab: null, pos: 'center',
    title: '👋 Bienvenido a Yve.01',
    text: 'El sistema de finanzas hoteleras que automatiza AR, AP, DRR y reporting. Este tour te lleva por cada módulo de izquierda a derecha — 3 minutos y ya lo dominas todo. Arrástrame si te estorbo: me acoplo solo donde me sueltes.'
  },
  {
    el: '#ar-stats-section', tab: 'ar', pos: 'auto',
    title: '📥 AR — Comisiones OTA',
    text: 'Verifica automáticamente las comisiones de Booking.com y Expedia. Facturas procesadas, importe total, discrepancias reclamables y certificados DI pendientes. El número rojo son euros que puedes recuperar.'
  },
  {
    el: '#stats-ap-grid', tab: 'ap', pos: 'auto',
    title: '📦 AP — 3-way Matching',
    text: 'Para cada factura de proveedor, Yve cruza 3 documentos: factura, pedido (PO) y albarán. Si cuadra todo → Match OK automático. Si hay diferencia → alerta y email al proveedor generado con IA.'
  },
  {
    el: '#drr-metrics', tab: 'drr', pos: 'auto',
    title: '📊 DRR — Daily Revenue Report',
    text: 'Arrastra tu archivo .xlsm aquí. Yve extrae RevPAR, ADR, GOP%, ocupación y las 7.000+ líneas del Trial Balance en segundos. Detecta Out of Balance automáticamente y te avisa al instante.'
  },
  {
    el: '#banco-stats', tab: 'banco', pos: 'auto',
    title: '🏦 Banco — Conciliación',
    text: 'Cruza automáticamente el extracto bancario con las facturas de proveedores. Identifica movimientos no conciliados, diferencias de importe y pagos duplicados. Desde 8 horas a 2 minutos.'
  },
  {
    el: '#panel-notif', tab: 'notif', pos: 'auto',
    title: '🔔 Notificaciones',
    text: 'Configura alertas automáticas por email o Telegram: discrepancias OTA, facturas sin firmar, Out of Balance en el DRR o stock bajo en F&B. Yve te avisa proactivamente.'
  },
  {
    el: '#fb-resumen', tab: 'fb', pos: 'auto',
    title: '🍽️ F&B Cost Control',
    text: 'Calcula el Food Cost real vs teórico por categoría. Conecta los datos POS, recetas e inventario. Detecta mermas, identifica qué platos tienen mejor margen y optimiza el rendimiento del restaurante.'
  },
  {
    el: '#ar-real-stats', tab: 'ar_real', pos: 'auto',
    title: '🏢 AR Real — Grupos Corporativos',
    text: 'Gestión completa de clientes corporativos: emite facturas, controla el aging (0-30 / 31-60 / +90 días), cobra con un clic y envía recordatorios automáticos por email.'
  },
  {
    el: '#mh-kpis', tab: 'multi_hotel', pos: 'auto',
    title: '🌍 Multi-Hotel — Vista de Grupo',
    text: 'Para el Financial Controller del grupo: KPIs consolidados, ranking de performance por hotel, tendencia de 6 meses y alertas centralizadas. Una pantalla, todo el grupo.',
    action: function() {
      if (typeof loadMultiHotel === 'function') { _mh_loaded = false; loadMultiHotel(); }
    }
  }
];

// ── Tour state ────────────────────────────────────────────────────────
var _tourActive = false;
var _tourStep   = 0;
var _tourBoxPos = 'center';  // current position: center|tl|tr|bl|br
var _tourScrollHandler = null;
var _tourResizeHandler = null;
var _tourCurrentTarget = null;

// ── Position presets ──────────────────────────────────────────────────

function _tourBoxCoords(pos, bw, bh) {
  var vw = window.innerWidth, vh = window.innerHeight;
  var pad = 20;
  // las posiciones superiores empiezan debajo de la barra de pestañas para no tapar el nombre del apartado
  var tabs = document.querySelector('.tabs');
  var navH = tabs ? Math.max(56, Math.round(tabs.getBoundingClientRect().bottom + 10)) : 56;
  switch(pos) {
    case 'center': return { top: Math.round((vh - bh)/2), left: Math.round((vw - bw)/2) };
    case 'tl':     return { top: navH + pad, left: pad };
    case 'tr':     return { top: navH + pad, left: vw - bw - pad };
    case 'bl':     return { top: vh - bh - pad, left: pad };
    case 'br':     return { top: vh - bh - pad, left: vw - bw - pad };
    case 'tc':     return { top: navH + pad, left: Math.round((vw - bw)/2) };
    case 'bc':     return { top: vh - bh - pad, left: Math.round((vw - bw)/2) };
    case 'lc':     return { top: Math.round((vh - bh)/2), left: pad };
    case 'rc':     return { top: Math.round((vh - bh)/2), left: vw - bw - pad };
    default:       return { top: Math.round((vh - bh)/2), left: Math.round((vw - bw)/2) };
  }
}

// ── Choose best auto-position avoiding the highlighted element ────────
function _autoPickPos(targetRect) {
  var candidatos = ['tl', 'tc', 'tr', 'lc', 'rc', 'bl', 'bc', 'br'];
  var pool = candidatos.filter(function(p) { return p !== _tourBoxPos; });
  return pool[Math.floor(Math.random() * pool.length)];
}



// ── Redraw the spotlight canvas ───────────────────────────────────────
function _drawSpotlight(target) {
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (!canvas) return;
  canvas.width  = window.innerWidth;
  canvas.height = window.innerHeight;
  var ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (target) {
    // ── Expand to the nearest full panel/section ─────────────────────
    // Walk up from the target to find a panel-level container
    var highlight = target;
    var el = target.parentElement;
    while (el && el !== document.body) {
      var id = el.id || '';
      var cls = el.className || '';
      // Stop at panel-*, card, or tab content containers
      if (/^panel-/.test(id) || cls.indexOf('panel') >= 0 ||
          id === 'app-body' || el.tagName === 'MAIN') {
        highlight = el;
        break;
      }
      // Also stop if the element is wide enough to be a section
      var r0 = el.getBoundingClientRect();
      if (r0.width > window.innerWidth * 0.7 && r0.height > 100) {
        highlight = el;
        break;
      }
      el = el.parentElement;
    }

    var rect = highlight.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) {
      ctx.fillStyle = 'rgba(0,0,0,0.78)';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      return;
    }

    var pad = 8;
    // Clamp to viewport
    var x = Math.max(0, rect.left - pad);
    var y = Math.max(0, rect.top - pad);
    var w = Math.min(canvas.width  - x, rect.width  + pad*2);
    var h = Math.min(canvas.height - y, rect.height + pad*2);
    var r = 16;

    // Dark overlay
    ctx.fillStyle = 'rgba(0,0,0,0.72)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Cut hole for the highlighted section
    ctx.save();
    ctx.globalCompositeOperation = 'destination-out';
    ctx.beginPath();
    if (ctx.roundRect) {
      ctx.roundRect(x, y, w, h, r);
    } else {
      ctx.moveTo(x+r,y); ctx.lineTo(x+w-r,y); ctx.arcTo(x+w,y,x+w,y+r,r);
      ctx.lineTo(x+w,y+h-r); ctx.arcTo(x+w,y+h,x+w-r,y+h,r);
      ctx.lineTo(x+r,y+h); ctx.arcTo(x,y+h,x,y+h-r,r);
      ctx.lineTo(x,y+r); ctx.arcTo(x,y,x+r,y,r);
      ctx.closePath();
    }
    ctx.fill();
    ctx.restore();

    // Blue glow ring around the section
    ctx.shadowColor = '#3b82f6';
    ctx.shadowBlur = 24;
    ctx.strokeStyle = 'rgba(59,130,246,0.85)';
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    if (ctx.roundRect) ctx.roundRect(x, y, w, h, r);
    else ctx.rect(x, y, w, h);
    ctx.stroke();
    ctx.shadowBlur = 0;
  } else {
    ctx.fillStyle = 'rgba(0,0,0,0.72)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
}

// ── Build/update the tour box HTML ───────────────────────────────────
function _renderTourBox(step) {
  var box = document.getElementById('tour-box');
  var isNew = !box;
  if (!box) {
    box = document.createElement('div');
    box.id = 'tour-box';
    box.style.cssText =
      'position:fixed;background:#0f172a;border:2px solid var(--acc,#3b82f6);border-radius:16px;' +
      'padding:18px 20px 16px;max-width:360px;width:calc(100vw - 32px);z-index:10000;' +
      'box-shadow:0 20px 60px rgba(0,0,0,.85),0 0 60px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.08);' +
      'pointer-events:all;font-family:Inter,system-ui,sans-serif;color:#f1f5f9;' +
      'animation:tourBoxIn .3s cubic-bezier(.34,1.56,.64,1);user-select:none';
    document.body.appendChild(box);
  }
  if (!isNew) {
    // fade suave del contenido al cambiar de paso
    box.style.transition = 'opacity .15s ease';
    box.style.opacity = '0.4';
    setTimeout(function() { box.style.opacity = '1'; }, 160);
  }

  box.innerHTML =
    // Header row: drag grip + close
    '<div style="display:flex;align-items:center;justify-content:flex-end;margin-bottom:8px">' +
      '<button onclick="endTour()" style="background:none;border:none;color:#475569;font-size:20px;' +
        'cursor:pointer;padding:0 2px;line-height:1;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s" ' +
        'onmouseover="this.style.color=\'#94a3b8\'" onmouseout="this.style.color=\'#475569\'">×</button>' +
    '</div>' +
    // Title
    '<div style="font-size:17px;font-weight:800;color:#f1f5f9;margin-bottom:8px;line-height:1.3">' +
      step.title + '</div>' +
    // Body
    '<div style="font-size:13px;color:#94a3b8;line-height:1.7;margin-bottom:16px">' +
      step.text + '</div>' +
    // Progress dots + nav buttons
    '<div style="display:flex;align-items:center;justify-content:space-between">' +
      '<div style="display:flex;gap:5px;align-items:center">' +
        _tourSteps.map(function(_,i) {
          var active = i === _tourStep;
          return '<div style="transition:background-color .25s,border-color .25s,color .25s,box-shadow .25s,transform .25s,opacity .25s;border-radius:' + (active ? '4px' : '50%') + ';' +
            'background:' + (active ? 'var(--acc,#3b82f6)' : 'rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.25)') + ';' +
            'width:' + (active ? '18px' : '7px') + ';height:7px"></div>';
        }).join('') +
      '</div>' +
      '<div style="display:flex;gap:6px">' +
        (_tourStep > 0 ?
          '<button onclick="prevTourStep()" style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);' +
          'color:#94a3b8;padding:7px 13px;border-radius:9px;font-size:13px;cursor:pointer">← Atrás</button>' : '') +
        '<button onclick="nextTourStep()" style="background:var(--acc,#3b82f6);border:none;color:#fff;' +
          'padding:7px 16px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;' +
          'box-shadow:0 4px 12px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)">' +
          (_tourStep === _tourSteps.length - 1 ? '✓ Finalizar' : 'Siguiente →') +
        '</button>' +
      '</div>' +
    '</div>';
  _initTourDrag(box);
  if (typeof _i18nAfterRender === 'function') _i18nAfterRender();
}

// ── Drag & imán: arrastra la burbuja y se acopla sola a la zona más cercana ──
function _initTourDrag(box) {
  box.style.cursor = 'grab';
  function _startDrag(cx, cy) {
    var rect = box.getBoundingClientRect();
    box.style.transition = 'none';
    box.style.animation = 'none';
    box.style.left = rect.left + 'px';
    box.style.top = rect.top + 'px';
    box.style.transform = 'none';
    box.style.willChange = 'transform';
    return { x: cx, y: cy };
  }
  function _makeMover(start) {
    var dx = 0, dy = 0, raf = null, moved = false;
    function paint() { raf = null; box.style.transform = 'translate3d(' + dx + 'px,' + dy + 'px,0)'; }
    return {
      move: function(cx, cy) {
        moved = true;
        dx = cx - start.x; dy = cy - start.y;
        if (!raf) raf = requestAnimationFrame(paint);
      },
      end: function() {
        if (raf) cancelAnimationFrame(raf);
        box.style.willChange = '';
        if (!moved) return;
        box.style.transform = 'translate3d(' + dx + 'px,' + dy + 'px,0)';
        var r = box.getBoundingClientRect();
        box.style.transform = 'none';
        box.style.left = r.left + 'px';
        box.style.top = r.top + 'px';
        _snapTourBox(box);
      }
    };
  }
  box.onmousedown = function(e) {
    if (e.target.closest('button')) return;
    e.preventDefault();
    box.style.cursor = 'grabbing';
    var m = _makeMover(_startDrag(e.clientX, e.clientY));
    function mv(ev) { m.move(ev.clientX, ev.clientY); }
    function up() {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      box.style.cursor = 'grab';
      m.end();
    }
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  };
  box.ontouchstart = function(e) {
    if (e.target.closest('button')) return;
    var t0 = e.touches[0];
    var m = _makeMover(_startDrag(t0.clientX, t0.clientY));
    function mv(ev) { m.move(ev.touches[0].clientX, ev.touches[0].clientY); ev.preventDefault(); }
    function up() {
      box.removeEventListener('touchmove', mv);
      box.removeEventListener('touchend', up);
      m.end();
    }
    box.addEventListener('touchmove', mv, {passive:false});
    box.addEventListener('touchend', up);
  };
}

function _snapTourBox(box) {
  var vw = window.innerWidth, vh = window.innerHeight;
  var r = box.getBoundingClientRect();
  var cx = r.left + r.width/2, cy = r.top + r.height/2;
  var col = cx < vw/3 ? 'l' : (cx > vw*2/3 ? 'r' : 'c');
  var row = cy < vh/3 ? 't' : (cy > vh*2/3 ? 'b' : 'c');
  var pos;
  if (col === 'c' && row === 'c') pos = 'center';
  else if (row === 'c') pos = (col === 'l') ? 'lc' : 'rc';
  else if (col === 'c') pos = (row === 't') ? 'tc' : 'bc';
  else pos = row + col;
  _tourBoxPos = pos;
  var c = _tourBoxCoords(pos, r.width, r.height);
  box.style.transition = 'top .22s cubic-bezier(.22,.9,.35,1), left .22s cubic-bezier(.22,.9,.35,1)';
  box.style.top = c.top + 'px';
  box.style.left = c.left + 'px';
  setTimeout(function(){ box.style.transition = 'none'; }, 260);
}

// ── Apply position to box ─────────────────────────────────────────────
function _applyTourBoxPos(targetRect) {
  var box = document.getElementById('tour-box');
  if (!box) return;
  var bw = box.offsetWidth  || 360;
  var bh = box.offsetHeight || 280;

  // For step 0 always center; for others auto or stored pos
  var pos = _tourBoxPos;
  if (pos === 'auto') pos = _autoPickPos(targetRect || null);
  _tourBoxPos = pos;

  var coords = _tourBoxCoords(pos, bw, bh);
  // si pisa la zona iluminada, deslizar justo debajo (o encima) manteniendo el carril horizontal
  if (targetRect) {
    var vh = window.innerHeight;
    var solapan = !(coords.left > targetRect.right || coords.left + bw < targetRect.left ||
                    coords.top > targetRect.bottom || coords.top + bh < targetRect.top);
    if (solapan) {
      var debajo = targetRect.bottom + 14;
      var encima = targetRect.top - bh - 14;
      if (debajo + bh <= vh - 12) coords.top = Math.round(debajo);
      else if (encima >= 64) coords.top = Math.round(encima);
      else coords.top = Math.max(64, vh - bh - 20);
    }
  }
  // desliza suavemente hasta la nueva posición (si ya estaba colocada)
  if (box.style.top) {
    box.style.transition = 'top .38s cubic-bezier(.22,.9,.35,1), left .38s cubic-bezier(.22,.9,.35,1), opacity .15s ease';
    setTimeout(function() { box.style.transition = 'none'; }, 420);
  }
  box.style.top  = coords.top  + 'px';
  box.style.left = coords.left + 'px';
  box.style.transform = '';
}



// ── Main step renderer ────────────────────────────────────────────────
function _showTourStep() {
  if (_tourStep >= _tourSteps.length) { endTour(); return; }
  var step = _tourSteps[_tourStep];
  _tourCurrentTarget = null;

  // Set position for this step
  if (_tourStep === 0) {
    _tourBoxPos = 'center';  // Always center for welcome
  } else if (step.pos && step.pos !== 'auto') {
    _tourBoxPos = step.pos;
  } else {
    _tourBoxPos = 'auto';    // Will auto-pick after finding target
  }

  // Clear previous highlight
  document.querySelectorAll('[data-tour-active]').forEach(function(el) {
    el.style.outline = '';
    el.style.zIndex = '';
    el.style.position = '';
    el.removeAttribute('data-tour-active');
  });

  // Switch tab first
  if (step.tab) {
    var tabEl = document.getElementById('tab-' + step.tab) ||
                document.getElementById('tab-' + step.tab.replace(/_/g, '-'));
    if (tabEl) {
      switchTab(step.tab, tabEl);
      // iluminar también el nombre/emoji de la pestaña activa
      tabEl.setAttribute('data-tour-active', '1');
      tabEl.style.position = 'relative';
      tabEl.style.zIndex = '9950';
    }
  }

  // Run step action
  if (step.action && typeof step.action === 'function') {
    setTimeout(function() { try { step.action(); } catch(e) {} }, 150);
  }

  // Tabs with async data loading need longer delays
  var _asyncTabs = {'ar_real': 900, 'multi_hotel': 1000};
  var delay = step.tab ? (_asyncTabs[step.tab] || 400) : 50;
  setTimeout(function() {
    var target = step.el ? document.querySelector(step.el) : null;
    _tourCurrentTarget = target;

    if (target) {
      // Scroll target into view
      target.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
      target.setAttribute('data-tour-active', '1');
      target.style.position = 'relative';
      target.style.zIndex   = '9950';
    }

    // Wait for scroll to settle, then draw (retry if el not populated yet)
    var _drawAttempt = function(attemptsLeft) {
      if (step.el) {
        var fresh = document.querySelector(step.el);
        // For async tabs, wait until the element has visible content
        var isEmpty = fresh && (fresh.children.length === 0 || fresh.innerHTML.trim() === '');
        if ((!fresh || isEmpty) && attemptsLeft > 0) {
          setTimeout(function() { _drawAttempt(attemptsLeft - 1); }, 200);
          return;
        }
        if (fresh) { _tourCurrentTarget = fresh; target = fresh; }
      }
      _drawSpotlight(target);
      _renderTourBox(step);
      // setTimeout en vez de rAF: rAF se congela si la ventana está tapada/minimizada
      setTimeout(function() {
        _applyTourBoxPos(target ? target.getBoundingClientRect() : null);
      }, 0);
    };
    setTimeout(function() { _drawAttempt(6); }, target ? 150 : 0);
  }, delay);
}

// ── Navigation ────────────────────────────────────────────────────────
function nextTourStep() {
  _tourStep++;
  _showTourStep();
}

function prevTourStep() {
  if (_tourStep > 0) { _tourStep--; _showTourStep(); }
}

// ── Start / End ───────────────────────────────────────────────────────
function startTour() {
  _tourActive = true;
  _tourStep   = 0;
  _tourBoxPos = 'center';

  // Close menu
  var mm = document.getElementById('main-menu');
  if (mm) mm.classList.remove('open');

  // Remove old elements
  var oldBox = document.getElementById('tour-box');
  if (oldBox) oldBox.remove();

  // Overlay
  var overlay = document.getElementById('tour-overlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'tour-overlay';
    document.body.appendChild(overlay);
  }
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9900;pointer-events:all;cursor:default;display:block;background:transparent';
  overlay.onclick = null;  // clicking outside does NOTHING

  // Spotlight canvas
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (!canvas) {
    canvas = document.createElement('canvas');
    canvas.id = 'tour-spotlight-canvas';
    canvas.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:9901';
    document.body.appendChild(canvas);
  }
  canvas.style.display = 'block';

  // Scroll listener — redraw spotlight when user scrolls
  if (_tourScrollHandler) window.removeEventListener('scroll', _tourScrollHandler, true);
  _tourScrollHandler = function() {
    if (!_tourActive || !_tourCurrentTarget) return;
    _drawSpotlight(_tourCurrentTarget);
    _applyTourBoxPos(_tourCurrentTarget.getBoundingClientRect());
  };
  window.addEventListener('scroll', _tourScrollHandler, { passive: true, capture: true });

  // Resize listener — redraw on window resize
  if (_tourResizeHandler) window.removeEventListener('resize', _tourResizeHandler);
  _tourResizeHandler = function() {
    if (!_tourActive) return;
    var canvas = document.getElementById('tour-spotlight-canvas');
    if (canvas) { canvas.width = window.innerWidth; canvas.height = window.innerHeight; }
    _drawSpotlight(_tourCurrentTarget);
    _applyTourBoxPos(_tourCurrentTarget ? _tourCurrentTarget.getBoundingClientRect() : null);
  };
  window.addEventListener('resize', _tourResizeHandler);

  _showTourStep();
}

// ── Confetti celebration ─────────────────────────────────────────────
function _paletteColors() {
  var cs = getComputedStyle(document.documentElement);
  var acc  = (cs.getPropertyValue('--acc')  || '#3b82f6').trim() || '#3b82f6';
  var acc2 = (cs.getPropertyValue('--acc2') || '#60a5fa').trim() || '#60a5fa';
  var acc3 = (cs.getPropertyValue('--acc3') || '#93c5fd').trim() || '#93c5fd';
  return [acc, acc2, acc3, '#f1f5f9', '#fbbf24', acc];
}

function _launchConfetti() {
  var colors = _paletteColors();
  var count = 0, max = 160;
  var interval = setInterval(function() {
    if (count++ > max) { clearInterval(interval); return; }
    var el = document.createElement('div');
    var sz = 5 + Math.random()*10;
    var c = colors[Math.floor(Math.random()*colors.length)];
    var shape = Math.random();
    el.style.cssText = 'position:fixed;top:' + (-20 - Math.random()*30) + 'px;left:' + (Math.random()*100) + '%;' +
      'width:' + sz + 'px;height:' + (shape > .7 ? sz*0.4 : sz) + 'px;' +
      'border-radius:' + (shape > .5 ? '50%' : '2px') + ';' +
      'background:' + c + ';box-shadow:0 0 ' + (4+Math.random()*8) + 'px ' + c + ';' +
      'opacity:1;z-index:99999;pointer-events:none;' +
      'animation:confettiFall ' + (1.6+Math.random()*1.6) + 's cubic-bezier(.3,.1,.6,1) forwards;' +
      'animation-delay:' + (Math.random()*0.25) + 's';
    document.body.appendChild(el);
    setTimeout(function(){ el.remove(); }, 4000);
  }, 12);
  // onda expansiva central con el color de la paleta
  var ring = document.createElement('div');
  ring.style.cssText = 'position:fixed;top:50%;left:50%;width:20px;height:20px;border-radius:50%;' +
    'border:3px solid var(--acc,#3b82f6);transform:translate(-50%,-50%);z-index:99998;pointer-events:none;' +
    'animation:tourRing .9s ease-out forwards';
  document.body.appendChild(ring);
  setTimeout(function(){ ring.remove(); }, 1000);
}
// ─────────────────────────────────────────────────────────────────────
function endTour() {
  _tourActive = false;
  _tourStep   = 0;
  _tourCurrentTarget = null;

  // Remove listeners
  if (_tourScrollHandler) { window.removeEventListener('scroll', _tourScrollHandler, true); _tourScrollHandler = null; }
  if (_tourResizeHandler) { window.removeEventListener('resize', _tourResizeHandler); _tourResizeHandler = null; }

  // Clear highlights
  document.querySelectorAll('[data-tour-active]').forEach(function(el) {
    el.style.outline = '';
    el.style.zIndex  = '';
    el.style.position = '';
    el.removeAttribute('data-tour-active');
  });

  // Hide overlay + canvas
  var ov = document.getElementById('tour-overlay');
  if (ov) ov.style.display = 'none';
  var canvas = document.getElementById('tour-spotlight-canvas');
  if (canvas) {
    var ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    canvas.style.display = 'none';
  }

  // Remove tour box
  var box = document.getElementById('tour-box');
  if (box) box.remove();

  // Mark done
  localStorage.setItem('tour_done', _TOUR_VER);

  // Mover a global scope para que onclick funcione
  window._closeCongrats = function() { var c=document.getElementById('tour-congrats'); if(c) c.remove(); };
  window._startAR = function() { window._closeCongrats(); var t=document.getElementById('tab-ar'); if(t) switchTab('ar',t); };
// ── Celebración final: tarjeta de bienvenida sobre el dashboard ──
  _launchConfetti();
  setTimeout(function() {
    var card = document.createElement('div');
    card.id = 'tour-congrats';
    card.style.cssText = 'position:fixed;inset:0;z-index:10001;display:flex;align-items:center;' +
      'justify-content:center;background:rgba(0,0,0,.7);animation:tourBoxIn .4s ease';
    card.innerHTML =
      '<div style="background:#0f172a;border:2px solid var(--acc,#3b82f6);border-radius:20px;padding:36px 40px;' +
        'text-align:center;max-width:420px;width:calc(100% - 40px);' +
        'box-shadow:0 24px 80px rgba(0,0,0,.9),0 0 60px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.25)">' +
        '<div style="font-size:52px;margin-bottom:12px;animation:tourBoxIn .6s cubic-bezier(.34,1.56,.64,1)">🎉</div>' +
        '<div style="font-size:22px;font-weight:800;color:#f1f5f9;margin-bottom:8px">¡Ya conoces Yve.01!</div>' +
        '<div style="font-size:14px;color:#94a3b8;line-height:1.7;margin-bottom:24px">' +
          'El sistema está listo para automatizar las finanzas de tu hotel.<br>' +
          'El primer paso: procesa las facturas OTA del mes en <b style="color:var(--acc2,#60a5fa)">AR — OTAs</b>.' +
        '</div>' +
        '<div style="display:flex;gap:10px;justify-content:center">' +
          '<button onclick="_closeCongrats()" ' +
            'style="background:transparent;border:1px solid #334155;color:#64748b;' +
            'padding:10px 20px;border-radius:10px;font-size:13px;cursor:pointer">Cerrar</button>' +
          '<button onclick="_startAR()" ' +
            'style="background:linear-gradient(135deg,var(--acc,#3b82f6),var(--acc-dark,#7c3aed));border:none;color:#fff;' +
            'padding:10px 22px;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;' +
            'box-shadow:0 4px 16px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)">' +
            'Empezar con AR →</button>' +
        '</div>' +
      '</div>';
    document.body.appendChild(card);
    card.addEventListener('click', function(e) {
      if (e.target === card) card.remove();
    });
    if (typeof _i18nAfterRender === 'function') _i18nAfterRender();
    // Auto-close after 12s
    setTimeout(function() { if (card.parentNode) card.remove(); }, 12000);
  }, 300);
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

// ── El hotel de destino ──────────────────────────────────────────────────
// El servidor ya no deja procesar sin hotel cuando hay 2 o mas (devuelve 409).
// Esto es para que no llegues al 409 de sorpresa despues de elegir 20 ficheros:
// se pregunta antes, y el boton no se enciende hasta que esta contestado.
// La garantia la sigue dando el servidor; esto es la experiencia.
var _uploadCenso = { hoteles: [], hotel: '' };
// Una vez enseñada la fila en esta apertura del modal, se queda. Si no,
// desapareceria justo al elegir hotel —porque elegir pone hotel activo— y el
// selector se esfumaria en la cara del que lo esta usando.
var _uploadFilaVista = false;

function _hayVariosHoteles(){ return _uploadCenso.hoteles.length >= 2; }
function _dentroDeUnHotel(){ return !!_uploadCenso.hotel; }

// La fila del hotel SOLO se enseña en la vista de grupo con 2+ hoteles, que es
// el unico caso en el que hay algo que decidir. Dentro de un hotel no se
// pregunta: el documento va a ese y punto, igual que no se pregunta cuando
// solo hay uno.
function _hotelObligatorio(){ return _hayVariosHoteles() && !_dentroDeUnHotel(); }

// Se puede procesar si: no hay ambiguedad (0 o 1 hotel), o ya hay uno elegido
// —venga del selector de la cabecera o del propio modal—.
function _hotelResuelto(){ return !_hayVariosHoteles() || _dentroDeUnHotel(); }

async function _cargarCensoSubida(){
  try {
    var d = await fetch('/api/hotel_activo').then(function(r){ return r.json(); });
    _uploadCenso = { hoteles: (d && d.hoteles) || [], hotel: (d && d.hotel) || '' };
  } catch(e) { _uploadCenso = { hoteles: [], hotel: '' }; }

  var row = document.getElementById('upload-hotel-row');
  var sel = document.getElementById('upload-hotel-sel');
  if (!row || !sel) return;
  if (_hotelObligatorio()) _uploadFilaVista = true;
  if (!_uploadFilaVista) { row.style.display = 'none'; return; }

  row.style.display = '';
  sel.innerHTML = '<option value="">' + t('upload.eligeHotel', '— Elige un hotel —') + '</option>' +
    _uploadCenso.hoteles.map(function(h){
      var id = String(h.id || ''), nom = String(h.nombre || '');
      return '<option value="' + id.replace(/"/g,'&quot;') + '"' +
             (_uploadCenso.hotel === id ? ' selected' : '') + '>🏨 ' +
             nom.replace(/</g,'&lt;') + '</option>';
    }).join('');
  _pintarAvisoHotel();
}

function _pintarAvisoHotel(){
  var av = document.getElementById('upload-hotel-aviso');
  if (!av) return;
  if (_hotelResuelto()) { av.style.display = 'none'; return; }
  av.style.display = '';
  av.textContent = t('upload.hotelPendiente',
    'Elige el hotel antes de procesar. Un documento sin hotel no es «del hotel principal», es un documento del que no sabemos el hotel.');
}

async function _elegirHotelSubida(id){
  // Reusa el mismo endpoint que el selector de la cabecera: un solo concepto
  // de "hotel activo", no dos que puedan desincronizarse.
  try { await _postJson('/api/hotel_activo', {hotel: id || ''}); } catch(e){}
  _uploadCenso.hotel = id || '';
  if (typeof _initHotelActivo === 'function') { try { _initHotelActivo(); } catch(e){} }
  _pintarAvisoHotel();
  _renderFileList();
}

// Si alguien llega igualmente al 409 (pestaña vieja, API a pelo), que lo vea.
function _avisar409(d){
  var msg = (d && d.error) || t('upload.hotelPendiente', 'Elige el hotel antes de procesar.');
  showNotification('🏨 ' + msg, 'error');
  _cargarCensoSubida();
  return false;
}

// ── M7: el fondo se queda quieto mientras hay un modal abierto ───────
// `overflow:hidden` a secas NO vale en iOS: el fondo sigue moviendose con el
// dedo. Lo que si funciona es fijar el body y compensar el desplazamiento,
// devolviendolo al cerrar para que la pagina no pegue un salto.
var _scrollFondo = 0;
function _bloquearFondo(si) {
  var b = document.body;
  if (!b) return;
  if (si) {
    if (b.dataset.fondoFijo) return;            // ya estaba, no re-guardar
    _scrollFondo = window.scrollY || window.pageYOffset || 0;
    b.dataset.fondoFijo = '1';
    b.style.position = 'fixed';
    b.style.top = (-_scrollFondo) + 'px';
    b.style.left = '0';
    b.style.right = '0';
    b.style.width = '100%';
  } else if (b.dataset.fondoFijo) {
    delete b.dataset.fondoFijo;
    b.style.position = ''; b.style.top = ''; b.style.left = '';
    b.style.right = ''; b.style.width = '';
    window.scrollTo(0, _scrollFondo);
  }
}

async function openUploadModal() {
  // Reset state
  _uploadFiles = [];
  document.getElementById('upload-file-list').style.display = 'none';
  document.getElementById('upload-files-container').innerHTML = '';
  document.getElementById('upload-count-new').textContent =
    t('upload.nuevosN', '{n} nuevos').replace('{n}', 0);
  document.getElementById('upload-count-dup').textContent =
    t('upload.yaProcN', '{n} ya procesados (se saltarán)').replace('{n}', 0);
  var procBtn = document.getElementById('btn-upload-procesar');
  procBtn.disabled = true; procBtn.style.opacity = '.4'; procBtn.style.cursor = 'not-allowed';

  var _avDesc = document.getElementById('upload-aviso-descartes');
  if (_avDesc) { _avDesc.style.display = 'none'; _avDesc.innerHTML = ''; }

  _uploadFilaVista = false;      // cada apertura decide de cero
  await _cargarCensoSubida();

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
      sList.innerHTML = '';
      serverFiles.forEach(function(f) {
        var row = document.createElement('div');
        row.id = 'file-row-' + f.nombre.replace(/[^a-zA-Z0-9]/g,'_');
        row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 10px;background:' + 
          (f.procesado ? 'rgba(34,197,94,.05)' : 'var(--bg)') + ';border-radius:7px;border:1px solid ' +
          (f.procesado ? 'rgba(34,197,94,.2)' : 'var(--s2)') + ';font-size:12px;margin-bottom:4px';
        
        var icon = document.createElement('span');
        icon.textContent = f.nombre.endsWith('.xlsm') ? '📊' : '📄';
        
        var nombre = document.createElement('span');
        nombre.style.cssText = 'flex:1;color:var(--tx)';
        nombre.textContent = f.nombre;
        
        var tamano = document.createElement('span');
        tamano.style.cssText = 'color:var(--dim);font-size:11px';
        tamano.textContent = f.tamano_str;
        
        var estado = document.createElement('span');
        estado.style.cssText = 'font-size:11px;padding:2px 7px;border-radius:5px;background:' +
          (f.procesado ? 'rgba(34,197,94,.1)' : 'rgba(245,158,11,.1)') + ';color:' +
          (f.procesado ? '#22c55e' : '#f59e0b');
        estado.textContent = f.procesado ? '✓ Procesado' : '⏳ Pendiente';
        
        var btnX = document.createElement('button');
        btnX.textContent = '✕';
        btnX.title = 'Eliminar archivo';
        btnX.style.cssText = 'background:transparent;border:1px solid rgba(239,68,68,.3);color:#e05252;width:22px;height:22px;border-radius:50%;cursor:pointer;font-size:11px;line-height:1;padding:0;flex-shrink:0;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s';
        btnX.onmouseover = function(){ this.style.background='rgba(239,68,68,.15)'; };
        btnX.onmouseout = function(){ this.style.background='transparent'; };
        btnX.onclick = function() { eliminarArchivoServidor(f.nombre, row); };
        
        row.appendChild(icon);
        row.appendChild(nombre);
        row.appendChild(tamano);
        row.appendChild(estado);
        row.appendChild(btnX);
        sList.appendChild(row);
      });
    } else if (sSection) {
      sSection.style.display = 'none';
    }
  } catch(e) {}

  var modal = document.getElementById('upload-modal');
  modal.style.display = 'flex';
  _bloquearFondo(true);          // M7
}

function procesarPendientesServidor() {
  // Process files already on server that haven't been processed yet
  fetch('/api/archivos_estado')
    .then(function(r){ return r.json(); })
    .then(function(d) {
      var pendientes = (d.files || []).filter(function(f){ return !f.procesado; }).map(function(f){ return f.nombre; });
      if (!pendientes.length) { showNotification('No hay archivos pendientes en el servidor', 'info'); return; }
      // Este boton es la otra forma de disparar el lote, y se saltaba el
      // candado del modal. Misma regla para los dos caminos.
      if (!_hotelResuelto()) { _pintarAvisoHotel(); return _avisar409(null); }
      closeUploadModal();
      showNotification('⏳ Procesando ' + pendientes.length + ' archivo(s) del servidor...', 'info');
      _runBatchPipeline(pendientes);
    })
    .catch(function(e){ showNotification('Error: ' + e.message, 'error'); });
}

function closeUploadModal() {
  document.getElementById('upload-modal').style.display = 'none';
  _bloquearFondo(false);         // M7
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

function _pareceDocumento(f) {
  return /\.(pdf|xlsm|xlsx|xls|csv|jpe?g|png|webp|heic)$/i.test(f.name || '') ||
         (f.type || '').indexOf('image/') === 0;
}

// ── DOS FOTOS CON EL MISMO NOMBRE SON DOS DOCUMENTOS ─────────────────
// El movil llama `image.jpg` a TODAS las fotos de camara. Antes la 2a y
// siguientes se descartaban por nombre repetido: por eso la camara solo
// dejaba añadir una. Ahora se renombran para que entren todas.
//
// La clave de choque NO es el nombre que trae, es el que va a tener DESPUES
// de comprimirse: `_comprimirImagen` renombra toda foto comprimida a
// `<base>.jpg`, asi que `recibo.png` y `recibo.jpeg` acabarian siendo el
// mismo `recibo.jpg` aunque entren con nombres distintos. Y ese nombre es la
// clave con la que el servidor deduplica las facturas: dos iguales y una
// desaparece del Excel sin avisar.
function _claveNombre(nombre) {
  var n = String(nombre || 'foto');
  if (/\.(jpe?g|png|webp|heic)$/i.test(n)) return n.replace(/\.\w+$/, '').toLowerCase() + '.jpg';
  return n.toLowerCase();
}

// "Es EXACTAMENTE el mismo fichero" (el mismo de la galeria elegido dos
// veces) frente a "dos fotos distintas que se llaman igual". Por el nombre no
// se distinguen; con el tamaño y la fecha, si. Se guarda el nombre original
// en el fichero renombrado para que la huella siga funcionando a la tercera.
function _huellaFichero(f) {
  return (f._nombreOriginal || f.name || '') + '|' + (f.size || 0) + '|' + (f.lastModified || 0);
}

function _nombreLibre(nombre, usadas) {
  if (!usadas.has(_claveNombre(nombre))) return nombre;
  var m = String(nombre || 'foto').match(/^(.*?)(\.[^.]+)?$/);
  var base = m[1] || 'foto', ext = m[2] || '';
  for (var i = 2; i < 1000; i++) {
    var cand = base + '_' + i + ext;
    if (!usadas.has(_claveNombre(cand))) return cand;
  }
  return base + '_' + Date.now() + ext;
}

function _esImagen(f) {
  return /\.(jpe?g|png|webp|heic)$/i.test(f.name || '') || (f.type || '').indexOf('image/') === 0;
}

function handleUploadFiles(fileList, input) {
  var todos = Array.from(fileList || []);
  var buenos = [], desconocidos = [];
  todos.forEach(function(f) { (_pareceDocumento(f) ? buenos : desconocidos).push(f); });

  var r = _addFilesToList(buenos);
  _avisoDescartes(todos.length, r.anadidos, desconocidos, r.repetidos, r.renombrados);

  // EL ARREGLO DE LA CAMARA: el input no se limpiaba NUNCA. Si el navegador
  // considera que el valor no ha cambiado, no vuelve a disparar 'change' y la
  // siguiente foto no llega a ejecutar nada. Se limpia DESPUES de haber
  // sacado los File del FileList: los objetos ya extraidos siguen siendo
  // validos, lo que se vacia es la seleccion del input.
  if (input) { try { input.value = ''; } catch(e) {} }
}

function _addFilesToList(newFiles) {
  var usadas = new Set();
  _uploadFiles.forEach(function(f) { usadas.add(_claveNombre(f.name)); });
  var huellas = new Set(_uploadFiles.map(_huellaFichero));

  var anadidos = 0, repetidos = [], renombrados = [];
  (newFiles || []).forEach(function(f) {
    // El MISMO fichero otra vez: eso si es un duplicado de verdad.
    if (huellas.has(_huellaFichero(f))) { repetidos.push(f.name); return; }

    // Las fotos ademas esquivan los nombres YA PROCESADOS. Si no, una foto de
    // camara nueva llamada `image.jpg` heredaria el "ya procesado" de la de
    // ayer y `uploadAndProcess` la filtraria sin decir nada. Los documentos
    // (pdf, excel) conservan ese aviso, que ahi si es util: no vienen de una
    // camara y repetir nombre suele significar repetir fichero.
    var evita = usadas;
    if (_esImagen(f) && _processedNames && _processedNames.size) {
      evita = new Set(usadas);
      _processedNames.forEach(function(n) { evita.add(_claveNombre(n)); });
    }

    var nombre = _nombreLibre(f.name || 'foto', evita);
    var entra = f;
    if (nombre !== f.name) {
      try {
        entra = new File([f], nombre, { type: f.type, lastModified: f.lastModified });
        entra._nombreOriginal = f.name;
      } catch (e) { entra = f; }     // navegador sin constructor de File: entra tal cual
      if (entra !== f) renombrados.push(f.name + ' → ' + nombre);
    }
    usadas.add(_claveNombre(entra.name));
    huellas.add(_huellaFichero(entra));
    _uploadFiles.push(entra);
    anadidos++;
  });
  _renderFileList();
  return { anadidos: anadidos, repetidos: repetidos, renombrados: renombrados };
}

// Lo que llega y no entra, dicho en voz alta. Antes desaparecia sin rastro:
// la pantalla se quedaba igual y no habia forma de saber por que.
function _avisoDescartes(recibidos, anadidos, desconocidos, repetidos, renombrados) {
  var caja = document.getElementById('upload-aviso-descartes');
  if (!caja) return;
  desconocidos = desconocidos || []; repetidos = repetidos || []; renombrados = renombrados || [];
  var fuera = desconocidos.length + repetidos.length;

  // Renombrar NO se avisa, A PROPOSITO. Con la camara TODAS las fotos se
  // llaman `image.jpg`, asi que el aviso saltaria en cada tanda y acabaria
  // siendo ruido que se ignora — que es como se pierde el credito de los
  // avisos que SI importan. La foto entra igual, con su nombre unico, y eso
  // ya se ve en la lista. `renombrados` se sigue recibiendo: las pruebas
  // comprueban con el que el renombrado ocurre aunque no se enseñe.
  caja.style.borderColor = 'rgba(245,158,11,.35)';
  caja.style.background = 'rgba(245,158,11,.08)';
  if (!fuera) { caja.style.display = 'none'; caja.innerHTML = ''; return; }

  var nombres = function(lista) {
    var v = lista.slice(0, 3).map(function(x) { return typeof x === 'string' ? x : (x.name || '(sin nombre)'); });
    return v.join(', ') + (lista.length > 3 ? ' y ' + (lista.length - 3) + ' mas' : '');
  };
  var partes = [];
  if (repetidos.length) {
    partes.push('<div style="margin-top:5px">· ' +
      t(repetidos.length === 1 ? 'upload.mismoFicheroUno' : 'upload.mismoFicheroN',
        repetidos.length === 1 ? '1 era el mismo fichero que ya estaba en la lista: '
                               : '{n} eran el mismo fichero que ya estaba en la lista: ').replace('{n}', repetidos.length) +
      '<b>' + nombres(repetidos) + '</b></div>');
  }
  if (desconocidos.length) {
    partes.push('<div style="margin-top:5px">· ' +
      t('upload.ilegiblesN', '{n} que el navegador no identifica ni como documento ni como foto: ')
        .replace('{n}', desconocidos.length) +
      '<b>' + nombres(desconocidos) + '</b></div>');
  }
  caja.innerHTML =
    '<div style="font-weight:700">⚠ ' +
      (recibidos === 1
        ? t('upload.llegoUno', 'Ha llegado 1 archivo y no se ha añadido.')
        : t(anadidos === 1 ? 'upload.llegaronUnoDentro' : 'upload.llegaronN',
            anadidos === 1 ? 'Han llegado {r} archivos y se ha añadido 1.'
                           : 'Han llegado {r} archivos y se han añadido {a}.')
            .replace('{r}', recibidos).replace('{a}', anadidos)) +
    '</div>' + partes.join('');
  caja.style.display = 'block';
}

function _detectType(fname) {
  var n = fname.toLowerCase();
  // DRR
  if (n.includes('drr') || n.includes('daily revenue') || n.includes('revenue report')) return 'DRR';
  // OTA
  if (n.includes('booking') || n.includes('expedia') || n.includes('hotelbeds') || n.includes('ota') || n.includes('comision') || n.includes('commission')) return 'AR — OTA';
  // Banco
  if (n.includes('extracto') || n.includes('bank') || n.includes('statement') || n.includes('movimientos')) return 'Banco';
  // F&B / Ventas
  if (n.includes('ventas') || n.includes('sales') || n.includes('pos') || n.includes('tpv') || n.includes('ticket') || n.includes('restaurante')) return 'F&B';
  // Inventario
  if (n.includes('inventario') || n.includes('inventory') || n.includes('stock') || n.includes('almacen')) return 'Inventario';
  // Mermas
  if (n.includes('merma') || n.includes('waste') || n.includes('pérdida')) return 'Mermas';
  // Rooming
  if (n.includes('rooming') || n.includes('room list') || n.includes('guest list')) return 'Rooming';
  // No-procesables conocidos
  if (n.includes('beo') || n.includes('banquet event')) return 'BEO';
  if (n.includes(' tm ') || n.includes('technical')) return 'TM';
  if (n.includes('contrato') || n.includes('contract') || n.includes('sow') || n.includes('scope of work')) return 'Contrato';
  if (n.includes('agenda') || n.includes('logo')) return 'Omitir';
  // Fotos → OCR con IA
  if (n.match(/\.(jpe?g|png|webp|heic)$/)) return 'Foto';
  // Factura por defecto para PDFs
  if (n.endsWith('.pdf')) return 'Factura';
  if (n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv')) return 'Datos';
  return 'Archivo';
}

function _typeColor(t) {
  if (t === 'DRR') return '#a78bfa';
  if (t.includes('OTA') || t.includes('AR')) return '#60a5fa';
  if (t === 'Factura' || t.includes('AP')) return '#f59e0b';
  if (t === 'Banco') return '#22c55e';
  if (t === 'Foto') return '#a855f7';
  if (t === 'F&B' || t === 'Inventario' || t === 'Mermas') return '#f97316';
  if (t === 'Rooming') return '#06b6d4';
  if (t === 'BEO' || t === 'TM' || t === 'Contrato') return '#a78bfa';
  if (t === 'Omitir') return '#64748b';
  return 'var(--mut)';
}

// ── QUE FOTOS SON PAGINAS DEL MISMO DOCUMENTO ────────────────────────
// Por defecto NINGUNA: cada foto es su propio documento (grupo de 1). Solo
// se agrupa lo que el usuario une a mano. Antes esto se decidia contando
// fotos, asi que subir dos facturas distintas las mandaba juntas al lector
// de contratos de grupo. Aqui ya no adivina nadie.
//
// El numero de grupo viaja PEGADO al objeto File (`_grp`), no en un array
// paralelo: quitar una foto de la lista mueve los indices y un array
// paralelo se desalinearia en silencio.
var _grpSeq = 0;

function _esFotoSubida(f) {
  var n = f.name || '';
  // La EXTENSION manda, y manda para decir que NO. Algunos selectores de
  // Android devuelven `type` vacio o hasta `image/...` para un PDF escaneado:
  // fiarse del MIME hacia que un documento se tratase como foto, saliera
  // marcado "Foto" y se ofreciera unirlo con otro.
  if (/\.(pdf|xlsx?|xlsm|csv|docx?|txt)$/i.test(n)) return false;
  if (/\.(jpe?g|png|webp|heic|heif)$/i.test(n)) return true;
  return (f.type || '').indexOf('image/') === 0;
}

// Fotos que todavia se pueden marcar: las ya procesadas no se tocan.
function _fotosAgrupables() {
  return _uploadFiles.filter(function(f) {
    return _esFotoSubida(f) && !_processedNames.has(f.name);
  });
}

function _fotosSeleccionadas() {
  return _fotosAgrupables().filter(function(f) { return f._sel; });
}

function _toggleSelFoto(idx) {
  var f = _uploadFiles[idx];
  if (!f || !_esFotoSubida(f) || _processedNames.has(f.name)) return;
  f._sel = !f._sel;
  _renderFileList();
}

function _unirSeleccionadas() {
  var sel = _fotosSeleccionadas();
  if (sel.length < 2) return;
  _grpSeq++;
  var g = _grpSeq;
  sel.forEach(function(f) { f._grp = g; f._sel = false; });
  _renderFileList();
}

function _deshacerGrupo(g) {
  _uploadFiles.forEach(function(f) { if (f._grp === g) { f._grp = 0; f._sel = false; } });
  _renderFileList();
}

function _limpiarSeleccion() {
  _uploadFiles.forEach(function(f) { f._sel = false; });
  _renderFileList();
}

// Reparte las fotos de la tanda en documentos.
// Devuelve {grupos:[[f,f],...], sueltas:[f,...]}, las sueltas en el orden de
// la lista. Un grupo que se quedo con una sola foto NO es un documento de
// varias paginas: vuelve a ser una foto suelta.
function _repartirFotos(imgs) {
  var porGrupo = {}, orden = [];
  imgs.forEach(function(f) {
    if (!f._grp) return;
    if (!porGrupo[f._grp]) { porGrupo[f._grp] = []; orden.push(f._grp); }
    porGrupo[f._grp].push(f);
  });
  var grupos = [], validos = {};
  orden.forEach(function(g) {
    if (porGrupo[g].length >= 2) { grupos.push(porGrupo[g]); validos[g] = true; }
  });
  var sueltas = imgs.filter(function(f) { return !(f._grp && validos[f._grp]); });
  return { grupos: grupos, sueltas: sueltas };
}

// La barra de agrupar. Solo aparece cuando hay algo que unir de verdad.
function _pintarBarraUnir() {
  var bar = document.getElementById('upload-unir-bar');
  if (!bar) return;
  var agrup = _fotosAgrupables();
  if (agrup.length < 2) { bar.style.display = 'none'; bar.innerHTML = ''; return; }
  bar.style.display = 'block';
  var sel = _fotosSeleccionadas().length;
  if (sel >= 2) {
    bar.innerHTML =
      '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">' +
        '<button onclick="_unirSeleccionadas()" style="flex:1;min-width:200px;background:var(--acc);border:none;color:#fff;' +
          'padding:13px 18px;border-radius:11px;font-size:14px;font-weight:700;cursor:pointer">' +
          t('upload.unirN', '🔗 Unir las {n} en un documento').replace('{n}', sel) + '</button>' +
        '<button onclick="_limpiarSeleccion()" style="background:var(--s2);border:1px solid var(--s3);color:var(--tx);' +
          'padding:13px 16px;border-radius:11px;font-size:13px;font-weight:600;cursor:pointer">' +
          t('upload.quitarMarcas', 'Quitar marcas') + '</button>' +
      '</div>';
  } else {
    bar.innerHTML =
      '<div style="font-size:12px;color:var(--dim);line-height:1.5;padding:2px 2px">' +
        (sel === 1
          ? t('upload.pistaOtra', '☝ Marca al menos otra foto para unirlas como un solo documento.')
          : t('upload.pistaUnir', '¿Varias fotos de un mismo documento? Marcalas y pulsa <b style="color:var(--tx)">Unir</b>. Si no marcas nada, cada foto es un documento aparte.')) +
      '</div>';
  }
}

function _renderFileList() {
  var cont = document.getElementById('upload-files-container');
  var list = document.getElementById('upload-file-list');
  if (!_uploadFiles.length) { list.style.display = 'none'; return; }
  list.style.display = 'block';

  var newCount = 0, dupCount = 0;
  var _tam = function(n) { return n < 1024*1024 ? Math.round(n/1024) + 'KB' : (n/1024/1024).toFixed(1) + 'MB'; };
  var _pintados = {};   // grupos ya pintados: el bloque va donde su PRIMERA foto
  var _piezas = [];

  _uploadFiles.forEach(function(f, i) {
    var isProc = _processedNames.has(f.name);
    if (isProc) dupCount++; else newCount++;

    // ¿Esta foto va dentro de un documento de varias paginas?
    var g = (!isProc && f._grp) ? f._grp : 0;
    if (g) {
      var hermanas = _uploadFiles.filter(function(x) { return x._grp === g && !_processedNames.has(x.name); });
      if (hermanas.length >= 2) {
        if (_pintados[g]) return;         // ya se pinto con la primera
        _pintados[g] = true;
        var nombres = hermanas.map(function(x) { return x.name; }).join(' · ');
        var bytes = hermanas.reduce(function(s, x) { return s + x.size; }, 0);
        _piezas.push(
          '<div style="padding:10px 12px;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.07);' +
            'border-radius:10px;border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.35)">' +
            '<div style="display:flex;align-items:center;gap:10px">' +
              '<div style="font-size:20px">📄</div>' +
              '<div style="flex:1;min-width:0">' +
                '<div style="font-size:13px;font-weight:700;color:var(--acc2)">' +
                  t('upload.docPaginas', 'Documento de {n} páginas').replace('{n}', hermanas.length) + '</div>' +
                '<div style="font-size:11px;color:var(--dim);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + nombres + '</div>' +
                '<div style="font-size:11px;color:var(--dim)">' + _tam(bytes) + '</div>' +
              '</div>' +
              '<button onclick="_deshacerGrupo(' + g + ')" title="' + t('upload.separarGrupo', 'Separar en fotos sueltas') + '" ' +
                'style="background:none;border:1px solid var(--s3);color:var(--dim);cursor:pointer;font-size:16px;' +
                'min-width:44px;min-height:44px;display:flex;align-items:center;justify-content:center;' +
                'border-radius:10px;line-height:1;flex:0 0 auto">✕</button>' +
            '</div>' +
          '</div>');
        return;
      }
      // Se quedo sola: deja de ser un documento de varias paginas.
      f._grp = 0;
    }

    // Fila normal. Las fotos nuevas llevan casilla grande y toda la fila es
    // zona de toque: en el movil una casilla de 14px no se acierta con el dedo.
    var puedeMarcar = !isProc && _esFotoSubida(f) && _fotosAgrupables().length >= 2;
    var sel = !!(puedeMarcar && f._sel);
    var casilla = puedeMarcar
      ? '<div style="width:26px;height:26px;flex:0 0 26px;border-radius:8px;display:flex;align-items:center;' +
          'justify-content:center;font-size:15px;font-weight:800;color:#fff;border:2px solid ' +
          (sel ? 'var(--acc)' : 'var(--s3)') + ';background:' + (sel ? 'var(--acc)' : 'transparent') + '">' +
          (sel ? '✓' : '') + '</div>'
      : '';
    _piezas.push(
      '<div' + (puedeMarcar ? ' onclick="_toggleSelFoto(' + i + ')"' : '') +
        ' style="display:flex;align-items:center;gap:10px;padding:10px 12px;background:' +
        (isProc ? 'rgba(245,158,11,.06)' : (sel ? 'rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10)' : 'var(--bg)')) +
        ';border-radius:8px;border:1px solid ' +
        (isProc ? 'rgba(245,158,11,.2)' : (sel ? 'var(--acc)' : 'var(--s2)')) +
        ';opacity:' + (isProc ? '.6' : '1') + (puedeMarcar ? ';cursor:pointer' : '') + '">' +
        casilla +
        '<div style="font-size:18px">' + (f.name.endsWith('.xlsm') ? '📊' : '📄') + '</div>' +
        '<div style="flex:1;min-width:0">' +
          '<div style="font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + f.name + '</div>' +
          '<div style="font-size:11px;color:var(--dim)">' + _tam(f.size) + ' · <span style="color:' + _typeColor(_detectType(f.name)) + '">' + _detectType(f.name) + '</span>' +
            (isProc ? ' · <span style="color:var(--ora)">⚠ Ya procesado</span>' : '') +
          '</div>' +
        '</div>' +
        '<button onclick="event.stopPropagation();_removeUploadFile(' + i + ')" title="' + t('upload.quitar', 'Quitar') + '" ' +
          'style="background:none;border:none;color:var(--dim);cursor:pointer;font-size:20px;line-height:1;' +
          'min-width:44px;min-height:44px;display:flex;align-items:center;justify-content:center;flex:0 0 auto">×</button>' +
      '</div>');
  });
  cont.innerHTML = _piezas.join('');
  _pintarBarraUnir();

  document.getElementById('upload-count-new').textContent =
    t(newCount === 1 ? 'upload.nuevoUno' : 'upload.nuevosN',
      newCount === 1 ? '1 nuevo' : '{n} nuevos').replace('{n}', newCount);
  document.getElementById('upload-count-dup').textContent =
    t(dupCount === 1 ? 'upload.yaProcUno' : 'upload.yaProcN',
      dupCount === 1 ? '1 ya procesado (se saltará)'
                     : '{n} ya procesados (se saltarán)').replace('{n}', dupCount);
  
  var procBtn = document.getElementById('btn-upload-procesar');
  // Hacen falta las dos cosas: ficheros nuevos Y hotel resuelto.
  var puede = newCount > 0 && _hotelResuelto();
  procBtn.textContent = !_hotelResuelto()
    ? t('upload.eligeHotelBoton', '🏨 Elige un hotel para continuar')
    : (newCount > 0
        ? t(newCount === 1 ? 'upload.procesarUno' : 'upload.procesarN',
            newCount === 1 ? '⚡ Procesar 1 archivo nuevo'
                           : '⚡ Procesar {n} archivos nuevos').replace('{n}', newCount)
        : t('upload.procesarSel', '⚡ Procesar seleccionados'));
  procBtn.disabled = !puede;
  procBtn.style.opacity = puede ? '1' : '.4';
  procBtn.style.cursor = puede ? 'pointer' : 'not-allowed';
  _pintarAvisoHotel();
}

function _removeUploadFile(idx) {
  _uploadFiles.splice(idx, 1);
  _renderFileList();
}

async function uploadAndProcess() {
  var newFiles = _uploadFiles.filter(function(f) { return !_processedNames.has(f.name); });
  if (!newFiles.length) { showNotification(_tSSE('No hay archivos nuevos que procesar'), 'info'); return; }
  // Ultima red antes de subir nada: si el censo cambio en otra pestaña
  // mientras este modal estaba abierto, el boton podria estar encendido.
  await _cargarCensoSubida();
  if (!_hotelResuelto()) { _renderFileList(); return _avisar409(null); }

  var isImg = function(f) { return /\.(jpe?g|png|webp|heic)$/i.test(f.name) || (f.type || '').indexOf('image/') === 0; };
  var imgs = newFiles.filter(isImg);
  var docs = newFiles.filter(function(f){ return !isImg(f); });

  var btn = document.getElementById('btn-upload-procesar');
  btn.disabled = true; btn.style.opacity = '.4';

  if (docs.length) {
    btn.textContent = '⏳ Subiendo archivos...';
    var formData = new FormData();
    docs.forEach(function(f) { formData.append('files', f, f.name); });
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
  }

  closeUploadModal();

  if (imgs.length) {
    var overlay = document.getElementById('overlay');
    var log = document.getElementById('log');
    var icon = document.getElementById('modal-icon');
    var title = document.getElementById('modal-title');
    var spin = document.getElementById('spin');
    var btnCl = document.getElementById('btn-cl');
    var runBtn = document.getElementById('btn-run');
    if (overlay) overlay.classList.add('on');
    if (log) log.innerHTML = '';
    if (spin) spin.style.display = 'block';
    if (icon) icon.textContent = '⚡';
    if (title) title.textContent = _tSSE('Procesando ' + newFiles.length + ' archivo(s)...');
    if (btnCl) btnCl.disabled = true;
    if (runBtn) runBtn.disabled = true;
    var addLine = function(txt, cls) {
      if (!log) return;
      var p = document.createElement('p');
      p.className = cls || 'l-dim';
      p.textContent = _tSSE(txt);
      log.appendChild(p);
      log.scrollTop = log.scrollHeight;
    };
    var errs = 0;
    var _cierreFotos = {};
    // Cada foto es su propio documento salvo que el usuario haya unido varias
    // a mano en el modal. Antes se decidia CONTANDO fotos (imgs.length >= 2):
    // dos facturas distintas se mandaban juntas al lector de contratos, que
    // cuesta una llamada entera a la IA antes de acabar acertando por el
    // camino de atras. Los dos endpoints y sus mensajes NO cambian.
    var _rep = _repartirFotos(imgs);
    var _soloUno = (_rep.grupos.length === 1 && !_rep.sueltas.length);
    for (var _gi = 0; _gi < _rep.grupos.length; _gi++) {
      // Con un unico documento el mensaje se queda EXACTAMENTE como estaba:
      // es el caso que hoy acierta y no se puede mover.
      var _pref = _soloUno ? '' : ('Documento ' + (_gi + 1) + ' de ' + _rep.grupos.length + ': ');
      errs += await _procesarGrupoFotos(_rep.grupos[_gi], _pref, addLine, _cierreFotos);
    }
    if (_rep.sueltas.length) {
      errs += await _procesarImagenes(_rep.sueltas, addLine, _cierreFotos);
    }
    if (docs.length) {
      // tanda mezclada: las fotos le pasan su pendiente al lote y se cierra
      // UNA vez al final, no dos.
      _runBatchPipeline(docs.map(function(f){ return f.name; }), true, _cierreFotos);
    } else {
      // Solo fotos. Aqui estaba el agujero: se acababa el proceso sin cruce y
      // sin asignador, asi que una foto de una factura se guardaba y no
      // llegaba NUNCA a Aprobaciones AP — nadie le ponia cuenta ni asiento.
      if (title) title.textContent = _tSSE('Cerrando el pipeline...');
      var _avCierre = await _correrCierre(_cierreFotos, addLine);
      if (_avCierre) errs++;
      if (icon) icon.textContent = errs ? '⚠️' : '✅';
      if (title) title.textContent = _tSSE(errs ? 'Procesado finalizado con avisos' : 'Procesado completado');
      if (spin) spin.style.display = 'none';
      if (btnCl) btnCl.disabled = false;
      if (runBtn) runBtn.disabled = false;
      if (log) _showTabBadges(log.textContent || '');
      // M6: 700 ms de espera fija DESPUES de que el cierre ya termino, o
      // sea con los datos ya escritos. No hay nada que esperar.
      setTimeout(function(){ loadAll(); if (typeof cargarARRealData === 'function') { try { cargarARRealData(); } catch(e){} } }, 60);
    }
    return;
  }

  _runBatchPipeline(docs.map(function(f){ return f.name; }));
}

// ── El paso de cierre, en su propia conexion ─────────────────────────
// El cruce factura<->albaran y el asignador de cuentas ya NO van al final de
// cada lote. Iban dentro del EventSource del lote, que este mismo fichero
// cierra a los 60 s, asi que en el ultimo lote no llegaban a arrancar: los dos
// son subprocesos con 180 s de margen cada uno. Ahora se llaman UNA vez, aqui,
// cuando ya esta todo guardado — venga de lotes, de fotos o de las dos cosas.
function _correrCierre(pasos, addLine) {
  return new Promise(function(resolve) {
    var lista = Object.keys(pasos || {}).filter(function(k){ return pasos[k]; });
    if (!lista.length) { resolve(false); return; }
    var es = new EventSource('/api/cerrar_pipeline_stream?pasos=' + encodeURIComponent(lista.join(',')));
    var aviso = false, hecho = false;
    function _acabar(malo) {
      if (hecho) return;
      hecho = true;
      clearTimeout(reloj);
      try { es.close(); } catch(e) {}
      resolve(malo);
    }
    // 7 min. El cierre son dos subprocesos con 180 s de margen cada uno, asi
    // que este reloj no tiene que cortar nada en un caso normal: esta para que
    // la pantalla no se quede colgada para siempre si algo se atasca.
    var reloj = setTimeout(function() {
      addLine('\u26a0 El cierre del pipeline tarda demasiado — vuelve a pulsar Procesar para reintentarlo', 'l-warn');
      _acabar(true);
    }, 420000);
    es.onmessage = function(ev) {
      var txt = ev.data;
      if (txt === 'CIERRE_COMPLETO') { _acabar(aviso); return; }
      if (txt === 'CIERRE_CON_ERRORES') { _acabar(true); return; }
      if (!txt) return;
      var c0 = txt.charAt(0), cls = 'l-dim';
      if (c0 === '\u2713' || c0 === '\u2705') cls = 'l-ok';
      else if (c0 === '\u2717') { cls = 'l-err'; aviso = true; }
      else if (c0 === '\u26a0') { cls = 'l-warn'; aviso = true; }
      else if (c0 === '\u2139') cls = 'l-info';
      else if (txt.indexOf('>>') === 0) cls = 'l-info';
      addLine(txt, cls);
    };
    es.onerror = function() {
      if (hecho) return;   // el servidor ha cerrado bien y EventSource reintenta
      addLine('\u26a0 Se ha cortado la conexion durante el cierre del pipeline', 'l-warn');
      _acabar(true);
    };
  });
}

function _runBatchPipeline(fileNames, keepLog, cierreInicial) {
  // Guardia para CUALQUIER camino que llegue hasta aqui. Importa porque el
  // lote va por EventSource, y un EventSource NO sabe leer un 409: la petición
  // falla y salta `onerror`, o sea que el usuario veria "⚡ Reconectando..."
  // en vez del motivo. El candado del servidor sigue siendo el que manda —
  // esto solo evita que el mensaje se pierda por el camino.
  if (!_hotelResuelto()) { _avisar409(null); return; }
  var overlay = document.getElementById('overlay');
  var log = document.getElementById('log');
  var btn = document.getElementById('btn-run');
  var spin = document.getElementById('spin');
  var lbl = document.getElementById('run-lbl');
  var btnCl = document.getElementById('btn-cl');
  var icon = document.getElementById('modal-icon');
  var title = document.getElementById('modal-title');

  if (overlay) overlay.classList.add('on');
  if (log && !keepLog) log.innerHTML = '';
  if (btn) btn.disabled = true;
  if (spin) spin.style.display = 'block';
  if (lbl) lbl.textContent = 'Procesando...';
  if (btnCl) btnCl.disabled = true;
  if (icon) icon.textContent = '⚡';
  if (title) title.textContent = _tSSE('Procesando ' + fileNames.length + ' archivo(s)...');

  var total = fileNames.length;
  var hadError = false;
  // Lo que hay que cerrar cuando acabe TODO. Puede venir sembrado desde las
  // fotos: una tanda mezclada de fotos y PDF cierra una sola vez, al final.
  var _cierrePend = cierreInicial || {};

  function _log(txt, cls) {
    if (!log) return;
    var p = document.createElement('p');
    p.className = cls || 'l-dim';
    p.textContent = _tSSE(txt);
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;
  }

  function _finishReal(ok) {
    if (icon) icon.textContent = ok ? '✅' : '⚠️';
    if (title) title.textContent = _tSSE(ok ? 'Procesado completado' : 'Procesado finalizado con avisos');
    // Mostrar badges verdes en los tabs que se actualizaron
    if (log) _showTabBadges(log.textContent || '');
    if (btn) btn.disabled = false;
    if (spin) spin.style.display = 'none';
    if (lbl) lbl.textContent = _tSSE('⚡ Procesar Archivos');
    if (btnCl) { btnCl.disabled = false; btnCl.textContent = 'Cerrar'; }
    var retryBtn = document.getElementById('btn-retry');
    if (retryBtn) retryBtn.style.display = 'none';
    // M6: igual que arriba — 800 ms de espera por si acaso, cuando el lote
    // ya ha terminado de escribir. Lo que queda de lentitud es el servidor
    // releyendo el Excel consolidado, que es otro trabajo.
    setTimeout(function(){ loadAll(); if (typeof cargarARRealData === 'function') { try { cargarARRealData(); } catch(e){} } }, 60);
  }

  // El cierre corre SIEMPRE antes de dar el proceso por acabado, y por
  // cualquier camino: lotes agotados, error de conexion o timeout. Antes iba
  // dentro del stream del ultimo lote, asi que un timeout justo ahi se llevaba
  // por delante el cruce y el asignador de TODA la tanda.
  function _finish(ok) {
    if (title) title.textContent = _tSSE('Cerrando el pipeline...');
    _correrCierre(_cierrePend, _log).then(function(aviso) {
      _finishReal(ok && !aviso);
    });
  }

  // Dividir en lotes de 4 para evitar timeout de Render (30s por conexión SSE)
  var BATCH_SIZE = 4;
  var batches = [];
  for (var bi = 0; bi < fileNames.length; bi += BATCH_SIZE) {
    batches.push(fileNames.slice(bi, bi + BATCH_SIZE));
  }
  var batchIdx = 0;

  function processBatch() {
    if (batchIdx >= batches.length) {
      _finish(!hadError);
      return;
    }
    if (batchIdx > 0) _log('⚡ Lote ' + (batchIdx+1) + '/' + batches.length + '...', 'l-dim');
    var batch = batches[batchIdx];
    var batchFiles = encodeURIComponent(JSON.stringify(batch));
    var evtSrc = new EventSource('/api/procesar_batch_stream?archivos=' + batchFiles);
    var timer = setTimeout(function() {
      evtSrc.close();
      _log('⚠ Timeout en lote ' + (batchIdx+1) + ' — continuando', 'l-warn');
      batchIdx++;
      processBatch();
    }, 60000);

    evtSrc.onmessage = function(ev) {
      var txt = ev.data;
      // Lo que este lote deja pendiente de cerrar. No se pinta: es para la
      // maquina, no para el usuario.
      if (txt && txt.indexOf('CIERRE_PENDIENTE:') === 0) {
        txt.slice(17).split(',').forEach(function(p){ if (p) _cierrePend[p] = true; });
        return;
      }
      if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
        try { if (typeof _invalidarPaneles === 'function') _invalidarPaneles(); } catch(e){}
        clearTimeout(timer);
        evtSrc.close();
        if (txt === 'PIPELINE_CON_ERRORES') hadError = true;
        batchIdx++;
        setTimeout(processBatch, 500);
      } else if (txt && txt !== '') {
        var cls = 'l-dim';
        if (txt.startsWith('✓') || txt.startsWith('✅')) cls = 'l-ok';
        else if (txt.startsWith('✗')) { cls = 'l-err'; hadError = true; }
        else if (txt.startsWith('⚠')) cls = 'l-warn';
        else if (txt.startsWith('ℹ')) cls = 'l-info';
        else if (txt.startsWith('>>')) cls = 'l-info';
        else if (txt.startsWith('📍')) cls = 'l-info';
        _log(txt, cls);
      }
    };
    evtSrc.onerror = function() {
      clearTimeout(timer);
      evtSrc.close();
      batchIdx++;
      if (batchIdx < batches.length) {
        _log('⚡ Reconectando...', 'l-dim');
        setTimeout(processBatch, 1000);
      } else {
        _finish(!hadError);
      }
    };
  }
  processBatch();
}



// ── Historial de procesado ──────────────────────────────────────────
async function mostrarHistorialProcesado() {
  try {
    var r = await fetch('/api/historial_procesado');
    var d = await r.json();
    if (!d.ok || !d.items.length) {
      showNotification(t('hist.vacio', 'No hay archivos procesados todavía.'), 'info');
      return;
    }
    // Crear modal con historial
    var existing = document.getElementById('historial-modal');
    if (existing) existing.remove();
    
    var rows = d.items.map(function(item) {
      var color = item.icono === '✓' ? '#4ade80' : (item.icono === '⚠' ? '#facc15' : (item.icono === 'ℹ' ? '#60a5fa' : '#f87171'));
      // M9: en el movil cada fila se convierte en una tarjeta (ver .hist-t en
      // el CSS). Los data-r sirven de etiqueta ahi, donde no hay cabecera.
      return '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">' +
        '<td style="padding:8px 12px;color:' + color + ';font-size:14px">' + item.icono + '</td>' +
        '<td data-r="Archivo" style="padding:8px 10px;font-size:12px;max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + item.archivo + '">' + item.archivo + '</td>' +
        '<td data-r="Dónde" style="padding:8px 10px;font-size:12px;color:var(--acc2)">' + item.tab + '</td>' +
        '<td data-r="Fecha" style="padding:8px 10px;font-size:11px;color:#64748b">' + item.fecha + '</td>' +
        '</tr>';
    }).join('');
    
    // Contar tabs actualizados
    var tabCounts = {};
    d.items.forEach(function(item) {
      if (item.tab !== 'Omitido' && item.tab !== 'Error' && item.tab !== '—') {
        tabCounts[item.tab] = (tabCounts[item.tab] || 0) + 1;
      }
    });
    var tabSummary = Object.entries(tabCounts).map(function(e) {
      return '<span style="display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.12);color:var(--acc2);margin:2px 3px">' + e[1] + ' ' + e[0] + '</span>';
    }).join('');
    
    var modal = document.createElement('div');
    modal.id = 'historial-modal';
    modal.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.7);backdrop-filter:blur(4px);display:flex;align-items:center;justify-content:center';
    modal.innerHTML = '<div style="background:#0f1729;border:1px solid rgba(255,255,255,.08);border-radius:16px;padding:24px;max-width:700px;width:95%;max-height:80vh;overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,.7);position:relative">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">' +
        '<h3 style="margin:0;font-size:16px;font-weight:700">📋 Historial de Procesado</h3>' +
        '<button onclick="this.closest(\'[id=historial-modal]\').remove()" style="background:none;border:none;color:var(--mut);font-size:20px;cursor:pointer">✕</button>' +
      '</div>' +
      '<div style="margin-bottom:14px">' + (tabSummary || '<span style="color:#64748b;font-size:12px">Sin datos procesados</span>') + '</div>' +
      '<table class="hist-t" style="width:100%;border-collapse:collapse">' +
        '<thead><tr style="border-bottom:1px solid rgba(255,255,255,.1)">' +
          '<th style="padding:6px 12px;text-align:left;font-size:11px;color:#64748b"></th>' +
          '<th style="padding:6px 10px;text-align:left;font-size:11px;color:#64748b">ARCHIVO</th>' +
          '<th style="padding:6px 10px;text-align:left;font-size:11px;color:#64748b">SECCIÓN</th>' +
          '<th style="padding:6px 10px;text-align:left;font-size:11px;color:#64748b">FECHA</th>' +
        '</tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +
    '</div>';
    modal.addEventListener('click', function(e) { if (e.target === modal) modal.remove(); });
    document.addEventListener('keydown', function _escH(e) { if (e.key === 'Escape') { var m = document.getElementById('historial-modal'); if(m){m.remove();document.removeEventListener('keydown',_escH);} }});
    document.body.appendChild(modal);
  } catch(e) {
    showNotification('✗ Historial: ' + e.message, 'error');
  }
}

// Tabs actualizados — puntos verdes + highlight de stats
// Persisten hasta que el usuario VISITA el tab y luego SALE de él
var _updatedTabs = {};

function _showTabBadges(logText) {
  var tabMap = {
    '✓ AP ': 'ap',
    '✓ AR ': 'ar',
    'OTA': 'ar',
    '✓ Banco': 'banco',
    '✓ F&B': 'fb',
    '✓ Inventario': 'fb',
    '✓ Mermas': 'fb',
    '✓ Rooming': 'fb',
    '✓ DRR': 'drr',
    '✓ Contrato': 'ar_real',
    'AR Real': 'ar_real',
  };
  
  for (var key in tabMap) {
    if (logText.includes(key)) {
      _updatedTabs[tabMap[key]] = true;
    }
  }
  _renderBadges();
  _highlightActiveStats();
  _marcarStatsActualizadas();
}

// Solo si el CONTORNO (acentuar-todo) está activo: pone en VERDE únicamente
// las stats de los apartados que se acaban de actualizar. El resto se queda
// con el color de contorno personalizado.
var _PANEL_DE_TAB = { ap:'panel-ap', ar:'panel-ar', banco:'panel-banco', fb:'panel-fb', drr:'panel-drr', ar_real:'panel-ar_real' };
function _statCardsDe(panelId) {
  var panel = document.getElementById(panelId);
  if (!panel) return [];
  return panel.querySelectorAll('.sc, .fb-kpi-card, .card');
}
function _marcarStatsActualizadas() {
  // 1) siempre limpiar el verde anterior
  document.querySelectorAll('.upd-green').forEach(function(c){ c.classList.remove('upd-green'); });
  // 2) solo aplicar si el usuario tiene el contorno activo en Personalizar
  if (!document.body.classList.contains('acentuar-todo')) return;
  for (var k in _updatedTabs) {
    if (!_updatedTabs[k]) continue;
    var pid = _PANEL_DE_TAB[k]; if (!pid) continue;
    _statCardsDe(pid).forEach(function(c){ c.classList.add('upd-green'); });
  }
}
function _limpiarStatsPanel(tabKey) {
  var pid = _PANEL_DE_TAB[tabKey]; if (!pid) return;
  _statCardsDe(pid).forEach(function(c){ c.classList.remove('upd-green'); });
}

function _renderBadges() {
  // Limpiar badges anteriores
  document.querySelectorAll('.proc-badge').forEach(function(b) { b.remove(); });
  
  // Tab name mapping para buscar elementos
  var nameMap = {
    'ap': ['proveedores', 'ap'],
    'ar': ['ota', 'ar'],
    'banco': ['banco'],
    'fb': ['f&b', 'fb'],
    'drr': ['drr'],
    'ar_real': ['ar real', 'real'],
  };
  
  var tabButtons = document.querySelectorAll('[id^="tab-"]');
  
  for (var tabKey in _updatedTabs) {
    if (!_updatedTabs[tabKey]) continue;
    
    var tabEl = document.getElementById('tab-' + tabKey);
    if (!tabEl && tabKey === 'ar_real') tabEl = document.getElementById('tab-ar-real');
    if (!tabEl && nameMap[tabKey]) {
      tabButtons.forEach(function(btn) {
        var txt = btn.textContent.toLowerCase();
        nameMap[tabKey].forEach(function(name) {
          if (txt.includes(name)) tabEl = btn;
        });
      });
    }
    if (tabEl) {
      var dot = document.createElement('span');
      dot.className = 'proc-badge';
      dot.dataset.tab = tabKey;
      dot.style.cssText = 'width:8px;height:8px;border-radius:50%;background:#22c55e;display:inline-block;margin-left:6px;vertical-align:middle;box-shadow:0 0 8px rgba(34,197,94,.6);animation:pulse-green 1.5s ease-in-out infinite';
      dot.title = 'Datos actualizados — visita el tab para marcar como visto';
      tabEl.appendChild(dot);
    }
  }
}

function _highlightActiveStats() {
  // Encontrar el tab activo y ver si tiene actualizaciones
  var activeTab = document.querySelector('.tab-active, [class*="tab"][class*="active"]');
  if (!activeTab) return;
  var tabText = activeTab.textContent.toLowerCase();
  var isUpdated = false;
  if (tabText.includes('proveedores') && _updatedTabs['ap']) isUpdated = true;
  if (tabText.includes('ota') && _updatedTabs['ar']) isUpdated = true;
  if (tabText.includes('banco') && _updatedTabs['banco']) isUpdated = true;
  if (tabText.includes('f&b') && _updatedTabs['fb']) isUpdated = true;
  if (tabText.includes('drr') && _updatedTabs['drr']) isUpdated = true;
  
  // Quitar highlight de todas las cards primero
  document.querySelectorAll('[data-highlighted="true"]').forEach(function(card) {
    card.style.boxShadow = '';
    card.style.borderColor = '';
    card.removeAttribute('data-highlighted');
  });
  
  if (isUpdated) {
    // Buscar stat cards en el contenido visible del tab
    var section = document.querySelector('[style*="display: block"], [style*="display:block"]');
    if (!section) section = document;
    var cards = section.querySelectorAll('[class*="stat"], [class*="kpi"], [class*="card-stat"]');
    if (cards.length === 0) {
      // Fallback: buscar divs que parezcan stat cards por estructura
      cards = document.querySelectorAll('.tab-content:not([style*="none"]) > div > div > div');
    }
    cards.forEach(function(card) {
      if (card.offsetHeight > 30 && card.offsetHeight < 200) {
        card.style.transition = 'box-shadow 0.5s, border-color 0.5s';
        card.style.boxShadow = '0 0 12px rgba(34,197,94,.25)';
        card.style.borderColor = 'rgba(34,197,94,.4)';
        card.setAttribute('data-highlighted', 'true');
      }
    });
  }
}

// Al cambiar de tab, marcar el tab anterior como "visto" y quitar su badge
var _prevVisitedTab = null;
var _origSwitchTab = typeof switchTab === 'function' ? switchTab : null;

function _onTabSwitch(newTab) {
  // Si el tab anterior tenía actualización, marcarla como vista
  if (_prevVisitedTab && _updatedTabs[_prevVisitedTab]) {
    _updatedTabs[_prevVisitedTab] = false;
    _limpiarStatsPanel(_prevVisitedTab);
    _renderBadges();
  }
  // Determinar qué tab key es el nuevo
  var txt = (newTab || '').toLowerCase();
  if (txt.includes('ap') || txt.includes('proveedores')) _prevVisitedTab = 'ap';
  else if (txt.includes('ar') || txt.includes('ota')) _prevVisitedTab = 'ar';
  else if (txt.includes('banco')) _prevVisitedTab = 'banco';
  else if (txt.includes('fb') || txt.includes('f&b')) _prevVisitedTab = 'fb';
  else if (txt.includes('drr')) _prevVisitedTab = 'drr';
  else _prevVisitedTab = null;
  
  // Highlight stats del nuevo tab si está actualizado
  setTimeout(_highlightActiveStats, 300);
}


// ── Escanear Documento Físico ───────────────────────────────────────
// ── Compresión de fotos en el navegador (las de móvil pesan 5-12MB) ──
async function _comprimirImagen(file) {
  try {
    var bmp = await createImageBitmap(file);
    var maxSide = 1800;
    var scale = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
    if (scale >= 1 && file.size < 1200000) { try { bmp.close(); } catch(e){} return file; } // ya es pequeña
    var c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(bmp.width * scale));
    c.height = Math.max(1, Math.round(bmp.height * scale));
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    try { bmp.close(); } catch(e){}
    var blob = await new Promise(function(res) { c.toBlob(res, 'image/jpeg', 0.85); });
    if (blob && blob.size < file.size) {
      return new File([blob], (file.name || 'foto').replace(/\.\w+$/, '') + '.jpg', { type: 'image/jpeg' });
    }
    return file;
  } catch(e) { return file; } // HEIC u otros que el navegador no decodifica: se envía tal cual
}

// ── UN grupo de fotos = las paginas de UN documento ───────────────────
// Es el camino del contrato de grupo tal cual estaba dentro de
// uploadAndProcess: misma llamada, mismo endpoint, mismos mensajes. Lo unico
// nuevo es que ahora se puede llamar VARIAS veces, una por documento, que es
// lo que hace falta para una tanda mezclada (contrato + factura + contrato).
// `pref` va vacio cuando la tanda trae un solo documento, para que ese caso
// —el que hoy acierta— salga byte a byte igual que antes.
async function _procesarGrupoFotos(grupo, pref, addLine, cierre) {
  var errs = 0;
  addLine('📄 ' + (pref || '') + 'Analizando ' + grupo.length + ' fotos como contrato de grupo (puede tardar ~1 min)...', 'l-info');
  var fdc = new FormData();
  for (var _k = 0; _k < grupo.length; _k++) {
    var _cf = await _comprimirImagen(grupo[_k]);
    fdc.append('files', _cf, _cf.name);
  }
  try {
    var _rc = await fetch('/api/ar_real/procesar_contrato', { method: 'POST', body: fdc, headers: { 'X-CSRF-Token': _csrfToken } });
    var _dc = await _rc.json();
    if (_dc && _dc.ok) {
      // BUG 8: el contrato tambien deja trabajo pendiente (la comision va a
      // AP y necesita cuenta y asiento). Antes solo las fotos sueltas pedian
      // el cierre, asi que la comision se quedaba sin contabilizar.
      if (cierre && _dc.cierre) { _dc.cierre.forEach(function(p){ if (p) cierre[p] = true; }); }
      var _money = function(v){ return '€' + (Number(v)||0).toLocaleString('es-ES', {minimumFractionDigits:2}); };
      addLine('✓ Contrato ' + (_dc.contrato || '') + ' · ' + (_dc.cliente || '') + ' · ' + _money(_dc.total_receivable) + ' → AR Real' + (_dc.beo_lineas ? ' · BEO con ' + _dc.beo_lineas + ' partidas' : ''), 'l-ok');
      var _di2 = _dc.distribucion || {};
      if (_di2.ap)    addLine('✓ AP comisión agencia: ' + _money(_di2.ap) + ' (pago pendiente)', 'l-ok');
      if (_di2.banco) addLine('✓ Banco depósito previsto: ' + _money(_di2.banco), 'l-ok');
      if (_di2.fb)    addLine('✓ F&B evento (banquete): ' + _money(_di2.fb), 'l-ok');
      if (_dc.requiere_certificado_di) addLine('⚠ Certificado de doble imposición pendiente', 'l-warn');
    } else if (_dc && _dc.reprocesar) {
      // LA MARCA, no el texto. El servidor pone `reprocesar` en dos casos
      // distintos ("no parecen un contrato" y "no se ha leido ningun dato
      // aprovechable") y antes solo se atrapaba el primero, porque se
      // comparaba el texto del error. El segundo se llevaba las fotos por
      // delante aunque el servidor estuviera pidiendo justo lo contrario.
      // El servidor manda un `message` DISTINTO para cada uno de los dos
      // casos que marca con `reprocesar`, y antes se enseñaba siempre el
      // mismo texto: "no son un contrato" cuando en realidad el sistema si
      // creyo que lo era y lo que fallo fue la lectura. Se usa el suyo.
      // Se le quita el "— revisar manualmente", que ya no es verdad: se
      // reprocesan solas.
      var _mm = String((_dc && _dc.message) || '').trim()
                  .replace(/\s*[—-]\s*revisar manualmente\.?$/i, '')
                  .replace(/\.$/, '');
      addLine((_mm || 'Las fotos no son un contrato') + ' — proceso cada una como documento suelto', 'l-info');
      errs += await _procesarImagenes(grupo, addLine, cierre);
    } else {
      // La IA fallo o devolvio algo que no se entiende. Las fotos siguen
      // siendo documentos: se les da su oportunidad una a una en vez de
      // descartarlas. Perder una factura en silencio es mucho peor que
      // gastar unas llamadas de mas.
      addLine('⚠ No se pudo leer el contrato: ' + ((_dc && _dc.error) || 'error') +
              ' — pruebo cada foto como documento suelto', 'l-warn');
      errs += await _procesarImagenes(grupo, addLine, cierre);
    }
  } catch(e) {
    // Se cayo la red entre el navegador y Render. Igual: no se tiran. Si la
    // red sigue caida, `_procesarImagenes` reintenta 3 veces por foto y
    // acabara cantando el error de cada una — que es lo que hay que ver.
    addLine('✗ ' + (e.message || 'error de red procesando el contrato') +
            ' — pruebo cada foto como documento suelto', 'l-err');
    errs += await _procesarImagenes(grupo, addLine, cierre);
  }
  return errs;
}

// ── Procesado de fotos de documentos, integrado en Procesar Archivos ──
function _mb(n) { return n > 950000 ? (n/1048576).toFixed(1) + 'MB' : Math.round(n/1024) + 'KB'; }

// ── CUANTAS FOTOS A LA VEZ ────────────────────────────────────────────
// De una en una, 12 fotos son ~2 minutos con la pantalla encendida, y ahi es
// donde el movil puede congelar la pestaña al cambiar de app. Lo que tarda es
// la IA (~6,5 s medidos por foto), no el navegador: la memoria ni se mueve.
// Con tres a la vez, esos 2 minutos son ~40 s.
//
// OJO: esto SOLO es seguro con el candado del servidor puesto. Sin el, tres
// guardados simultaneos se pisan el Excel y se pierden facturas (medido: 3 de
// 6). Si alguna vez Render se atraganta, se baja a 2 aqui y ya esta.
var _FOTOS_A_LA_VEZ = 3;

// Una linea de progreso viva, que se mueve al final del log para no perderse
// de vista. Solo aparece cuando hay algo que esperar (3 fotos o mas): con una
// o dos seria ruido, y con UNA el log tiene que salir identico al de siempre.
function _progresoFotos(total) {
  var log = document.getElementById('log');
  if (!log || total < 3) return { paso: function(){}, fin: function(){} };
  var p = document.createElement('p');
  p.className = 'l-info';
  p.style.fontWeight = '700';
  var t0 = Date.now(), hechas = 0;
  var pintar = function() {
    var txt = '⏳ ' + hechas + ' de ' + total;
    if (hechas > 0 && hechas < total) {
      // el tiempo medido por foto TERMINADA ya lleva dentro el paralelismo,
      // asi que multiplicar por las que quedan da el tiempo de reloj real.
      var seg = Math.round(((Date.now() - t0) / hechas) * (total - hechas) / 1000);
      txt += ' · quedan ~' + (seg >= 60 ? Math.ceil(seg / 60) + ' min' : Math.max(1, seg) + ' s');
    }
    p.textContent = _tSSE(txt);
    log.appendChild(p);                 // reaparece al final en cada paso
    log.scrollTop = log.scrollHeight;
  };
  pintar();
  return {
    paso: function() { hechas++; pintar(); },
    fin:  function() { try { p.remove(); } catch(e) {} }
  };
}

// El trabajo de UNA foto. Es el cuerpo del bucle de antes, tal cual: misma
// llamada, mismos 3 reintentos, mismos mensajes. `fi` sigue siendo el indice
// en la LISTA, no el de llegada, para que `[2/6]` signifique siempre la
// segunda foto que eligio el usuario aunque termine la cuarta.
async function _unaFoto(original, fi, total, addLine, acc) {
  var errors = 0;
  var file = await _comprimirImagen(original);
  var label = '[' + (fi+1) + '/' + total + '] ' + (file.name || 'foto');
  var sizeInfo = file.size < original.size ? ' (' + _mb(original.size) + ' → ' + _mb(file.size) + ')' : ' (' + _mb(file.size) + ')';
  addLine('🔍 ' + label + sizeInfo + '...', 'l-info');
  var success = false;
  for (var retry = 0; retry < 3 && !success; retry++) {
    try {
      var formData = new FormData();
      formData.append('image', file);
      var r = await fetch('/api/scan_documento', { method: 'POST', body: formData, headers: { 'X-CSRF-Token': _csrfToken } });
      var data = await r.json();
      if (data.ok) {
        var _ok = (data.guardado !== false);
        // lo que esta foto deja pendiente (cruce / asignador). Este era el
        // agujero: scan_documento guardaba el documento y ahi se acababa.
        if (acc && data.cierre) { data.cierre.forEach(function(p){ if (p) acc[p] = true; }); }
        addLine((_ok ? '✓ ' : '⚠ ') + (file.name || 'foto') + ': ' + (data.tipo || '—') + (data.mensaje ? ' — ' + data.mensaje : ''), _ok ? 'l-ok' : 'l-warn');
      } else {
        addLine('✗ ' + (file.name || 'foto') + ': ' + (data.error || 'error'), 'l-err');
        errors++;
      }
      success = true;
    } catch(e) {
      if (retry < 2) {
        addLine('⚠ ' + label + ' — ' + _tSSE('Reconectando') + '... (' + (retry+2) + '/3)', 'l-warn');
        await new Promise(function(res) { setTimeout(res, 2000 * (retry+1)); });
      } else {
        addLine('✗ ' + label + ' — ' + e.message, 'l-err');
        errors++; success = true;
      }
    }
  }
  return errors;
}

async function _procesarImagenes(imgs, addLine, acc) {
  var lista = imgs || [];
  var total = lista.length;
  if (!total) return 0;
  var errors = 0;
  var prog = _progresoFotos(total);
  // Un pozo de trabajo: cada obrero coge la siguiente foto que quede. Asi la
  // foto lenta no bloquea a las demas, que es lo que pasaria repartiendolas
  // en tandas fijas de tres.
  var siguiente = 0;
  var obreros = [];
  var aLaVez = Math.min(_FOTOS_A_LA_VEZ, total);
  for (var w = 0; w < aLaVez; w++) {
    obreros.push((async function() {
      while (true) {
        var i = siguiente++;
        if (i >= total) return;
        // OJO: `errors += await ...` NO vale con varios obreros. El `+=`
        // LEE `errors` ANTES de esperar y lo escribe despues, asi que dos
        // obreros que terminan a la vez se pisan la cuenta y los errores
        // desaparecen. Lo cazo el test: dos fotos fallaban y el contador
        // decia 0. Se lee DESPUES de la espera.
        var _errFoto = await _unaFoto(lista[i], i, total, addLine, acc);
        errors += _errFoto;
        prog.paso();
      }
    })());
  }
  await Promise.all(obreros);
  prog.fin();
  return errors;
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
  indicator.style.cssText = 'position:fixed;top:0;left:50%;transform:translateX(-50%);background:var(--acc);color:#fff;padding:6px 16px;border-radius:0 0 12px 12px;font-size:12px;font-weight:600;z-index:9999;display:none;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s';
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

// ── Hotel activo: filtra AR/AP/AR Real a un hotel concreto ────────────────
// FASE F · Multi-Hotel es una vista de GRUPO, asi que solo existe en el grupo.
//
// Dentro de un hotel no significa nada: enseñaria las tarjetas de todos los
// hoteles justo cuando el usuario ha dicho que quiere mirar uno. Y como el
// agregador NO depende del hotel de la sesion (se comprobo en la fase A), la
// pestaña seguiria enseñando el grupo entero — que es peor que no estar,
// porque parece que el filtro no funciona.
//
// En vista de grupo pasa a ser la PRIMERA: es la que se mira al entrar.
function _ordenarPestanaMultiHotel(hayHotelElegido) {
  var tab = document.getElementById('tab-multi-hotel');
  if (!tab) return;
  var barra = tab.parentElement;
  if (!barra) return;

  if (hayHotelElegido) {
    tab.style.display = 'none';
    // Si estabamos DENTRO de Multi-Hotel al elegir hotel, la pestaña se
    // esconde debajo de los pies. Hay que llevarse al usuario a algun sitio,
    // no dejarlo mirando un panel sin pestaña.
    var panel = document.getElementById('panel-multi_hotel');
    if (panel && panel.classList.contains('active')) {
      var destino = document.getElementById('tab-ar_otas') ||
                    barra.querySelector('.tab:not([style*="none"])');
      if (destino && typeof switchTab === 'function') {
        var m = (destino.getAttribute('onclick') || '').match(/switchTab\('([^']+)'/);
        if (m) switchTab(m[1], destino);
      }
    }
    return;
  }

  tab.style.display = '';
  if (barra.firstElementChild !== tab) barra.insertBefore(tab, barra.firstElementChild);
}

async function _initHotelActivo() {
  try {
    var d = await fetch('/api/hotel_activo').then(function(r){ return r.json(); });
    var sel = document.getElementById('hotel-activo-sel');
    _ordenarPestanaMultiHotel(!!(d && d.hotel));
    if (!sel) return;
    if (!d.hoteles || d.hoteles.length < 2) { sel.style.display = 'none'; return; }
    sel.style.display = '';
    // El value es el ID; lo que se lee es el nombre. Antes el value era el
    // nombre y por eso dos hoteles parecidos se confundian.
    // Sin emoji en las opciones: el desplegable del movil lo pinta el sistema
    // y cada icono le suma alto y ancho a una lista que ya es larga.
    sel.innerHTML = '<option value="">' + t('mh.todosHoteles', 'Todos los hoteles') + '</option>' +
      d.hoteles.map(function(h) {
        var _id = String((h && h.id) || ''), _nom = String((h && h.nombre) || '');
        return '<option value="' + _id.replace(/"/g, '&quot;') + '"' + (d.hotel === _id ? ' selected' : '') +
               '>' + _nom.replace(/</g, '&lt;') + '</option>';
      }).join('');
  } catch(e) {}
}

async function seleccionarHotelActivo(h, irATab) {
  try {
    // `h` es el ID; el nombre para el aviso lo devuelve el servidor, que es
    // quien tiene el censo.
    var _r = await _postJson('/api/hotel_activo', {hotel: h || ''});
    var _d = {}; try { _d = await _r.json(); } catch(e) {}
    var _nom = (_d && _d.nombre) || '';
    showNotification(_nom ? '🏨 ' + _nom : '🌍 ' + t('mh.vistaGrupo', 'Vista de grupo (todos los hoteles)'), 'info');
    _initHotelActivo();
    // Lo cacheado es del hotel ANTERIOR. Esto va antes de repintar nada.
    //
    // Sin esto solo se refrescaba lo que estuviera a la vista, y los paneles
    // con cargador propio (F&B, DRR, Multi-Hotel) se quedaban con los numeros
    // del hotel de antes hasta que los abrias... y ni asi, porque al abrirlos
    // `_cargarPanel` los daba por cargados y no volvia a pedir nada.
    //
    // AR Real no lo sufria por casualidad: se le llama a pelo aqui abajo,
    // saltandose `_cargarPanel`. Casualidad, no diseño — y por eso el arreglo
    // va en `_invalidarPaneles` y no en un `_refrescarFB()` suelto: asi cubre
    // tambien DRR y Multi-Hotel, que es el que manda a partir de la fase B.
    try { if (typeof _invalidarPaneles === 'function') _invalidarPaneles(); } catch(e){}
    loadAll(); loadAP(); loadBanco();
    if (typeof cargarARRealData === 'function') { try { cargarARRealData(); } catch(e) {} }
    if (h && irATab) {
      var tabEl = document.getElementById('tab-ar');
      if (tabEl) switchTab('ar', tabEl);
    }
  } catch(e) { showNotification('✗ ' + e.message, 'error'); }
}
_initHotelActivo();

function parseCadenasDemo(texto) {
  var cadenas = [];
  texto.split('\n').map(function(l){ return l.trim(); }).filter(Boolean).forEach(function(linea) {
    if (linea.indexOf(':') > -1) {
      var partes = linea.split(':');
      var nombre = partes[0].trim();
      var hoteles = partes.slice(1).join(':').split(',').map(function(h){ return h.trim(); }).filter(Boolean);
      if (nombre && hoteles.length) cadenas.push({nombre: nombre, hoteles: hoteles});
    } else {
      cadenas.push({nombre: linea, hoteles: [linea]});
    }
  });
  return cadenas;
}

async function generarDemo() {
  var cadenas = parseCadenasDemo(document.getElementById('demo-setup-input').value || '');
  if (!cadenas.length) { showNotification(t('demo.faltaNombre', 'Escribe al menos un hotel o cadena'), 'info'); return; }
  var btn = document.getElementById('btn-demo-generar');
  btn.disabled = true; btn.textContent = '⏳ ' + t('demo.generando', 'Generando datos...');
  try {
    var r = await _postJson('/api/demo/generar', {cadenas: cadenas});
    var d = await r.json();
    if (!d.ok) throw new Error(d.error || 'error');
    document.getElementById('demo-setup-modal').style.display = 'none';
    demoModeActive = true;
    var banner = document.getElementById('demo-banner');
    if (banner) { banner.style.display = 'block'; document.body.style.paddingTop = '36px'; }
    var btnD = document.getElementById('btn-demo');
    if (btnD) { btnD.style.color = '#f59e0b'; btnD.querySelector('span') && (btnD.querySelector('span').textContent = '🎭 Demo ON'); }
    showNotification('🎭 ' + d.hoteles + ' ' + t('demo.hotelesListos', 'hotel(es) con datos de ejemplo listos'), 'success');
    _mhClasicaLoaded = false;
    _initHotelActivo();
    loadAll(); loadAP(); loadBanco();
    var tabDestino = d.hoteles > 1 ? 'multi_hotel' : 'ar';
    var tabEl = document.getElementById('tab-' + tabDestino) || document.getElementById('tab-' + tabDestino.replace(/_/g, '-'));
    if (tabEl) switchTab(tabDestino, tabEl);
  } catch(e) {
    showNotification('✗ Demo: ' + e.message, 'error');
  }
  btn.disabled = false; btn.textContent = '🎭 ' + t('demo.generar', 'Generar demo');
}

async function toggleDemoMode() {
  // Activar → pedir nombres primero; el toggle real solo apaga
  if (!demoModeActive) {
    document.getElementById('main-menu')?.classList.remove('open');
    document.getElementById('demo-setup-modal').style.display = 'flex';
    setTimeout(function() { document.getElementById('demo-setup-input').focus(); }, 150);
    return;
  }
  try {
    const res = await _postJson('/api/demo/toggle', {});
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
        // Load Calipolis data
        // Show tour prompt after 1.5s
        setTimeout(() => {
          const tourMsg = document.createElement('div');
          tourMsg.style.cssText = 'position:fixed;bottom:80px;right:20px;background:linear-gradient(135deg,#1e293b,#0d1827);border:1px solid rgba(245,158,11,.4);border-radius:14px;padding:18px 20px;z-index:8500;max-width:280px;box-shadow:0 8px 32px rgba(0,0,0,.5)';
          document.body.appendChild(tourMsg);
          setTimeout(() => tourMsg.remove(), 12000);
        }, 1500);
      }, 300);
    } else {
      if (banner) { banner.style.display = 'none'; document.body.style.paddingTop = ''; }
      if (btn) { btn.style.color = ''; btn.querySelector('span') && (btn.querySelector('span').textContent = '🎭 Demo Mode'); }
      _mhClasicaLoaded = false; _mhGrupoActivo = '';
      seleccionarHotelActivo('');
      _initHotelActivo();
      showNotification(t('demo.off', 'Demo desactivado — datos de ejemplo eliminados'), 'info');
      loadAll(); loadAP(); loadBanco();
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
    const res = await _postJson(`/api/rol/cambiar/${newRole}`, {});
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
  // Atajo de tab por URL (?tab=ap|ar|drr|notif|fb|banco) — usado por los shortcuts de la PWA
  try {
    var _tab = new URLSearchParams(window.location.search).get('tab');
    if (_tab) {
      var _tb = document.querySelector('.tab[onclick*="switchTab(\'' + _tab + '\'"]');
      if (_tb) setTimeout(function(){ switchTab(_tab, _tb); }, 60);
    }
  } catch(e){}
  // Banner de instalación (iOS al instante; Android tras beforeinstallprompt)
  setTimeout(function(){ if (typeof yvePwaMaybeShowInstall === 'function') yvePwaMaybeShowInstall(); }, 1200);
});



function eliminarArchivoServidor(nombre, rowEl) {
  // M2: sin confirmacion. Es poco destructivo —el fichero sigue en
  // facturas-entrada— y el dialogo estorbaba justo cuando mas se usa: al
  // quitar varios ya procesados seguidos.
  _postJson('/api/eliminar_archivo', {nombre: nombre})
  .then(function(r){ return r.json(); })
  .then(function(d) {
    if (d.ok) {
      rowEl.style.opacity = '0';
      rowEl.style.transform = 'translateX(20px)';
      rowEl.style.transition = 'all .3s';
      setTimeout(function(){ rowEl.remove(); }, 300);
      showNotification('✓ ' + nombre + ' eliminado', 'info');
      // Actualizar contador del botón procesar pendientes
      var serverBtn = document.getElementById('btn-procesar-server');
      if (serverBtn) {
        var remaining = document.querySelectorAll('[id^="file-row-"]').length - 1;
        if (remaining <= 0) {
          serverBtn.style.display = 'none';
          document.getElementById('server-files-section').style.display = 'none';
        } else {
          serverBtn.textContent = '▶ Procesar ' + remaining + ' pendiente' + (remaining !== 1 ? 's' : '') + ' del servidor';
        }
      }
    } else {
      showNotification('Error: ' + (d.error || 'no se pudo eliminar'), 'error');
    }
  })
  .catch(function(e){ showNotification('Error de conexión', 'error'); });
}

function switchTab(tab, el) {
  if (typeof t !== 'function' && typeof _T_FN === 'function') { try { t = _T_FN; } catch(e){} }
  if (typeof _onTabSwitch === 'function') _onTabSwitch(tab);
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  var panel = document.getElementById('panel-' + tab);
  if (panel) panel.classList.add('active');
  // Salto INSTANTANEO, no suave. Medido: viniendo de otro apartado con la
  // pagina desplazada 600 px, el scroll suave hacia arriba se cruza con el
  // cambio de alto del panel (AR OTAs mide 1604 px, DRR 1172, Banco 751), el
  // navegador lo corta a medias y te deja en 423 en vez de arriba. Por eso solo
  // se notaba al entrar en los apartados largos: en los cortos no hay sitio
  // donde quedarse a medias.
  window.scrollTo(0, 0);
  // Mobile: scroll tab into view + highlight bottom nav
  if (IS_MOBILE && el) {
    setTimeout(function(){ el.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'}); }, 50);
  }
  // Lo que ya estuviera pintado se traduce YA, sin esperar al observer.
  if (panel) _pintarYa(panel);
  // La PRIMERA vez que se abre Banco (sin config elegida) sale el modal de
  // "¿cómo funciona tu banco?". Va aqui, no en el cargador, para que salga al
  // ABRIR la pestaña (no en la precarga de fondo) aunque el panel este cacheado.
  if (tab === 'banco' && typeof _checkBancoConfig === 'function') _checkBancoConfig();
  _cargarPanel(tab, panel, false);
}

// Cargadores por apartado, en un solo sitio. Antes estaban sueltos dentro de
// switchTab y se llamaban CADA VEZ que entrabas, repintando el panel entero.
var _CARGADORES = {
  fb:          function(){ return loadFBTab(); },
  ar_real:     function(){ return cargarARRealData(); },
  drr:         function(){ return loadDRR(); },
  banco:       function(){ return loadBanco(); },
  notif:       function(){ return loadNotifConfig(); },
  multi_hotel: function(){ return loadMultiHotel(); },
  cierre:      function(){ return loadCierre(); }
};

// ── Cuadre de banco por pestañas (Ola B·2) ───────────────────────────────
var _cbFiltro = '';
async function loadCuadreBanco(){
  var inp = document.getElementById('cierre-mes'); if (!inp) return;
  var mes = inp.value;
  var ex = document.getElementById('cbanco-excel'); if (ex) ex.href = '/api/exportar/cuadre_banco?mes=' + encodeURIComponent(mes);
  var pw = document.getElementById('cbanco-pestanas'), body = document.getElementById('cbanco-body'), sal = document.getElementById('cbanco-saldo');
  try {
    var r = await fetch('/api/cuadre_banco?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok_api) { body.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    if (sal) sal.textContent = (d.saldo_final!=null ? t('cbanco.saldo','saldo del extracto {f}: {s}').replace('{f}', d.fecha_saldo).replace('{s}', _cEur(d.saldo_final)) + ' · ' : '') + t('cbanco.movs','{n} movimientos · {p} sin conciliar').replace('{n}', d.n).replace('{p}', d.sin_conciliar);
    var LBL = {AR:t('cbanco.AR','AR · cobros'), AP:t('cbanco.AP','AP · pagos'), TARJETAS:t('cbanco.TARJETAS','Tarjetas'), CAJA:t('cbanco.CAJA','Income / caja'), VARIOS:t('cbanco.VARIOS','Varios'), SIN_CLASIFICAR:t('cbanco.SIN','Sin clasificar')};
    var COL = {CUADRA:'#22c55e', PENDIENTE:'#f59e0b', INFO:'var(--dim)', SIN_DATO:'var(--mut)'};
    var ST = {CUADRA:t('cierre.stCuadra','✓ cuadra'), PENDIENTE:t('cierre.stPend','pendiente'), INFO:'info', SIN_DATO:t('cierre.stSinDato','sin dato')};
    var ps = d.pestanas || {};
    pw.innerHTML = Object.keys(LBL).map(function(k){ var p = ps[k] || {}; var act = _cbFiltro===k;
      return '<div class="card" onclick="_cbFiltro=(_cbFiltro===\'' + k + '\'?\'\':\'' + k + '\');loadCuadreBanco()" style="padding:10px;border-radius:10px;cursor:pointer;' + (act?'outline:2px solid var(--acc,#3b82f6);':'') + '"><div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + _cEsc(LBL[k]) + '</div><div style="font-size:15px;font-weight:700">' + _cEur(p.total) + '</div><div style="font-size:11px;color:var(--dim)">' + (p.n||0) + ' mov.' + (p.justificado!=null ? ' · ' + t('cierre.hJust','Justificado') + ' ' + _cEur(p.justificado) : '') + '</div><div style="font-size:11px;font-weight:700;color:' + (COL[p.estado]||'var(--mut)') + '">' + _cEsc(ST[p.estado]||p.estado||'') + '</div></div>'; }).join('');
    var ms = (d.movimientos||[]).filter(function(m){ return !_cbFiltro || m.pestana===_cbFiltro; });
    var opts = Object.keys(LBL).map(function(k){ return '<option value="' + k + '">' + _cEsc(LBL[k]) + '</option>'; }).join('');
    body.innerHTML = ms.length ? '<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="color:var(--dim);text-align:left"><th>' + t('cierre.hFecha','Fecha') + '</th><th>' + t('cierre.hConceptoD','Concepto') + '</th><th style="text-align:right">' + t('cbanco.importe','Importe') + '</th><th>' + t('cbanco.factura','Factura') + '</th><th>' + t('cbanco.pestana','Pestaña') + '</th></tr></thead><tbody>' +
      ms.map(function(m){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:3px 4px;white-space:nowrap">' + _cEsc(m.fecha) + '</td><td style="padding:3px 4px">' + _cEsc(m.concepto) + '</td><td style="text-align:right;padding:3px 4px;color:' + (m.importe<0?'#f87171':'#22c55e') + '">' + _cEur(m.importe) + '</td><td style="padding:3px 4px;color:var(--dim)">' + _cEsc(m.factura_ref||'') + '</td><td style="padding:3px 4px"><select data-clave="' + _cEsc(m.clave) + '" onchange="_cbAsignar(this)" style="background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:3px;border-radius:6px;font-size:11px">' + opts.replace('value="' + m.pestana + '"', 'value="' + m.pestana + '" selected') + '</select>' + (m.via==='manual' ? ' <span title="manual">✎</span>' : '') + '</td></tr>'; }).join('') + '</tbody></table>' : '<div class="empty"><p>' + t('cbanco.vacio','Sin movimientos del extracto en este mes. Súbelo en la pestaña Banco.') + '</p></div>';
  } catch(e) { if (body) body.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
async function _cbAsignar(sel){
  try {
    var r = await _postJson('/api/cuadre_banco/asignar', {clave: sel.getAttribute('data-clave'), pestana: sel.value});
    var d = await r.json();
    if (!d || !d.ok) showNotification('✗ ' + ((d&&d.error)||'No se pudo asignar'), 'error');
    loadCuadreBanco();
  } catch(e){ showNotification('✗ ' + e.message, 'error'); }
}

// ── Inventarios de cierre (Ola B·3) ──────────────────────────────────────
async function loadInventarios(){
  var inp = document.getElementById('cierre-mes'); if (!inp) return;
  var mes = inp.value;
  var ex = document.getElementById('inv-excel'); if (ex) ex.href = '/api/exportar/inventarios?mes=' + encodeURIComponent(mes);
  var hj = document.getElementById('inv-hoja'); if (hj) hj.href = '/api/inventarios/hoja_recuento?mes=' + encodeURIComponent(mes);
  var rs = document.getElementById('inv-resumen'), body = document.getElementById('inv-body');
  try {
    var r = await fetch('/api/inventarios?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok) { body.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    var s = d.resumen || {};
    var tile = function(l, v, sub, col){ return '<div class="card" style="padding:10px;border-radius:10px"><div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + _cEsc(l) + '</div><div style="font-size:15px;font-weight:700;color:' + (col||'var(--tx)') + '">' + v + '</div>' + (sub?'<div style="font-size:11px;color:var(--dim)">' + _cEsc(sub) + '</div>':'') + '</div>'; };
    rs.innerHTML = tile(t('inv.kFinal','Existencias finales'), _cEur(s.valor_final), s.n_articulos + ' ' + t('inv.articulos','artículos')) +
      tile(t('inv.kCompras','Compras F&B del mes'), _cEur(s.compras_fb), (s.n_facturas_fb||0) + ' ' + t('inv.facturas','facturas')) +
      tile(t('inv.kReal','Consumo real F&B'), s.consumo_real_fb==null?'—':_cEur(s.consumo_real_fb), t('inv.formula','inicial + compras − final')) +
      tile(t('inv.kTeorico','Consumo teórico'), s.consumo_teorico_fb==null?t('inv.sinDato','sin dato'):_cEur(s.consumo_teorico_fb), t('inv.escandallo','escandallo × ventas')) +
      tile(t('inv.kDesv','Desviación'), s.desviacion_fb==null?'—':_cEur(s.desviacion_fb) + (s.desviacion_pct!=null?' (' + s.desviacion_pct + ' %)':''), s.n_revisar ? s.n_revisar + ' ' + t('inv.revisar','artículos a revisar') : '', s.desviacion_fb==null?'var(--tx)':(Math.abs(s.desviacion_pct||0)>5?'#f87171':'#22c55e'));
    var fams = d.familias || [];
    var LBL = {ALIMENTOS:t('inv.ALIMENTOS','Alimentos'), BEBIDAS:t('inv.BEBIDAS','Bebidas'), LICORES:t('inv.LICORES','Licores'), GUEST_SUPPLIES:t('inv.GUEST','Guest supplies'), OTROS:t('inv.OTROS','Otros')};
    var h = fams.length ? '<table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="color:var(--dim);text-align:left"><th>' + t('inv.familia','Familia') + '</th><th style="text-align:right">' + t('inv.articulos','artículos') + '</th><th style="text-align:right">' + t('inv.inicial','Inicial') + '</th><th style="text-align:right">' + t('inv.final','Final') + '</th><th style="text-align:right">' + t('inv.variacion','Variación') + '</th><th style="text-align:right">' + t('inv.revisar','a revisar') + '</th></tr></thead><tbody>' +
      fams.map(function(f){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:5px 4px;font-weight:700">' + _cEsc(LBL[f.familia]||f.familia) + '</td><td style="text-align:right;padding:5px 4px">' + f.n + '</td><td style="text-align:right;padding:5px 4px">' + _cEur(f.valor_inicial) + '</td><td style="text-align:right;padding:5px 4px">' + _cEur(f.valor_final) + '</td><td style="text-align:right;padding:5px 4px;color:' + (f.variacion<0?'#f87171':'#22c55e') + '">' + _cEur(f.variacion) + '</td><td style="text-align:right;padding:5px 4px">' + (f.revisar||0) + '</td></tr>'; }).join('') + '</tbody></table>' : '<div class="empty"><p>' + t('inv.vacio','Sin inventario. Descarga la hoja de recuento, cuéntalo y súbela, o procesa un inventario en Procesar Archivos.') + '</p></div>';
    var asi = d.asientos || [];
    if (asi.length) h += '<div style="margin-top:10px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)">' + t('inv.asientos','Asiento de variación de existencias') + '</div><table style="width:100%;border-collapse:collapse;font-size:11px">' + asi.map(function(a){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:3px 4px"><b>' + _cEsc(a.cuenta) + '</b> ' + _cEsc(a.desc_cuenta) + '</td><td style="padding:3px 4px">' + _cEsc(a.concepto) + '</td><td style="text-align:right;padding:3px 4px">' + (a.debe?_cEur(a.debe):'') + '</td><td style="text-align:right;padding:3px 4px">' + (a.haber?_cEur(a.haber):'') + '</td></tr>'; }).join('') + '</table>';
    var rev = d.revisar || [];
    if (rev.length) h += '<div style="margin-top:10px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#f59e0b">' + t('inv.revisarTit','Artículos a revisar') + '</div><div style="font-size:11px">' + rev.slice(0,30).map(function(a){ return _cEsc(a.articulo) + ' — ' + _cEsc(a.motivo); }).join('<br>') + '</div>';
    if (s.nota) h += '<div style="margin-top:8px;font-size:11px;color:var(--dim)">' + _cEsc(s.nota) + '</div>';
    body.innerHTML = h;
  } catch(e) { if (body) body.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
async function _invSubir(input){
  var f = input.files && input.files[0]; if (!f) return;
  var fd = new FormData(); fd.append('archivo', f);
  try {
    var tok = ''; try { tok = (await (await fetch('/api/csrf_token')).json()).token || ''; } catch(e){}
    var r = await fetch('/api/inventarios/recuento?mes=' + encodeURIComponent((document.getElementById('cierre-mes')||{}).value||''), {method:'POST', body: fd, headers: tok ? {'X-CSRF-Token': tok} : {}});
    var d = await r.json();
    if (d && d.ok) { showNotification(t('inv.subido','✓ Recuento subido: {n} artículos').replace('{n}', d.contados), 'success'); loadInventarios(); }
    else showNotification('✗ ' + ((d&&d.error)||'Error'), 'error');
  } catch(e){ showNotification('✗ ' + e.message, 'error'); }
  input.value = '';
}

// ── Inmovilizado y amortizaciones (Ola B·4) ──────────────────────────────
var _inmCats = null;
async function loadFiscal(){
  var inp = document.getElementById('cierre-mes'); if (!inp) return;
  var mes = inp.value;
  var ex = document.getElementById('fis-excel'); if (ex) ex.href = '/api/exportar/fiscal?mes=' + encodeURIComponent(mes);
  var rs = document.getElementById('fis-resumen'), body = document.getElementById('fis-body'), est = document.getElementById('fis-estado');
  try {
    var r = await fetch('/api/fiscal?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok) { body.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    var m = d.m303 || {}, s3 = d.m349 || {}, sii = d.sii || {};
    var col = {PREPARADO:'var(--ok)', PENDIENTE:'var(--warn)', SIN_DATO:'var(--mut)'}[d.estado] || 'var(--mut)';
    est.innerHTML = '<span style="color:' + col + '">' + _cEsc(d.estado||'') + '</span>';
    var tile = function(l, v, sub){ return '<div class="card" style="padding:10px;border-radius:10px"><div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + _cEsc(l) + '</div><div style="font-size:15px;font-weight:700">' + v + '</div>' + (sub?'<div style="font-size:11px;color:var(--dim)">' + _cEsc(sub) + '</div>':'') + '</div>'; };
    var lib = d.libro;
    rs.innerHTML = tile(t('fis.k303','303 · resultado'), _cEur(m.c46_resultado), (m.signo||'')) +
      tile(t('fis.kDev','IVA devengado (27)'), _cEur(m.c27_devengado), lib ? (t('fis.libro','libro 477') + ' ' + _cEur(lib.iva_repercutido_477)) : '') +
      tile(t('fis.kDed','IVA deducible (45)'), _cEur(m.c45_deducible), lib ? (t('fis.libro472','libro 472') + ' ' + _cEur(lib.iva_soportado_472)) : '') +
      tile(t('fis.k349','349 · operadores UE'), (s3.n||0), _cEur(s3.total_base)) +
      tile(t('fis.kSii','SII'), (sii.n_expedidas||0) + ' / ' + (sii.n_recibidas||0), t('fis.siiSub','expedidas / recibidas'));
    var h = '';
    if (lib && !lib.cuadra) h += '<div style="font-size:12px;color:var(--warn);margin-bottom:8px">⚠️ ' + _cEsc(t('fis.noCuadra','El 303 no cuadra con el libro del mes (477/472). Revisa los asientos antes de presentar.')) + '</div>';
    if (!sii.n_expedidas && !sii.n_recibidas) { h += _vacioCard(t('fis.vacio','Sin facturas ni ventas en este mes: no hay nada que declarar.')); body.innerHTML = h; return; }
    h += '<div style="overflow-x:auto"><table class="tbl" style="width:100%;font-size:12px"><thead><tr><th>' + _cEsc(t('fis.casilla','Casilla')) + '</th><th>' + _cEsc(t('fis.concepto','Concepto')) + '</th><th style="text-align:right">' + _cEsc(t('fis.base','Base')) + '</th><th style="text-align:right">' + _cEsc(t('fis.cuota','Cuota')) + '</th></tr></thead><tbody>';
    (m.casillas||[]).forEach(function(c){ if (!c.base && !c.cuota) return; h += '<tr><td>' + _cEsc(c.casilla_base + ' / ' + c.casilla_cuota) + '</td><td>' + _cEsc(c.concepto) + '</td><td style="text-align:right">' + _cEur(c.base) + '</td><td style="text-align:right">' + _cEur(c.cuota) + '</td></tr>'; });
    h += '<tr style="font-weight:700"><td>46</td><td>' + _cEsc(t('fis.resultado','Resultado')) + ' (' + _cEsc(m.signo||'') + ')</td><td></td><td style="text-align:right">' + _cEur(m.c46_resultado) + '</td></tr></tbody></table></div>';
    if ((s3.filas||[]).length) {
      h += '<div style="font-weight:700;font-size:12px;margin:12px 0 6px">' + _cEsc(t('fis.t349','Modelo 349 — operaciones intracomunitarias')) + '</div><div style="overflow-x:auto"><table class="tbl" style="width:100%;font-size:12px"><thead><tr><th>' + _cEsc(t('fis.operador','Operador')) + '</th><th>NIF-IVA</th><th>' + _cEsc(t('fis.clave','Clave')) + '</th><th style="text-align:right">' + _cEsc(t('fis.base','Base')) + '</th></tr></thead><tbody>';
      s3.filas.forEach(function(f){ h += '<tr><td>' + _cEsc(f.operador) + '</td><td>' + (f.nif ? _cEsc(f.nif) : '<span style="color:var(--warn)">' + _cEsc(t('fis.nifPend','pendiente')) + '</span>') + '</td><td>' + _cEsc(f.clave) + '</td><td style="text-align:right">' + _cEur(f.base) + '</td></tr>'; });
      h += '</tbody></table></div>';
    }
    if ((d.avisos||[]).length) h += '<div style="font-size:12px;color:var(--mut);margin-top:10px">' + d.avisos.map(function(a){ return '• ' + _cEsc(a); }).join('<br>') + '</div>';
    body.innerHTML = h;
  } catch(e) { body.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
async function loadInmovilizado(){
  var inp = document.getElementById('cierre-mes'); if (!inp) return;
  var mes = inp.value;
  var ex = document.getElementById('inm-excel'); if (ex) ex.href = '/api/exportar/inmovilizado?mes=' + encodeURIComponent(mes);
  var rs = document.getElementById('inm-resumen'), body = document.getElementById('inm-body');
  try {
    var r = await fetch('/api/inmovilizado?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok) { body.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    _inmCats = d.categorias || {};
    var s = d.resumen || {};
    var tile = function(l, v, sub){ return '<div class="card" style="padding:10px;border-radius:10px"><div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + _cEsc(l) + '</div><div style="font-size:15px;font-weight:700">' + v + '</div>' + (sub?'<div style="font-size:11px;color:var(--dim)">' + _cEsc(sub) + '</div>':'') + '</div>'; };
    rs.innerHTML = tile(t('inm.kActivos','Activos'), s.n_activos||0, (s.n_en_curso||0) + ' ' + t('inm.enCurso','en curso') + ' · ' + (s.n_amortizados||0) + ' ' + t('inm.amortizados','amortizados')) +
      tile(t('inm.kCoste','Coste total'), _cEur(s.coste_total)) + tile(t('inm.kAcum','Amortización acumulada'), _cEur(s.acumulada_total)) +
      tile(t('inm.kVnc','Valor neto contable'), _cEur(s.vnc_total)) + tile(t('inm.kCuota','Cuota del mes'), _cEur(s.cuota_mes), (s.altas_pendientes ? s.altas_pendientes + ' ' + t('inm.altasPend','posibles altas sin registrar') : ''));
    var acts = d.activos || [];
    var EST = {EN_CURSO:t('inm.enCurso','en curso'), AMORTIZADO:t('inm.amortizado','amortizado'), BAJA:t('inm.baja','baja'), NO_ALTA:t('inm.noAlta','alta posterior'), ERROR:'error'};
    var h = acts.length ? '<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="color:var(--dim);text-align:left"><th>' + t('inm.hDesc','Activo') + '</th><th>' + t('inm.hCat','Categoría') + '</th><th>' + t('inm.hAlta','Alta') + '</th><th style="text-align:right">' + t('inm.hCoste','Coste') + '</th><th style="text-align:right">' + t('inm.hCuota','Cuota mes') + '</th><th style="text-align:right">' + t('inm.hAcum','Acumulada') + '</th><th style="text-align:right">VNC</th><th></th><th></th></tr></thead><tbody>' +
      acts.map(function(a){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:4px">' + _cEsc(a.descripcion) + (a.error?'<div style="color:#f87171">' + _cEsc(a.error) + '</div>':'') + '</td><td style="padding:4px">' + _cEsc((_inmCats[a.categoria]||{}).nombre||a.categoria) + '</td><td style="padding:4px;white-space:nowrap">' + _cEsc(a.fecha_alta||'') + (a.fecha_baja?'<br><span style="color:#f87171">⤓ ' + _cEsc(a.fecha_baja) + '</span>':'') + '</td><td style="text-align:right;padding:4px">' + _cEur(a.coste) + '</td><td style="text-align:right;padding:4px">' + _cEur(a.cuota) + '</td><td style="text-align:right;padding:4px">' + _cEur(a.acumulada) + '</td><td style="text-align:right;padding:4px;font-weight:700">' + _cEur(a.vnc) + '</td><td style="padding:4px;color:var(--dim)">' + _cEsc(EST[a.estado]||a.estado) + '</td><td style="padding:4px">' + (a.estado==='EN_CURSO'||a.estado==='AMORTIZADO' ? '<button class="btn-ref" style="font-size:10px;padding:2px 6px" onclick="_inmBaja(\'' + _cEsc(a.id) + '\')">' + t('inm.darBaja','Baja') + '</button>' : '') + '</td></tr>'; }).join('') + '</tbody></table>' : '<div class="empty"><p>' + t('inm.vacio','Sin activos registrados. Da de alta el mobiliario, la maquinaria, los equipos… y Yve calcula la amortización de cada mes.') + '</p></div>';
    var asi = d.asientos || [];
    if (asi.length) h += '<div style="margin-top:10px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)">' + t('inm.asiento','Asiento de amortización del mes') + '</div><table style="width:100%;border-collapse:collapse;font-size:11px">' + asi.map(function(a){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:3px 4px"><b>' + _cEsc(a.cuenta) + '</b> ' + _cEsc(a.desc_cuenta) + '</td><td style="padding:3px 4px">' + _cEsc(a.concepto) + '</td><td style="text-align:right;padding:3px 4px">' + (a.debe?_cEur(a.debe):'') + '</td><td style="text-align:right;padding:3px 4px">' + (a.haber?_cEur(a.haber):'') + '</td></tr>'; }).join('') + '</table>';
    var pend = d.altas_pendientes || [];
    if (pend.length) h += '<div style="margin-top:10px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#f59e0b">' + t('inm.pendTit','Facturas del mes que podrían ser un activo') + '</div><div style="font-size:11px">' + pend.slice(0,20).map(function(p){ return _cEsc(p.numero_factura) + ' · ' + _cEsc(p.proveedor) + ' · ' + _cEur(p.base) + ' · ' + _cEsc(p.motivo) + ' <a href="#" onclick="_inmForm({descripcion:\'' + _cEsc(p.proveedor) + ' ' + _cEsc(p.numero_factura) + '\',coste:' + (p.base||0) + ',fecha:\'' + _cEsc(p.fecha) + '\',doc:\'' + _cEsc(p.numero_factura) + '\'});return false;">' + t('inm.registrar','registrar') + '</a>'; }).join('<br>') + '</div>';
    body.innerHTML = h;
  } catch(e) { if (body) body.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
function _inmForm(pre){
  var f = document.getElementById('inm-form'); if (!f) return;
  var sel = document.getElementById('inm-cat');
  if (sel && !sel.options.length && _inmCats) sel.innerHTML = Object.keys(_inmCats).map(function(k){ return '<option value="' + k + '">' + _cEsc(_inmCats[k].nombre) + ' (' + _inmCats[k].vida + ' a.)</option>'; }).join('');
  pre = pre || {};
  document.getElementById('inm-desc').value = pre.descripcion || '';
  document.getElementById('inm-coste').value = pre.coste || '';
  document.getElementById('inm-fecha').value = pre.fecha || '';
  document.getElementById('inm-doc').value = pre.doc || '';
  document.getElementById('inm-vida').value = '';
  f.style.display = 'block';
}
async function _inmGuardar(){
  var body = {descripcion: document.getElementById('inm-desc').value, categoria: document.getElementById('inm-cat').value, fecha_alta: document.getElementById('inm-fecha').value, coste: document.getElementById('inm-coste').value, vida_util_anios: document.getElementById('inm-vida').value, documento: document.getElementById('inm-doc').value};
  try {
    var r = await _postJson('/api/inmovilizado/alta', body); var d = await r.json();
    if (d && d.ok) { showNotification(t('inm.ok','✓ Activo dado de alta'), 'success'); document.getElementById('inm-form').style.display='none'; loadInmovilizado(); }
    else showNotification('✗ ' + ((d&&d.error)||'Error'), 'error');
  } catch(e){ showNotification('✗ ' + e.message, 'error'); }
}
async function _inmBaja(id){
  var f = prompt(t('inm.fechaBaja','Fecha de baja (aaaa-mm-dd):'), (document.getElementById('cierre-mes')||{}).value ? (document.getElementById('cierre-mes').value + '-01') : '');
  if (!f) return;
  try { var r = await _postJson('/api/inmovilizado/baja', {id: id, fecha_baja: f}); var d = await r.json(); if (!d || !d.ok) showNotification('✗ ' + ((d&&d.error)||'Error'), 'error'); loadInmovilizado(); } catch(e){ showNotification('✗ ' + e.message, 'error'); }
}

// ── Archivo de fin de mes para la central (Ola B·5) ──────────────────────
async function loadPaquete(){
  var inp = document.getElementById('cierre-mes'); if (!inp) return;
  var mes = inp.value;
  var ex = document.getElementById('paq-excel'); if (ex) ex.href = '/api/exportar/cierre_paquete?mes=' + encodeURIComponent(mes);
  var rs = document.getElementById('paq-resultado'), body = document.getElementById('paq-body'), est = document.getElementById('paq-estado');
  try {
    var r = await fetch('/api/cierre/paquete?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok) { body.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    if (est) { est.textContent = d.listo ? t('paq.listo','✓ listo para la central') : t('paq.pendiente','{n} bloque(s) pendiente(s)').replace('{n}', (d.resumen_checklist||{}).PENDIENTE||0); est.style.color = d.listo ? '#22c55e' : '#f59e0b'; }
    var rr = d.resultado || {};
    var tile = function(l, v, col){ return '<div class="card" style="padding:10px;border-radius:10px"><div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + _cEsc(l) + '</div><div style="font-size:15px;font-weight:700;color:' + (col||'var(--tx)') + '">' + v + '</div></div>'; };
    rs.innerHTML = tile(t('paq.ingresos','Ingresos asentados'), _cEur(rr.ingresos)) + tile(t('paq.gastos','Gastos asentados'), _cEur(rr.gastos)) + tile(t('paq.resultado','Resultado (según documentos)'), _cEur(rr.resultado), (rr.resultado||0) >= 0 ? '#22c55e' : '#f87171') + (rr.drr_rooms_revenue!=null ? tile(t('paq.drr','Rooms Revenue DRR (MTD)'), _cEur(rr.drr_rooms_revenue)) : '');
    var COL = {OK:'#22c55e', PENDIENTE:'#f59e0b', SIN_DATO:'var(--mut)'};
    var LBL = {OK:'✓ OK', PENDIENTE:t('cierre.stPend','pendiente'), SIN_DATO:t('cierre.stSinDato','sin dato')};
    var h = '<div style="margin-bottom:8px"><label style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px">' + t('paq.comentarioGeneral','Comentario general para la central') + '</label><textarea id="paq-com-resumen" rows="2" style="width:100%;box-sizing:border-box;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:6px;border-radius:8px;font-size:12px;font-family:inherit">' + _cEsc(d.comentario_general||'') + '</textarea><button class="btn-ref" style="font-size:11px;margin-top:4px" onclick="_paqGuardar(\'resumen\')">' + t('paq.guardar','Guardar comentario') + '</button></div>';
    h += '<table style="width:100%;border-collapse:collapse;font-size:12px"><tbody>' + (d.checklist||[]).map(function(c){
      return '<tr style="border-top:1px solid var(--s2)"><td style="padding:6px 4px;width:34%"><b>' + _cEsc(c.titulo) + '</b><div style="font-size:11px;color:var(--dim)">' + _cEsc(c.cifra||'') + (c.detalle ? '<br>' + _cEsc(c.detalle) : '') + '</div></td>' +
        '<td style="padding:6px 4px;width:90px;font-weight:700;color:' + (COL[c.estado]||'var(--mut)') + '">' + _cEsc(LBL[c.estado]||c.estado) + '</td>' +
        '<td style="padding:6px 4px"><textarea id="paq-com-' + _cEsc(c.clave) + '" rows="1" placeholder="' + t('paq.comentario','comentario para la central') + '" style="width:100%;box-sizing:border-box;background:var(--bg);border:1px solid var(--s2);color:var(--tx);padding:5px;border-radius:8px;font-size:11px;font-family:inherit" onblur="_paqGuardar(\'' + _cEsc(c.clave) + '\')">' + _cEsc(c.comentario||'') + '</textarea></td></tr>'; }).join('') + '</tbody></table>';
    h += '<div style="margin-top:8px;font-size:11px;color:var(--dim)">' + _cEsc(d.nota||'') + '</div>';
    body.innerHTML = h;
  } catch(e) { if (body) body.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}
async function _paqGuardar(seccion){
  var ta = document.getElementById('paq-com-' + seccion); if (!ta) return;
  var mes = (document.getElementById('cierre-mes')||{}).value || '';
  try {
    var r = await _postJson('/api/cierre/comentario', {mes: mes, seccion: seccion, texto: ta.value});
    var d = await r.json();
    if (d && d.ok) { if (seccion === 'resumen') showNotification(t('paq.guardado','✓ Comentario guardado'), 'success'); }
    else showNotification('✗ ' + ((d&&d.error)||'Error'), 'error');
  } catch(e){ showNotification('✗ ' + e.message, 'error'); }
}

// ── Cierre de mes (Ola B) ─────────────────────────────────────────────────
function _cEsc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function _cEur(v){ return (v==null||isNaN(Number(v))) ? '—' : Number(v).toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}) + ' €'; }
async function loadCierre(forzar){
  var inp = document.getElementById('cierre-mes');
  if (!inp) return;
  if (!inp.value) { var d=new Date(); inp.value = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
  var mes = inp.value;
  var ex = document.getElementById('cierre-excel'); if (ex) ex.href = '/api/exportar/cierre?mes=' + encodeURIComponent(mes);
  var rb = document.getElementById('cierre-recon-body'), mb = document.getElementById('cierre-mayor-body'), db = document.getElementById('cierre-diario-body'), av = document.getElementById('cierre-avisos');
  try { loadCuadreBanco(); } catch(e){}
  try { loadInventarios(); } catch(e){}
  try { loadInmovilizado(); } catch(e){}
  try { loadFiscal(); } catch(e){}
  try { loadPaquete(); } catch(e){}
  try {
    var r = await fetch('/api/cierre/asientos?mes=' + encodeURIComponent(mes));
    var d = await r.json();
    if (!d || !d.ok) { rb.innerHTML = '<div class="empty"><p>' + _cEsc((d&&d.error)||'Error') + '</p></div>'; return; }
    var hs = document.getElementById('cierre-hotel'); if (hs) hs.textContent = d.hotel ? '' : t('cierre.grupo','vista de grupo (todos los hoteles)');
    document.getElementById('cierre-k-asientos').textContent = d.n_asientos;
    document.getElementById('cierre-k-debe').textContent = _cEur(d.debe);
    document.getElementById('cierre-k-haber').textContent = _cEur(d.haber);
    var kc = document.getElementById('cierre-k-cuadre'); kc.textContent = d.cuadra ? t('cierre.cuadra','✓ cuadra') : t('cierre.noCuadra','✗ no cuadra'); kc.style.color = d.cuadra ? '#22c55e' : '#f87171';
    var f = d.fuentes || {};
    var fuentes = t('cierre.fuentes','{ap} facturas AP · {ota} comisiones OTA · {fb} días de TPV · {ar} facturas AR · {cob} cobros · {bk} mov. banco · {pv} provisiones')
      .replace('{ap}', f.ap||0).replace('{ota}', f.ar_ota||0).replace('{fb}', f.ventas_fb||0).replace('{ar}', f.ar_facturas||0).replace('{cob}', f.ar_cobros||0).replace('{bk}', f.banco||0).replace('{pv}', f.provisiones||0);
    av.innerHTML = '<div style="color:var(--dim)">' + _cEsc(fuentes) + '</div>' + (d.avisos||[]).map(function(a){ return '<div>⚠ ' + _cEsc(a) + '</div>'; }).join('');
    // reconciliacion
    var rec = d.reconciliacion || {}; var chk = rec.checks || [];
    var COL = {CUADRA:'#22c55e', DIFERENCIA:'#f87171', PENDIENTE:'#f59e0b', SIN_DATO:'var(--mut)', REVISAR:'#f87171', INFO:'var(--dim)'};
    var LBL = {CUADRA:t('cierre.stCuadra','✓ cuadra'), DIFERENCIA:t('cierre.stDif','⚠ diferencia'), PENDIENTE:t('cierre.stPend','pendiente'), SIN_DATO:t('cierre.stSinDato','sin dato'), REVISAR:t('cierre.stRevisar','revisar'), INFO:'info'};
    rb.innerHTML = (chk.length ? '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="color:var(--dim);text-align:left"><th>' + t('cierre.hCuenta','Cuenta') + '</th><th>' + t('cierre.hConcepto','Comprobación') + '</th><th style="text-align:right">' + t('cierre.hLibro','Libro') + '</th><th style="text-align:right">' + t('cierre.hJust','Justificado') + '</th><th style="text-align:right">' + t('cierre.hDif','Diferencia') + '</th><th></th></tr></thead><tbody>' +
      chk.map(function(c){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:6px 4px;font-weight:700">' + _cEsc(c.cuenta) + '</td><td style="padding:6px 4px">' + _cEsc(c.concepto) + (c.nota ? '<div style="font-size:11px;color:var(--dim)">' + _cEsc(c.nota) + '</div>' : '') + '</td><td style="text-align:right;padding:6px 4px">' + _cEur(c.libro) + '</td><td style="text-align:right;padding:6px 4px">' + (c.justificado==null?'—':_cEur(c.justificado)) + '</td><td style="text-align:right;padding:6px 4px">' + (c.diferencia==null?'—':_cEur(c.diferencia)) + '</td><td style="padding:6px 4px;text-align:right"><span style="font-size:11px;font-weight:700;color:' + (COL[c.estado]||'var(--mut)') + '">' + _cEsc(LBL[c.estado]||c.estado) + '</span></td></tr>'; }).join('') + '</tbody></table></div>' : '<div class="empty"><p>' + t('cierre.vacio','Sin datos en este mes.') + '</p></div>');
    // mayor
    var my = d.mayor || [];
    mb.innerHTML = my.length ? '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="color:var(--dim);text-align:left"><th>' + t('cierre.hCuenta','Cuenta') + '</th><th></th><th style="text-align:right">' + t('cierre.kDebe','Debe') + '</th><th style="text-align:right">' + t('cierre.kHaber','Haber') + '</th><th style="text-align:right">' + t('cierre.hSaldo','Saldo') + '</th></tr></thead><tbody>' +
      my.map(function(m){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:5px 4px;font-weight:700">' + _cEsc(m.cuenta) + '</td><td style="padding:5px 4px">' + _cEsc(m.descripcion) + '</td><td style="text-align:right;padding:5px 4px">' + _cEur(m.debe) + '</td><td style="text-align:right;padding:5px 4px">' + _cEur(m.haber) + '</td><td style="text-align:right;padding:5px 4px;font-weight:700">' + _cEur(m.saldo) + '</td></tr>'; }).join('') + '</tbody></table></div>' : '<div class="empty"><p>' + t('cierre.vacio','Sin datos en este mes.') + '</p></div>';
    // diario
    var as = d.asientos || [];
    db.innerHTML = as.length ? '<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="color:var(--dim);text-align:left"><th>#</th><th>' + t('cierre.hFecha','Fecha') + '</th><th>' + t('cierre.hCuenta','Cuenta') + '</th><th>' + t('cierre.hConceptoD','Concepto') + '</th><th style="text-align:right">' + t('cierre.kDebe','Debe') + '</th><th style="text-align:right">' + t('cierre.kHaber','Haber') + '</th><th>' + t('cierre.hOrigen','Origen') + '</th></tr></thead><tbody>' +
      as.map(function(a){ return '<tr style="border-top:1px solid var(--s2)"><td style="padding:3px 4px">' + a.num + '</td><td style="padding:3px 4px;white-space:nowrap">' + _cEsc(a.fecha) + '</td><td style="padding:3px 4px;white-space:nowrap"><b>' + _cEsc(a.cuenta) + '</b> ' + _cEsc(a.desc_cuenta) + '</td><td style="padding:3px 4px">' + _cEsc(a.concepto) + '</td><td style="text-align:right;padding:3px 4px">' + (a.debe?_cEur(a.debe):'') + '</td><td style="text-align:right;padding:3px 4px">' + (a.haber?_cEur(a.haber):'') + '</td><td style="padding:3px 4px;color:var(--dim)">' + _cEsc(a.origen) + '</td></tr>'; }).join('') + '</tbody></table>' + (d.truncado ? '<div style="color:var(--dim);margin-top:6px">' + t('cierre.truncado','Se muestran las primeras líneas; el Excel lleva todas.') + '</div>' : '') : '<div class="empty"><p>' + t('cierre.vacio','Sin datos en este mes.') + '</p></div>';
  } catch(e) { if (rb) rb.innerHTML = '<div class="empty"><p>' + _cEsc(e.message) + '</p></div>'; }
}

function _cargarPanel(tab, panel, forzar) {
  var fn = _CARGADORES[tab];
  if (!fn) return Promise.resolve();
  if (_panelCargado[tab] && !forzar) return Promise.resolve();
  _panelCargado[tab] = true;
  var p;
  try { p = fn(); } catch (e) { _panelCargado[tab] = false; return Promise.resolve(); }
  var el = panel || document.getElementById('panel-' + tab);
  // los seis cargadores son `async`, asi que devuelven promesa: se traduce
  // JUSTO cuando termina de pintar, no 250 ms despues
  return Promise.resolve(p).then(function(){ _pintarYa(el); })
                           .catch(function(){ _panelCargado[tab] = false; });
}

// Al entrar documentos nuevos hay que volver a poblar: se marca todo como no
// cargado y se repuebla el que este a la vista. El resto, la proxima vez.
//
// Tambien al CAMBIAR DE HOTEL, que tiene exactamente la misma consecuencia:
// todo lo que hay cacheado es de otro hotel.
function _invalidarPaneles() {
  _panelCargado = {};
  // Y las guardas de DENTRO. `_panelCargado` es la de fuera, pero F&B y
  // Multi-Hotel llevan ademas la suya propia. Con la de fuera limpia y la de
  // dentro puesta, el panel se vuelve a "cargar" sin volver a pedir nada: el
  // sintoma es que ensena los numeros del hotel anterior y parece que no pasa
  // nada. Medido en produccion antes de arreglarlo: hasta 40 s con las ventas
  // del otro hotel en pantalla, porque lo unico que lo corregia era el
  // setInterval(loadAll, 60000) — y solo si daba la casualidad de que F&B
  // estaba a la vista en ese tick.
  //
  // Van en try/catch por separado a proposito: si una peta, la otra tiene que
  // limpiarse igual. Juntas, un fallo en F&B dejaria Multi-Hotel sucio.
  try { ['resumen','inventario','mermas','recetas']
          .forEach(function(s){ _fbLoaded[s] = false; }); } catch(e){}
  try { _mh_loaded = false; _mhClasicaLoaded = false; } catch(e){}
  var act = document.querySelector('.panel.active');
  if (act && act.id && act.id.indexOf('panel-') === 0) {
    var t = act.id.slice(6);
    if (_CARGADORES[t]) _cargarPanel(t, act, true);
  }
}
window._invalidarPaneles = _invalidarPaneles;

// Precarga: al arrancar, poblar los seis apartados escalonados y ya traducidos,
// para que al pulsar no haya nada que pintar. Son las MISMAS 9 llamadas que hoy
// se hacen al entrar en cada uno; solo se adelantan. Escalonadas para no
// competir con el arranque en un Render frio.
function _precargarPaneles() {
  var tabs = ['banco', 'drr', 'multi_hotel', 'notif', 'ar_real', 'fb'];
  tabs.forEach(function(t, i) {
    setTimeout(function(){
      try { _cargarPanel(t, document.getElementById('panel-' + t), false); } catch(e){}
    }, 400 + i * 350);
  });
}
if (window.requestIdleCallback) requestIdleCallback(function(){ _precargarPaneles(); }, {timeout: 3000});
else setTimeout(_precargarPaneles, 800);
// ══ F&B COST CONTROL ══════════════════════════════════
function loadFB() { if (typeof cargarFB === 'function') cargarFB(); else if (typeof loadFBCost === 'function') loadFBCost(); }

var _fbLoaded = {resumen:false, inventario:false, mermas:false, recetas:false};
var _fbActive = 'resumen';

async function fbUploadRecetas(input) {
  const file = input.files[0];
  if (!file) return;
  const msg = document.getElementById('fb-upload-msg');
  msg.style.color = 'var(--mut)';
  msg.textContent = 'Subiendo ' + file.name + '...';
  const form = new FormData();
  form.append('file', file);
  try {
    const r = await fetch('/fb/api/upload_recetas', {method: 'POST', body: form,
                          headers: {'X-CSRF-Token': _csrfToken}});
    const d = await r.json();
    if (d.ok) {
      // Los avisos NO son un fallo, pero tampoco se pueden esconder: dicen que
      // parte del escandallo va a contar de menos.
      const hayAvisos = d.avisos && d.avisos.length;
      msg.style.color = hayAvisos ? 'var(--ora)' : 'var(--grn)';
      msg.textContent = (hayAvisos ? '⚠ ' : '✓ ') + d.recetas_importadas +
        ' recetas (' + d.ingredientes + ' ingredientes) — total ' + d.total_recetas +
        (hayAvisos ? ' · ' + d.avisos.join(' · ') : '');
      _fbLoaded.recetas = false; _fbLoaded.resumen = false;
      if (_fbActive === 'recetas') loadFBRecetas();
      else if (_fbActive === 'resumen') loadFBResumen();
    } else {
      msg.style.color = 'var(--red)';
      msg.textContent = '✗ ' + (d.error || 'Error al importar');
    }
  } catch(e) {
    msg.style.color = 'var(--red)';
    msg.textContent = '✗ ' + e.message;
  }
  input.value = '';
}

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
  // Cargar siempre si no está cargado O si el panel está vacío
  var _panel = document.getElementById(panels[sub]);
  var _isEmpty = _panel && (_panel.innerHTML.includes('Cargando') || _panel.innerHTML.trim() === '');
  if (!_fbLoaded[sub] || _isEmpty) {
    if (sub === 'resumen')    loadFBResumen();
    if (sub === 'inventario') loadFBInventario();
    if (sub === 'mermas')     loadFBMermas();
    if (sub === 'recetas')    loadFBRecetas();
  }
}

async function loadFBTab() {
  if (!_fbLoaded.resumen) loadFBResumen();
  // Precargar las otras sub-tabs para que estén listas al hacer clic
  setTimeout(function() {
    if (!_fbLoaded.inventario) loadFBInventario();
  }, 500);
  setTimeout(function() {
    if (!_fbLoaded.mermas) loadFBMermas();
  }, 1000);
  setTimeout(function() {
    if (!_fbLoaded.recetas) loadFBRecetas();
  }, 1500);
}

function _refrescarFB() {
  // Tira las cuatro banderas de "ya cargado" y vuelve a pintar el subtab
  // visible. Sin esto, cambiar de hotel dejaba los numeros del anterior: los
  // subtabs se cargan una vez y no se vuelven a pedir.
  ['resumen','inventario','mermas','recetas'].forEach(function(s){ _fbLoaded[s] = false; });
  var sub = (typeof _fbActive !== 'undefined' && _fbActive) ? _fbActive : 'resumen';
  if (sub === 'inventario' && typeof loadFBInventario === 'function') loadFBInventario();
  else if (sub === 'mermas' && typeof loadFBMermas === 'function') loadFBMermas();
  else if (sub === 'recetas' && typeof loadFBRecetas === 'function') loadFBRecetas();
  else if (typeof loadFBResumen === 'function') loadFBResumen();
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
      cont.innerHTML = _emptyState('🍽️', t('fb.vacioTitulo', 'Aún no hay datos de F&B'), t('fb.vacioSub', 'Sube ventas POS, inventario o mermas y Yve calculará el Food Cost automáticamente.'));
      return;
    }
    const r = data.resumen;
    // Cobertura: sobre qué parte de la facturación está calculado el food cost.
    // Un FC sin cobertura al lado es un número que tranquiliza sin haberse
    // ganado la tranquilidad — las ventas sin escandallo no cuentan para el
    // coste, así que sin este dato el porcentaje parece mejor de lo que es.
    const cob = data.cobertura || r.cobertura || null;
    const fcColor = r.alerta ? 'var(--red)' : (r.fc_real_pct <= r.fc_teorico_pct ? 'var(--grn)' : 'var(--ora)');
    const fcDiff  = (r.fc_real_pct - r.fc_teorico_pct).toFixed(2);
    const fcSign  = fcDiff > 0 ? '+' : '';

    // ── Header: título a la izquierda, botón recalcular a la derecha ──
    let html = '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px;gap:12px">';
    html += '<div><h2 style="font-size:18px;font-weight:700;margin:0">F&B Cost Control</h2>';
    html += '<div style="font-size:12px;color:var(--mut);margin-top:4px">' + (t('fb.datosReales', 'Datos calculados desde ventas reales')) + ' · ' + data.ventas_diarias.fechas.length + ' ' + (t('fb.dias', 'días')) + '</div></div>';
    html += '<button class="btn-ref" onclick="runFB()" style="font-size:12px;flex-shrink:0" data-i18n="btn.recalcular">↺ Recalcular</button>';
    html += '</div>';

    // ── KPIs: 4 cards en fila ──
    html += '<div class="fb-kpi-grid" style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px">';
    html += _fbKpi(t('fb.ventasFb', 'Ventas F&B'), '€' + Math.round(r.total_ventas).toLocaleString('es-ES'), t('fb.periodoCompleto', 'período completo'), 'var(--acc2)');
    html += _fbKpi(t('fb.fcTeorico', 'FC Teórico'), r.fc_teorico_pct + '%', _fbSobre(cob), 'var(--grn)');
    html += _fbKpi(t('fb.fcReal', 'FC Real'), r.fc_real_pct + '%', fcSign + fcDiff + ' ' + (t('fb.vsObjetivo', 'pp vs objetivo')), fcColor);
    html += _fbKpi(t('fb.mermasLabel', 'Mermas'), '€' + r.coste_mermas.toLocaleString('es-ES'), r.alerta ? t('fb.revisar', '⚠ Revisar') : t('fb.bajoControl', 'bajo control'), r.alerta ? 'var(--red)' : 'var(--mut)');
    html += '</div>';
    html += _fbAvisoCobertura(cob);

    // ── Fila: gráfico ventas (izq, ancho) + gauge FC% (der, estrecho) ──
    const maxG = Math.max(r.fc_teorico_pct, r.fc_real_pct) * 1.35;
    html += '<div class="fb-chart-grid" style="display:grid;grid-template-columns:1.5fr 1fr;gap:16px;margin-bottom:16px">';
    html += '<div class="card"><div class="card-title" data-i18n="card.ventasDiarias">Ventas diarias F&B</div>';
    html += '<div style="height:200px;position:relative"><canvas id="fb-ventas-chart"></canvas></div></div>';
    html += '<div class="card"><div class="card-title" style="margin-bottom:16px" data-i18n="fb.gaugeTitle">Food Cost % — Teórico vs Real</div>';
    html += _fcBar(t('fb.gaugeTeorico', 'Teórico'), r.fc_teorico_pct, maxG, 'var(--grn)');
    html += '<div style="height:14px"></div>';
    html += _fcBar(t('fb.gaugeReal', 'Real'),    r.fc_real_pct,    maxG, fcColor);
    html += '</div></div>';

    // ── Fila: categorías (izq) + top platos (der) ──
    html += '<div class="fb-chart-grid" style="display:grid;grid-template-columns:1.2fr 0.8fr;gap:16px">';
    html += '<div class="card"><div class="card-title" data-i18n="card.fcCategoria">Food Cost por Categoría</div>';
    html += '<div class="tbl-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>' + (t('fb.thCategoria', 'Categoría')) + '</th><th style="text-align:right">' + (t('fb.thVentas', 'Ventas')) + '</th><th style="text-align:right">FC%</th><th style="text-align:center">' + (t('fb.thEstado', 'Estado')) + '</th>';
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
    html += '<div class="tbl-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="min-width:0;width:100%"><thead><tr><th>' + (t('fb.thPlato', 'Plato')) + '</th><th style="text-align:right">€</th><th style="text-align:right">FC%</th></tr></thead><tbody>';
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
    if (!data.items || !data.items.length) { cont.innerHTML = _emptyState('📦', t('fb.invVacioTitulo', 'Inventario vacío'), t('fb.sinInventario', 'Sin datos de inventario. Sube un inventario con ⚡ Procesar Archivos.')); return; }

    const alertas = data.items.filter(i => i.alerta);
    let html = '<div class="fb-kpi-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;margin-bottom:20px">';
    html += _fbKpi(t('fb.itemsStock', 'Items en Stock'), data.items.length, t('fb.ingredientes', 'ingredientes'), 'var(--acc2)');
    html += _fbKpi(t('fb.valorInv', 'Valor Inventario'), '€' + data.valor_total.toLocaleString('es-ES'), t('fb.valorActual', 'valoración actual'), 'var(--grn)');
    html += _fbKpi(t('fb.alertasStock', 'Alertas Stock Bajo'), alertas.length, alertas.length > 0 ? 'revisar urgente' : t('fb.todoOk', 'todo OK'), alertas.length > 0 ? 'var(--red)' : 'var(--grn)');
    html += '</div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.stockIngredientes">Stock de Ingredientes</div>';
    html += '<div class="tbl-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="min-width:0;width:100%"><thead><tr>';
    html += '<th>' + (t('fb.thIngrediente', 'Ingrediente')) + '</th><th>' + (t('fb.thCategoria', 'Categoría')) + '</th><th>' + (t('th.proveedor', 'Proveedor')) + '</th>';
    html += '<th style="text-align:right">' + (t('fb.thActual', 'Actual')) + '</th><th style="text-align:right">€/u</th>';
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
    if (!data.mermas || !data.mermas.length) { cont.innerHTML = _emptyState('🗑️', t('fb.sinMermas', 'Sin mermas registradas.'), t('fb.merVacioSub', 'Cuando registres mermas o subas un archivo, aparecerán aquí con su coste y causa.')); return; }

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
    html += '<div class="fb-chart-grid" style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">';
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
    html += '<div class="tbl-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="min-width:0;width:100%"><thead><tr>';
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
  const r = await _postJson('/fb/api/registrar_merma', data);
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
    if (!data.recetas || !data.recetas.length) { cont.innerHTML = _emptyState('📖', t('fb.recVacioTitulo', 'Sin recetas cargadas'), t('fb.sinRecetas', 'Aún no hay recetario. Súbelo con «Importar recetario»: una fila por ingrediente (receta, ingrediente, cantidad, coste) y Yve calcula el food cost de cada plato.')); return; }

    const avg = data.recetas.length ? data.recetas.reduce((a,r)=>a+r.fc_pct,0)/data.recetas.length : 0;
    let html = '<div class="fb-kpi-grid" style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px">';
    html += _fbKpi((t('fb.recetas', 'Recetas activas')), data.recetas.length, t('fb.enCarta', 'recetas en carta'), 'var(--acc2)');
    if (data.avg_fc_pct) html += _fbKpi('FC% Medio', data.avg_fc_pct + '%', 'media del menú', data.avg_fc_pct <= 22 ? 'var(--grn)' : 'var(--ora)');
    if (data.best_margin) html += _fbKpi('Mejor margen', data.best_margin.split(' ').slice(0,2).join(' '), 'menor FC%', 'var(--grn)');
    html += _fbKpi('FC% promedio', avg.toFixed(1) + '%', 'media ponderada', avg < 30 ? 'var(--grn)' : 'var(--ora)');
    html += _fbKpi('Alertas FC alto', data.recetas.filter(r=>r.alerta).length, '>35% FC', 'var(--red)');
    html += '</div>';

    html += '<div class="card"><div class="card-title" data-i18n="card.fichaRecetas">Ficha de Recetas con Coste Teórico</div>';
    html += '<div class="tbl-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="min-width:0;width:100%"><thead><tr>';
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

function _emptyState(emoji, titulo, sub, conCta) {
  return '<div style="text-align:center;padding:44px 20px">' +
    '<div style="font-size:44px;margin-bottom:12px;opacity:.85">' + emoji + '</div>' +
    '<div style="font-size:15px;font-weight:700;color:var(--tx);margin-bottom:6px">' + titulo + '</div>' +
    '<div style="font-size:12.5px;color:var(--mut);max-width:340px;margin:0 auto 18px;line-height:1.6">' + sub + '</div>' +
    (conCta !== false ? '<button class="btn-run" onclick="openUploadModal()" style="margin:0 auto;font-size:13px">' + t('nav.procesar', '⚡ Procesar Archivos') + '</button>' : '') +
    '</div>';
}

function _fbSobre(cob) {
  // Subtítulo del food cost: sobre qué % de la facturación está calculado.
  if (!cob || typeof cob.pct !== 'number') return t('fb.objetivoCalc', 'objetivo calculado');
  if (cob.pct >= 99.95) return t('fb.sobreTodas', 'sobre todas las ventas');
  return t('fb.sobreEl', 'sobre el') + ' ' + cob.pct.toLocaleString('es-ES') + '% ' +
         t('fb.deLasVentas', 'de las ventas');
}

function _fbAvisoCobertura(cob) {
  // Solo aparece si hay ventas que no cruzan con ninguna receta. Decimos
  // cuánto dinero se queda fuera y qué platos son, que es lo accionable.
  if (!cob || !cob.n_platos_sin_receta) return '';
  var eurSin = Math.round(cob.ventas_sin_receta || 0).toLocaleString('es-ES');
  var lista  = (cob.platos_sin_receta || []).join(', ');
  var mas    = cob.n_platos_sin_receta > (cob.platos_sin_receta || []).length
             ? ' (+' + (cob.n_platos_sin_receta - cob.platos_sin_receta.length) + ')' : '';
  return '<div style="background:var(--s1);border:1px solid var(--s2);border-left:3px solid var(--ora);' +
         'border-radius:9px;padding:11px 14px;margin-bottom:16px;font-size:12px;color:var(--mut)">' +
         '<b style="color:var(--ora)">' + cob.n_platos_sin_receta + ' ' +
         (cob.n_platos_sin_receta === 1 ? t('fb.platoSinEscandallo', 'plato sin escandallo')
                                        : t('fb.platosSinEscandallo', 'platos sin escandallo')) +
         '</b> — €' + eurSin + ' ' +
         t('fb.noCuentanFc', 'de ventas que no cuentan para el food cost') + ': ' +
         lista + mas + '</div>';
}

function _fbKpi(lbl, val, sub, color) {
  return '<div class="fb-kpi-card" style="background:var(--s1);border:1px solid var(--s2);border-radius:13px;padding:18px 16px">' +
    '<div class="fb-kpi-lbl" style="font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px;font-weight:600">' + lbl + '</div>' +
    '<div class="fb-kpi-val" style="font-size:24px;font-weight:800;color:' + color + ';line-height:1;letter-spacing:-.5px">' + val + '</div>' +
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
  return new Intl.NumberFormat('es-ES', {minimumFractionDigits:2, maximumFractionDigits:2}).format(v) + ' €';
}

function estadoBadgeAP(est) {
  const m = {
    'MATCH_CORRECTO':'ok','MATCH_3WAY_OK':'ok',
    'DISCREPANCIA_PO':'disc','DISCREPANCIA':'disc',
    'SIN_PO':'sinpo',
    'ALERTA_CONSUMO':'alerta',
    'REVISAR_MANUAL':'manual',
    'NO_REQUIERE_ALBARAN':'sinpo',
    'PENDIENTE':''
  };
  const cls = m[est] || '';
  return `<span class="ap-badge ${cls}">${est || 'PENDIENTE'}</span>`;
}

// ── Provisiones de cierre (Ola A) ─────────────────────────────────────
function _provEsc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function _provFmt(v){ return (Number(v)||0).toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}) + ' €'; }
async function loadProvisiones() {
  const body = document.getElementById('prov-body');
  if (!body) return;
  const mesEl = document.getElementById('prov-mes');
  if (mesEl && !mesEl.value) {
    const d = new Date(); mesEl.value = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
  }
  const mes = mesEl ? mesEl.value : '';
  const dl = document.getElementById('prov-descarga');
  if (dl) dl.href = '/api/exportar/provisiones?mes=' + encodeURIComponent(mes);
  try {
    const d = await (await fetch('/api/provisiones?mes=' + encodeURIComponent(mes), {cache:'no-store'})).json();
    if (!d.ok) { body.innerHTML = '<div class="empty"><p>' + _provEsc(d.error || 'Error') + '</p></div>'; return; }
    const a = d.albaranes, c = d.comisiones;
    const bloque = (titulo, n, total, cuenta, filas, vacio, extra) =>
      '<div style="margin-bottom:14px">' +
        '<div style="display:flex;justify-content:space-between;align-items:baseline;gap:10px;flex-wrap:wrap">' +
          '<strong style="color:var(--tx)">' + titulo + '</strong>' +
          '<span>' + n + ' · <strong style="color:var(--tx)">' + _provFmt(total) + '</strong> · ' + t('prov.cuenta','cuenta') + ' <code>' + _provEsc(cuenta) + '</code></span>' +
        '</div>' +
        (filas.length ? '<div class="tbl-wrap" style="margin-top:6px"><table style="font-size:12px">' + filas.join('') + '</table></div>'
                      : '<div style="font-size:12px;color:var(--dim);margin-top:4px">' + vacio + '</div>') +
        (extra ? '<div style="font-size:11px;color:var(--dim);margin-top:4px">' + extra + '</div>' : '') +
      '</div>';
    const fa = a.por_proveedor.map(p => '<tr><td>' + _provEsc(p.nombre_proveedor) + '</td><td>' + p.n_albaranes + ' ' + t('prov.albaranes','albaranes') + '</td><td><code>' + _provEsc(p.cuenta_gasto) + '</code></td><td style="text-align:right">' + _provFmt(p.importe) + '</td></tr>');
    const fc = c.por_ota.map(p => '<tr><td>' + _provEsc(p.nombre_ota) + '</td><td>' + p.n_facturas + ' ' + t('prov.liquidaciones','liquidaciones') + '</td><td>' + t('prov.facturado','facturado') + ' ' + _provFmt(p.importe_facturado) + '</td><td style="text-align:right">' + _provFmt(p.importe_provision) + '</td></tr>');
    let extraA = '';
    if (a.sin_cruzar) extraA += t('prov.sinCruzar','{n} albaranes sin cruzar todavía: no entran hasta que corra el cruce.').replace('{n}', a.sin_cruzar) + ' ';
    if (a.sin_importe) extraA += t('prov.sinImporte','{n} sin importe legible (cuentan 0).').replace('{n}', a.sin_importe);
    const extraC = c.n ? t('prov.basePactado','{p} de {n} provisionadas por el % pactado; el resto por lo facturado.').replace('{p}', c.n_pactado).replace('{n}', c.n) : '';
    body.innerHTML =
      bloque(t('prov.albSinFactura','Albaranes sin factura'), a.n + ' ' + t('prov.albaranes','albaranes'), a.total, a.cuenta_provision.codigo, fa,
             t('prov.albVacio','Nada que provisionar: todo lo entregado hasta el corte tiene factura.'), extraA) +
      bloque(t('prov.comisiones','Comisiones OTA del mes'), c.n + ' ' + t('prov.liquidaciones','liquidaciones'), c.total, c.cuenta_provision.codigo, fc,
             t('prov.comVacio','Sin liquidaciones OTA con periodo en este mes.'), extraC) +
      '<div style="font-size:11px;color:var(--dim)">' + t('prov.nota','Solo lectura: los asientos van en el Excel del cierre. Las comisiones devengadas sin liquidación necesitan la producción OTA del PMS.') + '</div>';
    if (typeof _pintarYa === 'function') _pintarYa(body);
  } catch(e) { body.innerHTML = '<div class="empty"><p>Error</p></div>'; }
}

// ── Aging AP (Ola A) ─────────────────────────────────────────────────
async function loadAgingAP() {
  const body = document.getElementById('aging-body'), tr = document.getElementById('aging-tramos');
  if (!body || !tr) return;
  try {
    const d = await (await fetch('/api/aging_ap', {cache:'no-store'})).json();
    if (!d.ok) { body.innerHTML = '<div class="empty"><p>' + _provEsc(d.error || 'Error') + '</p></div>'; return; }
    const orden = ['0-30','31-60','61-90','>90','sin fecha'];
    const color = {'0-30':'var(--grn)','31-60':'var(--tx)','61-90':'#f59e0b','>90':'var(--red)','sin fecha':'var(--dim)'};
    tr.innerHTML = orden.filter(k => k !== 'sin fecha' || d.tramos[k] > 0).map(k =>
      '<div class="sc" style="padding:10px"><div class="sc-lbl">' + (k === 'sin fecha' ? t('aging.sinFecha','sin fecha') : k + ' ' + t('aging.dias','días')) + '</div>' +
      '<div class="sc-val" style="font-size:16px;color:' + color[k] + '">' + _provFmt(d.tramos[k]) + '</div></div>').join('') +
      '<div class="sc" style="padding:10px"><div class="sc-lbl">' + t('aging.total','pendiente') + '</div><div class="sc-val" style="font-size:16px">' + _provFmt(d.total) + '</div><div class="sc-sub">' + d.n + ' ' + t('lbl.facturas','facturas') + '</div></div>';
    if (!d.por_acreedor.length) {
      body.innerHTML = '<div style="font-size:12px;color:var(--dim)">' + t('aging.vacio','Nada pendiente de pago.') + '</div>';
      return;
    }
    body.innerHTML = '<div class="tbl-wrap"><table style="font-size:12px"><thead><tr><th>' + t('aging.acreedor','Acreedor') + '</th><th>' + t('aging.masAntigua','Más antigua') + '</th><th>0-30</th><th>31-60</th><th>61-90</th><th>&gt;90</th><th style="text-align:right">' + t('aging.total','pendiente') + '</th></tr></thead><tbody>' +
      d.por_acreedor.map(p => '<tr><td>' + _provEsc(p.acreedor) + ' <span style="color:var(--dim)">· ' + _provEsc(p.origen) + (p.sin_aprobar ? ' · ' + p.sin_aprobar + ' ' + t('aging.sinAprobar','sin aprobar') : '') + '</span></td>' +
        '<td>' + (p.mas_antigua ? _provEsc(p.mas_antigua) + ' <span style="color:' + (p.dias_max > 60 ? 'var(--red)' : 'var(--dim)') + '">(' + p.dias_max + ' ' + t('aging.dias','días') + ')</span>' : '—') + '</td>' +
        ['0-30','31-60','61-90','>90'].map(k => '<td>' + (p[k] ? _provFmt(p[k]) : '—') + '</td>').join('') +
        '<td style="text-align:right"><strong>' + _provFmt(p.importe) + '</strong></td></tr>').join('') +
      '</tbody></table></div><div style="font-size:11px;color:var(--dim);margin-top:6px">' + t('aging.nota','Pendiente = sin conciliar en el extracto bancario; sin extracto, todo cuenta como pendiente.') + '</div>';
    if (typeof _pintarYa === 'function') _pintarYa(body);
  } catch(e) { body.innerHTML = '<div class="empty"><p>Error</p></div>'; }
}

async function loadAP() {
  _skelOn(['ap-total','ap-importe','ap-matches','ap-disc','ap-sinpo','ap-aprobadas']);
  if (_oracleSim === null) _cargarModoOracle();
  loadProvisiones();
  loadAgingAP();
  cargarReclamacionesAP();
  try { cargarAlbaranes(); } catch(e){}
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
    _skelOff(['ap-total','ap-importe','ap-matches','ap-disc','ap-sinpo','ap-aprobadas']);
    setTimeout(() => injectSparklines(AP_SPARKS), 60);

    const tbody = el('ap-tbody');
    if (tbody) tbody.innerHTML = '';
    document.getElementById('ap-count').textContent = facts.length + ' ' + (t('lbl.facturas', 'facturas'));

    facts.forEach(f => {
      const tr = document.createElement('tr');
      tr.setAttribute('data-estado', f.estado || '');
      tr.setAttribute('data-clave', f.clave || '');
      tr.setAttribute('data-accion', f.accion || '');
      tr.style.cursor = 'pointer';
      tr.addEventListener('mouseover', function(){ this.style.background='rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.04)'; });
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

      const dupHtml = (f.duplicados > 1)
        ? ` <span class="ap-badge disc" title="${(f.duplicado_de||'').replace(/"/g,'&quot;')}" style="cursor:help">⚠ ${f.duplicados} documentos con este número</span>`
        : '';
      tr.innerHTML = `
        <td><strong>${f.numero_factura}</strong>${dupHtml}</td>
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
  title.textContent = _tSSE('Pipeline AP — Procesando...');
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
    p.textContent = _tSSE(txt);
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      // Han entrado documentos: lo precargado ya no vale. Sin esto, "no
      // repintar al volver" dejaria Banco, DRR o Multi-Hotel con los numeros
      // de antes de procesar — un parpadeo cosmetico cambiado por datos
      // viejos, que es peor.
      try { if (typeof _invalidarPaneles === 'function') _invalidarPaneles(); } catch(e){}
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = _tSSE(ok ? 'Pipeline AP completado' : 'Pipeline AP con errores');
      btn.disabled = false;
      spin.style.display = 'none';
      lbl.textContent = _tSSE('⚡ Procesar Archivos');
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
    lbl.textContent = _tSSE('⚡ Procesar Archivos');
    btnCl.disabled = false;
  };
}


// ══════════════════════════════════════════════════════════════
// MÓDULO ORACLE — JavaScript
// ══════════════════════════════════════════════════════════════

// BOMBA 2: el boton de Oracle existia como funcion y no como boton. Ahora esta
// en la barra de AP y dice en que modo trabaja. `_oracleSim` lo pone
// `_cargarModoOracle()` desde /api/oracle/status (que lee oracle_auth): en
// simulacion el pipeline genera un Excel de asientos y NO escribe en ningun
// Oracle real, y eso hay que decirlo antes y despues de pulsar.
var _oracleSim = null;
async function _cargarModoOracle() {
  try {
    const d = await (await fetch('/api/oracle/status', {cache: 'no-store'})).json();
    _oracleSim = !!d.simulacion;
  } catch (e) { _oracleSim = true; }
  const chip = document.getElementById('oracle-modo-chip');
  const btn = document.getElementById('btnOracle');
  if (chip) {
    chip.style.display = 'inline-block';
    chip.textContent = _oracleSim ? t('oracle.chipSim', 'simulación · sin Oracle real')
                                  : t('oracle.chipReal', 'Oracle real conectado');
  }
  if (btn) btn.title = _oracleSim
    ? t('oracle.titleSim', 'Modo simulación: genera los asientos en un Excel; no contabiliza en ningún Oracle real')
    : t('oracle.titleReal', 'Contabiliza las facturas aprobadas en Oracle Fusion');
  return _oracleSim;
}
function procesarOracle() {
  // Los ids del modal son los de HOY (`modal-icon`, `modal-title`, `btn-cl`,
  // como en runPipeline): esta funcion se escribio contra un modal viejo
  // (`spinner`, `btnLabel`, `modalIcon`...) y al pulsar el boton nuevo
  // reventaba en la primera linea. Encontrado en produccion, no leyendo.
  const btn   = document.getElementById('btnOracle');
  const log   = document.getElementById('log');
  const icon  = document.getElementById('modal-icon');
  const title = document.getElementById('modal-title');
  const btnCl = document.getElementById('btn-cl');
  const lbl   = { set textContent(v) { if (btn) btn.textContent = v; } };

  if (btn) btn.disabled = true;
  lbl.textContent = '⏳ ' + _tSSE('Contabilizando') + '...';
  log.innerHTML = '';
  btnCl.disabled = true;
  icon.textContent = '🔮';
  title.textContent = _tSSE('Oracle Pipeline — Contabilizando...');
  document.getElementById('overlay').classList.add('on');
  if (_oracleSim !== false) {
    const aviso = document.createElement('p');
    aviso.className = 'l-info';
    aviso.textContent = '⚠ ' + t('oracle.avisoSim', 'Modo simulación: se generan los asientos en un Excel de prueba. No se escribe nada en un Oracle real.');
    log.appendChild(aviso);
  }

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
    p.textContent = _tSSE(txt);
    log.appendChild(p);
    log.scrollTop = log.scrollHeight;

    if (txt === 'PIPELINE_COMPLETO' || txt === 'PIPELINE_CON_ERRORES') {
      // Han entrado documentos: lo precargado ya no vale. Sin esto, "no
      // repintar al volver" dejaria Banco, DRR o Multi-Hotel con los numeros
      // de antes de procesar — un parpadeo cosmetico cambiado por datos
      // viejos, que es peor.
      try { if (typeof _invalidarPaneles === 'function') _invalidarPaneles(); } catch(e){}
      src.close();
      const ok = txt === 'PIPELINE_COMPLETO';
      icon.textContent  = ok ? '✅' : '⚠️';
      title.textContent = (ok && _oracleSim !== false)
        ? t('oracle.simOk', 'Simulación terminada: asientos en Excel, nada contabilizado en Oracle')
        : _tSSE(ok ? 'Oracle: contabilización completada' : 'Oracle: pipeline con errores');
      if (btn) btn.disabled = false;
      lbl.textContent = _tSSE('🔮 Contabilizar en Oracle');
      btnCl.disabled = false;
      setTimeout(loadAP, 800);
    }
  };

  src.onerror = () => {
    src.close();
    const p = document.createElement('p');
    p.className = 'l-err';
    p.textContent = _tSSE('ERROR: conexión con servidor perdida');
    log.appendChild(p);
    if (btn) btn.disabled = false;
    lbl.textContent = _tSSE('🔮 Contabilizar en Oracle');
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
  var panel    = document.getElementById('chat-panel');
  var fab      = document.getElementById('chat-fab');
  var backdrop = document.getElementById('chat-backdrop');
  panel.classList.toggle('open', chatOpen);
  // Mobile: show backdrop + lock body scroll
  if (IS_MOBILE) {
    if (backdrop) backdrop.style.display = chatOpen ? 'block' : 'none';
    document.body.style.overflow = chatOpen ? 'hidden' : '';
  }
  // FAB: hide when open
  if (fab) fab.style.opacity = chatOpen ? '0' : '1';
  if (fab) fab.style.pointerEvents = chatOpen ? 'none' : 'auto';
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
  div.className = `msg-${role}`;
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
  thinkDiv.innerHTML = '<span class="typing" style="display:flex;gap:5px;padding:2px 0"><span class="dot-pulse"></span><span class="dot-pulse"></span><span class="dot-pulse"></span></span>';

  try {
    const resp = await _postJson('/api/chat', { messages: chatHistory });
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
      status.textContent = '✓ ' + file.name + ' · ' + diasStr + ' ' + (t('drr.diasLabel', 'días'));
      // renderDRR reconstruye el panel entero (grupos, budget, gráfico y el
      // Estado diario, que ya marca los OOB por día), así que no hace falta el
      // badge suelto ni recargar el gráfico aparte.
      await renderDRR(data.stats);
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

async function renderDRR(s) {
  const body = document.getElementById('drr-body');
  const statusEl = document.getElementById('drr-status');

  // Sin DRR para este hotel (o error): se vacia el cuerpo ENTERO y se destruye
  // el grafico. `loadDRR` llama a renderDRR SIEMPRE (ya no la protege con
  // `if (data)`), asi que con null el panel queda limpio en vez de heredar el
  // DRR del hotel anterior — Ribera/Faro ensenaban el de Costa Azul. El backend
  // ya filtra bien (drr_del_hotel devuelve None); el agujero era de repintado.
  if (!s || s.error) {
    const _msg = (s && s.error) ? ('Error: ' + s.error) : 'Sin DRR para este hotel.';
    if (body) body.innerHTML = '<div class="drr-metrics" id="drr-metrics" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px">'
      + '<div class="empty"><p>' + _msg + '</p></div></div>';
    if (statusEl) statusEl.textContent = '';
    window._drrChartData = null;
    try { if (_drrChart) { _drrChart.destroy(); _drrChart = null; } } catch(e) {}
    return;
  }
  if (!body) return;

  // Datos del grafico: se piden UNA vez y se comparten. Los usan el importe por
  // dia del Estado diario (abajo) y `renderDRRChart` (que lee este mismo global
  // en vez de volver a pedirlo).
  try { const _r = await fetch('/api/drr_daily_chart'); window._drrChartData = await _r.json(); }
  catch(e) { window._drrChartData = null; }

  const M = s.metricas || {};
  const _num = v => parseFloat(String(v == null ? '' : v).replace(/[^0-9.]/g,'')) || 0;
  const _eur = n => '€' + Math.round(n).toLocaleString('en-US');

  // Una tarjeta KPI: MTD grande (el numero del que se habla) y Hoy/Prev/Ppto en
  // una linea compacta debajo. Los cuatro periodos siguen; cambia la jerarquia.
  function tile(name, key, cls, tip) {
    const m = M[key] || {};
    return '<div class="drr-mc rd-tile ' + (cls || '') + '">'
      + '<div class="mc-name"' + (tip ? ' data-tip="' + tip + '"' : '') + '>' + name + '</div>'
      + '<div class="rd-hero">' + (m.mtd || 'N/D') + '<span class="per">MTD</span></div>'
      + '<div class="rd-mini"><div>Hoy<b>' + (m.today || 'N/D') + '</b></div>'
      + '<div>Prev.<b>' + (m.forecast || 'N/D') + '</b></div>'
      + '<div>Ppto.<b>' + (m.budget || 'N/D') + '</b></div></div></div>';
  }

  // Budget: MES (MTD) vs Presupuesto. Antes era HOY (un dia) / presupuesto del
  // MES entero -> ~3% siempre, sin sentido. Mismos numeros, ratio con sentido.
  const tr = M['Total Revenue'] || {};
  const _bp = Math.round(_num(tr.mtd) / (_num(tr.budget) || 1) * 100);
  const bpct = isFinite(_bp) ? _bp : 0;
  const budCol = bpct >= 95 ? 'var(--grn)' : bpct >= 75 ? 'var(--ora)' : 'var(--red)';
  const budBar = '<div class="drr-mc rd-bud">'
    + '<div style="display:flex;justify-content:space-between;font-size:11px;color:var(--mut);margin-bottom:9px"><span>Revenue del mes (MTD) vs Presupuesto</span><span style="color:' + budCol + ';font-weight:800">' + bpct + '%</span></div>'
    + '<div style="height:8px;border-radius:5px;background:var(--s2);overflow:hidden"><div style="height:100%;width:' + Math.min(100, Math.max(0, bpct)) + '%;background:' + budCol + ';border-radius:5px"></div></div>'
    + '<div style="font-size:10.5px;color:var(--dim);margin-top:8px">' + (tr.mtd || 'N/D') + ' de ' + (tr.budget || 'N/D') + ' presupuestado</div></div>';

  // Procedencia del GOP, en su propia burbuja (no pegada al %). Refleja lo que
  // trae `gop_procedencia`: medido / derivado / estimado / sin datos.
  const _proc = s.gop_procedencia || {};
  const _pv = _proc.mtd || _proc.today || _proc.forecast || 'sin_datos';
  const _provTxt = {medido:'● Medido — lo trae el DRR', derivado:'● Derivado de los datos del hotel', inventado:'● Estimado — sin datos del hotel', sin_datos:'● El DRR no trae el GOP'}[_pv] || '● Medido — lo trae el DRR';
  const _provCol = {medido:'var(--grn)', derivado:'var(--ora)', inventado:'var(--red)', sin_datos:'var(--dim)'}[_pv] || 'var(--grn)';

  // Tres grupos con jerarquia. #drr-metrics (ancla del tour) va en el primero.
  const groups = ''
    + '<div class="rd-group"><div class="rd-glabel">Ocupación y tarifa</div>'
    +   '<div class="rd-grid rate" id="drr-metrics">'
    +     tile('RevPAR', 'Revenue PAR', 'hero', 'Revenue Per Available Room — Rooms Revenue ÷ habitaciones disponibles. También = ADR × ocupación')
    +     tile('Ocupación', 'Occupancy %', '')
    +     tile('ADR', 'ADR', '', 'Average Daily Rate — Rooms Revenue ÷ habitaciones ocupadas')
    +     tile('Hab. ocupadas', 'Rooms Occupied', '', 'El denominador del ADR y el numerador de la ocupación')
    +   '</div></div>'
    + '<div class="rd-group"><div class="rd-glabel">Ingresos</div>'
    +   '<div class="rd-grid rev">'
    +     tile('Total Revenue', 'Total Revenue', 'hero')
    +     tile('Rooms Revenue', 'Rooms Revenue', '', 'El numerador del ADR y del RevPAR')
    +     budBar
    +   '</div></div>'
    + '<div class="rd-group"><div class="rd-glabel">Beneficio</div>'
    +   '<div class="rd-grid gop">'
    +     tile('GOP %', 'GOP %', 'hero grn')
    +     tile('GOP', 'GOP', '', 'Gross Operating Profit — beneficio bruto antes de deuda e impuestos')
    +     '<div class="drr-mc" style="display:flex;align-items:center"><div class="rd-prov" style="color:' + _provCol + '">' + _provTxt + '</div></div>'
    +   '</div></div>';

  const chartCard = '<div class="card" id="drr-chart-card" style="margin-bottom:20px">'
    + '<div class="card-title">Revenue Diario</div>'
    + '<div style="height:215px;position:relative"><canvas id="drr-revenue-chart"></canvas></div>'
    + '<div id="drr-chart-vacio" style="display:none;padding:14px 4px;font-size:12.5px;color:var(--mut);line-height:1.6"></div>'
    + '<div style="font-size:11px;color:var(--mut);margin-top:12px;display:flex;gap:18px;flex-wrap:wrap">'
    +   '<span style="display:inline-flex;align-items:center;gap:6px"><i style="width:14px;height:3px;border-radius:2px;background:var(--acc)"></i>Revenue por día</span>'
    +   '<span style="color:var(--red)">⚠ = día descuadrado</span></div></div>';

  // Estado diario: una tarjeta por dia CON DATOS, con el revenue del dia y, en
  // los descuadrados, el importe del descuadre. El importe sale del revenue
  // (limpio), NO del campo `diff`, que arrastra el "descuadre falso" en los dias
  // que cuadran; `diff` solo se usa para el descuadre real de los OOB.
  const _cd = window._drrChartData || {};
  const _rev = _cd.revenue || [], _cdias = _cd.dias || [];
  const revByDay = {};
  _cdias.forEach((dn, i) => { revByDay[dn] = _rev[i]; });
  const dcards = (s.dias || []).map(d => {
    const rev = revByDay[d.dia];
    const f = String(d.fecha || '').slice(5);
    return '<div class="rd-daycard ' + (d.oob ? 'oob' : 'ok') + '">'
      + '<div class="rd-dc-top">Día ' + d.dia + (f ? ' · ' + f : '') + '</div>'
      + '<div class="rd-dc-amt">' + (rev != null ? _eur(rev) : '—') + '</div>'
      + '<div class="rd-dc-sub">revenue del día</div>'
      + '<div class="rd-dc-st ' + (d.oob ? 'oob' : 'ok') + '">' + (d.oob ? ('⚠ Descuadre ' + _eur(Math.abs(d.diff || 0))) : '✓ Cuadra') + '</div></div>';
  }).join('');
  const trialCard = '<div class="card"><div class="card-title">Estado diario · Trial Balance</div>'
    + '<div class="rd-daywrap">' + (dcards || '<div class="empty"><p>Sin días con datos.</p></div>') + '</div></div>';

  body.innerHTML = groups + chartCard + trialCard;

  // Estado oculto: lo leen la subida de DRR y el paso de onboarding ("...días").
  if (statusEl) statusEl.textContent = (s.archivo || '') + ' · ' + s.total_dias + ' días · ' + s.dias_oob + ' OOB';

  // El grafico lee window._drrChartData (ya cargado arriba).
  renderDRRChart();
}

async function runConciliacion() {
  const btn = document.querySelector('button[onclick="runConciliacion()"]');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Conciliando...'; }
  try {
    const r = await _postJson('/api/conciliar', {});
    const d = await r.json();
    if (d.ok) {
      var _m = '✓ ' + (t('bk.conciliacionOk', 'Conciliación completada') ) + ': ' + d.conciliados + ' OK · ' + d.diferencias + ' dif. · ' + d.pendientes + ' pend.';
      if (d.facturas_disponibles === 0) _m += ' — ' + t('bk.sinFacturas', 'sin facturas AP/AR para cruzar: procesa facturas primero');
      showNotification(_m, d.facturas_disponibles === 0 ? 'info' : 'success');
      loadBanco();
    } else {
      showNotification('✗ Error conciliación: ' + (d.error||''), 'error');
    }
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
  if (btn) { btn.disabled = false; btn.textContent = '⚡ Conciliar'; }
}
async function aprobarMatchOK() {
  // BOMBA 1: el navegador ya NO decide que cuadra ni cuenta nada. Manda las
  // claves de las facturas sin decision y el servidor aprueba solo las que
  // tienen el cruce correcto; el mensaje sale de SUS cifras.
  const rows = document.querySelectorAll('#ap-tbody tr[data-clave]:not([data-accion="APROBADA"]):not([data-accion="RECHAZADA"])');
  const claves = [...rows].map(r => r.getAttribute('data-clave')).filter(Boolean);
  if (!claves.length) { showNotification(t('ap.loteNada', 'No hay facturas pendientes de aprobar'), 'info'); return; }
  showNotification('⏳ ' + t('ap.loteEnCurso', 'Comprobando qué facturas cuadran...'), 'info');
  try {
    const resp = await _postJson('/api/ap/aprobar_lote', {facturas: claves});
    const d = await resp.json();
    if (d.ok) {
      const n = d.aprobadas || 0, p1 = d.primera_firma || 0, esp = d.esperan_segunda || 0;
      const extra = (p1 || esp) ? ' · ' + t('ap.loteFirma', '{p} con primera firma (más de {u} €: falta otra persona)')
          .replace('{p}', p1 + esp).replace('{u}', d.umbral || 500) : '';
      if (n > 0 || p1 > 0) {
        showNotification((n > 0 ? '✓ ' + t('ap.loteOk', '{n} facturas aprobadas (cruce correcto)').replace('{n}', n) : '✍') + extra, n > 0 ? 'success' : 'info');
      } else {
        showNotification(t('ap.loteCero', 'Ninguna factura aprobada: {c} sin cruce correcto, {d} ya decididas')
          .replace('{c}', d.no_cuadran || 0).replace('{d}', d.ya_decididas || 0) + extra, 'info');
      }
      setTimeout(loadAP, 60);
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
  if (countEl) countEl.textContent = visible + ' ' + (t('lbl.facturas', 'facturas'));
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
    // Siempre, tambien con `data` null: renderDRR vacia el panel cuando no hay
    // DRR. Con el viejo `if (data)` el vacio no se repintaba y quedaba a la
    // vista el DRR del hotel anterior. Esa era la fuga de Ribera/Faro.
    await renderDRR(data);
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════════
// NOTIFICACIONES — JavaScript
// ══════════════════════════════════════════════════════════════

const NOTIF_CHANNELS = [
  {key:'email',    icon:'📧', name:'Email'},
  {key:'whatsapp', icon:'💬', name:'WhatsApp', hint:'Requiere cuenta Twilio (TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)'},
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
  {key:'stock_bajo',              labelKey:'notif.evStock', label:'Stock bajo en inventario F&B'},
];
var _notifConfig = null;

async function loadNotifConfig() {
  try {
    const ch = document.getElementById('notif-canales');
    if (ch && !ch.dataset.loaded) ch.innerHTML = skelCards(5, 'grid-template-columns:repeat(5,1fr)');
    const [rCfg, rSmtp] = await Promise.all([
      fetch('/api/notif_config'), fetch('/api/smtp_status')
    ]);
    _notifConfig = await rCfg.json();
    const smtp = await rSmtp.json();
    if (ch) ch.dataset.loaded = '1';

    // Mostrar banner de estado SMTP
    const banner = document.getElementById('notif-smtp-banner');
    if (banner) {
      if (smtp.ok) {
        banner.style.display = 'flex';
        banner.style.display = 'none';  // No mostrar banner verde — el estado se ve en el canal Email como "Activo" 
      } else {
        banner.style.display = 'flex';
        banner.innerHTML = '<div style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.3);border-radius:12px;padding:14px 18px;width:100%">' +
          '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">' +
          '<span style="font-size:20px">⚠️</span>' +
          '<div style="font-size:13px;font-weight:700;color:#f59e0b">SMTP no configurado — el email no funcionará</div></div>' +
          '<div style="font-size:12px;color:var(--mut);line-height:1.7">' +
          'Para activar notificaciones por email, añade estas variables en <strong>Render → grupo Yve → Environment</strong>:<br>' +
          '<code style="background:var(--s2);padding:2px 6px;border-radius:4px;margin:2px 0;display:inline-block">SMTP_SERVER = smtp.gmail.com</code><br>' +
          '<code style="background:var(--s2);padding:2px 6px;border-radius:4px;margin:2px 0;display:inline-block">SMTP_PORT = 587</code><br>' +
          '<code style="background:var(--s2);padding:2px 6px;border-radius:4px;margin:2px 0;display:inline-block">SMTP_USER = tu@gmail.com</code><br>' +
          '<code style="background:var(--s2);padding:2px 6px;border-radius:4px;margin:2px 0;display:inline-block">SMTP_PASSWORD = app_password_gmail</code><br>' +
          '<span style="font-size:11px;color:var(--dim)">💡 Usa una App Password de Gmail, no tu contraseña normal. ' +
          '<a href="https://myaccount.google.com/apppasswords" target="_blank" style="color:var(--acc2)">Crear App Password →</a></span></div></div>';
      }
    }

    // Pre-fill email from hotel config if empty
    if (!_notifConfig.email) {
      try {
        const cfg = await fetch('/api/hotel_config').then(r=>r.json()).catch(()=>({}));
        if (cfg.hotel_email) _notifConfig.email = cfg.hotel_email;
      } catch(e) {}
    }
  } catch(e) {
    _notifConfig = {canales:{email:true,push:true},email:'',whatsapp:'',alertas:{}};
  }
  renderNotifConfig();
}

function renderNotifConfig() {
  const c = _notifConfig;
  // Channels
  const cont = document.getElementById('notif-canales');
  if (cont) {
    cont.innerHTML = NOTIF_CHANNELS.filter(ch => ch.key !== 'push' || yvePushSupported()).map(ch => {
      const on = c.canales && c.canales[ch.key];
      return '<div onclick="toggleNotifCanal(\'' + ch.key + '\')" style="cursor:pointer;background:' +
        (on ? 'rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.1)' : 'var(--s2)') + ';border:1px solid ' +
        (on ? 'var(--acc)' : 'var(--s2)') + ';border-radius:12px;padding:14px;text-align:center;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s">' +
        '<div style="font-size:22px;margin-bottom:6px">' + ch.icon + '</div>' +
        '<div style="font-size:13px;font-weight:600;color:' + (on ? 'var(--acc2)' : 'var(--mut)') + '">' + ch.name + '</div>' +
        '<div style="font-size:10px;color:' + (on ? 'var(--grn)' : 'var(--dim)') + ';margin-top:4px">' + (on ? '● Activo' : '○ Inactivo') + '</div>' +
        '</div>';
    }).join('');
  }
  // Channel fields (only for active channels needing input)
  const fields = document.getElementById('notif-channel-fields');
  if (fields) {
    let html = '';
    if (c.canales && c.canales.email)
      html += notifField('email', 'Email de notificaciones', 'controller@hotel.com', c.email || '');
    if (c.canales && c.canales.whatsapp)
      html += notifField('whatsapp', 'Número WhatsApp destino (+34...)', '+34600123456', c.whatsapp || '') +
              '<div style="font-size:11px;color:var(--dim);margin-top:4px">Necesita TWILIO_ACCOUNT_SID + TWILIO_AUTH_TOKEN + TWILIO_WHATSAPP_FROM en Render</div>';

    if (c.canales && c.canales.slack)
      html += notifField('slack_webhook', 'Slack Webhook URL', 'https://hooks.slack.com/services/...', c.slack_webhook || '');
    if (c.canales && c.canales.push && yvePushSupported()) {
      var permTxt = ('Notification' in window)
        ? (Notification.permission === 'granted' ? '● Permiso concedido en este dispositivo'
           : (Notification.permission === 'denied' ? '⚠ Permiso bloqueado — actívalo en los ajustes del navegador'
           : 'Se pedirá permiso al activar el canal'))
        : 'Este navegador no soporta notificaciones push';
      html += '<div style="background:var(--bg);border:1px solid var(--s2);border-radius:9px;padding:12px 14px">' +
              '<div style="font-size:11px;color:var(--mut);text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px">Notificaciones push</div>' +
              '<div style="font-size:12px;color:var(--dim);margin-bottom:10px">' + permTxt + '</div>' +
              '<button onclick="yvePushTest()" class="btn-ref" style="font-size:12px">🔔 Enviar push de prueba</button></div>';
    }
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
  if (key === 'push') {
    var turningOn = !_notifConfig.canales.push;
    if (turningOn) {
      yvePushSubscribe().then(function(ok){
        _notifConfig.canales.push = !!ok;
        renderNotifConfig();
        if (ok) guardarNotifConfig();
      });
    } else {
      yvePushUnsubscribe();
      _notifConfig.canales.push = false;
      renderNotifConfig();
      guardarNotifConfig();
    }
    return;
  }
  _notifConfig.canales[key] = !_notifConfig.canales[key];
  renderNotifConfig();
}



// Resetear botón guardar al cambiar cualquier input en el panel de notificaciones
document.addEventListener('change', function(e) {
  if (e.target.closest('#panel-notif')) _resetGuardarBtn();
});
document.addEventListener('input', function(e) {
  if (e.target.closest('#panel-notif')) _resetGuardarBtn();
});

function _resetGuardarBtn() {
  var btn = document.getElementById('btn-save-notif');
  if (btn && btn._saved) {
    btn.textContent = '💾 Guardar configuración';
    btn.style.background = '';
    btn.style.borderColor = '';
    btn.style.color = '';
    btn._saved = false;
  }
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
    await _postJson('/api/notif_config', _notifConfig);
    btn.textContent = '✓ Guardado';
    btn.style.background = 'rgba(34,197,94,.15)';
    btn.style.borderColor = 'rgba(34,197,94,.3)';
    btn.style.color = '#22c55e';
    btn._saved = true;
  } catch(e) {
    btn.textContent = '⚠️ Error';
    setTimeout(() => { btn.textContent = '💾 Guardar configuración'; btn.style.cssText='font-size:12px'; }, 2000);
  }
}

async function probarNotif() {
  const btn = document.querySelector('[onclick="probarNotif()"]');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Enviando...'; }
  try {
    const r = await _postJson('/api/test_notif', {});
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
    // Deduplicar: una misma alerta enviada por varios canales cuenta como una sola
    const _seen = new Set();
    const rows = data.slice().reverse().filter(function(n){
      const k = (n.tipo||'') + '|' + (n.asunto||'') + '|' + ((n.fecha||'').slice(0,16));
      if (_seen.has(k)) return false; _seen.add(k); return true;
    });
    count.textContent = rows.length + ' ' + (t('lbl.registros', 'registros'));
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
    const r = await _postJson('/api/test_smtp', {});
    const d = await r.json();
    showNotification(d.ok ? '✓ SMTP funcionando: ' + (d.message||'OK') : '✗ SMTP: ' + (d.error||'Error'), d.ok ? 'success' : 'error');
  } catch(e) { showNotification('✗ Error probando SMTP', 'error'); }
}
async function enviarNotificaciones() {
  const btn = document.getElementById('btn-send-notif');
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  try {
    const r = await _postJson('/api/enviar_notificaciones', {});
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

var _mhView = 'cards';
var _mhClasicaLoaded = false;

function setMHView(view) {
  _mhView = view;
  localStorage.setItem('mh_view', view);
  var resumen = document.getElementById('mh-view-resumen');
  var clasica = document.getElementById('mh-view-clasica');
  var btnC = document.getElementById('mh-view-cards');
  var btnR = document.getElementById('mh-view-ranking');
  if (view === 'ranking') {
    if (resumen) resumen.style.display = 'none';
    if (clasica) clasica.style.display = 'block';
    if (btnC) { btnC.style.background = 'transparent'; btnC.style.color = 'var(--mut)'; }
    if (btnR) { btnR.style.background = 'var(--acc2)'; btnR.style.color = '#fff'; }
    loadMHClasica();
  } else {
    if (resumen) resumen.style.display = 'block';
    if (clasica) clasica.style.display = 'none';
    if (btnC) { btnC.style.background = 'var(--acc2)'; btnC.style.color = '#fff'; }
    if (btnR) { btnR.style.background = 'transparent'; btnR.style.color = 'var(--mut)'; }
  }
}

var _mhGrupoActivo = '';



async function loadMHClasica() {
  // FASE B: ya no pide nada. Antes hacia tres llamadas a /overview, /rankings y
  // /alertas, que leian `kpis_hoteles.xlsx` — el fichero del demo. Ahora la
  // vista de ranking la rellena `renderMHFinancieroClasico` con el MISMO
  // payload del agregador que pinta la vista de resumen, asi que las dos
  // perspectivas no pueden discrepar: es literalmente el mismo objeto.
  if (typeof loadMultiHotel === 'function') return loadMultiHotel();
}







// ── FASE B · Multi-Hotel con datos reales ────────────────────────────────
//
// Hasta aqui el panel leia `kpis_hoteles.xlsx`, cuyo unico escritor era el
// generador de demo. Ahora lee `/api/multi_hotel/agregado`, que sale de los
// documentos de verdad (fase A).
//
// Solo la FILA FINANCIERA: AP, AR/OTA, AR Real y F&B, mas el banco del grupo.
// La fila hotelera —ocupacion, ADR, RevPAR, GOP— sale del DRR y va en su fase.
// Mientras tanto cada tarjeta dice "sin DRR" en vez de enseñar un hueco mudo:
// un dato que falta tiene que decir POR QUE falta, o parece que vale cero.

// FASE F · aqui vivian siete funciones del panel antiguo —renderMHGrupos,
// filtrarMHGrupo, renderMHStatus, renderMHRankings, renderMHAlertasClasica,
// _calSparkline y renderMHTableFull— que quedaron huerfanas al reescribir
// Multi-Hotel en la fase B. No las llamaba nadie.
//
// En `renderMHTableFull` vivia el bug de las estrellas: hacia
// `'★'.repeat(h.stars)` con `h.stars` siendo el TEXTO '4★'. `repeat` convierte
// su argumento a numero, '4★' da NaN, y NaN se convierte a 0: la columna salia
// SIEMPRE vacia. Se arregla borrandolo, no parcheandolo — la categoria ahora
// sale del censo (ver `_mhTarjeta`), que es de donde tenia que haber salido
// siempre en vez de adivinarse del nombre del hotel.

function _mhEur(n) {
  n = Number(n) || 0;
  if (Math.abs(n) >= 10000) return '€' + Math.round(n / 1000) + 'K';
  return '€' + n.toLocaleString('es-ES', {minimumFractionDigits: 0, maximumFractionDigits: 0});
}

function _mhBloque(etiqueta, valor, pie, color) {
  return '<div style="background:var(--bg);border-radius:8px;padding:9px 10px">' +
    '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.3px;margin-bottom:4px">' + etiqueta + '</div>' +
    '<div style="font-size:16px;font-weight:800;line-height:1.15;color:' + (color || 'var(--txt)') + '">' + valor + '</div>' +
    '<div style="font-size:10px;color:var(--dim);margin-top:2px">' + (pie || '&nbsp;') + '</div>' +
  '</div>';
}

// Una tarjeta por caja. `tipo` cambia el borde y el pie, no el contenido:
// hotel / sin_asignar / desconocido enseñan exactamente los mismos numeros.
function _mhTarjeta(f, tipo) {
  var esHotel = (tipo === 'hotel');
  var borde = esHotel ? '' :
    'border-style:dashed;border-color:' + (tipo === 'desconocido' ? 'rgba(239,68,68,.45)' : 'rgba(148,163,184,.4)') + ';';
  var titulo = f.nombre || '(sin nombre)';
  var sub;
  if (tipo === 'sin_asignar') {
    sub = 'Documentos sin hotel · no se reparten';
  } else if (tipo === 'desconocido') {
    sub = 'Etiquetados con un hotel que ya no esta en el censo';
  } else {
    // FASE E: antes ponia "sin DRR" fijo, y desde que la fila hotelera existe
    // eso era mentira en las tarjetas que SI lo tienen. El subtitulo dice el
    // estado de verdad.
    var e = (f.drr || {}).estado;
    // FASE F: la categoria y las habitaciones salen del CENSO. Antes se
    // adivinaban del nombre ('5★' si llevaba un 5), asi que un "Hotel 5 de
    // Mayo" salia de cinco estrellas. Si el censo no las trae, no se ponen:
    // no hay nada que adivinar.
    var c = f.censo || {};
    var ident = [c.categoria, c.habitaciones ? c.habitaciones + ' hab.' : null, c.ciudad]
                  .filter(Boolean).join(' · ');
    sub = (ident ? ident + ' — ' : '') +
          (e === 'con_drr'   ? 'con DRR'
         : e === 'drr_viejo' ? 'DRR de hace ' + ((f.drr || {}).dias_drr || 0) + ' días'
         :                     'sin DRR');
  }
  var fc = Number(f.fb.food_cost_pct) || 0;
  var fcColor = fc === 0 ? 'var(--dim)' : fc > 35 ? '#ef4444' : fc > 30 ? '#f59e0b' : '#22c55e';
  var recl = Number(f.ar_ota.importe_reclamable) || 0;

  return '<div class="card"' + (esHotel ? ' style="padding:18px;cursor:pointer" title="Ver solo este hotel" onclick="seleccionarHotelActivo(\'' + String(f.hotel_id).replace(/'/g, "\\'") + '\', true)"' : ' style="padding:18px;' + borde + '"') + '>' +
    '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px">' +
      '<div>' +
        '<div style="font-size:15px;font-weight:700">' + (tipo === 'hotel' ? '' : (tipo === 'desconocido' ? '⚠ ' : '📄 ')) + titulo + '</div>' +
        '<div style="font-size:11px;color:var(--mut);margin-top:3px">' + sub + '</div>' +
      '</div>' +
    '</div>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">' +
      _mhBloque('Por pagar · AP', _mhEur(f.ap.importe),
                f.ap.facturas + ' fact.' + (f.ap.discrepancias ? ' · ' + f.ap.discrepancias + ' con incidencia' : ''),
                f.ap.discrepancias ? '#f59e0b' : null) +
      // El numero que justifica el producto, y por eso va en verde y con
      // nombre completo: "reclamable" a secas se confunde con el bruto.
      _mhBloque('Reclamable a OTAs', _mhEur(recl),
                f.ar_ota.facturas + ' fact. · bruto ' + _mhEur(f.ar_ota.importe_bruto),
                recl > 0 ? '#22c55e' : null) +
      _mhBloque('Por cobrar · AR Real', _mhEur(f.ar_real.pendiente),
                f.ar_real.facturas + ' fact.' + (f.ar_real.vencido ? ' · vencido ' + _mhEur(f.ar_real.vencido) : ''),
                f.ar_real.vencido ? '#ef4444' : null) +
      _mhBloque('Ventas F&B', _mhEur(f.fb.ventas),
                fc ? 'food cost ' + fc + '%' : 'sin escandallo', fcColor) +
    '</div>' +
    (esHotel ? _mhFilaHotelera(f.drr) : '') +
  '</div>';
}

// FASE E · la fila hotelera de la tarjeta: ocupacion, ADR, RevPAR y GOP.
//
// Tres estados, y los tres se DICEN. Un hueco mudo se lee como un cero, y un
// hotel que no ha subido un papel no es un hotel que va mal.
function _mhFilaHotelera(d) {
  var linea = 'border-top:1px solid var(--s2);padding-top:10px;margin-top:2px;';
  if (!d || d.estado === 'sin_drr') {
    return '<div style="' + linea + 'font-size:11px;color:var(--dim);display:flex;align-items:center;gap:6px">' +
      '<span style="opacity:.7">📊</span>' +
      '<span>Ocupación, ADR, RevPAR y GOP: <b style="color:var(--mut)">falta subir el DRR de este hotel</b></span>' +
    '</div>';
  }
  var celda = function(et, v, suf) {
    return '<div style="text-align:center">' +
      '<div style="font-size:9px;color:var(--mut);text-transform:uppercase;letter-spacing:.3px">' + et + '</div>' +
      '<div style="font-size:14px;font-weight:700">' +
        (v === null || v === undefined ? '<span style="color:var(--dim);font-weight:500">N/D</span>'
                                       : (suf === '%' ? v + '%' : '€' + Math.round(v))) +
      '</div></div>';
  };
  // El GOP lleva su procedencia pegada: si es derivado se dice, y si no hay
  // GOP sale N/D en vez de un numero de relleno (fase D).
  var pg = d.gop_procedencia;
  var notaGop = pg === 'derivado' ? ' <span style="color:#f59e0b">derivado</span>'
              : pg === 'medido'   ? ''
              : ' <span style="color:var(--dim)">el DRR no lo trae</span>';
  var viejo = d.estado === 'drr_viejo'
    ? '<span style="color:#f59e0b">⚠ DRR de hace ' + d.dias_drr + ' días</span>'
    : '<span style="color:var(--dim)">DRR de hace ' + (d.dias_drr || 0) + ' días</span>';

  return '<div style="' + linea + '">' +
    '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:8px">' +
      celda('Ocupación', d.ocupacion_pct, '%') +
      celda('ADR', d.adr) +
      celda('RevPAR', d.revpar) +
      celda('GOP %', d.gop_pct, '%') +
    '</div>' +
    '<div style="font-size:10px;color:var(--dim);display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap">' +
      '<span>' + viejo + (d.dias_oob ? ' · <span style="color:#ef4444">' + d.dias_oob + ' días fuera de balance</span>' : '') + '</span>' +
      '<span>GOP ' + (d.gop === null || d.gop === undefined ? 'N/D' : _mhEur(d.gop)) + notaGop + '</span>' +
    '</div>' +
  '</div>';
}

async function loadMultiHotel() {
  if (_mh_loaded) return;
  try {
    var mhTitle = document.querySelector('#panel-multi_hotel h2');
    var mhSub   = document.querySelector('#panel-multi_hotel h2 + div, #panel-multi_hotel .mh-sub');
    if (mhTitle) mhTitle.textContent = '🌍 Multi-Hotel';
    if (mhSub)   mhSub.textContent   = 'Vista consolidada del grupo · datos reales';

    // Ya no hay selector de mes ni graficos de tendencia: el agregador es una
    // FOTO de ahora, no una serie. Enseñar una tendencia de un solo punto seria
    // inventar los otros cinco meses.
    var selMes = document.getElementById('mh-mes-select');
    if (selMes) selMes.style.display = 'none';
    var trendRow = document.getElementById('mh-trend-row');
    if (trendRow) trendRow.style.display = 'none';

    var r = await fetch('/api/multi_hotel/agregado');
    var data = await r.json().catch(function(){ return {ok:false, error:'Sin datos'}; });
    if (!data.ok) throw new Error(data.error || 'Sin datos');

    var g   = data.grupo;
    var hs  = data.hoteles || [];
    var sa  = data.sin_asignar;
    var des = data.desconocido;
    var hayDesconocido = des && (des.ap.facturas || des.ar_ota.facturas ||
                                 des.ar_real.facturas || des.fb.ventas);
    var haySinAsignar  = sa && (sa.ap.facturas || sa.ar_ota.facturas ||
                                sa.ar_real.facturas || sa.fb.ventas);

    if (!hs.length && !haySinAsignar && !hayDesconocido) throw new Error('Sin hoteles');

    // ── Cabecera: los cuatro numeros del grupo ────────────────────────────
    var kEl = document.getElementById('mh-kpis');
    if (kEl) kEl.innerHTML =
      '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(200px,45%),1fr));gap:10px">' +
      [
        {l:'POR PAGAR · AP', v:_mhEur(g.ap.importe), c:'var(--acc2)',
         s:g.ap.facturas + ' facturas · ' + g.ap.discrepancias + ' con incidencia'},
        {l:'RECLAMABLE A OTAs', v:_mhEur(g.ar_ota.importe_reclamable), c:'#22c55e',
         s:g.ar_ota.discrepancias + ' discrepancias de ' + g.ar_ota.facturas + ' facturas'},
        {l:'POR COBRAR · AR REAL', v:_mhEur(g.ar_real.pendiente),
         c: g.ar_real.vencido ? '#ef4444' : '#f1f5f9',
         s: g.ar_real.vencido ? 'vencido ' + _mhEur(g.ar_real.vencido) : g.ar_real.facturas + ' facturas'},
        {l:'VENTAS F&B', v:_mhEur(g.fb.ventas), c:'#a78bfa',
         // Ponderado, no la media de los hoteles: cada hotel compra a su
         // precio y el inventario del grupo aplanaria el coste al del ultimo.
         s: g.fb.food_cost_pct ? 'food cost ' + g.fb.food_cost_pct + '% (ponderado)' : 'sin escandallo'},
      ].map(function(c) {
        return '<div class="sc">' +
          '<div class="sc-lbl" style="font-size:9px;letter-spacing:.5px">' + c.l + '</div>' +
          '<div class="sc-val" style="color:' + c.c + ';font-size:clamp(20px,4vw,32px);font-weight:900;line-height:1.1;margin:4px 0">' + c.v + '</div>' +
          '<div class="sc-sub" style="font-size:10px;color:var(--dim)">' + c.s + '</div>' +
        '</div>';
      }).join('') + '</div>';

    // ── Banco, fila hotelera del grupo, y cuadre ─────────────────────────
    var iEl = document.getElementById('mh-insights');
    if (iEl) {
      var trozos = [];
      var b = data.banco || {};

      // FASE E · las medias del grupo, PONDERADAS y con el denominador dicho.
      //
      // "Ocupación del grupo 78%" a secas es una trampa si en realidad es "de
      // los 2 hoteles que subieron el DRR". El "sobre X de N" es lo que hace
      // que el número sea defendible delante de un cliente.
      //
      // Y son DOS denominadores, no uno: un hotel puede tener DRR (y contar
      // para la ocupación) y no traer un GOP agregable.
      var h = data.hotelero || {};
      if (h.con_datos) {
        var ce = function(et, v, suf) {
          return '<div><div style="font-size:9px;color:var(--mut);text-transform:uppercase">' + et + '</div>' +
            '<div style="font-size:18px;font-weight:800">' +
            (v === null || v === undefined ? '<span style="color:var(--dim);font-size:14px">N/D</span>'
                                           : (suf === '%' ? v + '%' : '€' + Math.round(v))) +
            '</div></div>';
        };
        trozos.push(
          '<div class="card" style="border-left:3px solid #22c55e;padding:13px 16px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px">' +
            '<div>' +
              '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#22c55e;text-transform:uppercase">🛏 Del DRR — medias ponderadas por tamaño</div>' +
              '<div style="font-size:11px;color:var(--dim);margin-top:3px">' +
                'Sobre <b style="color:var(--mut)">' + h.con_datos + ' de ' + h.n_hoteles + '</b> hoteles' +
                (h.sin_drr ? ' · ' + h.sin_drr + ' sin DRR' : '') +
                (h.viejos ? ' · <span style="color:#f59e0b">' + h.viejos + ' con DRR viejo</span>' : '') +
                (h.gop_sobre !== h.con_datos ? ' · GOP sobre ' + h.gop_sobre : '') +
                (h.dias_oob ? ' · <span style="color:#ef4444">' + h.dias_oob + ' días fuera de balance</span>' : '') +
              '</div>' +
            '</div>' +
            '<div style="display:flex;gap:20px;flex-wrap:wrap">' +
              ce('Ocupación', h.ocupacion_pct, '%') + ce('ADR', h.adr) +
              ce('RevPAR', h.revpar) + ce('GOP %', h.gop_pct, '%') +
            '</div>' +
          '</div>');
      } else if (h.n_hoteles) {
        trozos.push(
          '<div class="card" style="border-left:3px solid var(--s2);padding:13px 16px">' +
            '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:var(--mut);text-transform:uppercase">🛏 Del DRR</div>' +
            '<div style="font-size:12px;color:var(--dim);margin-top:4px">Ningún hotel ha subido su DRR todavía. ' +
            'Ocupación, ADR, RevPAR y GOP llegan con él — no se estiman.</div>' +
          '</div>');
      }
      if (b.hay_datos) {
        // Separado de las tarjetas y etiquetado "del grupo" a proposito: el
        // extracto es de la cuenta de la sociedad, no del hotel. Si estuviera
        // al lado de las tarjetas, se sumaria mentalmente.
        trozos.push(
          '<div class="card" style="border-left:3px solid #60a5fa;padding:13px 16px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">' +
            '<div>' +
              '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#60a5fa;text-transform:uppercase">🏦 Banco — del grupo, no por hotel</div>' +
              '<div style="font-size:11px;color:var(--dim);margin-top:3px">El extracto es de la cuenta de la sociedad. Repartirlo entre hoteles seria inventar.</div>' +
            '</div>' +
            '<div style="display:flex;gap:18px;flex-wrap:wrap">' +
              '<div><div style="font-size:9px;color:var(--mut);text-transform:uppercase">Sin conciliar</div>' +
              '<div style="font-size:18px;font-weight:800;color:' + (b.pendientes ? '#f59e0b' : '#22c55e') + '">' + _mhEur(b.importe_pendiente) + '</div></div>' +
              '<div><div style="font-size:9px;color:var(--mut);text-transform:uppercase">Movimientos</div>' +
              '<div style="font-size:18px;font-weight:800">' + b.conciliados + '/' + b.total + '</div></div>' +
            '</div>' +
          '</div>');
      }
      if (!data.cuadra) {
        // Un descuadre se DICE. Enseñar el total y callar es como se pierden
        // filas sin que nadie se entere.
        var malas = (data.cuadre || []).filter(function(c){ return !c.cuadra; })
                      .map(function(c){ return c.metrica + ' (' + c.diferencia + ')'; }).join(', ');
        trozos.push(
          '<div class="card" style="border-left:3px solid #ef4444;padding:13px 16px">' +
            '<div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#ef4444;text-transform:uppercase">⚠ La suma de los hoteles no cuadra con el total</div>' +
            '<div style="font-size:12px;color:var(--mut);margin-top:4px">' + malas + '</div>' +
          '</div>');
      }
      iEl.innerHTML = trozos.join('');
      iEl.style.display = trozos.length ? '' : 'none';
    }

    // ── Las tarjetas ─────────────────────────────────────────────────────
    var cardsEl = document.getElementById('mh-hotel-cards');
    if (cardsEl) {
      var trozos = hs.map(function(f){ return _mhTarjeta(f, 'hotel'); });
      // Las dos especiales van AL FINAL y con borde discontinuo, para que no
      // se lean como un hotel mas. Y solo si tienen algo: una caja vacia de
      // "sin asignar" es ruido.
      if (haySinAsignar)  trozos.push(_mhTarjeta(sa,  'sin_asignar'));
      if (hayDesconocido) trozos.push(_mhTarjeta(des, 'desconocido'));
      cardsEl.innerHTML = trozos.join('');
    }

    // ── Vista ranking, del mismo sitio ───────────────────────────────────
    renderMHFinancieroClasico(data);

    _mh_loaded = true;
    var savedView = localStorage.getItem('mh_view');
    if (savedView === 'ranking') setMHView('ranking');
    if (_i18nLang && _i18nLang !== 'es') applyI18n(_i18nData);
  } catch(e) {
    var _msg = String(e && e.message || e);
    if (_msg.indexOf('Sin datos') === -1 && _msg.indexOf('Sin hoteles') === -1) console.error('MH Error:', e);
    var el = document.getElementById('mh-kpis');
    var esEstadoVacio = _msg.indexOf('Sin hoteles') >= 0 || _msg.indexOf('Sin datos') >= 0;
    if (el) {
      if (esEstadoVacio) {
        el.innerHTML = _emptyState('🏨', 'Todavia no hay nada que consolidar',
          'Da de alta los hoteles del grupo y sube sus documentos: aqui veras lo que hay por pagar, lo reclamable a las OTAs, lo pendiente de cobro y las ventas de F&B, hotel por hotel.', false);
      } else {
        el.innerHTML = '<div style="color:#ef4444;padding:20px;font-size:13px">⚠ Error cargando datos: ' + _msg + '</div>';
      }
    }
    var cardsEl2 = document.getElementById('mh-hotel-cards');
    if (cardsEl2) cardsEl2.innerHTML = '';
    var insEl2 = document.getElementById('mh-insights');
    if (insEl2) insEl2.innerHTML = '';
  }
}

// La vista "Ranking" sale del MISMO payload, no de otra llamada. Antes pedia
// /rankings y /alertas, que leian el fichero de demo.
function renderMHFinancieroClasico(data) {
  var hs = (data.hoteles || []).slice();
  var g  = data.grupo;

  // Estado: un hotel esta "al dia" si no tiene incidencias de AP, ni facturas
  // vencidas, ni discrepancias de OTA sin resolver.
  var incidencias = function(f) {
    return (f.ap.discrepancias || 0) + (f.ap.revisar || 0) +
           (f.ar_ota.discrepancias || 0) + (f.ar_real.vencido ? 1 : 0);
  };
  var st = document.getElementById('mh-status');
  if (st) {
    var ok = hs.filter(function(f){ return incidencias(f) === 0; }).length;
    var wa = hs.filter(function(f){ var n = incidencias(f); return n >= 1 && n <= 2; }).length;
    var cr = hs.filter(function(f){ return incidencias(f) > 2; }).length;
    var caja = function(color, icono, n, etiqueta) {
      return '<div style="background:rgba(' + color + ',.08);border:1px solid rgba(' + color + ',.25);border-radius:12px;padding:16px;display:flex;align-items:center;gap:12px">' +
        '<span style="font-size:26px">' + icono + '</span>' +
        '<div><div style="font-size:24px;font-weight:800;color:rgb(' + color + ')">' + n + '</div>' +
        '<div style="font-size:12px;color:var(--mut)">' + etiqueta + '</div></div></div>';
    };
    st.innerHTML = caja('34,197,94', '✅', ok, 'Hoteles al dia') +
                   caja('245,158,11', '⚠️', wa, 'Con incidencias') +
                   caja('239,68,68', '🚨', cr, 'Criticos');
  }

  // Top: por dinero reclamable a las OTAs, que es lo accionable de verdad.
  var rk = document.getElementById('mh-rankings');
  if (rk) {
    var orden = hs.slice().sort(function(a,b){
      return (b.ar_ota.importe_reclamable || 0) - (a.ar_ota.importe_reclamable || 0); });
    var conAlgo = orden.filter(function(f){ return (f.ar_ota.importe_reclamable || 0) > 0; });
    var titulo = document.querySelector('#mh-view-clasica .card-title');
    if (titulo) titulo.textContent = '💰 Mas reclamable a las OTAs';
    rk.innerHTML = conAlgo.length ? conAlgo.map(function(f, i) {
      var medalla = i === 0 ? '#FFD700' : i === 1 ? '#C0C0C0' : i === 2 ? '#CD7F32' : 'var(--dim)';
      return '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--s2)">' +
        '<div style="display:flex;align-items:center;gap:12px">' +
          '<span style="font-size:17px;font-weight:800;color:' + medalla + ';min-width:20px">' + (i+1) + '</span>' +
          '<div style="font-weight:600;font-size:13px">' + f.nombre + '</div></div>' +
        '<div style="font-weight:700;color:#22c55e">' + _mhEur(f.ar_ota.importe_reclamable) + '</div></div>';
    }).join('') : '<div style="color:var(--dim);font-size:13px;padding:8px">Nada reclamable ahora mismo</div>';
  }

  var al = document.getElementById('mh-alertas');
  if (al) {
    var avisos = [];
    hs.forEach(function(f) {
      if (f.ap.discrepancias) avisos.push({h:f.nombre, m:f.ap.discrepancias + ' facturas AP con discrepancia'});
      if (f.ap.revisar)       avisos.push({h:f.nombre, m:f.ap.revisar + ' facturas AP a revisar a mano'});
      if (f.ar_ota.discrepancias) avisos.push({h:f.nombre, m:f.ar_ota.discrepancias + ' comisiones de OTA a reclamar'});
      if (f.ar_ota.di_pendientes) avisos.push({h:f.nombre, m:f.ar_ota.di_pendientes + ' certificados de doble imposicion pendientes'});
      if (f.ar_real.vencido)  avisos.push({h:f.nombre, m:'cobro vencido: ' + _mhEur(f.ar_real.vencido)});
    });
    al.innerHTML = avisos.length ? avisos.slice(0,8).map(function(a) {
      return '<div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid var(--s2)">' +
        '<span style="color:#f59e0b">▲</span>' +
        '<div style="font-size:12px"><span style="font-weight:600">' + a.h + '</span> ' +
        '<span style="color:var(--mut)">' + a.m + '</span></div></div>';
    }).join('') : '<div style="color:#22c55e;font-size:13px;padding:8px">✓ Sin incidencias</div>';
  }

  // La tabla, con las columnas que Yve SI sabe.
  var thead = document.querySelector('#mh-view-clasica table thead tr');
  if (thead) thead.innerHTML =
    '<th>Hotel</th><th style="text-align:right">AP €</th><th style="text-align:right">Fact. AP</th>' +
    '<th style="text-align:right">Reclamable OTAs</th><th style="text-align:right">Por cobrar</th>' +
    '<th style="text-align:right">Ventas F&B</th><th style="text-align:right">Food cost</th>' +
    '<th style="text-align:center">Estado</th>';

  var tb = document.getElementById('mh-tbody-full');
  if (tb) {
    var filas = hs.map(function(f){ return {f:f, tipo:'hotel'}; });
    if (data.sin_asignar && (data.sin_asignar.ap.facturas || data.sin_asignar.ar_ota.facturas ||
        data.sin_asignar.ar_real.facturas || data.sin_asignar.fb.ventas))
      filas.push({f:data.sin_asignar, tipo:'especial'});
    if (data.desconocido && (data.desconocido.ap.facturas || data.desconocido.ar_ota.facturas ||
        data.desconocido.ar_real.facturas || data.desconocido.fb.ventas))
      filas.push({f:data.desconocido, tipo:'especial'});

    tb.innerHTML = filas.map(function(x) {
      var f = x.f, esp = x.tipo === 'especial';
      var n = incidencias(f);
      var color = n === 0 ? '#22c55e' : n <= 2 ? '#f59e0b' : '#ef4444';
      var icono = n === 0 ? '●' : n <= 2 ? '▲' : '■';
      return '<tr style="' + (esp ? 'opacity:.75;font-style:italic' : 'cursor:pointer') + '"' +
        (esp ? '' : ' onclick="seleccionarHotelActivo(\'' + String(f.hotel_id).replace(/'/g, "\\'") + '\', true)"') + '>' +
        '<td style="font-weight:600">' + f.nombre + '</td>' +
        '<td style="text-align:right">' + _mhEur(f.ap.importe) + '</td>' +
        '<td style="text-align:right">' + f.ap.facturas + '</td>' +
        '<td style="text-align:right;font-weight:600;color:#22c55e">' + _mhEur(f.ar_ota.importe_reclamable) + '</td>' +
        '<td style="text-align:right">' + _mhEur(f.ar_real.pendiente) + '</td>' +
        '<td style="text-align:right">' + _mhEur(f.fb.ventas) + '</td>' +
        '<td style="text-align:right">' + (f.fb.food_cost_pct ? f.fb.food_cost_pct + '%' : '—') + '</td>' +
        '<td style="text-align:center;color:' + color + '">' + (esp ? '' : icono) + '</td></tr>';
    }).join('') +
    // El total, del df entero. Va en la tabla para que se vea al lado de las
    // partes: si no cuadrase, se nota aqui antes que en ningun sitio.
    '<tr style="border-top:2px solid var(--s2);font-weight:800">' +
      '<td>GRUPO</td>' +
      '<td style="text-align:right">' + _mhEur(g.ap.importe) + '</td>' +
      '<td style="text-align:right">' + g.ap.facturas + '</td>' +
      '<td style="text-align:right;color:#22c55e">' + _mhEur(g.ar_ota.importe_reclamable) + '</td>' +
      '<td style="text-align:right">' + _mhEur(g.ar_real.pendiente) + '</td>' +
      '<td style="text-align:right">' + _mhEur(g.fb.ventas) + '</td>' +
      '<td style="text-align:right">' + (g.fb.food_cost_pct ? g.fb.food_cost_pct + '%' : '—') + '</td>' +
      '<td></td></tr>';
  }
}


function loadNotificaciones() { if (typeof cargarNotificaciones === 'function') cargarNotificaciones(); else if (typeof loadNotif === 'function') loadNotif(); }

// ── Banco: cómo funciona (grupo vs por hotel), elegido por el usuario ──────
async function _cargarConfigBanco() {
  try {
    const d = await fetch('/api/config_banco').then(r => r.json());
    window._bancoModo = (d && d.elegido) ? d.modo : '';
  } catch(e) { window._bancoModo = ''; }
  return window._bancoModo;
}
function abrirModoBanco() {
  var c = document.getElementById('banco-modal-cancelar'); if (c) c.style.display = 'inline';
  var m = document.getElementById('modal-banco-config'); if (m) m.style.display = 'flex';
}
function cerrarModoBanco() {
  var m = document.getElementById('modal-banco-config'); if (m) m.style.display = 'none';
}
// La PRIMERA vez (sin elección) el modal es obligatorio: sin botón de cancelar.
async function _checkBancoConfig() {
  var modo = window._bancoModo;
  if (modo === undefined) modo = await _cargarConfigBanco();
  if (!modo) {
    var c = document.getElementById('banco-modal-cancelar'); if (c) c.style.display = 'none';
    var m = document.getElementById('modal-banco-config'); if (m) m.style.display = 'flex';
  }
}
async function elegirModoBanco(modo) {
  try {
    const r = await _postJson('/api/config_banco', { modo: modo });
    const d = await r.json();
    if (!d || !d.ok) { showNotification('✗ ' + ((d && d.error) || 'no se pudo guardar'), 'error'); return; }
    window._bancoModo = d.modo;
    cerrarModoBanco();
    showNotification(d.modo === 'grupo' ? '🏛️ Banco del grupo' : '🏨 Banco por hotel', 'info');
    loadBanco();
  } catch(e) { showNotification('✗ ' + e.message, 'error'); }
}

async function loadBanco() {
  // Etiqueta de modo (si ya está elegido). El modal de primera vez lo dispara
  // switchTab al ABRIR la pestaña, no este cargador de datos.
  var modo = window._bancoModo;
  if (modo === undefined) modo = await _cargarConfigBanco();
  var chip = document.getElementById('banco-modo-chip');
  var camb = document.getElementById('banco-modo-cambiar');
  if (modo) {
    if (chip) { chip.style.display = 'inline-block'; chip.textContent = modo === 'grupo' ? '🏛️ Banco del grupo' : '🏨 Banco por hotel'; }
    if (camb) camb.style.display = 'inline';
  } else {
    if (chip) chip.style.display = 'none';
    if (camb) camb.style.display = 'none';
  }
  try {
    var r = await fetch('/api/stats_banco');
    var d = await r.json();
    if (!d) {
      ['bk-total','bk-conc','bk-pend','bk-diff'].forEach(function(id){ var e=document.getElementById(id); if(e) e.textContent='0'; });
      var ip = document.getElementById('bk-imp-pend'); if (ip) ip.textContent = '—';
      var ba0 = document.getElementById('bk-alertas'); if (ba0) ba0.innerHTML = '<div class="empty"><p>Sin movimientos bancarios.</p></div>';
      var pb0 = document.getElementById('banco-progress-bar'); if (pb0) pb0.style.display = 'none';
      return;
    }
    document.getElementById('bk-total').textContent = d.total || '0';
    document.getElementById('bk-conc').textContent = d.conciliados || '0';
    var _bT = d.total||0, _bC = d.conciliados||0, _pEl = document.getElementById('banco-progress-bar');
    if (_pEl && _bT > 0) { var _pct = Math.round(_bC/_bT*100), _col = _pct>=80?'var(--grn)':_pct>=50?'var(--ora)':'var(--red)'; _pEl.style.display='block'; _pEl.innerHTML = '<div style="display:flex;align-items:center;gap:12px;font-size:12px"><span style="color:var(--mut);white-space:nowrap">Conciliado:</span><div style="flex:1;background:var(--s2);border-radius:4px;height:8px;overflow:hidden"><div style="height:100%;border-radius:4px;background:'+_col+';width:'+_pct+'%;transition:width .6s ease"></div></div><span style="color:'+_col+';font-weight:700;min-width:60px">'+_bC+'/'+_bT+' ('+_pct+'%)</span></div>'; }
    else if (_pEl) { _pEl.style.display='none'; }
    document.getElementById('bk-pend').textContent = d.pendientes || '0';
    document.getElementById('bk-diff').textContent = d.diferencias || '0';
    document.getElementById('bk-imp-pend').textContent = d.importe_pendiente ? eur(d.importe_pendiente) + ' pend.' : '—';

    var el = document.getElementById('bk-alertas');
    var _html = (d.alertas && d.alertas.length)
      ? d.alertas.map(function(a) { return '<div class="act-item"><div class="adot r"></div><div class="atxt"><b>' + a.dias + ' ' + t('bk.dias', 'días') + '</b> ' + t('bk.sinConciliar', 'sin conciliar:') + ' ' + a.concepto + ' — ' + eur(a.importe) + '</div></div>'; }).join('')
      : '<div class="empty"><p>Sin alertas bancarias pendientes.</p></div>';
    // Modo por hotel: lo que no está asignado a ningún hotel NO se esconde, se avisa.
    if (modo === 'por_hotel' && d.sin_asignar) {
      _html = '<div class="act-item"><div class="adot" style="background:var(--ora)"></div><div class="atxt"><b>' + d.sin_asignar + '</b> movimiento(s) sin asignar a un hotel — súbelos dentro del hotel que corresponda</div></div>' + _html;
    }
    el.innerHTML = _html;
  } catch(e) {
    console.warn('Error banco:', e);
    var el2 = document.getElementById('bk-alertas');
    if (el2 && el2.innerHTML.includes('—')) el2.innerHTML = '<div class="empty"><p>Sin alertas bancarias.</p></div>';
    ['bk-total','bk-conc','bk-pend','bk-diff'].forEach(function(id){ var e=document.getElementById(id); if (e && e.textContent === '—') e.textContent = '0'; });
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
// ── Alta de un cliente de credito ────────────────────────────────────
// Sin clientes no hay limite ni aviso de riesgo, y no habia forma de meter
// ninguno: el generador de demo era el unico que escribia ese fichero.
function abrirNuevoCliente() {
  var v = document.getElementById('nuevo-cliente-modal');
  if (v) v.remove();
  var m = document.createElement('div');
  m.id = 'nuevo-cliente-modal';
  m.style.cssText = 'position:fixed;inset:0;z-index:9500;background:rgba(0,0,0,.7);display:flex;align-items:center;justify-content:center;padding:16px';
  var campo = function(id, etiq, tipo, ph) {
    return '<label style="display:block;margin-bottom:10px">' +
      '<span style="display:block;font-size:11px;color:var(--mut);margin-bottom:4px">' + etiq + '</span>' +
      '<input id="' + id + '" type="' + tipo + '" placeholder="' + (ph || '') + '" ' +
      'style="width:100%;box-sizing:border-box;background:var(--bg);border:1px solid var(--s3);color:var(--tx);' +
      'padding:11px 12px;border-radius:9px;font-size:14px;outline:none"></label>';
  };
  m.innerHTML = '<div style="background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:22px;width:min(420px,100%);max-height:88vh;overflow:auto">' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">' +
      '<h3 style="margin:0;font-size:16px;font-weight:800">➕ ' + t('ar.nuevoCliente', 'Nuevo cliente de crédito') + '</h3>' +
      '<button onclick="cerrarNuevoCliente()" style="background:none;border:none;color:var(--mut);font-size:22px;cursor:pointer;min-width:44px;min-height:44px">✕</button>' +
    '</div>' +
    campo('ncl-nombre', t('ar.clNombre', 'Nombre del cliente') + ' *', 'text', 'Viajes Meridiano S.A.') +
    campo('ncl-nif', t('ar.clNif', 'NIF / CIF'), 'text', 'A28004556') +
    campo('ncl-limite', t('ar.clLimite', 'Límite de crédito (€)') + ' *', 'number', '25000') +
    campo('ncl-dias', t('ar.clDias', 'Días de pago'), 'number', '30') +
    campo('ncl-email', t('ar.clEmail', 'Email'), 'email', 'cuentas@cliente.com') +
    campo('ncl-tel', t('ar.clTel', 'Teléfono'), 'text', '') +
    '<div id="ncl-err" style="display:none;font-size:12px;color:var(--red);margin:4px 0 10px"></div>' +
    '<div style="display:flex;gap:10px;justify-content:flex-end;margin-top:6px">' +
      '<button onclick="cerrarNuevoCliente()" class="btn-ref" style="min-height:44px">' + t('js.cancelar', 'Cancelar') + '</button>' +
      '<button id="ncl-ok" onclick="guardarNuevoCliente()" style="background:var(--acc);border:none;color:#fff;padding:12px 20px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;min-height:44px">' + t('js.guardar', 'Guardar') + '</button>' +
    '</div></div>';
  m.addEventListener('click', function(e) { if (e.target === m) cerrarNuevoCliente(); });
  document.body.appendChild(m);
  if (typeof _bloquearFondo === 'function') _bloquearFondo(true);
  var n = document.getElementById('ncl-nombre'); if (n) n.focus();
}

function cerrarNuevoCliente() {
  var m = document.getElementById('nuevo-cliente-modal');
  if (m) m.remove();
  if (typeof _bloquearFondo === 'function') _bloquearFondo(false);
}

async function guardarNuevoCliente() {
  var val = function(id) { var e = document.getElementById(id); return e ? e.value.trim() : ''; };
  var err = document.getElementById('ncl-err');
  var pinta = function(txt) { if (err) { err.textContent = txt; err.style.display = txt ? 'block' : 'none'; } };
  pinta('');
  var nombre = val('ncl-nombre'), limite = val('ncl-limite');
  if (!nombre) { pinta(t('ar.faltaNombre', 'Pon el nombre del cliente.')); return; }
  if (!limite || Number(limite) <= 0) { pinta(t('ar.faltaLimite', 'El límite de crédito tiene que ser mayor que 0.')); return; }
  var btn = document.getElementById('ncl-ok');
  if (btn) { btn.disabled = true; btn.style.opacity = '.5'; }
  try {
    var r = await _postJson('/api/ar_real/cliente', {
      nombre: nombre, nif: val('ncl-nif'), limite: limite,
      dias_pago: val('ncl-dias') || 30, email: val('ncl-email'), telefono: val('ncl-tel')
    });
    var d = await r.json();
    if (!d.ok) throw new Error(d.error || 'error');
    cerrarNuevoCliente();
    showNotification('✓ ' + d.cliente, 'success');
    if (typeof cargarARRealData === 'function') cargarARRealData();
  } catch(e) {
    pinta('✗ ' + (e.message || 'error'));
    if (btn) { btn.disabled = false; btn.style.opacity = '1'; }
  }
}

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
  if (!cliente) { showNotification(tt('ar.selCliente', 'Selecciona un cliente'), 'info'); return; }
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
// ── Direct bill: bono de agencia vs factura a credito ──
function _bonoEsc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function _bonoEur(v){ return (v==null||isNaN(Number(v))) ? '—' : Number(v).toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}) + ' €'; }
async function cargarBonosAR(){
  var wrap = document.getElementById('ar-bonos-list');
  var res = document.getElementById('ar-bonos-resumen');
  if (!wrap) return;
  try {
    var r = await fetch('/api/ar_real/bonos');
    var d = await r.json();
    var bonos = (d && d.bonos) || [];
    var sin = (d && d.facturas_sin_bono) || [];
    var rs = (d && d.resumen) || {};
    if (res) res.textContent = (bonos.length || sin.length)
      ? t('bonos.resumen', '{ok} cuadran · {dif} con diferencia · {sf} sin factura · {sb} facturas sin bono')
          .replace('{ok}', rs.CUADRA||0).replace('{dif}', (rs.DIFERENCIA_IMPORTE||0)+(rs.DIFERENCIA_FECHAS||0)).replace('{sf}', rs.SIN_FACTURA||0).replace('{sb}', rs.FACTURA_SIN_BONO||0)
      : '';
    if (!bonos.length && !sin.length) { wrap.innerHTML = _vacioCard(t('bonos.vacio', 'Sube el bono de la agencia (voucher) en Procesar Archivos y aquí verás si la factura a crédito cuadra con lo autorizado.')); return; }
    var COL = {CUADRA:['#22c55e','rgba(34,197,94,.12)'], DIFERENCIA_IMPORTE:['#f87171','rgba(239,68,68,.12)'], DIFERENCIA_FECHAS:['#f59e0b','rgba(245,158,11,.12)'], SIN_FACTURA:['var(--mut)','rgba(148,163,184,.12)']};
    var LBL = {CUADRA:t('bonos.cuadra','✓ Cuadra'), DIFERENCIA_IMPORTE:t('bonos.difImporte','⚠ Importe distinto'), DIFERENCIA_FECHAS:t('bonos.difFechas','⚠ Fechas distintas'), SIN_FACTURA:t('bonos.sinFactura','Sin factura aún')};
    var h = bonos.map(function(b){
      var c = COL[b.estado] || COL.SIN_FACTURA;
      return '<div class="card" style="padding:10px 12px;border-radius:10px;display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;align-items:center">' +
        '<div style="min-width:0"><div style="font-weight:700;font-size:13px">' + _bonoEsc(b.agencia||'—') + ' · ' + t('bonos.bono','bono') + ' ' + _bonoEsc(b.numero_bono||'s/n') + '</div>' +
        '<div style="font-size:11px;color:var(--dim)">' + (b.huesped?_bonoEsc(b.huesped)+' · ':'') + _bonoEsc(b.fecha_entrada||'') + (b.fecha_salida?' → '+_bonoEsc(b.fecha_salida):'') +
        ' · ' + t('bonos.autorizado','autorizado') + ' <b>' + _bonoEur(b.importe_bono) + '</b>' +
        (b.numero_factura ? ' · ' + t('bonos.factura','factura') + ' ' + _bonoEsc(b.numero_factura) + ' ' + _bonoEur(b.importe_factura) : '') +
        (b.detalle ? ' · ' + _bonoEsc(b.detalle) : '') + '</div></div>' +
        '<span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;background:' + c[1] + ';color:' + c[0] + '">' + _bonoEsc(LBL[b.estado]||b.estado) + '</span></div>';
    });
    if (sin.length) {
      h.push('<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#f87171;margin-top:6px">' + t('bonos.sinBonoTitulo','Facturas a crédito sin bono que las respalde') + '</div>');
      sin.forEach(function(f){
        h.push('<div class="card" style="padding:8px 12px;border-radius:10px;font-size:12px;display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><span>' + _bonoEsc(f.numero) + ' · ' + _bonoEsc(f.cliente) + (f.fecha_entrada?' · '+_bonoEsc(f.fecha_entrada):'') + '</span><b>' + _bonoEur(f.total) + '</b></div>');
      });
    }
    wrap.innerHTML = h.join('');
  } catch(e) {}
}
async function cargarBeosAR() {
  var wrap = document.getElementById('ar-beos-list');
  var cnt = document.getElementById('ar-beos-count');
  if (!wrap) return;
  try {
    var r = await fetch('/api/ar_real/beos');
    var d = await r.json();
    var beos = (d && d.beos) || [];
    if (cnt) cnt.textContent = beos.length ? '(' + beos.length + ')' : '';
    // Igual que en reclamaciones y en el grafico: sin datos, limpiar.
    if (!beos.length) {
      wrap.innerHTML = _vacioCard(t('beos.vacio', 'Procesa un contrato de grupo en <b>Procesar Archivos</b> y aquí verás su BEO con el cotejo de la factura.'));
      return;
    }
    var eur = function(v){ return '€' + (Number(v)||0).toLocaleString('es-ES',{minimumFractionDigits:2}); };
    wrap.innerHTML = beos.map(function(b){
      var c = b.cotejo || {};
      var badge, bg, col;
      if (c.estado === 'cuadra') { badge = '✓ Factura cuadra'; bg='rgba(34,197,94,.12)'; col='#22c55e'; }
      else if (c.estado === 'discrepancia') { badge = '⚠ ' + (c.diff_pct||0) + '% (' + eur(c.total_factura) + ' vs ' + eur(c.total_beo) + ')'; bg='rgba(239,68,68,.12)'; col='#f87171'; }
      else { badge = 'Sin factura aún'; bg='rgba(148,163,184,.12)'; col='var(--mut)'; }
      var lineas = (b.lineas||[]).map(function(l){
        return '<div style="display:flex;justify-content:space-between;gap:10px;font-size:12px;padding:5px 0;border-bottom:1px solid var(--s2)">' +
          '<span style="color:var(--tx)">' + (l.concepto||'') + ' <span style="color:var(--dim);font-size:11px">' + (l.detalle||'') + '</span></span>' +
          '<span style="color:var(--tx);font-weight:600;white-space:nowrap">' + eur(l.importe) + '</span></div>';
      }).join('');
      return '<div style="border:1px solid var(--s2);border-radius:12px;padding:14px;background:var(--s1)">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px;flex-wrap:wrap">' +
          '<div style="min-width:0"><div style="font-weight:700;font-size:13px">' + (b.evento||'Evento') + '</div>' +
          '<div style="font-size:11px;color:var(--dim)">' + (b.cliente||'') + (b.contrato ? ' · contrato ' + b.contrato : '') + (b.pax ? ' · ' + b.pax + ' pax' : '') + '</div></div>' +
          '<span style="font-size:11px;font-weight:700;padding:4px 10px;border-radius:20px;background:' + bg + ';color:' + col + '">' + badge + '</span>' +
        '</div>' + lineas +
        '<div style="display:flex;justify-content:space-between;font-size:13px;font-weight:800;padding-top:8px"><span>TOTAL BEO</span><span style="color:var(--acc2)">' + eur(b.total) + '</span></div>' +
        '<div style="font-size:10px;color:var(--dim);margin-top:6px">BEO generado automáticamente del contrato · ' + (b.fecha_generado||'') + '</div>' +
      '</div>';
    }).join('');
  } catch(e) {}
}

async function cargarARRealData() {
  // Show skeleton on KPIs while loading
  _skelOn(['arp-pendiente','arp-vencido','arp-cobrado','arp-nclientes']);
  try { cargarBeosAR(); } catch(e){}
  try { cargarBonosAR(); } catch(e){}
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
      _setText('arp-nclientes', dc.ok ? dc.clientes.length + ' ' + tt('ar.registrados', 'registrados') : '—');
      _skelOff(['arp-pendiente','arp-vencido','arp-cobrado','arp-nclientes']);

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
          return '<div class="card" style="padding:12px;cursor:pointer;transition:background-color .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s,opacity .15s" ' +
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
  if (el) { el.textContent = val; el.classList.remove('skeleton'); }
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
    tbody.innerHTML = '<tr><td colspan="4">' + _emptyState('📋', tt('ar.vacioTitulo', 'Sin facturas todavía'), tt('ar.vacioSub', 'Crea una factura con “Nueva factura” o procesa documentos de grupos y aparecerán aquí.'), false) + '</td></tr>';
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
      '<td style="text-align:center;padding:8px">' +
        (f.days_pending > 0 ? '<span style="font-size:13px;font-weight:700;color:' + (f.days_pending > 60 ? 'var(--red)' : f.days_pending > 30 ? 'var(--ora)' : 'var(--grn)') + '">' + f.days_pending + 'd</span>' : '<span style="color:var(--dim)">—</span>') +
      '</td>' +
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
  // Inline confirmation — no blocking dialog
  showConfirmAction(
    '¿Marcar como cobrada?',
    'Factura ' + numero,
    '💰 Confirmar cobro',
    async function() {
      try {
        const r = await _postJson('/api/ar_real/cobrar', {numero});
        const d = await r.json();
        if (d.ok) { showNotification('✓ Factura ' + numero + ' cobrada', 'success'); cargarARRealData(); }
        else showNotification('✗ ' + (d.error||'Error'), 'error');
      } catch(e) { showNotification('✗ Error de conexión', 'error'); }
    }
  );
}

async function recordatorioAR(numero) {
  try {
    const r = await _postJson('/api/ar_real/recordatorio', {numero});
    const d = await r.json();
    showNotification(d.ok ? '✓ ' + d.message : '✗ ' + (d.error||'Error'), d.ok ? 'success' : 'error');
  } catch(e) { showNotification('✗ Error de conexión', 'error'); }
}

async function procesarARReal() {
  cargarARRealData();
}


// FASE C: aqui estaban `openHotelDetail` y su modal. Llamaban a
// /api/multi_hotel/hotel/<id>, uno de los cuatro endpoints del demo que la
// fase C ha borrado. Su unico llamador era `renderMHMap`, que tampoco llamaba
// nadie: codigo muerto encadenado a un fichero de simulacion.


// ═══════════════════════════════════════════════════════════════════
// CALIPOLIS DASHBOARD
// ═══════════════════════════════════════════════════════════════════
// Calipolis code removed — product is generic

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
      '<div style="font-size:18px;color:var(--acc2)">🔗</div>' +
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
        <div style="color:var(--acc2);font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px">Junio 2026 — v1.5</div>
        <ul style="color:#94a3b8;padding-left:16px;line-height:1.8">
          <li>🔐 Protección CSRF en todas las rutas API autenticadas</li>
          <li>💳 Billing Stripe — checkout real con plan automático por habitaciones</li>
          <li>🏨 /unirse — registro self-service para nuevos hoteles</li>
          <li>📊 DRR GOP% — estimación automática cuando Excel tiene fórmulas</li>
          <li>🏢 Multi-Hotel — gráficos aislados, KPI cards siempre visibles</li>
          <li>💬 Chat Yve — abre desde nav, panel full-screen en móvil</li>
          <li>🧾 Signup → redirige a checkout automáticamente</li>
          <li>💰 Página pricing con CTAs directos a Stripe</li>
          <li>🏢 AR Real: flujo completo — clientes, aging, cobro, recordatorios</li>
          <li>🔍 Búsqueda global Ctrl+K · ⌨ Atajos 1-9</li>
          <li>⚖️ RGPD completo · 📝 Blog SEO 10 artículos</li>
        </ul>
      </div>
      <div>
        <div style="color:#a78bfa;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px">Mayo 2026</div>
        <ul style="color:#94a3b8;padding-left:16px;line-height:1.8">
          <li>🌍 Multi-idioma: ES/EN/CA/FR/DE/IT/PT</li>
          <li>💳 Integración Stripe (simulación)</li>
          <li>📊 Blog SEO con 6 artículos</li>
          <li>📱 Responsive móvil mejorado</li>
        </ul>
      </div>
    </div>
  </div>
</div>
<script>function showChangelog() {
  localStorage.setItem('changelog_seen', '2026-06-v3');
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
    {label:'Procesar archivos (⚡)', check: () => (parseInt(document.getElementById('sc-procesadas')?.textContent)||0) > 0, link:null, action:'Tab AR → Procesar Archivos'},
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
<!-- mobile bottom nav removed -->
<script>
if (window.innerWidth <= 768) {
  document.body.style.paddingBottom = '0';
}
</script>

<button id="back-top" onclick="window.scrollTo({top:0,behavior:'smooth'})" 
  style="display:none;position:fixed;bottom:88px;right:20px;background:var(--s1);border:1px solid var(--s2);color:var(--mut);width:36px;height:36px;border-radius:50%;font-size:16px;cursor:pointer;z-index:500;transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s;box-shadow:0 2px 8px rgba(0,0,0,.3)"
  onmouseover="this.style.borderColor='var(--acc)';this.style.color='var(--acc)'"
  onmouseout="this.style.borderColor='var(--s2)';this.style.color='var(--mut)'">↑</button>
<script>
window.addEventListener('scroll', () => {
  const btn = document.getElementById('back-top');
  if (btn) btn.style.display = window.scrollY > 300 ? 'flex' : 'none';
  if (btn) btn.style.alignItems = 'center'; if (btn) btn.style.justifyContent = 'center';
}, {passive: true});
</script>

<!-- Modal Escanear Documento -->

<script src="/static/yve-icons.js?v=__ASSETS_V__"></script>
</body>
</html>"""


@app.route('/api/demo/generar', methods=['POST'])
@login_required
def api_demo_generar():
    """Genera datos demo con los nombres de hotel/cadena(s) que da el usuario."""
    global DEMO_MODE
    data = request.get_json(force=True) or {}
    cadenas = data.get('cadenas') or []
    cadenas = [c for c in cadenas if c.get('nombre') and c.get('hoteles')]
    if not cadenas:
        return jsonify({'ok': False, 'error': 'Indica al menos un hotel o cadena'}), 400
    try:
        from demo_generator import generar_demo
        resumen = generar_demo(cadenas)
        DEMO_MODE = True
        _audit('DEMO_ON', f"{resumen['hoteles']} hoteles, {resumen['cadenas']} cadenas")
        return jsonify({'ok': True, 'demo_mode': True, **resumen})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500


@app.route('/api/demo/toggle', methods=['POST'])
@login_required
def toggle_demo():
    """Desactiva demo mode y limpia los datos generados."""
    global DEMO_MODE
    if DEMO_MODE:
        DEMO_MODE = False
        try:
            from demo_generator import limpiar_demo
            limpiar_demo()
        except Exception:
            pass
        return jsonify({"demo_mode": False, "status": "desactivado"})
    DEMO_MODE = True
    return jsonify({"demo_mode": True, "status": "activado"})

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
        from notificaciones import enviar_pendientes
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
        with open(os.path.join(_ddir(), 'notificaciones_historial.json'), 'r') as f:
            historial = json.load(f)
        return jsonify(historial[-50:])  # Últimas 50
    except:
        return jsonify([])



@app.route('/api/reportes/ejecutivo.pdf')
def reporte_ejecutivo():
    """Descarga reporte ejecutivo en PDF"""
    try:
        from exportador_final import generar_reporte_ejecutivo
        pdf = generar_reporte_ejecutivo()
        return send_file(os.path.abspath(pdf), mimetype='application/pdf', as_attachment=True, download_name='Reporte_Ejecutivo.pdf')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/reportes/consolidado.xlsx')
def reporte_consolidado():
    """Descarga reporte consolidado en Excel"""
    try:
        from exportador_final import generar_excel_consolidado
        xlsx = generar_excel_consolidado()
        return send_file(os.path.abspath(xlsx), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Consolidado.xlsx')
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



@app.route('/api/eliminar_archivo', methods=['POST'])
@login_required
def api_eliminar_archivo():
    """Elimina un archivo de facturas-entrada."""
    data = request.json or {}
    fname = data.get('nombre', '').strip()
    if not fname or '/' in fname or '..' in fname:
        return jsonify({"error": "nombre inválido"}), 400
    fpath = os.path.join(_edir(), fname)
    if os.path.exists(fpath):
        os.remove(fpath)
        # Quitar del log de procesados también
        log = _load_proc_log()
        log.pop(fname, None)
        _save_proc_log(log)
        return jsonify({"ok": True})
    return jsonify({"error": "archivo no encontrado"}), 404


@app.route('/api/reset_datos', methods=['POST'])
@login_required
def api_reset_datos():
    """Borra todos los datos procesados para empezar desde cero."""
    import glob, json as _json
    borrados = 0
    # facturas-procesadas
    for f in glob.glob(os.path.join(BASE_DIR, 'facturas-procesadas', '*.xlsx')):
        os.remove(f); borrados += 1
    # reportes de datos
    for pattern in ['doble_imposicion_*', 'matching_*', 'verificacion_*', 'conciliacion_*']:
        for f in glob.glob(os.path.join(BASE_DIR, 'reportes', pattern)):
            os.remove(f); borrados += 1
    # reset log
    _save_proc_log({})
    return jsonify({"ok": True, "borrados": borrados})

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
