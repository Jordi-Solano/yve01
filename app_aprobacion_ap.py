"""
app_aprobacion_ap.py — Yve.01 Módulo AP
Dashboard de aprobaciones AP para jefes de departamento. Puerto 5002.
Ejecutar: python app_aprobacion_ap.py
"""

import os, glob, json
from datetime import date, datetime
import pandas as pd
from flask import Blueprint, jsonify, request, Response
from flask_login import login_required
from version_estaticos import SELLO as SELLO_ESTATICOS

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
PROCESADAS_DIR   = os.path.join(BASE_DIR, "facturas-procesadas")
APROBACIONES_DIR = os.path.join(BASE_DIR, "aprobaciones")
os.makedirs(APROBACIONES_DIR, exist_ok=True)

APRO_FILE = os.path.join(APROBACIONES_DIR, "aprobaciones_ap.xlsx")
NF        = "NO_ENCONTRADO"

bp = Blueprint("aprob_ap", __name__, url_prefix="/aprobaciones-ap")

@bp.before_request
@login_required
def _require_login():
    """Protege todas las rutas del blueprint: exige sesión iniciada."""
    pass


# ── Datos ─────────────────────────────────────────────────────────────────

def cargar_facturas_ap():
    excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_contabilizadas_*.xlsx")), reverse=True)
    if not excels:
        excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_ap_*.xlsx")), reverse=True)
    if not excels:
        return pd.DataFrame()
    return pd.read_excel(excels[0])

def cargar_aprobaciones():
    if not os.path.exists(APRO_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(APRO_FILE)
    except Exception:
        return pd.DataFrame()

def safe_str(v):
    if v is None or str(v).strip() in (NF, "nan", "None", ""):
        return ""
    return str(v).strip()

def _cuenta_str(v):
    """El codigo contable sin el '.0' que le pega el viaje por Excel.

    Mismo criterio que dashboard._cuenta_str: cuenta_debe_gasto vuelve como
    float64 y str(6001.0) es '6001.0'. Un numero de cuenta con decimal no es
    un numero de cuenta.
    """
    s = safe_str(v)
    if not s:
        return ""
    try:
        f = float(s)
    except (TypeError, ValueError):
        return s                      # 'REVISAR_MANUAL' y demas: tal cual
    return str(int(f)) if f == int(f) else s


# Estados del cruce agrupados por lo que significan PARA QUIEN APRUEBA.
# Se quedaron sin actualizar desde la Fase 3b y por eso los tiles decian 0.
# Decision del usuario: "factura sin PO" es informativo, NO una incidencia —
# muchas compras de hotel no llevan PO y marcarlas todas seria ruido.
_ESTADOS_OK = ("MATCH_3WAY_OK", "MATCH_CORRECTO", "MATCH_ALBARAN_OK", "MATCH_PO_OK")
_ESTADOS_INCIDENCIA = ("DISCREPANCIA", "DISCREPANCIA_PO", "DIFERENCIA_IMPORTE",
                       "DIFERENCIA_LINEA", "DIFERENCIA_PO_IMPORTE",
                       "FACTURA_SIN_ALBARAN", "ALERTA_CONSUMO")


def clave_factura(fila):
    """Identidad de una factura en el panel. NUNCA vacia.

    El numero de factura no vale por si solo: una factura sin numero vuelve de
    Excel como NaN y `groupby` tira las claves NaN, asi que su aprobacion
    desaparecia. `archivo` es la columna con la que el propio Excel de facturas
    deduplica, o sea que ya es unica por factura — incluso cuando un documento
    trae varias ("hoja.csv#FP-2026-101").

    Asi, aprobar una factura sin numero no afecta a las demas sin numero.
    """
    num = safe_str(fila.get("numero_factura"))
    return num or safe_str(fila.get("archivo"))


def _acciones_por_clave():
    """{clave: accion} con la ULTIMA accion de cada factura.

    Se ordena por `fecha_hora` y gana la ultima, EXACTAMENTE igual que
    `oracle_lector_facturas.cargar_aprobaciones_ap()`. `fecha_hora` es texto
    dd/mm/aaaa y ordenarlo es incorrecto en cuanto cruzas meses, pero los dos se
    equivocarian IGUAL: es lo que hoy los mantiene de acuerdo. Arreglarlo hay
    que hacerlo en los dos a la vez, y eso es tocar Oracle.

    En Python plano, sin groupby: es justo el groupby el que tiraba las claves
    vacias (regla 3, la familia de los NaN).
    """
    df = cargar_aprobaciones()
    if df.empty or "accion" not in df.columns:
        return {}
    filas = df.to_dict("records")
    if "fecha_hora" in df.columns:
        filas.sort(key=lambda r: safe_str(r.get("fecha_hora")))
    mapa = {}
    for r in filas:
        # las filas viejas no tienen clave_factura: siguen valiendo por numero
        k = safe_str(r.get("clave_factura")) or safe_str(r.get("numero_factura"))
        if k:
            mapa[k] = safe_str(r.get("accion")).upper()
    return mapa


def facturas_a_lista(df):
    if df.empty:
        return []
    apro_map = _acciones_por_clave()

    filas = []
    for _, r in df.iterrows():
        num = safe_str(r.get("numero_factura", ""))
        _clave = clave_factura(r)
        filas.append({
            "clave":             _clave,
            "archivo":           safe_str(r.get("archivo")),
            "numero_factura":    num,
            "nombre_proveedor":  safe_str(r.get("nombre_proveedor")),
            "NIF_proveedor":     safe_str(r.get("NIF_proveedor")),
            "tipo_proveedor":    safe_str(r.get("tipo_proveedor","OTRAS")),
            "descripcion":       safe_str(r.get("descripcion_concepto")),
            "base_imponible":    safe_str(r.get("base_imponible")),
            "cuota_iva":         safe_str(r.get("cuota_iva")),
            "total_factura":     safe_str(r.get("total_factura")),
            "cuenta_debe":       _cuenta_str(r.get("cuenta_debe_gasto")),
            "estado_asignacion": safe_str(r.get("estado_asignacion")),
            "estado_matching":   safe_str(r.get("estado_matching","")),
            "alerta_detalle":    safe_str(r.get("detalle_matching",r.get("alerta_detalle",""))),
            "departamento_po":   safe_str(r.get("departamento_po","General")),
            "accion":            apro_map.get(_clave, ""),
        })
    return filas

@login_required
@bp.route("/api/facturas")
def api_facturas():
    """Por defecto SOLO las pendientes.

    Una factura aprobada o rechazada ya no es trabajo por hacer: su sitio es el
    historial. Antes se devolvian todas y la vista solo les quitaba los botones,
    asi que una aprobada seguia ocupando sitio en "por aprobar".
    `?estado=resueltas` o `?estado=todas` para lo demas.
    """
    df   = cargar_facturas_ap()
    dept = request.args.get("departamento","").strip().lower()
    estado = request.args.get("estado","pendientes").strip().lower()
    rows = facturas_a_lista(df)
    if dept:
        rows = [r for r in rows if dept in r.get("departamento_po","").lower()
                or dept == "todos"]
    if estado == "pendientes":
        rows = [r for r in rows if not r.get("accion")]
    elif estado == "resueltas":
        rows = [r for r in rows if r.get("accion")]
    return jsonify(rows)


@bp.route("/api/historial")
def api_historial():
    """Lo aprobado y lo rechazado, del SERVIDOR.

    Antes el historial era un array en memoria del navegador: al recargar se
    vaciaba y la factura solo aparecia en "por aprobar". El dato siempre estuvo
    en aprobaciones_ap.xlsx; lo que faltaba era poder leerlo.

    SOLO LECTURA. No escribe en el fichero que lee el gate de Oracle.
    """
    df = cargar_aprobaciones()
    if df.empty or "accion" not in df.columns:
        return jsonify([])
    vigentes = _acciones_por_clave()
    filas = df.to_dict("records")
    if "fecha_hora" in df.columns:
        filas.sort(key=lambda r: safe_str(r.get("fecha_hora")), reverse=True)
    out = []
    for r in filas:
        k = safe_str(r.get("clave_factura")) or safe_str(r.get("numero_factura"))
        accion = safe_str(r.get("accion")).upper()
        out.append({
            "clave":          k,
            "numero_factura": safe_str(r.get("numero_factura")) or k,
            "accion":         accion,
            "fecha_hora":     safe_str(r.get("fecha_hora")),
            "comentario":     safe_str(r.get("comentario")),
            "departamento":   safe_str(r.get("departamento")),
            "aprobador":      safe_str(r.get("aprobador")),
            # si despues se rectifico, esta entrada ya no es la que manda
            "vigente":        vigentes.get(k, "") == accion,
        })
    return jsonify(out)

@bp.route("/api/stats")
def api_stats():
    """Los contadores cuentan TRABAJO PENDIENTE.

    Antes `total` incluia lo ya aprobado, asi que el tile seguia contando
    trabajo hecho; y `match_ok`/`discrepancias` se quedaron en los estados de
    antes de la Fase 3b, o sea que MATCH_ALBARAN_OK, DIFERENCIA_IMPORTE,
    DIFERENCIA_LINEA y FACTURA_SIN_ALBARAN no los contaba nadie. Medido en
    produccion: una factura que cuadraba con su albaran y el tile en 0.
    """
    df   = cargar_facturas_ap()
    rows = facturas_a_lista(df)
    pend = [r for r in rows if not r.get("accion")]
    def cnt(key, val, sobre=None):
        return sum(1 for r in (pend if sobre is None else sobre) if r.get(key,"") == val)
    def cnt_en(estados, sobre):
        return sum(1 for r in sobre if r.get("estado_matching","") in estados)
    return jsonify({
        # lo que hay por hacer
        "pendientes":    len(pend),
        "match_ok":      cnt_en(_ESTADOS_OK, pend),
        "discrepancias": cnt_en(_ESTADOS_INCIDENCIA, pend),
        "alertas":       cnt("estado_matching","ALERTA_CONSUMO"),
        "sin_po":        cnt("estado_matching","SIN_PO"),
        "manuales":      cnt("estado_asignacion","SIN_REGLA"),
        # y lo ya resuelto, para el historial
        "aprobadas":     cnt("accion","APROBADA", rows),
        "rechazadas":    cnt("accion","RECHAZADA", rows),
        "total":         len(rows),
    })

@login_required
@bp.route("/api/accion", methods=["POST"])
def api_accion():
    data = request.get_json(force=True)
    num_fac     = safe_str(data.get("numero_factura",""))
    clave       = safe_str(data.get("clave","")) or num_fac
    accion      = data.get("accion","")
    comentario  = data.get("comentario","")
    departamento= data.get("departamento","")
    aprobador   = data.get("aprobador","Jefe de Departamento")

    # Se exige la CLAVE, no el numero: una factura sin numero tambien se aprueba,
    # y antes daba 400. La clave nunca esta vacia (ver clave_factura).
    if not clave or not accion or not comentario or not departamento:
        return jsonify({"ok": False, "error": "Faltan campos obligatorios"}), 400

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nueva = pd.DataFrame([{
        "fecha_hora":    now_str,
        # ORACLE lee ESTA columna. Si la factura tiene numero se escribe tal
        # cual, como siempre. Si no lo tiene se escribe la clave (el nombre del
        # documento): Oracle no encontrara correspondencia y NO contabilizara.
        # Falla en cerrado a proposito — una factura sin numero no tiene
        # referencia con la que hacer el asiento.
        "numero_factura":num_fac or clave,
        "clave_factura": clave,
        "accion":        accion,
        "comentario":    comentario,
        "departamento":  departamento,
        "aprobador":     aprobador,
    }])

    if os.path.exists(APRO_FILE):
        df_ex = pd.read_excel(APRO_FILE)
        df_ex = pd.concat([df_ex, nueva], ignore_index=True)
    else:
        df_ex = nueva

    with pd.ExcelWriter(APRO_FILE, engine="openpyxl") as w:
        df_ex.to_excel(w, index=False, sheet_name="Aprobaciones_AP")
    return jsonify({"ok": True})

@bp.route("/")
def index():
    return HTML.replace("__ASSETS_V__", SELLO_ESTATICOS)

HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/static/yve.css">
<title>Yve.01 — Facturas por aprobar</title>
<style>
/* Paleta y tipografia salen de /static/yve.css — aqui NO se redefinen,
   asi el color personalizado del usuario (yve_accent/yve_bg) tambien manda aqui. */
*{box-sizing:border-box;margin:0;padding:0}
html{overflow-x:hidden}
body{overflow-x:hidden;background:var(--bg);color:var(--tx);font-family:var(--font);
  min-height:100vh;line-height:1.5;position:relative}
/* Ambiente igual que el dashboard y el login */
body::before{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;
  background:
    radial-gradient(900px 500px at 90% -5%,rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10),transparent 60%),
    radial-gradient(700px 400px at -5% 105%,rgba(139,92,246,.08),transparent 55%)}

/* ── NAV ── */
.nav{position:sticky;top:0;z-index:200;display:flex;align-items:center;gap:12px;
  height:60px;padding:0 22px;
  background:rgba(var(--bg-r,15),var(--bg-g,23),var(--bg-b,42),.92);
  backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  border-bottom:1px solid var(--s2)}
.logo{display:flex;align-items:center;gap:10px;flex-shrink:0}
.logo-dot{width:10px;height:10px;border-radius:50%;background:var(--acc);flex-shrink:0;
  box-shadow:0 0 6px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.6),
             0 0 14px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.35),
             0 0 28px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)}
.logo-name{font-size:19px;font-weight:800;color:#fff;letter-spacing:-.3px}
.logo-name span{color:var(--acc2)}
.logo-tag{font-size:11px;color:var(--mut);white-space:nowrap;padding-left:4px}
.nm{flex:1}
.btn-back{display:inline-flex;align-items:center;gap:7px;text-decoration:none;
  background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10);
  border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.32);
  color:var(--acc2);padding:8px 15px;border-radius:9px;font-size:12.5px;font-weight:700;
  white-space:nowrap}
.btn-back:hover{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.20);
  border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.65)}
.sel-dept{background:rgba(255,255,255,.04);color:var(--tx);border:1px solid var(--s2);
  padding:7px 11px;border-radius:9px;font-size:12px;cursor:pointer;font-family:var(--font)}
.sel-dept:hover{border-color:var(--s3)}
.btn-ref{background:none;border:1px solid var(--s2);color:var(--mut);width:34px;height:34px;
  border-radius:9px;font-size:14px;cursor:pointer;flex-shrink:0}
.btn-ref:hover{border-color:var(--acc);color:var(--acc2)}

/* ── TABS ── */
.tabs{display:flex;gap:4px;padding:0 22px;border-bottom:1px solid var(--s2);
  background:rgba(var(--bg-r,15),var(--bg-g,23),var(--bg-b,42),.6);position:relative;z-index:1}
.tab{padding:13px 20px;font-size:13px;font-weight:600;color:var(--mut);cursor:pointer;
  border-bottom:2px solid transparent;margin-bottom:-1px}
.tab:hover{color:var(--tx)}
.tab.on{color:var(--acc2);border-bottom-color:var(--acc)}

/* ── LAYOUT ── */
.main{position:relative;z-index:1;padding:22px;max-width:860px;margin:0 auto}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px}
.sc{background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:18px 16px;
  transition:background-color .2s,border-color .2s,color .2s,box-shadow .2s,transform .2s,opacity .2s}
.sc-lbl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;
  margin-bottom:10px}
.sc-v{font-size:28px;font-weight:800;line-height:1;letter-spacing:-1px;
  font-variant-numeric:tabular-nums}
.c-b{color:var(--acc2)}.c-g{color:var(--grn)}.c-o{color:var(--ora)}

/* ── TARJETA DE FACTURA ── */
.card{background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:20px;
  margin-bottom:12px;box-shadow:var(--shadow-soft);transition:border-color .18s,transform .18s}
.card:hover{border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.35)}
.card-top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;
  margin-bottom:16px}
.prov-name{font-size:16px;font-weight:700;color:var(--tx);letter-spacing:-.2px}
.prov-num{font-size:11.5px;color:var(--dim);margin-top:3px;font-family:var(--mono)}
.badge{display:inline-flex;align-items:center;gap:4px;font-size:9.5px;font-weight:700;
  padding:4px 10px;border-radius:20px;text-transform:uppercase;letter-spacing:.4px;
  white-space:nowrap}
.b-fb{background:rgba(234,179,8,.14);color:#fde047;border:1px solid rgba(234,179,8,.30)}
.b-otras{background:rgba(139,92,246,.14);color:#c4b5fd;border:1px solid rgba(139,92,246,.30)}
.b-ok{background:rgba(34,197,94,.14);color:#4ade80;border:1px solid rgba(34,197,94,.30)}
.b-disc{background:rgba(239,68,68,.14);color:#fca5a5;border:1px solid rgba(239,68,68,.30)}
.b-alert{background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.14);
  color:var(--acc3);border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.30)}
.b-nopo{background:rgba(249,115,22,.14);color:#fdba74;border:1px solid rgba(249,115,22,.30)}
.b-man{background:rgba(139,92,246,.14);color:#c4b5fd;border:1px solid rgba(139,92,246,.30)}
.b-apr{background:rgba(34,197,94,.14);color:#4ade80;border:1px solid rgba(34,197,94,.30)}
.b-rec{background:rgba(239,68,68,.14);color:#fca5a5;border:1px solid rgba(239,68,68,.30)}
.b-pen{background:rgba(148,163,184,.10);color:var(--mut);border:1px solid var(--s2)}
.info-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px 18px;margin-bottom:14px}
.card-acc{display:flex;align-items:center;gap:8px;flex-shrink:0}
.chev{background:none;border:1px solid var(--s2);color:var(--mut);width:28px;height:28px;
  border-radius:8px;cursor:pointer;font-size:10px;line-height:1;display:flex;
  align-items:center;justify-content:center;flex-shrink:0;font-family:var(--font);
  transition:transform .18s,border-color .15s,color .15s}
.chev:hover{border-color:var(--acc);color:var(--acc2)}
.chev.abierta{transform:rotate(180deg);color:var(--acc2);
  border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.45)}
.detalle{display:none;border-top:1px solid var(--s2);margin-bottom:14px;padding-top:14px}
.detalle.abierta{display:block}
.detalle .info-grid{margin-bottom:0}
.det-full{grid-column:1 / -1}
.det-full .val{font-size:13px;font-weight:600;line-height:1.5}
.det-arch{font-family:var(--mono);font-size:11px;color:var(--dim);word-break:break-all}
.ii .lbl{font-size:9.5px;color:var(--mut);text-transform:uppercase;letter-spacing:.6px;
  margin-bottom:4px}
.ii .val{font-size:15px;font-weight:700;color:var(--tx);font-variant-numeric:tabular-nums}
.cuenta-tag{display:inline-flex;align-items:center;gap:6px;
  background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.10);
  border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.25);
  border-radius:8px;padding:5px 10px;font-size:11px;color:var(--acc3);font-weight:600;
  margin-bottom:14px}
.alerta-box{border-radius:10px;padding:10px 12px;font-size:11.5px;margin-bottom:14px;
  line-height:1.55;background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.08);
  border:1px solid rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.22);color:var(--acc3)}
.alerta-box.warn{background:rgba(249,115,22,.08);border-color:rgba(249,115,22,.25);color:#fdba74}
.alerta-box.err{background:rgba(239,68,68,.08);border-color:rgba(239,68,68,.25);color:#fca5a5}
.estado-row{padding-top:4px}

/* ── FORMULARIO DE APROBACION ── */
.sep{height:1px;background:var(--s2);margin:16px 0 14px}
.dept-row{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.dept-label{font-size:11.5px;color:var(--mut);font-weight:600;white-space:nowrap}
.dept-select{flex:1;background:rgba(255,255,255,.03);border:1px solid var(--s2);color:var(--tx);
  padding:9px 11px;border-radius:9px;font-size:12.5px;font-family:var(--font);cursor:pointer}
.dept-select:focus{border-color:var(--acc);outline:none;
  box-shadow:0 0 0 3px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)}
textarea{width:100%;background:rgba(255,255,255,.03);border:1px solid var(--s2);color:var(--tx);
  border-radius:9px;padding:10px 12px;font-size:12.5px;resize:none;height:58px;
  font-family:var(--font);line-height:1.5}
textarea::placeholder{color:var(--dim)}
textarea:focus{border-color:var(--acc);outline:none;
  box-shadow:0 0 0 3px rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.15)}
.btn-row{display:flex;gap:10px;margin-top:12px}
.btn{padding:12px 16px;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;
  font-family:var(--font);border:1px solid transparent}
.btn.ok{flex:2;background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;
  box-shadow:0 4px 14px rgba(34,197,94,.22)}
.btn.ok:hover{background:linear-gradient(135deg,#16a34a,#15803d)}
.btn.ko{flex:1;background:rgba(239,68,68,.10);color:#fca5a5;border-color:rgba(239,68,68,.35)}
.btn.ko:hover{background:rgba(239,68,68,.20);border-color:rgba(239,68,68,.6)}
.btn.disabled{opacity:.35;cursor:not-allowed;box-shadow:none}

/* ── VACIO / TOAST / HISTORIAL ── */
.empty{text-align:center;padding:54px 24px;color:var(--mut);font-size:13px;
  background:var(--s1);border:1px dashed var(--s2);border-radius:16px}
.empty .emo{font-size:32px;display:block;margin-bottom:12px;opacity:.75}
.empty .tit{font-size:15px;font-weight:700;color:var(--tx);margin-bottom:6px}
.toast{position:fixed;bottom:26px;left:50%;transform:translateX(-50%);padding:11px 22px;
  border-radius:24px;font-size:12.5px;font-weight:700;color:#fff;white-space:nowrap;opacity:0;
  transition:opacity .3s;pointer-events:none;z-index:300;box-shadow:var(--shadow-lift)}
.toast.on{opacity:1}
.hist-item{background:var(--s1);border:1px solid var(--s2);border-radius:12px;padding:12px 14px;
  margin-bottom:8px;display:flex;align-items:center;gap:12px}
.hist-icon{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:14px;flex-shrink:0}
.hist-icon.a{background:rgba(34,197,94,.15);color:#4ade80}
.hist-icon.r{background:rgba(239,68,68,.15);color:#fca5a5}
.hist-info .n{font-size:13px;font-weight:600;color:var(--tx)}
.hist-info .d{font-size:11px;color:var(--dim);margin-top:1px}
.hist-accion{margin-left:auto;font-size:9.5px;font-weight:700;padding:4px 9px;border-radius:6px;
  letter-spacing:.4px}
.hist-item.pasada{opacity:.55}
.hist-old{font-size:9px;font-weight:700;color:var(--mut);border:1px solid var(--s2);
  border-radius:5px;padding:1px 5px;margin-left:6px;text-transform:uppercase}
.hist-accion.a{background:rgba(34,197,94,.15);color:#4ade80}
.hist-accion.r{background:rgba(239,68,68,.15);color:#fca5a5}

@media(max-width:640px){
  .nav{padding:0 12px;gap:8px;height:54px}
  .logo-tag{display:none}
  .btn-back{padding:7px 11px;font-size:12px}
  .btn-back .bk-txt{display:none}
  .tabs{padding:0 12px}
  .tab{padding:12px 14px;font-size:12.5px}
  .main{padding:14px}
  .stats{gap:8px}
  .sc{padding:14px 12px}
  .sc-v{font-size:23px}
  .card{padding:16px}
  .info-grid{gap:12px}
}
svg.yvi{width:1em;height:1em;vertical-align:-0.125em;flex-shrink:0;display:inline-block}
</style>
<script src="/static/yve-tema.js?v=__ASSETS_V__"></script>
</head>
<body>
<nav class="nav">
  <div class="logo">
    <span class="logo-dot"></span>
    <span class="logo-name">Yve<span>.01</span></span>
    <span class="logo-tag">Facturas por aprobar</span>
  </div>
  <a class="btn-back" href="/" title="Volver al panel principal">← <span class="bk-txt">Panel principal</span></a>
  <div class="nm"></div>
  <select class="sel-dept" id="dept-filter" onchange="loadData()">
    <option value="">Todos los departamentos</option>
    <option value="f&b">F&B</option>
    <option value="administracion">Administración</option>
    <option value="mantenimiento">Mantenimiento</option>
    <option value="housekeeping">Housekeeping</option>
    <option value="seguridad">Seguridad</option>
  </select>
  <button class="btn-ref" onclick="loadData()" title="Actualizar">↻</button>
</nav>

<div class="tabs">
  <div class="tab on" id="tab-fact" onclick="showTab('facturas',this)">Facturas</div>
  <div class="tab" id="tab-hist" onclick="showTab('historial',this)">Historial</div>
</div>

<div class="main">
  <div class="stats" id="stats-bar">
    <div class="sc"><div class="sc-lbl">Por aprobar</div><div class="sc-v c-b" id="s-pend">—</div></div>
    <div class="sc"><div class="sc-lbl">Cuadran</div><div class="sc-v c-g" id="s-ok">—</div></div>
    <div class="sc"><div class="sc-lbl">Con incidencia</div><div class="sc-v c-o" id="s-tot">—</div></div>
  </div>

  <div id="tab-facturas">
    <div id="lista"></div>
  </div>

  <div id="tab-historial" style="display:none">
    <div id="hist-empty" class="empty"><span class="emo">🗂️</span>
      <div class="tit">Sin aprobaciones todavía</div>
      Lo que apruebes o rechaces en esta sesión aparecerá aquí.</div>
    <div id="hist-list"></div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
const historial = [];
// Que tarjetas estan desplegadas, por NUMERO DE FACTURA y no por indice: la
// lista se repinta entera tras aprobar, y con indices se quedaria abierta la
// tarjeta equivocada.
const _abiertas = new Set();

function toggleDet(i) {
  const card = document.getElementById('card-'+i);
  const det  = document.getElementById('det-'+i);
  const chev = document.getElementById('chev-'+i);
  if (!card || !det) return;
  const abierta = det.classList.toggle('abierta');
  if (chev) chev.classList.toggle('abierta', abierta);
  const num = card.dataset.num || '';
  if (abierta) _abiertas.add(num); else _abiertas.delete(num);
}

// Importe en formato español, y NUNCA 'nan' en pantalla (regla del NaN)
function eur(v) {
  var s = (v === null || v === undefined) ? '' : String(v).trim();
  if (!s || /^(nan|none|nat|null|-)$/i.test(s)) return '—';
  var n = s.indexOf(',') >= 0 ? parseFloat(s.replace(/\\./g,'').replace(',','.')) : parseFloat(s);
  if (!isFinite(n)) return '—';
  return n.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}) + ' €';
}
function txt(v) {
  var s = (v === null || v === undefined) ? '' : String(v).trim();
  return (!s || /^(nan|none|nat|null)$/i.test(s)) ? '—' : s;
}

function showTab(tab, el) {
  if (tab === 'historial') cargarHist();
  document.getElementById('tab-facturas').style.display = tab==='facturas'?'':'none';
  document.getElementById('tab-historial').style.display = tab==='historial'?'':'none';
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  el.classList.add('on');
}

function bMatch(m) {
  const mp = {
    MATCH_3WAY_OK:['b-ok','✓ 3-Way OK'], MATCH_CORRECTO:['b-ok','✓ Match OK'],
    DISCREPANCIA_PO:['b-disc','✗ Discrepancia PO'], DISCREPANCIA:['b-disc','✗ Discrepancia'],
    ALERTA_CONSUMO:['b-alert','⚠ Alerta Consumo'], SIN_PO:['b-nopo','? Sin PO'],
    SIN_DATOS_POS:['b-nopo','~ Sin POS'], SIN_IMPORTE:['b-man','~ Sin Importe'],
    MATCH_ALBARAN_OK:['b-ok','✓ Cuadra con albarán'],
    DIFERENCIA_IMPORTE:['b-disc','✗ Diferencia de importe'],
    DIFERENCIA_LINEA:['b-disc','✗ Diferencia en una línea'],
    FACTURA_SIN_ALBARAN:['b-nopo','⚠ Sin albarán'],
    ALBARAN_SIN_FACTURAR:['b-alert','⚠ Albarán sin facturar'],
    ANTERIOR_AL_REGISTRO:['b-pen','· Anterior al registro'],
    PENDIENTE:['b-pen','· Sin cruzar'],
    SIN_IMPORTE_FACTURA:['b-man','~ Sin importe'],
    SIN_NOMBRE_PROVEEDOR:['b-man','~ Sin proveedor'],
  };
  const[c,l]=mp[m]||['b-pen',m||'—'];
  return '<span class="badge '+c+'">'+l+'</span>';
}

function alertClass(m) {
  if(m==='ALERTA_CONSUMO')return 'warn';
  if(m==='DISCREPANCIA_PO'||m==='DISCREPANCIA')return 'err';
  return '';
}

async function loadData() {
  const dept = document.getElementById('dept-filter').value;
  // solo las PENDIENTES: una aprobada ya no es trabajo por hacer, su sitio es
  // el historial. Antes venian todas y solo se les quitaban los botones.
  const url  = '/aprobaciones-ap/api/facturas?estado=pendientes' + (dept ? '&departamento='+encodeURIComponent(dept) : '');
  const [fr, sr] = await Promise.all([fetch(url), fetch('/aprobaciones-ap/api/stats')]);
  const rows = await fr.json();
  const stats= await sr.json();

  document.getElementById('s-pend').textContent = stats.pendientes??'—';
  document.getElementById('s-ok').textContent   = stats.match_ok??'—';
  document.getElementById('s-tot').textContent  = stats.discrepancias??'—';

  const lista = document.getElementById('lista');
  if(!rows.length){lista.innerHTML='<div class="empty"><span class="emo">📭</span>'
    +'<div class="tit">No queda nada por aprobar</div>'
    +'Lo aprobado y lo rechazado está en la solapa Historial.</div>';return;}

  lista.innerHTML = rows.map((r,i) => {
    const needsAlert = r.estado_matching && r.estado_matching !== 'MATCH_3WAY_OK' && r.estado_matching !== 'MATCH_CORRECTO' && r.alerta_detalle;
    const alertHtml  = needsAlert ? '<div class="alerta-box '+alertClass(r.estado_matching)+'">'+r.alerta_detalle+'</div>' : '';
    const _num = txt(r.numero_factura);
    const _ab = _abiertas.has(_num);
    return '<div class="card" id="card-'+i+'" data-num="'+_num+'" data-clave="'+txt(r.clave).replace(/"/g,'&quot;')+'">' +
      '<div class="card-top">' +
        '<div><div class="prov-name">'+txt(r.nombre_proveedor)+'</div><div class="prov-num">'+_num+'</div></div>' +
        '<div class="card-acc">' +
          '<span class="badge '+(r.tipo_proveedor==='FB'?'b-fb':'b-otras')+'">'+txt(r.tipo_proveedor)+'</span>' +
          (r.accion ? '' :
            '<button class="chev'+(_ab?' abierta':'')+'" id="chev-'+i+'" onclick="toggleDet('+i+')" title="Ver la factura completa">▼</button>') +
        '</div>' +
      '</div>' +
      '<div class="info-grid">' +
        '<div class="ii"><div class="lbl">Total</div><div class="val">'+eur(r.total_factura)+'</div></div>' +
        '<div class="ii"><div class="lbl">IVA</div><div class="val">'+eur(r.cuota_iva)+'</div></div>' +
        '<div class="ii"><div class="lbl">Matching</div><div class="val">'+bMatch(r.estado_matching)+'</div></div>' +
        '<div class="ii"><div class="lbl">Departamento</div><div class="val">'+txt(r.departamento_po)+'</div></div>' +
      '</div>' +
      (r.accion ? '' :
      '<div class="detalle'+(_ab?' abierta':'')+'" id="det-'+i+'">' +
        '<div class="info-grid">' +
          '<div class="ii"><div class="lbl">Base imponible</div><div class="val">'+eur(r.base_imponible)+'</div></div>' +
          '<div class="ii"><div class="lbl">NIF del proveedor</div><div class="val">'+txt(r.NIF_proveedor)+'</div></div>' +
          '<div class="ii det-full"><div class="lbl">Concepto</div><div class="val">'+txt(r.descripcion)+'</div></div>' +
          '<div class="ii det-full"><div class="lbl">Documento de origen</div><div class="val det-arch">'+txt(r.archivo)+'</div></div>' +
        '</div>' +
      '</div>') +
      (r.cuenta_debe ? '<div class="cuenta-tag">📒 Cuenta ' + r.cuenta_debe + ' — ' + txt(r.estado_asignacion) + '</div>' : '') +
      alertHtml +
      (r.accion ? '<div class="estado-row"><span class="badge '+(r.accion==='APROBADA'?'b-apr':'b-rec')+'">'+(r.accion==='APROBADA'?'✓ ':'✗ ')+r.accion+'</span></div>' : '') +
      (!r.accion ? (
        '<div class="sep"></div>' +
        '<div class="dept-row"><span class="dept-label">Departamento:</span>' +
        '<select class="dept-select" id="dept-'+i+'" onchange="chk('+i+')"><option value="">— Selecciona —</option>' +
        '<option value="F&B">F&B</option><option value="Administracion">Administración</option>' +
        '<option value="Mantenimiento">Mantenimiento</option><option value="Housekeeping">Housekeeping</option>' +
        '<option value="Seguridad">Seguridad</option></select></div>' +
        '<textarea id="c-'+i+'" placeholder="Comentario obligatorio..." oninput="chk('+i+')"></textarea>' +
        '<div class="btn-row">' +
          '<button class="btn ok disabled" id="a-'+i+'" onclick="accion('+i+',\\'APROBADA\\',\\''+r.numero_factura+'\\')">✓ Aprobar</button>' +
          '<button class="btn ko disabled" id="k-'+i+'" onclick="accion('+i+',\\'RECHAZADA\\',\\''+r.numero_factura+'\\')">✗ Rechazar</button>' +
        '</div>'
      ) : '') +
    '</div>';
  }).join('');
}

function chk(i) {
  const ok = document.getElementById('c-'+i).value.trim().length > 0 &&
             document.getElementById('dept-'+i)?.value;
  document.getElementById('a-'+i).classList.toggle('disabled', !ok);
  document.getElementById('k-'+i).classList.toggle('disabled', !ok);
}

async function accion(i, tipo, numFac) {
  const c    = document.getElementById('c-'+i).value.trim();
  const dept = document.getElementById('dept-'+i)?.value;
  if(!c || !dept) return;
  // la CLAVE identifica la factura aunque no tenga numero; el numero se manda
  // igual porque es lo que lee el gate de Oracle
  const clave = (document.getElementById('card-'+i)||{dataset:{}}).dataset.clave || numFac;
  const res = await fetch('/aprobaciones-ap/api/accion',{method:'POST',headers:{'Content-Type':'application/json'},
    body: JSON.stringify({clave:clave,numero_factura:numFac,accion:tipo,comentario:c,departamento:dept})});
  const d = await res.json();
  if(d.ok) {
    const card = document.getElementById('card-'+i);
    card.style.transition='opacity .3s,transform .3s';
    card.style.opacity='0';
    card.style.transform='translateX('+(tipo==='APROBADA'?'20px':'-20px')+')';
    setTimeout(()=>{card.style.display='none'},300);
    cargarHist();
    showToast(tipo==='APROBADA'?'✓ Aprobada y guardada':'✗ Rechazada y guardada', tipo==='APROBADA'?'#16a34a':'#dc2626');
    loadData();
  } else {
    showToast('✗ '+(d.error||'No se ha podido guardar'), '#dc2626');
  }
}

async function cargarHist() {
  // del SERVIDOR, no de memoria: antes era un array del navegador y al recargar
  // la pagina el historial se vaciaba aunque el dato estuviera guardado.
  let h = [];
  try { h = await (await fetch('/aprobaciones-ap/api/historial')).json(); } catch(e) {}
  historial.length = 0; h.forEach(x=>historial.push(x));
  renderHist();
}

function renderHist() {
  const el=document.getElementById('hist-list');
  document.getElementById('hist-empty').style.display=historial.length?'none':'';
  el.innerHTML=historial.map(h=>{
    const ap = h.accion==='APROBADA';
    // una entrada rectificada despues deja de ser la que manda: se dice
    const viejo = (h.vigente===false) ? ' <span class="hist-old">rectificada</span>' : '';
    return '<div class="hist-item'+(h.vigente===false?' pasada':'')+'">' +
      '<div class="hist-icon '+(ap?'a':'r')+'">'+(ap?'✓':'✗')+'</div>' +
      '<div class="hist-info"><div class="n">'+txt(h.numero_factura)+viejo+'</div>' +
        '<div class="d">'+txt(h.fecha_hora)+(h.departamento?' · '+txt(h.departamento):'')+
        (h.comentario?' · '+txt(h.comentario):'')+'</div></div>' +
      '<span class="hist-accion '+(ap?'a':'r')+'">'+txt(h.accion)+'</span></div>';
  }).join('');
}

function showToast(msg, color) {
  const t=document.getElementById('toast');
  t.textContent=msg; t.style.background=color; t.classList.add('on');
  setTimeout(()=>t.classList.remove('on'),2500);
}

loadData();
cargarHist();
</script>
<script src="/static/yve-icons.js?v=__ASSETS_V__"></script>
</body>
</html>"""
