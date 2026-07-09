"""
lector_contratos_grupo.py — Yve.01
Procesa CONTRATOS DE GRUPO / EVENTOS (Hilton BEO + contrato de grupos) a partir
de fotos/escaneos de varias páginas, y los convierte en:
  · Cliente corporativo (AR Real)
  · Factura/receivable (habitaciones + F&B + salas)  -> reservas_credito.xlsx
  · Registro de COMISIÓN que el hotel paga a la agencia (base sin IVA × %)
  · Alerta de certificado de doble imposición (cliente extranjero)

Extracción: API de visión de Claude (multipágina). Si no hay API/red,
degrada de forma elegante marcando el contrato como "pendiente de revisión"
(nunca lanza excepción hacia arriba).
"""
import os, json, base64, glob, mimetypes
from datetime import datetime

MODEL = "claude-sonnet-4-6"

# Esquema que pedimos a la visión (una sola pasada con todas las páginas)
_PROMPT = """Eres un experto en contratación hotelera. Te paso las fotos de TODAS las páginas
de un contrato de grupo/eventos de hotel (contrato + BEO + anexos). Devuelve SOLO un JSON válido
(sin texto alrededor) con esta estructura exacta (usa null si un dato no aparece):
{
 "evento": {"id": "", "nombre": ""},
 "contrato_numero": "",
 "fecha_contrato": "YYYY-MM-DD",
 "hotel": {"nombre": "", "cif": ""},
 "cliente": {"nombre": "", "cif": "", "email": "", "contacto": "", "pais": ""},
 "agencia": {"nombre": "", "contacto": "", "email": "", "telefono": ""},
 "alojamiento": {"fecha_entrada":"YYYY-MM-DD","fecha_salida":"YYYY-MM-DD","noches":0,
                 "habitaciones":0,"tarifa_dui":0,"tarifa_doble":0,
                 "total_habitaciones":0,"iva_pct":10,"total_pernoctaciones":0},
 "tasa_turistica": {"por_persona_noche":0,"max_noches":0},
 "fb": {"total":0,"por_persona_dia":0,"pax":0,"dias":0,"iva_pct":10,"detalle":""},
 "salas": {"total":0},
 "comisiones": {"alojamiento_pct":0,"salas_pct":0,"fb_pct":0,"ddr_pct":0,"misc_pct":0},
 "deposito": {"pct":0,"cuando":"","iban":"","beneficiario":"","referencia":""},
 "doble_imposicion": false,
 "beos": []
}
Importante: los importes son numéricos (sin símbolo €, punto decimal). Las fechas en formato ISO.
"doble_imposicion" = true si el cliente es extranjero o el contrato menciona doble imposición / withholding / certificado de residencia fiscal."""


def _api_key():
    return os.environ.get("ANTHROPIC_API_KEY", "") or _env_file_key()


def _env_file_key():
    try:
        p = os.path.join(os.path.dirname(__file__), ".env")
        if os.path.exists(p):
            for ln in open(p, encoding="utf-8"):
                if ln.strip().startswith("ANTHROPIC_API_KEY"):
                    return ln.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return ""


def _img_block(path):
    mt = mimetypes.guess_type(path)[0] or "image/jpeg"
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode()
    return {"type": "image", "source": {"type": "base64", "media_type": mt, "data": data}}


def extraer_contrato_grupo(image_paths):
    """Extrae los datos del contrato con la visión de Claude. Nunca lanza; si falla
    devuelve {'_needs_review': True, '_error': ...}."""
    key = _api_key()
    if not key:
        return {"_needs_review": True, "_error": "sin ANTHROPIC_API_KEY (extracción no disponible en este entorno)"}
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        content = [{"type": "text", "text": _PROMPT}]
        for p in image_paths[:30]:
            content.append(_img_block(p))
        resp = client.messages.create(model=MODEL, max_tokens=1500,
                                      messages=[{"role": "user", "content": content}])
        txt = resp.content[0].text.strip()
        # aislar el JSON
        i, j = txt.find("{"), txt.rfind("}")
        datos = json.loads(txt[i:j + 1])
        datos["_needs_review"] = False
        return datos
    except Exception as e:
        return {"_needs_review": True, "_error": str(e)[:200]}


def _f(v, d=0.0):
    try:
        if v in (None, "", "null"):
            return d
        return float(str(v).replace("€", "").replace(",", "").strip())
    except Exception:
        return d


def calcular_comisiones(datos):
    """Comisión que el hotel paga a la agencia = base SIN IVA × % por concepto."""
    def base(total, iva):
        total = _f(total); iva = _f(iva, 10)
        return total / (1 + iva / 100) if total else 0.0
    aloj = datos.get("alojamiento", {}) or {}
    fb = datos.get("fb", {}) or {}
    salas = datos.get("salas", {}) or {}
    com = datos.get("comisiones", {}) or {}
    base_aloj = base(aloj.get("total_habitaciones"), aloj.get("iva_pct", 10))
    base_fb = base(fb.get("total"), fb.get("iva_pct", 10))
    base_salas = base(salas.get("total"), 21)
    c_aloj = round(base_aloj * _f(com.get("alojamiento_pct")) / 100, 2)
    c_fb = round(base_fb * _f(com.get("fb_pct")) / 100, 2)
    c_salas = round(base_salas * _f(com.get("salas_pct")) / 100, 2)
    return {
        "alojamiento": c_aloj, "fb": c_fb, "salas": c_salas,
        "total": round(c_aloj + c_fb + c_salas, 2),
        "base_alojamiento": round(base_aloj, 2), "base_fb": round(base_fb, 2),
    }


def transformar(datos):
    """Convierte los datos extraídos en filas para AR Real + comisión + flag DI.
    Testeable sin API."""
    ev = datos.get("evento", {}) or {}
    cli = datos.get("cliente", {}) or {}
    aloj = datos.get("alojamiento", {}) or {}
    fb = datos.get("fb", {}) or {}
    salas = datos.get("salas", {}) or {}
    contrato = str(datos.get("contrato_numero") or "").strip()

    imp_hab = _f(aloj.get("total_habitaciones"))
    imp_fb = _f(fb.get("total"))
    imp_salas = _f(salas.get("total"))
    total = round(imp_hab + imp_fb + imp_salas, 2)
    comis = calcular_comisiones(datos)
    di = bool(datos.get("doble_imposicion")) or (str(cli.get("pais", "")).lower() not in ("", "españa", "espana", "spain"))

    nombre_cli = (cli.get("nombre") or "Cliente grupo").strip()
    cliente_row = {
        "nombre_cliente": nombre_cli, "NIF": cli.get("cif") or "", "nif": cli.get("cif") or "",
        "email": cli.get("email") or "", "telefono": (datos.get("agencia", {}) or {}).get("telefono", "") or "",
        "dias_pago": 30, "limite_credito": 100000, "credito_limite": 100000, "credito_usado": 0,
    }
    numero = f"GRP-{contrato}" if contrato else f"GRP-{datetime.now().strftime('%Y%m%d%H%M')}"
    noches = int(_f(aloj.get("noches"))) or 0
    reserva_row = {
        "numero_reserva": numero, "numero": numero, "cliente": nombre_cli,
        "fecha_entrada": aloj.get("fecha_entrada") or "", "fecha_salida": aloj.get("fecha_salida") or "",
        "fecha_emision": "", "habitaciones": int(_f(aloj.get("habitaciones"))), "noches": noches,
        "importe_habitaciones": imp_hab, "importe_fb": imp_fb, "importe_extras": imp_salas,
        "importe": total, "total": total, "estado": "PENDIENTE_FACTURA",
        "evento": (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip(),
        "contrato": contrato, "comision_total": comis["total"],
        "requiere_certificado_di": di, "tipo": "CONTRATO_GRUPO",
    }
    return {
        "cliente": cliente_row, "reserva": reserva_row, "comisiones": comis,
        "doble_imposicion": di,
        "resumen": {
            "evento": reserva_row["evento"], "contrato": contrato, "cliente": nombre_cli,
            "total_receivable": total, "habitaciones": imp_hab, "fb": imp_fb, "salas": imp_salas,
            "comision_total": comis["total"], "requiere_certificado_di": di,
            "numero": numero,
        },
    }


def _datos_dir():
    try:
        from tenant_dirs import datos_dir
        return datos_dir()
    except Exception:
        return os.path.join(os.path.dirname(__file__), "datos-referencia")


def guardar(transformado, datos_dir=None):
    """Añade/actualiza cliente y reserva en los xlsx de AR Real. Dedup por nombre/numero."""
    import pandas as pd
    dd = datos_dir or _datos_dir()
    os.makedirs(dd, exist_ok=True)
    # Clientes
    pc = os.path.join(dd, "clientes_credito.xlsx")
    dfc = pd.read_excel(pc) if os.path.exists(pc) else pd.DataFrame()
    nom = transformado["cliente"]["nombre_cliente"]
    if not (len(dfc) and (dfc.get("nombre_cliente", pd.Series(dtype=str)).astype(str) == nom).any()):
        dfc = pd.concat([dfc, pd.DataFrame([transformado["cliente"]])], ignore_index=True)
        dfc.to_excel(pc, index=False)
    # Reservas / facturas
    pr = os.path.join(dd, "reservas_credito.xlsx")
    dfr = pd.read_excel(pr) if os.path.exists(pr) else pd.DataFrame()
    num = transformado["reserva"]["numero_reserva"]
    col = "numero_reserva" if "numero_reserva" in dfr.columns else "numero"
    if len(dfr) and col in dfr.columns:
        dfr = dfr[dfr[col].astype(str) != num]  # reemplaza si ya existía
    dfr = pd.concat([dfr, pd.DataFrame([transformado["reserva"]])], ignore_index=True)
    dfr.to_excel(pr, index=False)
    return {"clientes": pc, "reservas": pr}


def procesar_contrato_grupo(image_paths, datos_dir=None, guardar_datos=True):
    """Pipeline completo: extraer -> transformar -> guardar. Devuelve resumen."""
    if isinstance(image_paths, str):
        if os.path.isdir(image_paths):
            image_paths = sorted(glob.glob(os.path.join(image_paths, "*")))
        else:
            image_paths = [image_paths]
    image_paths = [p for p in image_paths if str(p).lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".heic"))]
    datos = extraer_contrato_grupo(image_paths)
    if datos.get("_needs_review"):
        return {"ok": False, "needs_review": True, "error": datos.get("_error", ""),
                "message": "No se pudo extraer automáticamente; guardado para revisión manual."}
    t = transformar(datos)
    if guardar_datos:
        t["_paths"] = guardar(t, datos_dir)
    r = t["resumen"]; r["ok"] = True
    return r


if __name__ == "__main__":
    import sys
    print(json.dumps(procesar_contrato_grupo(sys.argv[1] if len(sys.argv) > 1 else ".", guardar_datos=False),
                     ensure_ascii=False, indent=2))
