"""
lector_facturas_ap.py — Yve.01 Módulo AP
Lee facturas de proveedores (no OTAs) con Claude API para extracción estructurada.
Ejecutar: python lector_facturas_ap.py
"""

import os, glob, json, re
from datetime import date
import pdfplumber
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
ENTRADA_DIR      = os.path.join(BASE_DIR, "facturas-entrada")
SALIDA_DIR       = os.path.join(BASE_DIR, "facturas-procesadas")
REFERENCIA_DIR   = os.path.join(BASE_DIR, "datos-referencia")
os.makedirs(SALIDA_DIR, exist_ok=True)

FECHA_HOY      = date.today().strftime("%Y%m%d")
SALIDA_EXCEL   = os.path.join(SALIDA_DIR, f"facturas_ap_{FECHA_HOY}.xlsx")
PROV_FILE      = os.path.join(REFERENCIA_DIR, "proveedores.xlsx")
NF             = "NO_ENCONTRADO"

OTAS_CONOCIDAS = {"booking.com","booking.es","expedia","hotels.com","despegar",
                  "airbnb","agoda","trip.com","trivago","hrs"}

# ── Cargar proveedores ────────────────────────────────────────────────────

def cargar_proveedores():
    if not os.path.exists(PROV_FILE):
        return {}
    df = pd.read_excel(PROV_FILE)
    return {row["nombre_proveedor"].strip().lower(): row.to_dict()
            for _, row in df.iterrows()}

# ── Extracción de texto ───────────────────────────────────────────────────

def extraer_texto(pdf_path):
    textos = []
    with pdfplumber.open(pdf_path) as pdf:
        for pag in pdf.pages:
            t = pag.extract_text()
            if t:
                textos.append(t)
    return "\n".join(textos)

def es_ota(texto):
    txt_lower = texto.lower()
    return any(ota in txt_lower for ota in OTAS_CONOCIDAS)

# ── Extracción con Claude API ─────────────────────────────────────────────

def extraer_con_claude(texto, nombre_archivo):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("    AVISO: ANTHROPIC_API_KEY no encontrada — usando extracción regex")
        return extraer_con_regex(texto)
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        prompt = (
            "Eres un asistente de contabilidad hotelera. Extrae los campos de esta factura "
            "y devuelve EXCLUSIVAMENTE un objeto JSON con estas claves exactas:\n"
            "numero_factura, fecha, nombre_proveedor, NIF_proveedor, descripcion_concepto, "
            "base_imponible, porcentaje_iva, cuota_iva, total_factura\n\n"
            "Reglas:\n"
            "- base_imponible, cuota_iva y total_factura deben ser números decimales (float)\n"
            "- porcentaje_iva debe ser número (ej: 21, no '21%')\n"
            "- fecha en formato DD/MM/YYYY\n"
            "- Si no encuentras un campo usa null\n"
            "- Devuelve SOLO el JSON, sin explicaciones\n\n"
            f"TEXTO DE LA FACTURA:\n{texto[:3000]}"
        )
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role":"user","content":prompt}]
        )
        raw = resp.content[0].text.strip()
        # Limpiar posible markdown
        raw = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
        datos = json.loads(raw)
        return datos
    except Exception as e:
        print(f"    AVISO Claude API: {e} — usando regex")
        return extraer_con_regex(texto)

# ── Extracción fallback con regex ─────────────────────────────────────────

def _num(s):
    try:
        s = str(s).replace("EUR","").replace("€","").replace(" ","").replace("\xa0","").strip()
        if "," in s and "." in s:
            s = s.replace(",","") if s.rfind(".") > s.rfind(",") else s.replace(".","").replace(",",".")
        elif "," in s:
            s = s.replace(",",".") if not re.search(r",\d{3}$", s) else s.replace(",","")
        return float(s)
    except Exception:
        return None

def extraer_con_regex(texto):
    datos = {k: None for k in ["numero_factura","fecha","nombre_proveedor","NIF_proveedor",
                                "descripcion_concepto","base_imponible","porcentaje_iva",
                                "cuota_iva","total_factura"]}
    # Número de factura
    m = re.search(r"(?:factura|invoice|n[uú]m\.?|number)[:\s#]*([A-Z0-9][\w\-\/]{2,25})", texto, re.I)
    if m: datos["numero_factura"] = m.group(1).strip()

    # Fecha
    m = re.search(r"(?:fecha|date)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", texto, re.I)
    if m: datos["fecha"] = m.group(1).strip()

    # NIF / CIF
    m = re.search(r"\b([A-Z]\d{8}|\d{8}[A-Z])\b", texto)
    if m: datos["NIF_proveedor"] = m.group(1)

    # IVA porcentaje
    m = re.search(r"(?:IVA|VAT)[^\n%]{0,20}?(\d{1,2})\s*%", texto, re.I)
    if m: datos["porcentaje_iva"] = float(m.group(1))

    # Importes — buscar patrones EUR/€ seguidos de número
    _EUR = r"(?:EUR|€)\s*"
    _AMT = r"([0-9]{1,3}(?:[.,][0-9]{3})*[.,][0-9]{2}|[0-9]+[.,][0-9]{2})"
    # Base imponible
    m = re.search(r"(?:base\s+imponible|subtotal|base)[:\s]+" + _EUR + _AMT, texto, re.I)
    if m: datos["base_imponible"] = _num(m.group(1))
    # Total
    m = re.search(r"(?:total\s+(?:factura|a\s+pagar|invoice)|importe\s+total)[:\s]+" + _EUR + _AMT, texto, re.I)
    if m: datos["total_factura"] = _num(m.group(1))
    # Cuota IVA
    if datos["base_imponible"] and datos["porcentaje_iva"] and not datos["cuota_iva"]:
        datos["cuota_iva"] = round(datos["base_imponible"] * datos["porcentaje_iva"] / 100, 2)

    return datos

# ── Clasificación de proveedor ────────────────────────────────────────────

def clasificar_proveedor(nombre_proveedor, proveedores):
    if not nombre_proveedor:
        return "OTRAS", NF
    norm = str(nombre_proveedor).strip().lower()
    # Búsqueda exacta
    if norm in proveedores:
        p = proveedores[norm]
        return p.get("tipo","OTRAS"), p.get("cuenta_contable",NF)
    # Búsqueda parcial
    for key, p in proveedores.items():
        if key in norm or norm in key:
            return p.get("tipo","OTRAS"), p.get("cuenta_contable",NF)
    return "OTRAS", NF

# ── Procesado principal ───────────────────────────────────────────────────

def procesar_factura_ap(pdf_path, proveedores):
    nombre = os.path.basename(pdf_path)
    print(f"  Procesando: {nombre}")
    try:
        texto = extraer_texto(pdf_path)
    except Exception as e:
        print(f"    ERROR al leer PDF: {e}")
        return {"archivo": nombre, "error": str(e)}

    if es_ota(texto):
        print(f"    [SKIP] Es factura de OTA — omitida")
        return None

    datos = extraer_con_claude(texto, nombre)

    tipo_prov, cuenta = clasificar_proveedor(datos.get("nombre_proveedor"), proveedores)

    resultado = {
        "archivo":            nombre,
        "numero_factura":     datos.get("numero_factura") or NF,
        "fecha":              datos.get("fecha") or NF,
        "nombre_proveedor":   datos.get("nombre_proveedor") or NF,
        "NIF_proveedor":      datos.get("NIF_proveedor") or NF,
        "descripcion_concepto": datos.get("descripcion_concepto") or NF,
        "base_imponible":     datos.get("base_imponible") or NF,
        "porcentaje_iva":     datos.get("porcentaje_iva") or NF,
        "cuota_iva":          datos.get("cuota_iva") or NF,
        "total_factura":      datos.get("total_factura") or NF,
        "tipo_proveedor":     tipo_prov,
        "cuenta_contable":    cuenta,
        "error":              "",
    }

    for k, v in resultado.items():
        if k != "error":
            icono = "✓" if v not in (NF, None, "") else "✗"
            if k in ("archivo","tipo_proveedor","cuenta_contable","error"):
                continue
            print(f"    [{icono}] {k}: {v}")
    return resultado

def guardar_excel(registros, ruta):
    df = pd.DataFrame(registros)
    cols = ["archivo","numero_factura","fecha","nombre_proveedor","NIF_proveedor",
            "descripcion_concepto","base_imponible","porcentaje_iva","cuota_iva",
            "total_factura","tipo_proveedor","cuenta_contable","error"]
    cols = [c for c in cols if c in df.columns]
    df = df[cols]
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas_AP")
        ws = writer.sheets["Facturas_AP"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col)+4, 40)
    print(f"\n✅ Excel guardado: {ruta}")

def main():
    print("="*60)
    print("  Yve.01 — Lector de Facturas AP")
    print("="*60)
    proveedores = cargar_proveedores()
    print(f"  Proveedores en tabla: {len(proveedores)}")

    pdfs = sorted(f for f in glob.glob(os.path.join(ENTRADA_DIR,"*.pdf")))
    if not pdfs:
        print(f"\n⚠️  No hay PDFs en {ENTRADA_DIR}")
        return

    print(f"  PDFs encontrados: {len(pdfs)}\n")
    registros = []
    omitidas = 0
    for pdf in pdfs:
        res = procesar_factura_ap(pdf, proveedores)
        if res is None:
            omitidas += 1
        else:
            registros.append(res)
        print()

    if not registros:
        print("No hay facturas AP para procesar (todas eran OTAs o errores).")
        return

    guardar_excel(registros, SALIDA_EXCEL)
    print(f"\n  Total AP procesadas: {len(registros)}")
    print(f"  Omitidas (OTA):      {omitidas}")
    print("="*60)

if __name__ == "__main__":
    main()
