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

    # ── Bloque estructurado DATOS FACTURA SISTEMA (formato Yve.01 PDFs) ──
    if "DATOS FACTURA SISTEMA:" in texto or "NUMERO_FACTURA=" in texto:
        def _campo(clave):
            m = re.search(rf"^{clave}=(.+)$", texto, re.MULTILINE)
            return m.group(1).strip() if m else None
        datos["numero_factura"]    = _campo("NUMERO_FACTURA")
        datos["fecha"]             = _campo("FECHA")
        datos["nombre_proveedor"]  = _campo("PROVEEDOR")
        datos["NIF_proveedor"]     = _campo("NIF")
        datos["descripcion_concepto"] = _campo("CONCEPTO")
        v = _campo("BASE_IMPONIBLE"); datos["base_imponible"]  = _num(v) if v else None
        v = _campo("IVA_PORCENTAJE"); datos["porcentaje_iva"]  = float(v) if v else None
        v = _campo("CUOTA_IVA");      datos["cuota_iva"]       = _num(v) if v else None
        v = _campo("TOTAL");          datos["total_factura"]   = _num(v) if v else None
        return datos

    # ── Regex genérico (facturas externas) ────────────────────────────────
    # Número de factura
    m = re.search(r"(?:Numero|n[uú]m\.?|number|factura\s+n[oº]?)[:\s#]*([A-Z0-9][\w\-\/]{2,25})", texto, re.I)
    if m: datos["numero_factura"] = m.group(1).strip()

    # Fecha
    m = re.search(r"(?:fecha\s+emision|fecha|date)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", texto, re.I)
    if m: datos["fecha"] = m.group(1).strip()

    # NIF proveedor — buscar etiqueta "NIF proveedor:" (no el del cliente)
    m = re.search(r"NIF\s+proveedor[:\s]+([A-Z0-9\-]{8,12})", texto, re.I)
    if m:
        datos["NIF_proveedor"] = m.group(1).strip()
    else:
        m = re.search(r"\b([A-Z]\-?\d{7,8}|\d{8}[A-Z])\b", texto)
        if m: datos["NIF_proveedor"] = m.group(1)

    # Proveedor — primera línea que no sea vacía ni "FACTURA"
    for linea in texto.splitlines():
        linea = linea.strip()
        if linea and linea.upper() not in ("FACTURA", "FACTURADO A:") and len(linea) > 4:
            datos["nombre_proveedor"] = linea
            break

    # Concepto
    m = re.search(r"DESCRIPCION DEL SERVICIO\s*\n(.+)", texto, re.I)
    if m: datos["descripcion_concepto"] = m.group(1).strip()

    # IVA porcentaje
    m = re.search(r"(?:IVA|VAT|Cuota\s+IVA)\s+(\d{1,2})\s*(?:por\s*ciento|%)", texto, re.I)
    if m: datos["porcentaje_iva"] = float(m.group(1))

    # Base imponible
    m = re.search(r"Base\s+imponible\s+EUR[:\s]*([0-9]+[.,][0-9]{2})", texto, re.I)
    if m: datos["base_imponible"] = _num(m.group(1))

    # Total
    m = re.search(r"TOTAL\s+FACTURA\s+EUR[:\s]*([0-9]+[.,][0-9]{2})", texto, re.I)
    if m: datos["total_factura"] = _num(m.group(1))

    # Cuota IVA calculada si no se encontró
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
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    df.to_excel(ruta, index=False)
    return df


def main():
    print("=" * 60)
    print("  Yve.01 — Lector de Facturas AP")
    print("=" * 60)

    # Cargar proveedores de referencia
    proveedores = {}
    if os.path.exists(PROV_FILE):
        df_prov = pd.read_excel(PROV_FILE)
        for _, row in df_prov.iterrows():
            nombre = str(row.get("nombre_proveedor","")).strip().lower()
            proveedores[nombre] = {
                "tipo": str(row.get("tipo","OTRAS")).strip().upper(),
                "cuenta_contable": str(row.get("cuenta_contable","629")).strip(),
                "email_contacto": str(row.get("email_contacto","")).strip(),
            }
        print(f"  Proveedores en tabla: {len(proveedores)}")

    # Buscar PDFs
    pdfs = sorted(glob.glob(os.path.join(ENTRADA_DIR, "*.pdf")))
    print(f"  PDFs encontrados: {len(pdfs)}")

    registros = []
    skipped   = 0
    for pdf_path in pdfs:
        reg = procesar_factura_ap(pdf_path, proveedores)
        if reg is None:
            skipped += 1
        elif not reg.get("error"):
            registros.append(reg)

    if not registros:
        print("\n  No se procesaron facturas AP.")
        return

    # Guardar Excel
    ruta_excel = os.path.join(SALIDA_DIR, f"facturas_ap_{FECHA_HOY}.xlsx")
    guardar_excel(registros, ruta_excel)
    print(f"\n\u2705 Excel guardado: {ruta_excel}")
    print(f"\n  Total AP procesadas: {len(registros)}")
    print(f"  Omitidas (OTA):      {skipped}")
    print("=" * 60)


if __name__ == "__main__":
    import sys
    # Soporte para --file archivo.pdf (procesar un solo archivo)
    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        if idx + 1 < len(sys.argv):
            file_path = sys.argv[idx + 1]
            if not os.path.exists(file_path):
                print(f"ERROR: archivo no encontrado: {file_path}")
                sys.exit(1)
            proveedores = {}
            if os.path.exists(PROV_FILE):
                df_prov = pd.read_excel(PROV_FILE)
                for _, row in df_prov.iterrows():
                    nombre = str(row.get("nombre_proveedor","")).strip().lower()
                    proveedores[nombre] = {
                        "tipo": str(row.get("tipo","OTRAS")).strip().upper(),
                        "cuenta_contable": str(row.get("cuenta_contable","629")).strip(),
                        "email_contacto": str(row.get("email_contacto","")).strip(),
                    }
            reg = procesar_factura_ap(file_path, proveedores)
            if reg and not reg.get("error"):
                ruta_excel = os.path.join(SALIDA_DIR, f"facturas_ap_{FECHA_HOY}.xlsx")
                if os.path.exists(ruta_excel):
                    df_existing = pd.read_excel(ruta_excel)
                    df_new = pd.DataFrame([reg])
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                    # Deduplicar por ARCHIVO (no por numero_factura porque puede ser NO_ENCONTRADO)
                    df_combined.drop_duplicates(subset=["archivo"], keep="last", inplace=True)
                    df_combined.to_excel(ruta_excel, index=False)
                else:
                    guardar_excel([reg], ruta_excel)
                print(f"OK: {os.path.basename(file_path)} procesado correctamente")
                sys.exit(0)
            else:
                err = reg.get("error", "error desconocido") if reg else "sin resultado"
                print(f"ERROR: {err}")
                sys.exit(1)
    else:
        main()