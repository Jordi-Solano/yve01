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
    """Extrae texto de PDF. Si el PDF es escaneado (sin texto), intenta OCR."""
    textos = []
    with pdfplumber.open(pdf_path) as pdf:
        for pag in pdf.pages:
            t = pag.extract_text()
            if t:
                textos.append(t)
    texto = "\n".join(textos)
    
    # Si el PDF tiene poco texto, puede ser escaneado — intentar OCR
    if len(texto.strip()) < 100:
        try:
            import subprocess
            # pytesseract via subprocess (más compatible)
            r = subprocess.run(
                ['python3', '-c', f'''
import pdfplumber, io
from PIL import Image
with pdfplumber.open("{pdf_path}") as pdf:
    for p in pdf.pages:
        img = p.to_image(resolution=200)
        # Guardar como imagen temporal
        img.save("/tmp/yve_ocr_page.png")
        break
'''],
                capture_output=True, text=True, timeout=15
            )
            if os.path.exists("/tmp/yve_ocr_page.png"):
                r2 = subprocess.run(
                    ['tesseract', '/tmp/yve_ocr_page.png', 'stdout', '-l', 'spa+eng'],
                    capture_output=True, text=True, timeout=15
                )
                if r2.stdout.strip():
                    texto = r2.stdout.strip()
                os.remove("/tmp/yve_ocr_page.png")
        except Exception:
            pass  # OCR no disponible — continuar con lo que tengamos
    
    return texto

def es_ota(texto):
    txt_lower = texto.lower()
    return any(ota in txt_lower for ota in OTAS_CONOCIDAS)


# Documentos que NO son facturas — pre-filtro por nombre de archivo
NO_FACTURA_KEYWORDS = {
    'rooming', 'room list', 'room block', 'guest list',
    'agenda', 'logo', 'signage', 'banner',
    'beo', 'banquet event', 'banquet order',
    'sow', 'statement of work', 'scope of work',
    'contract', 'contrato', 'acuerdo',
    'proposal', 'presupuesto', 'quotation', 'quote',
    'menu ', 'wine list', 'carta de',
    'floorplan', 'floor plan', 'setup', 'plano',
    'technical manual', ' tm ', '_tm_', ' tm.',
    'resume', 'presentation', 'powerpoint', 'ppt',
    'meeting notes', 'acta ', 'minuta',
    'checklist', 'planning', 'schedule', 'timeline',
    'itinerary', 'itinerario', 'programa',
    'certificate', 'certificado', 'diploma',
    'running order', 'master onsite', 'event order',
}

NO_FACTURA_EXTENSIONS = {'.doc', '.docx', '.ppt', '.pptx', '.png', '.jpg', '.jpeg', '.gif', '.svg', '.mp4', '.zip', '.rar'}


def es_no_factura_por_nombre(nombre_archivo):
    """Pre-filtro rápido por nombre de archivo — evita gastar tokens de Claude."""
    nombre_lower = nombre_archivo.lower()
    ext = os.path.splitext(nombre_lower)[1]
    if ext in NO_FACTURA_EXTENSIONS:
        return True, f"extensión {ext} no es factura"
    for kw in NO_FACTURA_KEYWORDS:
        if kw in nombre_lower:
            return True, f"contiene '{kw}'"
    return False, ""


def es_no_factura_por_contenido(texto):
    """Pre-filtro CONSERVADOR. Solo bloquea si está MUY claro que no es factura.
    En caso de duda, deja pasar a Claude (que es mejor juez)."""
    if not texto or len(texto.strip()) < 30:
        return True, "documento vacío (posible PDF escaneado o imagen)"
    txt_lower = texto[:3000].lower()
    
    # Indicadores FUERTES de que SÍ es factura → nunca bloquear
    factura_signals = ['factura', 'invoice', 'rechnung', 'fattura', 'facture',
                       'base imponible', 'iva', 'vat', 'mwst', 'tva',
                       'total a pagar', 'importe total', 'amount due', 'total due',
                       'nif', 'cif', 'tax id', 'deposit', 'depósito', 'anticipo',
                       'fecha de emisión', 'invoice number', 'invoice no',
                       'payment terms', 'forma de pago', 'bank transfer',
                       'iban', 'subtotal', 'net amount', 'gross amount',
                       'proforma', 'pro forma', 'advance payment', '€', 'eur ',
                       'importe', 'precio', 'amount', 'price', 'fee']
    if any(s in txt_lower for s in factura_signals):
        return False, ""
    
    # Solo bloquear si hay señales MUY claras de no-factura Y ninguna de factura
    no_factura_signals = ['banquet event order', 'rooming list',
                          'meeting room setup', 'floor plan layout',
                          'technical rider', 'audio visual requirements']
    for s in no_factura_signals:
        if s in txt_lower:
            return True, f"contiene '{s}'"
    
    # Si no hay señales claras de nada, dejar pasar a Claude
    return False, ""

# ── Extracción con Claude API ─────────────────────────────────────────────

def extraer_con_claude(texto, nombre_archivo):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("    AVISO: ANTHROPIC_API_KEY no encontrada — usando extracción regex")
        return extraer_con_regex(texto)
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        # Enviar más texto si el documento es largo (facturas grandes)
        max_chars = min(len(texto), 6000)
        prompt = (
            "Eres un experto en contabilidad hotelera. Analiza este documento.\n\n"
            "PASO 1: ¿Es una factura real (invoice) con importes económicos?\n"
            "Un documento ES factura si tiene: número de factura, fecha, importes, IVA/VAT.\n"
            "Un documento NO es factura si es: email, contrato, BEO, rooming list, agenda, "
            "manual técnico, propuesta, presupuesto sin comprometer, certificado, acta.\n"
            "Un depósito/anticipo (deposit/advance payment) SÍ es factura.\n"
            "Una factura proforma SÍ es factura.\n\n"
            'Si NO es factura, devuelve SOLO: {"es_factura": false, "tipo_documento": "descripcion breve"}\n\n'
            "PASO 2: Si SÍ es factura, extrae TODOS los campos posibles:\n"
            '{"es_factura": true, '
            '"numero_factura": "string o null", '
            '"fecha": "DD/MM/YYYY o null", '
            '"nombre_proveedor": "nombre empresa emisora", '
            '"NIF_proveedor": "NIF/CIF/VAT ID o null", '
            '"descripcion_concepto": "resumen en español de qué se factura", '
            '"base_imponible": 0.00, '
            '"porcentaje_iva": 21, '
            '"cuota_iva": 0.00, '
            '"total_factura": 0.00, '
            '"moneda": "EUR"}\n\n'
            "REGLAS:\n"
            "- Importes SIEMPRE como float (1234.56), nunca strings\n"
            "- Si solo ves el total sin desglose IVA: base=total/1.21, iva=21, cuota=total-base\n"
            "- Si el IVA es 0% (intracomunitaria/export): base=total, iva=0, cuota=0\n"
            "- porcentaje_iva: número entero (21, 10, 0), no decimal ni string\n"
            "- Monedas: EUR, USD, GBP, PLN, etc. — extrae los números sin convertir\n"
            "- Si hay un depósito parcial (1st deposit, anticipo), el total es el depósito\n"
            "- Si no encuentras un campo, pon null — NUNCA inventes datos\n"
            "- SOLO JSON, sin markdown, sin explicaciones, sin ```\n\n"
            f"ARCHIVO: {nombre_archivo}\n"
            f"TEXTO:\n{texto[:max_chars]}"
        )
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role":"user","content":prompt}]
        )
        raw = resp.content[0].text.strip()
        # Limpiar posible markdown y texto extra
        raw = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
        # A veces Claude añade texto antes/después del JSON
        json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', raw)
        if json_match:
            raw = json_match.group()
        
        try:
            datos = json.loads(raw)
        except json.JSONDecodeError:
            # Intentar arreglar JSON común: trailing commas, single quotes
            fixed = raw.replace("'", '"').rstrip(',').rstrip(',}') + '}'
            fixed = re.sub(r',\s*}', '}', fixed)
            fixed = re.sub(r',\s*]', ']', fixed)
            try:
                datos = json.loads(fixed)
            except json.JSONDecodeError:
                print(f"    AVISO: JSON inválido de Claude — usando regex")
                print(f"    Raw: {raw[:200]}")
                return extraer_con_regex(texto)
        
        # Si Claude dice que no es factura, devolver None
        if datos.get("es_factura") is False:
            tipo_doc = datos.get("tipo_documento", "desconocido")
            print(f"    [INFO] Tipo: {tipo_doc}")
            return None
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
    """Clasifica proveedor con matching fuzzy en 3 niveles."""
    if not nombre_proveedor:
        return "OTRAS", NF
    norm = str(nombre_proveedor).strip().lower()
    
    # Nivel 1: Búsqueda exacta
    if norm in proveedores:
        p = proveedores[norm]
        return p.get("tipo","OTRAS"), str(p.get("cuenta_contable",NF))
    
    # Nivel 2: Búsqueda parcial (substring)
    for key, p in proveedores.items():
        if key in norm or norm in key:
            return p.get("tipo","OTRAS"), str(p.get("cuenta_contable",NF))
    
    # Nivel 3: Matching por palabras clave (al menos 2 palabras coinciden)
    norm_words = set(norm.replace(',','').replace('.','').split())
    for key, p in proveedores.items():
        key_words = set(key.replace(',','').replace('.','').split())
        common = norm_words & key_words
        # Ignorar palabras genéricas
        common -= {'sl', 'sa', 'slu', 'sll', 'de', 'del', 'la', 'el', 'los', 'las', 'y', 'e', 'ltd', 'gmbh', 'sp', 'zoo'}
        if len(common) >= 2:
            return p.get("tipo","OTRAS"), str(p.get("cuenta_contable",NF))
    
    return "OTRAS", NF

# ── Helpers de validación ────────────────────────────────────────────

def _safe_float(v):
    """Convierte a float de forma segura, devuelve None si no puede."""
    if v is None or v == NF or v == '':
        return None
    try:
        if isinstance(v, str):
            v = v.replace('€','').replace('$','').replace(' ','').replace(' ','')
            v = v.replace(',','.') if ',' in v and '.' not in v else v
        return round(float(v), 2)
    except (ValueError, TypeError):
        return None


def _auto_cuenta_pgc(concepto, proveedor=None):
    """Asigna automáticamente la cuenta contable PGC según el concepto."""
    if not concepto:
        return NF
    c = (concepto + ' ' + (proveedor or '')).lower()
    
    # F&B / alimentación
    if any(x in c for x in ['aliment','comida','bebida','food','beverage','catering',
                             'restaura','cocina','menu','coffee','café','bar ','minibar']):
        return '600'  # Compras mercaderías
    # Limpieza
    if any(x in c for x in ['limpieza','cleaning','lavandera','lavandería','housekeeping']):
        return '629'  # Otros servicios
    # Energía / suministros
    if any(x in c for x in ['electric','energía','gas ','agua ','water','utility','suministro']):
        return '628'  # Suministros
    # Mantenimiento
    if any(x in c for x in ['mantenimiento','maintenance','reparac','repair','conservac']):
        return '622'  # Reparaciones y conservación
    # Seguros
    if any(x in c for x in ['seguro','insurance','póliza']):
        return '625'  # Primas de seguros
    # Alquiler
    if any(x in c for x in ['alquiler','rent','arrendamiento','leasing']):
        return '621'  # Arrendamientos y cánones
    # Profesionales / consultoría
    if any(x in c for x in ['consultor','asesor','abogado','legal','audit','contab','lawyer',
                             'advisory','consulting','professional']):
        return '623'  # Servicios profesionales
    # Publicidad / marketing
    if any(x in c for x in ['publicidad','marketing','advertising','promo','campaign','diseño']):
        return '627'  # Publicidad y propaganda
    # Telecomunicaciones
    if any(x in c for x in ['telefon','telecom','internet','wifi','fibra','mobile']):
        return '629'  # Otros servicios
    # Transporte
    if any(x in c for x in ['transport','courier','mensajer','envío','shipping','logistic']):
        return '624'  # Transportes
    # Comisiones OTA / agencias
    if any(x in c for x in ['comisión','commission','booking','expedia','agencia','ota']):
        return '628'  # Comisiones agencias
    # Eventos
    if any(x in c for x in ['event','evento','congres','conferenc','meeting','audiovisual',
                             'decorac','flores','flower','signage','producción']):
        return '629'  # Otros servicios
    
    return '629'  # Default: otros servicios



# ── Procesado principal ───────────────────────────────────────────────────

def procesar_factura_ap(pdf_path, proveedores):
    nombre = os.path.basename(pdf_path)
    print(f"  Procesando: {nombre}")

    # Pre-filtro 1: por nombre de archivo (gratis, sin tokens)
    skip, motivo = es_no_factura_por_nombre(nombre)
    if skip:
        print(f"    [SKIP] {nombre}: documento no procesable — {motivo}")
        return None

    try:
        texto = extraer_texto(pdf_path)
    except Exception as e:
        print(f"    ERROR al leer PDF: {e}")
        return {"archivo": nombre, "error": str(e)}

    if es_ota(texto):
        print(f"    [SKIP] Es factura de OTA — omitida")
        return None

    # Pre-filtro 2: por contenido del PDF (gratis, sin tokens)
    skip2, motivo2 = es_no_factura_por_contenido(texto)
    if skip2:
        print(f"    [SKIP] {nombre}: documento no procesable — {motivo2}")
        return None

    datos = extraer_con_claude(texto, nombre)
    
    # Si Claude dice que no es factura
    if datos is None:
        print(f"    [SKIP] {nombre}: no contiene datos financieros extraíbles")
        return None

    tipo_prov, cuenta = clasificar_proveedor(datos.get("nombre_proveedor"), proveedores)

    # ── Validar y auto-calcular campos ──────────────────────────────────
    base = _safe_float(datos.get("base_imponible"))
    iva_pct = _safe_float(datos.get("porcentaje_iva"))
    cuota = _safe_float(datos.get("cuota_iva"))
    total = _safe_float(datos.get("total_factura"))

    # Auto-cálculo: si falta algún campo, intentar derivarlo
    if total and not base and iva_pct:
        base = round(total / (1 + iva_pct/100), 2)
        cuota = round(total - base, 2)
    elif total and not base and not iva_pct:
        # Asumir IVA 21% España si no se especifica
        iva_pct = 21
        base = round(total / 1.21, 2)
        cuota = round(total - base, 2)
    elif base and iva_pct and not total:
        cuota = round(base * iva_pct / 100, 2)
        total = round(base + cuota, 2)
    elif base and not cuota and iva_pct:
        cuota = round(base * iva_pct / 100, 2)
        if not total:
            total = round(base + cuota, 2)
    elif base and total and not iva_pct:
        cuota = round(total - base, 2)
        iva_pct = round(cuota / base * 100) if base > 0 else 0

    # Validación: importes deben ser positivos y razonables
    if total and total < 0:
        total = abs(total)  # Facturas negativas → abono
    if base and base < 0:
        base = abs(base)

    # Auto-asignar cuenta contable PGC por tipo de concepto
    if cuenta == NF and datos.get("descripcion_concepto"):
        cuenta = _auto_cuenta_pgc(datos.get("descripcion_concepto"), datos.get("nombre_proveedor"))

    resultado = {
        "archivo":            nombre,
        "numero_factura":     datos.get("numero_factura") or NF,
        "fecha":              datos.get("fecha") or NF,
        "nombre_proveedor":   datos.get("nombre_proveedor") or NF,
        "NIF_proveedor":      datos.get("NIF_proveedor") or NF,
        "descripcion_concepto": datos.get("descripcion_concepto") or NF,
        "base_imponible":     base or NF,
        "porcentaje_iva":     iva_pct or NF,
        "cuota_iva":          cuota or NF,
        "total_factura":      total or NF,
        "tipo_proveedor":     tipo_prov,
        "cuenta_contable":    cuenta,
        "moneda":             datos.get("moneda", "EUR"),
        "error":              "",
    }

    campos_ok = sum(1 for k,v in resultado.items() 
                    if k not in ("archivo","tipo_proveedor","cuenta_contable","moneda","error") 
                    and v not in (NF, None, ""))
    campos_total = 9  # campos de factura
    print(f"    Extraídos: {campos_ok}/{campos_total} campos")
    for k, v in resultado.items():
        if k in ("archivo","tipo_proveedor","cuenta_contable","moneda","error"):
            continue
        icono = "✓" if v not in (NF, None, "") else "✗"
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
            if reg is None:
                # No es una factura — salir con código 2 (distinto de error)
                print(f"documento no procesable — saltando")
                sys.exit(2)
            elif reg and not reg.get("error"):
                ruta_excel = os.path.join(SALIDA_DIR, f"facturas_ap_{FECHA_HOY}.xlsx")
                if os.path.exists(ruta_excel):
                    df_existing = pd.read_excel(ruta_excel)
                    df_new = pd.DataFrame([reg])
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                    # Deduplicar por archivo Y por numero_factura+proveedor
                    df_combined.drop_duplicates(subset=["archivo"], keep="last", inplace=True)
                    if reg.get("numero_factura") and reg["numero_factura"] != NF:
                        mask = (df_combined["numero_factura"] == reg["numero_factura"]) & \
                               (df_combined["nombre_proveedor"] == reg.get("nombre_proveedor",""))
                        if mask.sum() > 1:
                            # Mantener solo la última entrada del duplicado
                            idx_to_drop = df_combined[mask].index[:-1]
                            df_combined = df_combined.drop(idx_to_drop)
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