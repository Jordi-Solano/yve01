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

def _get_ap_client():
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        return None
    try:
        import anthropic
        return anthropic.Anthropic(api_key=key)
    except Exception:
        return None

def extraer_texto(pdf_path):
    """Extrae texto de un PDF o de una FOTO (imagen). Las imágenes se leen con Claude Vision."""
    _ext = os.path.splitext(pdf_path)[1].lower()
    if _ext in ('.jpg', '.jpeg', '.png', '.webp', '.heic'):
        cli = _get_ap_client()
        if cli is None:
            return ""
        try:
            import base64 as _b64, mimetypes as _mt
            media = _mt.guess_type(pdf_path)[0] or 'image/jpeg'
            with open(pdf_path, 'rb') as _f:
                _data = _b64.b64encode(_f.read()).decode()
            resp = cli.messages.create(model="claude-sonnet-4-6", max_tokens=2000,
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media, "data": _data}},
                    {"type": "text", "text": "Extrae TODO el texto visible de este documento (factura, albaran, etc.). Solo el texto, sin explicaciones."}]}])
            return resp.content[0].text.strip()
        except Exception as _e:
            print(f"    [OCR imagen] {_e}")
            return ""
    textos = []
    with pdfplumber.open(pdf_path) as pdf:
        for pag in pdf.pages:
            t = pag.extract_text()
            if t:
                textos.append(t)
    texto = "\n".join(textos)
    
    # Si el PDF tiene poco texto, puede ser escaneado — usar Claude Vision
    if len(texto.strip()) < 100:
        try:
            import base64
            with pdfplumber.open(pdf_path) as pdf:
                if pdf.pages:
                    img = pdf.pages[0].to_image(resolution=150)
                    img.save("/tmp/yve_ocr_page.png")
                    with open("/tmp/yve_ocr_page.png", "rb") as f:
                        img_b64 = base64.b64encode(f.read()).decode()
                    os.remove("/tmp/yve_ocr_page.png")
                    # Enviar imagen a Claude Vision para extraer texto
                    resp = client.messages.create(
                        model="claude-sonnet-4-6",
                        max_tokens=2000,
                        messages=[{"role":"user","content":[
                            {"type":"image","source":{"type":"base64","media_type":"image/png","data":img_b64}},
                            {"type":"text","text":"Extrae TODO el texto visible de esta imagen. Solo devuelve el texto, sin explicaciones."}
                        ]}]
                    )
                    ocr_text = resp.content[0].text.strip()
                    if len(ocr_text) > 50:
                        texto = ocr_text
                        print(f"    [OCR] Claude Vision extrajo {len(ocr_text)} chars")
        except Exception as e:
            print(f"    [OCR] Vision no disponible: {e}")
    
    return texto

def es_ota(texto):
    """Detecta si es una FACTURA de OTA (no solo mención de OTA en el texto).
    Un extracto bancario puede mencionar Booking.com sin ser una factura de OTA."""
    txt_lower = texto[:3000].lower()
    tiene_ota = any(ota in txt_lower for ota in OTAS_CONOCIDAS)
    if not tiene_ota:
        return False
    # Debe tener señales de factura DE la OTA, no solo mencionarla
    ota_invoice_signals = ['commission', 'comisión', 'comision', 'invoice', 'factura',
                           'amount due', 'total due', 'payment due', 'remittance']
    tiene_factura_ota = any(s in txt_lower for s in ota_invoice_signals)
    # Si menciona OTA pero también parece extracto bancario, NO es factura OTA
    bank_signals = ['extracto', 'saldo', 'movimiento', 'bank statement', 'cuenta corriente',
                    'transferencia', 'cargo', 'abono', 'balance']
    parece_banco = sum(1 for s in bank_signals if s in txt_lower) >= 3
    if parece_banco:
        return False
    return tiene_factura_ota


# Documentos que NO son facturas — pre-filtro por nombre de archivo
NO_FACTURA_KEYWORDS = {'menu ', 'checklist', 'minuta', 'diploma', 'schedule', 'quotation', 'meeting notes', 'certificado', 'signage', 'setup', 'powerpoint', 'programa', 'proposal', 'plano', 'itinerario', 'certificate', 'agenda', 'ppt', 'logo', 'floor plan', 'master onsite', 'acta ', 'timeline', 'carta de', 'presentation', 'floorplan', 'running order', 'quote', 'wine list', 'planning', 'banner', 'resume', 'itinerary', 'presupuesto'}

NO_FACTURA_EXTENSIONS = {'.doc', '.docx', '.ppt', '.pptx', '.gif', '.svg', '.mp4', '.zip', '.rar'}  # imágenes NO: se leen por OCR


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
    """Pre-filtro ULTRA conservador. Solo bloquea si el PDF está vacío.
    Todo lo demás pasa a Claude — él es mucho mejor juez que regex."""
    if not texto or len(texto.strip()) < 20:
        return True, "PDF sin texto extraíble (escaneado sin OCR)"
    # Todo lo que tenga texto → pasa a Claude
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
        # ── Prompt de clasificación y extracción universal ──
        max_chars = min(len(texto), 8000)
        
        prompt = """Eres un experto en operaciones y finanzas hoteleras.
Analiza el documento y haz DOS cosas: CLASIFICAR y EXTRAER.

CLASIFICACIÓN — Lee TODO el contenido antes de decidir:
• Lista de productos/ingredientes con stock/cantidades → INVENTARIO  
• Pérdidas/mermas/desperdicios con costes → MERMAS
• Ventas de restaurante/platos vendidos/tickets TPV → VENTAS_POS
• Movimientos bancarios con fechas, importes y saldos → EXTRACTO_BANCO
• Comisiones que una OTA (Booking/Expedia...) TE FACTURA, con nº de factura,
  periodo facturado e importes ya devengados → COMISIONES_OTA
• Tarifas de comisión PACTADAS con una OTA (contrato, acuerdo, anexo de
  condiciones): porcentajes y vigencia, SIN nº de factura ni importes
  facturados → CONTRATO_OTA. Ojo: esto NO es una factura, es lo acordado.
• Lista de habitaciones/huéspedes de un grupo → ROOMING
• Número de factura + proveedor + IVA/VAT + total → FACTURA
• Depósito/anticipo/proforma con importe → FACTURA
• BEO/Banquet Event Order con desglose de servicios y precios → BEO
• TM/Technical Manual/requisitos técnicos de evento → TM
• Contrato/acuerdo de servicios con importes → CONTRATO
• Agenda/email/logo/manual técnico sin datos financieros → OTRO

EXTRACCIÓN — Devuelve SOLO JSON según el tipo:

FACTURA:
{"es_factura":true,"numero_factura":"X","fecha":"DD/MM/YYYY","nombre_proveedor":"X","NIF_proveedor":"X","descripcion_concepto":"X","base_imponible":0.0,"porcentaje_iva":21,"cuota_iva":0.0,"total_factura":0.0,"moneda":"EUR"}

INVENTARIO:
{"tipo_documento":"INVENTARIO","items":[{"ingrediente":"nombre","categoria":"tipo","coste_unitario":0.0,"stock_actual_kg_l":0.0,"stock_inicial_kg_l":0.0,"unidad":"kg","proveedor":"nombre","critico":false}]}

MERMAS:
{"tipo_documento":"MERMAS","items":[{"fecha":"DD/MM/YYYY","ingrediente":"nombre","categoria":"tipo","cantidad_merma":0.0,"unidad":"kg","causa":"motivo","coste_unitario":0.0,"coste_merma":0.0}],"total_mermas":0.0}

VENTAS_POS:
{"tipo_documento":"VENTAS_POS","fecha":"YYYY-MM-DD","total_ventas":0.0,"num_tickets":0,"platos":[{"nombre_plato":"nombre","categoria":"tipo","unidades_vendidas":0,"total_venta":0.0}]}

EXTRACTO_BANCO:
{"tipo_documento":"EXTRACTO_BANCO","movimientos":[{"fecha":"DD/MM/YYYY","concepto":"descripción","importe":0.0,"saldo":0.0}]}

COMISIONES_OTA (una entrada en "facturas" por CADA factura/hotel que veas):
{"tipo_documento":"COMISIONES_OTA","ota":"nombre","periodo":"mes/año","importe_bruto":0.0,"comision":0.0,"porcentaje":0.0,"facturas":[{"numero_factura":"X","nombre_hotel":"X","fecha":"DD/MM/YYYY","periodo_inicio":"DD/MM/YYYY","periodo_fin":"DD/MM/YYYY","importe_bruto":0.0,"porcentaje_comision":0.0,"importe_comision":0.0,"importe_neto":0.0}]}

CONTRATO_OTA (una entrada en "tarifas" por CADA hotel/mercado pactado):
{"tipo_documento":"CONTRATO_OTA","ota":"nombre","tarifas":[{"nombre_hotel":"X","porcentaje_pactado":0.0,"mercado":"Nacional|Internacional|...","vigencia_inicio":"DD/MM/YYYY","vigencia_fin":"DD/MM/YYYY"}]}

ROOMING:
{"tipo_documento":"ROOMING","grupo":"nombre","num_habitaciones":0,"checkin":"DD/MM/YYYY","checkout":"DD/MM/YYYY","tarifa_media":0.0}

BEO:
{"tipo_documento":"BEO","evento":"nombre del evento","cliente":"empresa cliente","fecha_evento":"DD/MM/YYYY","num_asistentes":0,"items":[{"concepto":"descripción","cantidad":0,"precio_unitario":0.0,"total":0.0}],"total_estimado":0.0,"notas":"observaciones"}

TM:
{"tipo_documento":"TM","evento":"nombre del evento","cliente":"empresa","requisitos":[{"tipo":"AV/Sala/Catering/Decoración","descripcion":"detalle","coste_estimado":0.0}],"total_estimado":0.0}

CONTRATO:
{"tipo_documento":"CONTRATO","evento":"nombre del evento","cliente":"empresa cliente","NIF_cliente":"X","fecha_firma":"DD/MM/YYYY","importe_total":0.0,"deposito":0.0,"condiciones_pago":"descripción","items":[{"concepto":"descripción","importe":0.0}],"vigencia":"periodo"}

OTRO:
{"tipo_documento":"OTRO","descripcion":"qué es el documento"}

REGLAS:
- Importes SIEMPRE como float (1234.56)
- Si es factura y solo hay total: base=total/1.21, iva=21, cuota=total-base
- IVA 0% intracomunitaria: base=total, iva=0, cuota=0
- Extrae TODOS los items/movimientos/platos que veas, no solo los primeros
- Si no encuentras un campo → null. NUNCA inventes datos
- Un % de comisión NO convierte un documento en factura: sin nº de factura
  ni importes facturados, es CONTRATO_OTA, no COMISIONES_OTA
- Responde SOLO con JSON, sin markdown, sin explicaciones, sin ```

""" + f"ARCHIVO: {nombre_archivo}\nTEXTO:\n{texto[:max_chars]}"

        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,  # Suficiente para listas largas
            messages=[{"role":"user","content":prompt}]
        )
        raw = resp.content[0].text.strip()
        
        # Limpiar markdown
        raw = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
        
        # Extraer JSON — buscar el primer { y el último }
        first_brace = raw.find('{')
        last_brace = raw.rfind('}')
        if first_brace >= 0 and last_brace > first_brace:
            raw = raw[first_brace:last_brace+1]
        
        try:
            datos = json.loads(raw)
        except json.JSONDecodeError:
            # Intentar arreglar JSON común
            fixed = re.sub(r',\s*}', '}', raw)
            fixed = re.sub(r',\s*]', ']', fixed)
            try:
                datos = json.loads(fixed)
            except json.JSONDecodeError:
                print(f"    AVISO: JSON inválido — usando regex")
                print(f"    Raw (primeros 200): {raw[:200]}")
                return extraer_con_regex(texto)
        
        # Clasificación: si tiene tipo_documento, es un tipo no-factura
        tipo_doc = datos.get("tipo_documento")
        if tipo_doc and tipo_doc not in ("FACTURA",):
            items = datos.get("items", datos.get("movimientos", datos.get("platos", [])))
            n_items = len(items) if isinstance(items, list) else 0
            print(f"    [CLASIFICADO] Tipo: {tipo_doc} ({n_items} items)")
            return datos
        
        # Si Claude dice que no es factura
        if datos.get("es_factura") is False:
            tipo_doc = datos.get("tipo_documento", "OTRO")
            desc = datos.get("descripcion", "documento no financiero")
            print(f"    [INFO] Tipo: {tipo_doc} — {desc}")
            return {"_skip": True, "_motivo": f"Claude: {desc}", "tipo_documento": tipo_doc}
        
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
        print(f"    [SKIP] {nombre}: {motivo}")
        return {"_skip": True, "_motivo": f"nombre: {motivo}"}

    try:
        texto = extraer_texto(pdf_path)
    except Exception as e:
        print(f"    ERROR al leer PDF: {e}")
        return {"archivo": nombre, "error": str(e)}

    if es_ota(texto):
        print(f"    [SKIP] Es factura de OTA — omitida")
        return {"_skip": True, "_motivo": "factura OTA (va a AR, no AP)"}

    # Pre-filtro 2: por contenido del PDF (gratis, sin tokens)
    skip2, motivo2 = es_no_factura_por_contenido(texto)
    if skip2:
        print(f"    [SKIP] {nombre}: {motivo2}")
        return {"_skip": True, "_motivo": f"contenido: {motivo2}"}

    datos = extraer_con_claude(texto, nombre)
    
    # Si Claude no devolvió nada
    if datos is None:
        print(f"    [SKIP] {nombre}: sin datos extraíbles")
        return {"_skip": True, "_motivo": "sin datos extraíbles"}

    # Si Claude clasificó como skip
    if isinstance(datos, dict) and datos.get('_skip'):
        return datos

    # Si Claude clasificó como otro tipo (no factura) → pasar al handler para enrutar
    tipo_doc = datos.get('tipo_documento')
    if tipo_doc and tipo_doc not in ('FACTURA',):
        print(f"    [CLASIFICADO] {nombre}: tipo={tipo_doc}")
        return datos  # El streaming handler lo enrutará al módulo correcto

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