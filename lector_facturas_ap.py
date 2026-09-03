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

# ── Rutas por tenant, evaluadas EN CADA USO ───────────────────────────────────
# Estaban clavadas a la raiz: con un segundo cliente, sus facturas AP y su
# `proveedores.xlsx` salian del arbol del primero.
#
# Aqui NO vale resolverlas al importar, como hacen `lector_ota` o
# `verificador_comisiones`: ese patron sirve porque esos dos solo corren como
# subproceso. Este modulo lo importa `dashboard` EN PROCESO
# (`from lector_facturas_ap import procesar_factura_ap, cargar_proveedores,
# guardar_excel, ...`), y un proceso sirve a VARIOS tenants: la ruta quedaria
# congelada con el tenant que hubiera en el primer import y todos los demas
# leerian y escribirian ahi. Por eso se resuelven en cada uso.
#
# Es lo mismo que ya hacen `conciliacion_bancaria` y `tab_ar_real`, pero con
# `__fspath__` porque aqui las rutas se usan con `os.path`, no con `Path`.
from tenant_dirs import (entrada_dir as _t_edir, procesadas_dir as _t_pdir,
                         datos_dir as _t_ddir)


class _TDir:
    """Carpeta del tenant, resuelta en cada uso (no al importar)."""
    def __init__(self, fn): self._fn = fn
    def __fspath__(self): return self._fn()
    def __str__(self): return self._fn()
    def __repr__(self): return self._fn()
    def __add__(self, o): return self._fn() + o
    def __radd__(self, o): return o + self._fn()
    def __truediv__(self, o): return os.path.join(self._fn(), o)


class _TFile:
    """Fichero dentro de una carpeta del tenant, resuelto en cada uso."""
    def __init__(self, fn, nombre): self._fn, self._n = fn, nombre
    def __fspath__(self): return os.path.join(self._fn(), self._n)
    def __str__(self): return os.path.join(self._fn(), self._n)
    def __repr__(self): return str(self)


ENTRADA_DIR      = _TDir(_t_edir)
SALIDA_DIR       = _TDir(_t_pdir)
REFERENCIA_DIR   = _TDir(_t_ddir)
os.makedirs(SALIDA_DIR, exist_ok=True)

FECHA_HOY      = date.today().strftime("%Y%m%d")
SALIDA_EXCEL   = _TFile(_t_pdir, f"facturas_ap_{FECHA_HOY}.xlsx")
PROV_FILE      = _TFile(_t_ddir, "proveedores.xlsx")
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

# ── Prompt de clasificacion COMPARTIDO ────────────────────────────────────
# UNICA fuente de verdad de la clasificacion. Lo usan las dos entradas:
#   - documentos con texto (PDF, hojas de calculo) -> prompt_documento()
#   - fotos de documentos fisicos                  -> prompt_foto()
# Antes habia dos prompts distintos (este y otro incrustado en
# /api/scan_documento) y se desincronizaron: el de las fotos decia "mismos
# schemas que para PDFs" pero solo incluia 4 de los 12, asi que una foto de un
# inventario o de un extracto nunca se extraia bien.
PROMPT_CLASIFICACION = """CLASIFICACIÓN — Lee TODO el contenido antes de decidir:
• Lista de productos/ingredientes con stock/cantidades → INVENTARIO  
• Pérdidas/mermas/desperdicios con costes → MERMAS
• Ventas de restaurante/platos vendidos/tickets TPV → VENTAS_POS
• Movimientos bancarios con fechas, importes y saldos → EXTRACTO_BANCO
• Informe diario de ingresos de un hotel (Daily Revenue Report): tiene una hoja
  DAILY_MASTER con Occupancy/ADR/RevPAR/GOP, hojas numeradas 1..31 con el Trial
  Balance del dia (ASSETS/LIABILITIES/INCOME/EXPENSES, debe y haber) y a menudo
  una hoja CtaCble → DRR.
  OJO: un DRR SE PARECE a un extracto bancario (filas con fechas, debe y haber)
  pero NO lo es. Si ves secciones contables o hojas por dia numeradas, es DRR.
• Comisiones que una OTA (Booking/Expedia...) TE FACTURA, con nº de factura,
  periodo facturado e importes ya devengados → COMISIONES_OTA
• Tarifas de comisión PACTADAS con una OTA (contrato, acuerdo, anexo de
  condiciones): porcentajes y vigencia, SIN nº de factura ni importes
  facturados → CONTRATO_OTA. Ojo: esto NO es una factura, es lo acordado.
• Lista de habitaciones/huéspedes de un grupo → ROOMING
• Bono / voucher de agencia o empresa: el documento que una agencia, touroperador
  o empresa entrega para que el hotel le FACTURE a crédito la estancia de un
  huésped (direct bill). Lleva nº de bono/voucher, quién paga (la agencia o
  empresa), el huésped, fechas de entrada y salida, régimen y el precio pactado,
  y NO lleva IVA desglosado ni nº de factura → BONO.
  OJO: un bono SE PARECE a una ROOMING (huéspedes y fechas) y a una FACTURA
  (importe), pero es una AUTORIZACIÓN de cobro a crédito: ni es una venta hecha
  ni es una lista de grupo. Si dudas entre BONO y FACTURA: sin IVA y sin nº de
  factura, con "bono"/"voucher"/"a cargo de", es BONO.
• Orden de compra / pedido / purchase order: lo que el hotel PIDE, ANTES de
  recibir nada. Lleva número de pedido u orden, suele decir "aprobado por" y a
  qué departamento se carga, y NO acredita ni entrega ni cobro → ORDEN_COMPRA.
  OJO: un pedido SE PARECE a un albarán y a una factura (mismo proveedor,
  fechas, importes, a veces el detalle de artículos) pero NO es ninguno de los
  dos. Es un COMPROMISO de gasto, no un gasto.
  Si dudas entre ORDEN_COMPRA y ALBARAN: el pedido no tiene firma de recepción
  y habla de "pedido"/"orden"/"solicitud"; el albarán habla de lo ENTREGADO.
  Si dudas entre ORDEN_COMPRA y FACTURA: sin número de factura y sin IVA no es
  una factura. Un importe "aprobado" o "presupuestado" es de un pedido.
• Nota de entrega / albarán / delivery note: lo que el proveedor dice que ha
  ENTREGADO, con cantidades por producto y a menudo un hueco para la firma de
  quien lo recibe → ALBARAN.
  OJO: un albarán SE PARECE a una factura (mismo proveedor, fechas, precios)
  pero NO lo es. Señales de albarán: NO lleva IVA/VAT, NO lleva número de
  factura, habla de "entrega"/"envío"/"remito" y detalla CANTIDADES servidas.
  Si dudas entre FACTURA y ALBARAN: sin IVA y sin número de factura, es ALBARAN.
• Número de factura + proveedor + IVA/VAT + total → FACTURA.
  Al revés también: si hay número de factura E IVA, es FACTURA aunque venga con
  el detalle de lo entregado.
• Depósito/anticipo/proforma con importe → FACTURA
• BEO/Banquet Event Order con desglose de servicios y precios → BEO
• TM/Technical Manual/requisitos técnicos de evento → TM
• Contrato/acuerdo de servicios con importes → CONTRATO
• Agenda/email/logo/manual técnico sin datos financieros → OTRO

EXTRACCIÓN — Devuelve SOLO JSON según el tipo:

FACTURA (UNA sola factura en el documento — el caso normal de un PDF o una foto).
"lineas" SOLO si la factura detalla lo servido producto por producto; si es una
factura de un concepto suelto (luz, alquiler, una comisión), OMITE "lineas".
Tener líneas NO convierte la factura en albarán: manda la regla de arriba (con
número de factura E IVA es FACTURA):
{"es_factura":true,"numero_factura":"X","fecha":"DD/MM/YYYY","nombre_proveedor":"X","NIF_proveedor":"X","descripcion_concepto":"X","base_imponible":0.0,"porcentaje_iva":21,"cuota_iva":0.0,"total_factura":0.0,"moneda":"EUR","lineas":[{"descripcion":"X","cantidad":0.0,"unidad":"kg|ud|l|caja","precio_unitario":0.0,"importe":0.0}]}

FACTURA (VARIAS facturas en el MISMO documento, tipico en una hoja de cálculo con
una factura por fila. Usa esta forma SOLO si de verdad hay más de una; si hay una
sola, usa la de arriba):
{"es_factura":true,"facturas":[{"numero_factura":"X","fecha":"DD/MM/YYYY","nombre_proveedor":"X","NIF_proveedor":"X","descripcion_concepto":"X","base_imponible":0.0,"porcentaje_iva":21,"cuota_iva":0.0,"total_factura":0.0,"moneda":"EUR","lineas":[{"descripcion":"X","cantidad":0.0,"unidad":"kg|ud|l|caja","precio_unitario":0.0,"importe":0.0}]}]}

ORDEN_COMPRA (lo PEDIDO, no lo entregado ni lo cobrado).
"importe_aprobado" es el total del pedido. "iva_incluido": true SOLO si ese
importe ya lleva el IVA dentro; si el pedido se aprueba sin IVA —lo normal— pon
false. Si no se puede saber, omite el campo, NO lo adivines.
"lineas" solo si el pedido detalla artículos:
{"tipo_documento":"ORDEN_COMPRA","numero_po":"X","nombre_proveedor":"X","NIF_proveedor":"X","fecha_pedido":"DD/MM/YYYY","departamento":"X","importe_aprobado":0.0,"iva_incluido":false,"moneda":"EUR","estado":"ABIERTO","lineas":[{"descripcion":"X","cantidad":0.0,"unidad":"kg|ud|l|caja","precio_unitario":0.0,"importe":0.0}]}

ALBARAN (una entrada en "lineas" por CADA producto entregado.
referencia_pedido y referencia_factura solo si el albarán las trae):
{"tipo_documento":"ALBARAN","numero_albaran":"X","nombre_proveedor":"X","NIF_proveedor":"X","fecha_entrega":"DD/MM/YYYY","referencia_pedido":"X","referencia_factura":"X","lineas":[{"descripcion":"X","cantidad":0.0,"unidad":"kg|ud|l|caja","precio_unitario":0.0,"importe":0.0}],"total_albaran":0.0}

INVENTARIO:
{"tipo_documento":"INVENTARIO","items":[{"ingrediente":"nombre","categoria":"tipo","coste_unitario":0.0,"stock_actual_kg_l":0.0,"stock_inicial_kg_l":0.0,"unidad":"kg","proveedor":"nombre","critico":false}]}

MERMAS:
{"tipo_documento":"MERMAS","items":[{"fecha":"DD/MM/YYYY","ingrediente":"nombre","categoria":"tipo","cantidad_merma":0.0,"unidad":"kg","causa":"motivo","coste_unitario":0.0,"coste_merma":0.0}],"total_mermas":0.0}

VENTAS_POS:
{"tipo_documento":"VENTAS_POS","fecha":"YYYY-MM-DD","total_ventas":0.0,"num_tickets":0,"platos":[{"nombre_plato":"nombre","categoria":"tipo","unidades_vendidas":0,"total_venta":0.0}]}

EXTRACTO_BANCO:
{"tipo_documento":"EXTRACTO_BANCO","movimientos":[{"fecha":"DD/MM/YYYY","concepto":"descripción","importe":0.0,"saldo":0.0}]}

DRR (NO extraigas nada: el fichero es demasiado grande y lo lee un lector
propio. Basta con identificarlo):
{"tipo_documento":"DRR"}

COMISIONES_OTA (una entrada en "facturas" por CADA factura/hotel que veas):
{"tipo_documento":"COMISIONES_OTA","ota":"nombre","periodo":"mes/año","importe_bruto":0.0,"comision":0.0,"porcentaje":0.0,"facturas":[{"numero_factura":"X","nombre_hotel":"X","fecha":"DD/MM/YYYY","periodo_inicio":"DD/MM/YYYY","periodo_fin":"DD/MM/YYYY","importe_bruto":0.0,"porcentaje_comision":0.0,"importe_comision":0.0,"importe_neto":0.0}]}

CONTRATO_OTA (una entrada en "tarifas" por CADA (OTA, hotel/mercado) pactado; si el contrato cubre VARIAS OTAs —p.ej. Booking Y Expedia—, cada tarifa lleva SU "ota" y el "ota" de arriba se deja vacio):
{"tipo_documento":"CONTRATO_OTA","ota":"la OTA si el contrato es de una sola; vacio si cubre varias","tarifas":[{"ota":"la OTA de esta tarifa","nombre_hotel":"X","porcentaje_pactado":0.0,"mercado":"Nacional|Internacional|...","vigencia_inicio":"DD/MM/YYYY","vigencia_fin":"DD/MM/YYYY"}]}

BONO (lo que la agencia/empresa AUTORIZA facturarle; "importe_total" es el total
pactado de la estancia si el bono lo trae, y "precio_noche" el precio por noche
y habitación si lo trae. Si solo hay uno de los dos, deja el otro en null, NO lo
calcules. "referencia_reserva" solo si el bono cita el localizador/reserva):
{"tipo_documento":"BONO","numero_bono":"X","agencia":"quien paga (agencia/empresa)","NIF_agencia":"X","huesped":"X","nombre_hotel":"X","fecha_entrada":"DD/MM/YYYY","fecha_salida":"DD/MM/YYYY","noches":0,"habitaciones":0,"regimen":"SA|AD|MP|PC","precio_noche":0.0,"importe_total":0.0,"moneda":"EUR","referencia_reserva":"X"}

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

"""

_CABECERA_DOC = """Eres un experto en operaciones y finanzas hoteleras.
Analiza el documento y haz DOS cosas: CLASIFICAR y EXTRAER.

"""

_CABECERA_FOTO = """Eres un experto en operaciones y finanzas hoteleras.
Esta es una FOTO de un documento fisico. Lee TODO el texto visible que aparezca
(aunque este torcido, con sombras o parcialmente cortado) y despues haz DOS
cosas: CLASIFICAR y EXTRAER.

"""

_AVISO_OTA = """
IMPORTANTE — este documento lo emite una OTA (Booking, Expedia, Hotels.com...).
Una factura emitida POR una OTA al hotel NO es una factura de proveedor: es el
cargo de comision del canal. Clasificala como COMISIONES_OTA, nunca como
FACTURA, y rellena "facturas" con una entrada por cada linea/hotel que veas.
"""


def prompt_documento(texto, nombre_archivo, max_chars=8000, es_ota=False):
    """Prompt para documentos con texto (PDF, CSV, Excel)."""
    return (_CABECERA_DOC + PROMPT_CLASIFICACION
            + (_AVISO_OTA if es_ota else "")
            + f"\nARCHIVO: {nombre_archivo}\nTEXTO:\n{texto[:max_chars]}")


def prompt_foto(nombre_archivo=""):
    """Prompt para fotos de documentos fisicos (Claude Vision)."""
    return (_CABECERA_FOTO + PROMPT_CLASIFICACION
            + (f"\nARCHIVO: {nombre_archivo}\n" if nombre_archivo else ""))


def extraer_con_claude(texto, nombre_archivo, es_ota=False):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("    AVISO: ANTHROPIC_API_KEY no encontrada — usando extracción regex")
        return extraer_con_regex(texto)
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        # ── Prompt de clasificación y extracción universal ──
        max_chars = min(len(texto), 8000)
        
        prompt = prompt_documento(texto, nombre_archivo, max_chars, es_ota=es_ota)

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
        # BOMBA 3: devolvia 628, que en plan_cuentas.xlsx es "Comisiones de
        # agencias y OTAs" — la factura de la luz acababa en la cuenta de
        # Booking. En este plan la energia va a 629 ("Otros servicios
        # (telecom, energia)"), igual que en asignador_cuentas.
        return '629'  # Suministros / energia
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

def _detectar_ota(texto):
    """Nombre CANONICO de la OTA que aparece en el texto, si es conocida.

    Devuelve el nombre tal y como aparece en comisiones_pactadas.xlsx
    ("Booking.com"), no como venga en la factura ("Booking.com B.V."):
    verificador_comisiones casa por nombre exacto, asi que un sufijo societario
    convierte la reclamacion en OTA_DESCONOCIDA.
    Recorre de mas largo a mas corto (y en orden estable: OTAS_CONOCIDAS es un
    set) para que "booking.com" gane a "booking" y el resultado no dependa del
    orden de iteracion.
    """
    t = texto[:3000].lower()
    for ota in sorted(OTAS_CONOCIDAS, key=lambda s: (-len(s), s)):
        if ota in t:
            return ota.title().replace(".Com", ".com").replace(".Es", ".es")
    return None


def _factura_a_comisiones_ota(datos, texto, nombre):
    """Reencuadra una FACTURA emitida por una OTA como COMISIONES_OTA.

    Red de seguridad: si pese al aviso del prompt la IA la devuelve como
    FACTURA normal, se traduce aqui para que el enrutado la lleve a AR (y al
    verificador de comisiones) en vez de contabilizarla como gasto de
    proveedor. Antes este caso se descartaba entero.
    """
    # El nombre canonico manda sobre el que ponga la factura: el verificador
    # cruza por nombre exacto contra comisiones_pactadas.xlsx.
    ota = _detectar_ota(texto) or datos.get("nombre_proveedor") or "OTA"
    bruto = _safe_float(datos.get("base_imponible")) or _safe_float(datos.get("total_factura"))
    return {
        "tipo_documento": "COMISIONES_OTA",
        "ota": ota,
        "periodo": datos.get("fecha") or "",
        "importe_bruto": bruto,
        "comision": _safe_float(datos.get("total_factura")),
        "porcentaje": _safe_float(datos.get("porcentaje_comision")),
        "facturas": [{
            "numero_factura": datos.get("numero_factura"),
            "nombre_hotel": datos.get("nombre_cliente") or datos.get("nombre_hotel"),
            "fecha": datos.get("fecha"),
            "importe_bruto": bruto,
            "porcentaje_comision": _safe_float(datos.get("porcentaje_comision")),
            "importe_comision": _safe_float(datos.get("total_factura")),
        }],
        "_reencuadrado_desde_factura": True,
    }


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

    # Una factura emitida POR una OTA no es un gasto de proveedor: es el cargo
    # de comision del canal, y va a AR. ANTES se descartaba aqui mismo con
    # _skip, asi que una factura de Booking cuyo nombre de fichero no llevara
    # palabra clave se detectaba correctamente... y se tiraba. Ahora se marca y
    # se extrae como COMISIONES_OTA.
    _doc_es_ota = es_ota(texto)
    if _doc_es_ota:
        print(f"    [OTA] Factura de OTA -> se extrae como comisiones (AR), no como gasto AP")

    # Pre-filtro 2: por contenido del PDF (gratis, sin tokens)
    skip2, motivo2 = es_no_factura_por_contenido(texto)
    if skip2:
        print(f"    [SKIP] {nombre}: {motivo2}")
        return {"_skip": True, "_motivo": f"contenido: {motivo2}"}

    datos = extraer_con_claude(texto, nombre, es_ota=_doc_es_ota)
    
    # Si Claude no devolvió nada
    if datos is None:
        print(f"    [SKIP] {nombre}: sin datos extraíbles")
        return {"_skip": True, "_motivo": "sin datos extraíbles"}

    # Si Claude clasificó como skip
    if isinstance(datos, dict) and datos.get('_skip'):
        return datos

    # Documento de OTA que la IA ha devuelto como factura normal: reencuadrar
    # para que acabe en AR y no contabilizado como compra a proveedor.
    if _doc_es_ota and not datos.get('tipo_documento'):
        print(f"    [OTA] Reencuadrada como COMISIONES_OTA")
        return _factura_a_comisiones_ota(datos, texto, nombre)

    # Si Claude clasificó como otro tipo (no factura) → pasar al handler para enrutar
    tipo_doc = datos.get('tipo_documento')
    if tipo_doc and tipo_doc not in ('FACTURA',):
        print(f"    [CLASIFICADO] {nombre}: tipo={tipo_doc}")
        return datos  # El streaming handler lo enrutará al módulo correcto

    return facturas_de_respuesta(datos, nombre, proveedores, como_dict=True)


# Campos que describen al DOCUMENTO y no a una factura concreta: cuando el
# clasificador devuelve varias facturas, estos pueden venir una sola vez arriba.
_COMUNES_FACTURA = ("nombre_proveedor", "NIF_proveedor", "moneda")


def facturas_de_respuesta(datos, nombre, proveedores, como_dict=False):
    """Convierte UNA respuesta del clasificador en las facturas que contiene.

    Punto UNICO de normalizacion de facturas AP: lo usan los tres caminos de
    entrada (PDF, foto y hoja de calculo). Antes cada uno normalizaba por su
    cuenta y las copias se habian desincronizado.

    Un documento puede traer VARIAS facturas —una hoja de calculo con una
    factura por fila es el caso normal—, asi que el prompt admite "facturas".
    Si viene esa lista se normalizan TODAS; si no, es una sola y se devuelve
    igual que siempre.

    como_dict=True conserva el contrato historico de procesar_factura_ap:
    devuelve un dict. Si habia varias, la lista completa viaja en "_facturas"
    para que quien sepa guardarlas las guarde todas y no se pierda ninguna en
    silencio.
    """
    lista = datos.get("facturas") if isinstance(datos, dict) else None
    if isinstance(lista, list) and any(isinstance(f, dict) for f in lista):
        comunes = {k: datos[k] for k in _COMUNES_FACTURA
                   if datos.get(k) not in (None, "")}
        filas = []
        _origen = []            # el dict ORIGINAL de cada factura, para sus lineas
        for i, f in enumerate(lista):
            if not isinstance(f, dict):
                continue
            filas.append(normalizar_factura_ap({**comunes, **f}, nombre, proveedores))
            # Cada factura necesita su propio "archivo": el Excel deduplica por
            # esa columna, asi que repetir el nombre se comeria todas menos la
            # ultima. Se usa el numero de factura para que reprocesar el mismo
            # fichero actualice sus filas en vez de duplicarlas.
            marca = str(f.get("numero_factura") or "").strip()
            filas[-1]["archivo"] = "%s#%s" % (nombre, marca or (i + 1))
            _origen.append(f)
        if filas:
            if len(filas) == 1:
                filas[0]["archivo"] = nombre        # una sola: como toda la vida
            # Fase 3c: las lineas viajan en una clave PRIVADA. _guardar_factura_ap
            # ya filtra las que empiezan por '_', asi que la hoja plana de
            # facturas no cambia ni una columna. Las lineas se leen de CADA
            # factura, nunca de las comunes del documento: copiarlas a todas
            # multiplicaria la mercancia por el numero de facturas de la hoja.
            for _fila, _f in zip(filas, _origen):
                _lin = lineas_factura(_f, _fila, nombre)
                if _lin:
                    _fila["_lineas"] = _lin
            if not como_dict:
                return filas
            return dict(filas[0], _facturas=filas) if len(filas) > 1 else filas[0]

    una = normalizar_factura_ap(datos, nombre, proveedores)
    _lin = lineas_factura(datos, una, nombre)
    if _lin:
        una["_lineas"] = _lin
    return una if como_dict else [una]


def normalizar_factura_ap(datos, nombre, proveedores):
    """Los campos de UNA factura -> la fila que se guarda en facturas_ap_*.xlsx.

    Movido tal cual desde procesar_factura_ap: clasificacion del proveedor,
    autocalculo de base/IVA/total cuando falta alguno, y cuenta PGC.
    """
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

def _txt_alb(v):
    """Texto comparable para las claves de albaran. NaN y NO_ENCONTRADO = vacio."""
    s = "" if v is None else str(v)
    s = " ".join(s.split()).strip().lower()
    return "" if s in ("", "nan", "none", "nat", "<na>", "no_encontrado", "null") else s


def clave_albaran(numero, proveedor, archivo=""):
    """Identidad de un albaran: numero + proveedor.

    Mismo criterio que `almacen_datos`: si el campo principal (el numero) viene
    vacio, la fila NO puede deduplicarse con ninguna otra, asi que la clave
    incluye el fichero de origen. Preferimos un albaran repetido a dos albaranes
    distintos fusionados en uno — con mercancia de por medio, fusionar dos
    entregas es perder una.
    """
    n, p = _txt_alb(numero), _txt_alb(proveedor)
    return f"{n}|{p}" if n else f"|{p}|{_txt_alb(archivo)}"


def lineas_factura(datos, fila, nombre):
    """Las lineas de UNA factura -> filas para la hoja `Lineas`. Punto UNICO.

    Fase 3c. Hermana de las lineas del albaran, y por los mismos motivos: el
    nivel 3 del cruce compara cantidades y precios producto a producto, y sin
    importe por linea no se puede comparar nada. Se autocompletan los huecos
    aritmeticos (importe = cantidad x precio y al reves); lo que NO se puede
    derivar se queda vacio, nunca inventado.

    La clave es `archivo`, la MISMA columna con la que la hoja de facturas
    deduplica: asi reprocesar un documento se lleva sus lineas viejas en vez de
    dejarlas huerfanas sumando mercancia que no se facturo.

    Una factura de un concepto suelto (luz, alquiler, una comision) no trae
    lineas y no pasa nada: el nivel 3 simplemente no se puede aplicar ahi, y
    decirlo es mejor que inventarse una linea con el total dentro.
    """
    if not isinstance(datos, dict):
        return []
    archivo = str((fila or {}).get("archivo") or nombre or "").strip()
    numero  = str((fila or {}).get("numero_factura")
                  or datos.get("numero_factura") or "").strip()
    prov    = str((fila or {}).get("nombre_proveedor")
                  or datos.get("nombre_proveedor") or "").strip()

    filas = []
    for i, ln in enumerate(datos.get("lineas") or []):
        if not isinstance(ln, dict):
            continue
        desc = str(ln.get("descripcion") or "").strip()
        cant = _safe_float(ln.get("cantidad"))
        prec = _safe_float(ln.get("precio_unitario"))
        imp  = _safe_float(ln.get("importe"))
        if imp is None and cant is not None and prec is not None:
            imp = round(cant * prec, 2)
        elif prec is None and imp is not None and cant:
            prec = round(imp / cant, 4)
        elif cant is None and imp is not None and prec:
            cant = round(imp / prec, 3)
        # una linea sin descripcion, sin cantidad y sin importe no es una linea
        if not desc and cant is None and imp is None:
            continue
        filas.append({
            "archivo":          archivo,
            "numero_factura":   numero,
            "nombre_proveedor": prov,
            "n_linea":          i + 1,
            "descripcion":      desc,
            "cantidad":         cant,
            "unidad":           str(ln.get("unidad") or "").strip(),
            "precio_unitario":  prec,
            "importe":          imp,
        })
    return filas


def clave_po(numero, proveedor, archivo=""):
    """Identidad de un PO. Si el numero viene vacio, la clave incluye el fichero.

    Mismo criterio que `clave_albaran`, y por un motivo equivalente: **fusionar
    dos pedidos es perder un compromiso de gasto**. Dos pedidos sin numero del
    mismo proveedor son dos pedidos, no uno.
    """
    n = _txt_alb(numero)
    p = _txt_alb(proveedor)
    return f"{n}|{p}" if n else f"|{p}|{_txt_alb(archivo)}"


def orden_compra_de_respuesta(datos, nombre):
    """Una ORDEN_COMPRA del clasificador -> (cabecera, lineas). Punto UNICO.

    Clonado de `albaran_de_respuesta` a proposito: un pedido tambien es una
    cabecera con N lineas y se guarda en dos hojas unidas por `clave`.

    Las LINEAS se guardan pero **el cruce por totales NO las usa**: son para el
    dia que se haga la comparacion articulo por articulo. Guardarlas ahora sale
    gratis porque la infraestructura de lineas ya existe, y evita volver a tocar
    el esquema compartido mas adelante.

    `iva_incluido` es la pieza que evita la trampa del IVA: un pedido se aprueba
    normalmente SIN IVA y la factura llega CON IVA. Medido antes de escribir
    esto: comparar el total de la factura contra el importe aprobado marcaba como
    discrepancia el 100% de las facturas correctas. Si el clasificador no lo
    dice, se queda en None y quien compare tendra que avisarlo, nunca adivinarlo.
    """
    if not isinstance(datos, dict):
        return {}, []
    numero = str(datos.get("numero_po") or datos.get("numero_pedido") or "").strip()
    prov   = str(datos.get("nombre_proveedor") or "").strip()
    clave  = clave_po(numero, prov, nombre)

    lineas = []
    for i, ln in enumerate(datos.get("lineas") or []):
        if not isinstance(ln, dict):
            continue
        desc = str(ln.get("descripcion") or "").strip()
        cant = _safe_float(ln.get("cantidad"))
        prec = _safe_float(ln.get("precio_unitario"))
        imp  = _safe_float(ln.get("importe"))
        if imp is None and cant is not None and prec is not None:
            imp = round(cant * prec, 2)
        elif prec is None and imp is not None and cant:
            prec = round(imp / cant, 4)
        elif cant is None and imp is not None and prec:
            cant = round(imp / prec, 3)
        if not desc and cant is None and imp is None:
            continue
        lineas.append({
            "clave":            clave,
            "numero_po":        numero,
            "nombre_proveedor": prov,
            "n_linea":          i + 1,
            "descripcion":      desc,
            "cantidad":         cant,
            "unidad":           str(ln.get("unidad") or "").strip(),
            "precio_unitario":  prec,
            "importe":          imp,
        })

    importe = _safe_float(datos.get("importe_aprobado"))
    if importe is None:
        importe = _safe_float(datos.get("total_pedido") or datos.get("total"))
    if importe is None and lineas:
        _sum = [l["importe"] for l in lineas if l["importe"] is not None]
        importe = round(sum(_sum), 2) if _sum else None

    # tri-estado a proposito: True / False / None ("no se sabe")
    _iva = datos.get("iva_incluido")
    iva_incluido = bool(_iva) if isinstance(_iva, bool) else None

    cabecera = {
        "clave":            clave,
        "archivo":          nombre,
        "numero_po":        numero,
        "nombre_proveedor": prov,
        "NIF_proveedor":    str(datos.get("NIF_proveedor") or "").strip(),
        "fecha_pedido":     str(datos.get("fecha_pedido") or datos.get("fecha") or "").strip(),
        "departamento":     str(datos.get("departamento") or "").strip(),
        "importe_aprobado": importe,
        "iva_incluido":     iva_incluido,
        "moneda":           str(datos.get("moneda") or "EUR").strip(),
        "estado":           str(datos.get("estado") or "").strip().upper(),
        "n_lineas":         len(lineas),
    }
    return cabecera, lineas


def po_tiene_datos(cabecera, lineas):
    """True si del pedido ha salido algo con lo que se pueda cruzar despues.

    Misma regla de producto que `albaran_tiene_datos` y `_ap_tiene_datos`: si no
    hay nada aprovechable NO se dice "✓ Orden de compra". Un pedido sirve para
    cruzar con **numero + proveedor + importe**; las lineas son un extra.
    """
    if _txt_alb(cabecera.get("numero_po")) and cabecera.get("importe_aprobado") is not None:
        return True
    if _txt_alb(cabecera.get("nombre_proveedor")) and cabecera.get("importe_aprobado") is not None:
        return True
    return any((l.get("cantidad") is not None or l.get("importe") is not None)
               and (l.get("descripcion") or l.get("importe") is not None) for l in lineas)


def albaran_de_respuesta(datos, nombre):
    """Un ALBARAN del clasificador -> (cabecera, lineas). Punto UNICO.

    Un albaran es una cabecera con N lineas, asi que se guarda en dos hojas
    unidas por `clave` (ver _guardar_albaran en dashboard.py). Aqui solo se
    normaliza: nada de disco.

    Se autocompletan los huecos aritmeticos porque el cruce factura-albaran
    compara cantidades y precios: sin importe por linea no se puede comparar
    nada. Lo que NO se puede derivar se queda vacio, nunca inventado.
    """
    if not isinstance(datos, dict):
        return {}, []
    numero = str(datos.get("numero_albaran") or "").strip()
    prov   = str(datos.get("nombre_proveedor") or "").strip()
    clave  = clave_albaran(numero, prov, nombre)

    lineas = []
    for i, ln in enumerate(datos.get("lineas") or []):
        if not isinstance(ln, dict):
            continue
        desc = str(ln.get("descripcion") or "").strip()
        cant = _safe_float(ln.get("cantidad"))
        prec = _safe_float(ln.get("precio_unitario"))
        imp  = _safe_float(ln.get("importe"))
        if imp is None and cant is not None and prec is not None:
            imp = round(cant * prec, 2)
        elif prec is None and imp is not None and cant:
            prec = round(imp / cant, 4)
        elif cant is None and imp is not None and prec:
            cant = round(imp / prec, 3)
        # una linea sin descripcion, sin cantidad y sin importe no es una linea
        if not desc and cant is None and imp is None:
            continue
        lineas.append({
            "clave":           clave,
            "numero_albaran":  numero,
            "nombre_proveedor": prov,
            "n_linea":         i + 1,
            "descripcion":     desc,
            "cantidad":        cant,
            "unidad":          str(ln.get("unidad") or "").strip(),
            "precio_unitario": prec,
            "importe":         imp,
        })

    total = _safe_float(datos.get("total_albaran"))
    if total is None and lineas:
        _sum = [l["importe"] for l in lineas if l["importe"] is not None]
        total = round(sum(_sum), 2) if _sum else None

    cabecera = {
        "clave":              clave,
        "archivo":            nombre,
        "numero_albaran":     numero,
        "nombre_proveedor":   prov,
        "NIF_proveedor":      str(datos.get("NIF_proveedor") or "").strip(),
        "fecha_entrega":      str(datos.get("fecha_entrega") or "").strip(),
        "referencia_pedido":  str(datos.get("referencia_pedido") or "").strip(),
        "referencia_factura": str(datos.get("referencia_factura") or "").strip(),
        "total_albaran":      total,
        "n_lineas":           len(lineas),
    }
    return cabecera, lineas


def clave_bono(numero, agencia, archivo=""):
    """Identidad de un bono: numero + quien paga; sin numero, el archivo."""
    n = _txt_alb(numero)
    a = _txt_alb(agencia).lower()
    if n:
        return f"{n}|{a}"
    return f"{_txt_alb(archivo)}|{a}"


def bono_de_respuesta(datos, nombre):
    """Un BONO del clasificador -> fila plana. Punto UNICO. Nada de disco.

    Si el bono trae precio por noche, noches y habitaciones pero no total, el
    total se deriva (es aritmetica, no una suposicion). Lo que no se puede
    derivar se queda vacio, nunca inventado.
    """
    if not isinstance(datos, dict):
        return {}
    numero  = str(datos.get("numero_bono") or "").strip()
    agencia = str(datos.get("agencia") or "").strip()
    noches  = _safe_float(datos.get("noches"))
    habs    = _safe_float(datos.get("habitaciones"))
    precio  = _safe_float(datos.get("precio_noche"))
    total   = _safe_float(datos.get("importe_total"))
    if total is None and precio is not None and noches:
        total = round(precio * noches * (habs or 1), 2)
    return {
        "clave":              clave_bono(numero, agencia, nombre),
        "archivo":            nombre,
        "numero_bono":        numero,
        "agencia":            agencia,
        "NIF_agencia":        str(datos.get("NIF_agencia") or "").strip(),
        "huesped":            str(datos.get("huesped") or "").strip(),
        "nombre_hotel":       str(datos.get("nombre_hotel") or "").strip(),
        "fecha_entrada":      str(datos.get("fecha_entrada") or "").strip(),
        "fecha_salida":       str(datos.get("fecha_salida") or "").strip(),
        "noches":             int(noches) if noches is not None else None,
        "habitaciones":       int(habs) if habs is not None else None,
        "regimen":            str(datos.get("regimen") or "").strip(),
        "precio_noche":       precio,
        "importe_total":      total,
        "moneda":             str(datos.get("moneda") or "EUR").strip() or "EUR",
        "referencia_reserva": str(datos.get("referencia_reserva") or "").strip(),
    }


def bono_tiene_datos(fila):
    """True si hay con que cotejar: quien paga + (importe o fechas)."""
    if not fila or not _txt_alb(fila.get("agencia")):
        return False
    return fila.get("importe_total") is not None or bool(_txt_alb(fila.get("fecha_entrada")))


def albaran_tiene_datos(cabecera, lineas):
    """True si del albaran ha salido algo con lo que se pueda cruzar despues.

    Misma regla de producto que `_ap_tiene_datos`: si no hay nada aprovechable,
    NO se dice "✓ Albaran". Vale con una linea util, o con numero + total.
    """
    if any((l.get("cantidad") is not None or l.get("importe") is not None)
           and (l.get("descripcion") or l.get("importe") is not None) for l in lineas):
        return True
    return bool(_txt_alb(cabecera.get("numero_albaran"))
                and cabecera.get("total_albaran") is not None)


def _filas_limpias(reg):
    """Las facturas de un resultado, sin las claves internas (_facturas, _skip).

    procesar_factura_ap devuelve un dict por compatibilidad; si el documento
    traia varias, la lista completa va en "_facturas". Esto las despliega.
    """
    if not isinstance(reg, dict):
        return []
    filas = reg.get("_facturas") or [reg]
    return [{k: v for k, v in f.items() if not str(k).startswith("_")}
            for f in filas if isinstance(f, dict)]


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
            # un documento puede traer varias facturas: guardarlas todas
            registros.extend(_filas_limpias(reg))

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
                filas = _filas_limpias(reg)
                ruta_excel = os.path.join(SALIDA_DIR, f"facturas_ap_{FECHA_HOY}.xlsx")
                if os.path.exists(ruta_excel):
                    df_existing = pd.read_excel(ruta_excel)
                    df_new = pd.DataFrame(filas)
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
                    guardar_excel(filas, ruta_excel)
                print(f"OK: {os.path.basename(file_path)} procesado correctamente")
                sys.exit(0)
            else:
                err = reg.get("error", "error desconocido") if reg else "sin resultado"
                print(f"ERROR: {err}")
                sys.exit(1)
    else:
        main()