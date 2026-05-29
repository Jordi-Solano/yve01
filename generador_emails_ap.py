"""
generador_emails_ap.py — Yve.01
Genera emails profesionales en español para incidencias del módulo AP:
  - DISCREPANCIA_PO   → solicitar factura rectificativa
  - SIN_PO            → informar que no existe orden de compra aprobada
  - ALERTA_CONSUMO    → solicitar explicación de diferencia de consumo F&B
Guarda en reportes/emails_pendientes_ap/
"""

import os
import re
import glob
import json
from datetime import datetime
from pathlib import Path

import pandas as pd

# ── Anthropic (opcional — funciona sin internet vía fallback) ──────────────────
try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent
REPORTES_DIR   = BASE_DIR / "reportes"
EMAILS_DIR     = REPORTES_DIR / "emails_pendientes_ap"
FACTURAS_DIR   = BASE_DIR / "facturas-procesadas"
HOY            = datetime.now().strftime("%Y%m%d")
HOTEL_NOMBRE   = "Hilton Barcelona"
HOTEL_EMAIL    = "finanzas@hiltonbarcelona.com"
FIRMANTE       = "Departamento de Finanzas"

EMAILS_DIR.mkdir(parents=True, exist_ok=True)

# ── Cliente Anthropic ─────────────────────────────────────────────────────────
client = None
if ANTHROPIC_AVAILABLE:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        env_path = BASE_DIR / ".env"
        if env_path.exists():
            for line in env_path.read_text().splitlines():
                if line.startswith("ANTHROPIC_API_KEY="):
                    api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
    if api_key:
        client = anthropic.Anthropic(api_key=api_key)


# ═══════════════════════════════════════════════════════════════════════════════
# TEMPLATES DE FALLBACK
# ═══════════════════════════════════════════════════════════════════════════════

def template_discrepancia_po(factura: dict) -> str:
    proveedor     = factura.get("nombre_proveedor", "Proveedor")
    email_prov    = factura.get("email_contacto", "")
    num_factura   = factura.get("numero_factura", "N/D")
    total_factura = factura.get("total_factura", 0)
    importe_po    = factura.get("importe_po", 0)
    def _sf2(v):
        try: return float(str(v).replace("NO_ENCONTRADO","").replace("%","").strip() or 0)
        except: return 0.0
    diferencia    = abs(_sf2(total_factura) - _sf2(importe_po)) if importe_po else 0
    fecha_hoy     = datetime.now().strftime("%d/%m/%Y")

    return f"""Asunto: Discrepancia en factura {num_factura} — Solicitud de factura rectificativa

Estimados señores de {proveedor},

Nos ponemos en contacto con ustedes en relación con la factura número {num_factura} \
recibida con fecha {fecha_hoy}.

Durante el proceso de verificación contable de {HOTEL_NOMBRE}, hemos detectado \
una discrepancia entre el importe facturado y el importe de la Orden de Compra \
aprobada correspondiente:

  • Importe facturado:            {float(total_factura):,.2f} EUR
  • Importe Orden de Compra:      {float(importe_po):,.2f} EUR
  • Diferencia detectada:         {diferencia:,.2f} EUR

Conforme a nuestro procedimiento interno de control de pagos, no es posible \
tramitar el abono de facturas que no se ajusten al importe previamente aprobado.

Por este motivo, les solicitamos amablemente que emitan una factura rectificativa \
por el importe correspondiente a la Orden de Compra aprobada, o bien que nos \
remitan la documentación justificativa que ampare la diferencia indicada.

En caso de cualquier duda, pueden contactarnos en {HOTEL_EMAIL}.

Agradecemos su comprensión y colaboración.

Atentamente,

{FIRMANTE}
{HOTEL_NOMBRE}
{HOTEL_EMAIL}
Fecha: {fecha_hoy}
"""


def template_sin_po(factura: dict) -> str:
    proveedor   = factura.get("nombre_proveedor", "Proveedor")
    num_factura = factura.get("numero_factura", "N/D")
    total       = factura.get("total_factura", 0)
    concepto    = factura.get("descripcion_concepto", "servicios prestados")
    fecha_hoy   = datetime.now().strftime("%d/%m/%Y")

    return f"""Asunto: Factura {num_factura} recibida sin Orden de Compra aprobada

Estimados señores de {proveedor},

Hemos recibido su factura número {num_factura} por importe de {float(total):,.2f} EUR, \
correspondiente a: {concepto}.

Sin embargo, en nuestros registros no consta ninguna Orden de Compra aprobada \
que cubra este suministro o servicio.

La política de pagos de {HOTEL_NOMBRE} exige que toda factura esté respaldada \
por una Orden de Compra emitida y autorizada previamente por el departamento \
correspondiente. Sin este documento, no es posible proceder al pago.

Les solicitamos que contacten con nuestro departamento de Compras para \
regularizar la situación:

  - Si el servicio fue prestado bajo autorización verbal, necesitamos que nos \
faciliten el nombre del responsable que les autorizó para solicitar la OC \
de forma retroactiva.
  - Si existe un contrato marco, por favor remítannos el número de referencia.

En caso de cualquier duda, pueden contactarnos en {HOTEL_EMAIL}.

Quedamos a su disposición para resolver esta situación a la mayor brevedad posible.

Atentamente,

{FIRMANTE}
{HOTEL_NOMBRE}
{HOTEL_EMAIL}
Fecha: {fecha_hoy}
"""


def template_alerta_consumo(factura: dict) -> str:
    proveedor      = factura.get("nombre_proveedor", "Proveedor")
    num_factura    = factura.get("numero_factura", "N/D")
    total_factura  = factura.get("total_factura", 0)
    coste_pos      = factura.get("coste_pos_mes", 0)
    diferencia_pct = factura.get("diferencia_pos_pct", 0)
    fecha_hoy      = datetime.now().strftime("%d/%m/%Y")
    mes_actual     = datetime.now().strftime("%B %Y")

    return f"""Asunto: Alerta de consumo — Factura {num_factura} — Solicitud de aclaración

Estimados señores de {proveedor},

El sistema de control interno de {HOTEL_NOMBRE} ha detectado una discrepancia \
significativa entre el importe de su factura {num_factura} y los datos de \
consumo registrados en nuestro sistema de punto de venta (POS) para el \
período de {mes_actual}:

  • Importe facturado:            {float(total_factura):,.2f} EUR
  • Coste registrado en POS:      {float(coste_pos):,.2f} EUR
  • Diferencia porcentual:        {float(str(diferencia_pct).replace("%","").strip() or 0):.1f}%

Esta discrepancia supera el umbral del 15% establecido en nuestros \
procedimientos de control de inventario y mermas.

Las posibles causas que estamos analizando incluyen:
  a) Diferencias en el inventario físico no registradas en el sistema
  b) Mermas, roturas o caducidades no documentadas
  c) Posibles errores de registro en el POS
  d) Diferencias en las unidades de medida o conversión de precios

Les rogamos que nos remitan a la mayor brevedad:
  1. Albarán o documentación de entrega del período correspondiente
  2. Cualquier nota de crédito o ajuste pendiente de emitir
  3. Detalle del desglose de productos si la factura es agregada

En caso de que la discrepancia corresponda a un error en nuestra parte, \
procederemos a corregir los registros internos.

Pueden contactarnos en {HOTEL_EMAIL} para cualquier aclaración.

Gracias por su colaboración.

Atentamente,

{FIRMANTE}
{HOTEL_NOMBRE}
{HOTEL_EMAIL}
Fecha: {fecha_hoy}
"""


# ═══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN CON CLAUDE API
# ═══════════════════════════════════════════════════════════════════════════════

def generar_con_claude(tipo_email: str, factura: dict) -> str:
    """Genera un email profesional usando Claude API."""
    def _sf(v, default=0.0):
        try:
            x = str(v).replace("NO_ENCONTRADO","").replace("nan","").strip()
            return float(x) if x else default
        except Exception:
            return default
    contextos = {
        "DISCREPANCIA_PO": (
            f"La factura nº {factura.get('numero_factura')} de {factura.get('nombre_proveedor')} "
            f"tiene un importe de {factura.get('total_factura')} EUR pero la Orden de Compra "
            f"aprobada es de {factura.get('importe_po', 'desconocido')} EUR. "
            f"Diferencia: {abs(_sf(factura.get('total_factura') or 0) - _sf(factura.get('importe_po') or 0)):,.2f} EUR."
        ),
        "SIN_PO": (
            f"La factura nº {factura.get('numero_factura')} de {factura.get('nombre_proveedor')} "
            f"por {factura.get('total_factura')} EUR no tiene Orden de Compra asociada. "
            f"Concepto: {factura.get('descripcion_concepto', 'sin descripción')}."
        ),
        "ALERTA_CONSUMO": (
            f"La factura nº {factura.get('numero_factura')} de {factura.get('nombre_proveedor')} "
            f"es de {factura.get('total_factura')} EUR pero el coste registrado en POS es "
            f"{factura.get('coste_pos_mes', 'desconocido')} EUR "
            f"(diferencia del {_sf(factura.get('diferencia_pos_pct', 0)):.1f}%)."
        ),
    }

    instrucciones = {
        "DISCREPANCIA_PO": "Redacta un email formal en español al proveedor solicitando una factura rectificativa.",
        "SIN_PO":          "Redacta un email formal en español al proveedor informando que no existe OC aprobada y que no se puede pagar sin ella.",
        "ALERTA_CONSUMO":  "Redacta un email formal en español al proveedor solicitando aclaración sobre la diferencia entre lo facturado y el consumo registrado en POS.",
    }

    prompt = f"""Eres el asistente de contabilidad del {HOTEL_NOMBRE}.
Situación: {contextos.get(tipo_email, '')}
Tarea: {instrucciones.get(tipo_email, '')}
El email debe ser profesional, conciso, en tono cortés pero firme, y firmar como "{FIRMANTE}, {HOTEL_NOMBRE}".
Incluye asunto en la primera línea con formato: "Asunto: ..."
No añadas comentarios adicionales fuera del texto del email."""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=800,
            messages=[{"role": "user", "content": prompt}]
        )
        return resp.content[0].text.strip()
    except Exception as e:
        print(f"    ⚠  Claude API no disponible ({e}), usando template.")
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ═══════════════════════════════════════════════════════════════════════════════

def cargar_facturas_con_incidencias() -> pd.DataFrame:
    """Carga las facturas AP que tienen incidencias y necesitan email."""
    # 1. Buscar archivo de facturas contabilizadas (más completo)
    patron_cont = str(FACTURAS_DIR / "facturas_contabilizadas_*.xlsx")
    archivos_cont = sorted(glob.glob(patron_cont), reverse=True)

    # 2. Buscar archivo de facturas AP base
    patron_ap = str(FACTURAS_DIR / "facturas_ap_*.xlsx")
    archivos_ap = sorted(glob.glob(patron_ap), reverse=True)

    df = None
    if archivos_cont:
        try:
            df = pd.read_excel(archivos_cont[0])
            print(f"✓ Cargado: {Path(archivos_cont[0]).name} ({len(df)} facturas)")
        except Exception as e:
            print(f"  Error leyendo {archivos_cont[0]}: {e}")

    if df is None and archivos_ap:
        try:
            df = pd.read_excel(archivos_ap[0])
            print(f"✓ Cargado: {Path(archivos_ap[0]).name} ({len(df)} facturas)")
        except Exception as e:
            print(f"  Error leyendo {archivos_ap[0]}: {e}")

    if df is None or df.empty:
        print("⚠  No se encontraron facturas AP procesadas.")
        return pd.DataFrame()

    # Normalizar columnas
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df


def cargar_matching_otras() -> pd.DataFrame:
    """Carga el último reporte de matching OTRAS."""
    patron = str(REPORTES_DIR / "matching_otras_*.xlsx")
    archivos = sorted(glob.glob(patron), reverse=True)
    if not archivos:
        return pd.DataFrame()
    try:
        df = pd.read_excel(archivos[0], sheet_name="Detalle_OTRAS")
        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def cargar_matching_fb() -> pd.DataFrame:
    """Carga el último reporte de matching F&B."""
    patron = str(REPORTES_DIR / "matching_fb_*.xlsx")
    archivos = sorted(glob.glob(patron), reverse=True)
    if not archivos:
        return pd.DataFrame()
    try:
        df = pd.read_excel(archivos[0], sheet_name="Detalle_FB")
        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


# ═══════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def determinar_email_a_enviar(row_factura: pd.Series,
                               df_otras: pd.DataFrame,
                               df_fb: pd.DataFrame) -> tuple:
    """
    Determina el tipo de email a enviar basado en el estado de matching.
    Primero busca estado_matching en la propia fila (facturas_contabilizadas),
    luego en los reportes de matching por separado.
    Retorna: (tipo_email, datos_extra) o (None, None) si no requiere email.
    """
    num_factura = str(row_factura.get("numero_factura", "")).strip()
    tipo_prov   = str(row_factura.get("tipo_proveedor", "")).strip().upper()

    datos_extra = {}
    estado = None

    # Siempre buscar en los reportes de matching (tienen importe_po, coste_pos_periodo)
    df_match = df_fb if tipo_prov == "FB" else df_otras
    if not df_match.empty:
        for col in ("numero_factura", "archivo"):
            if col in df_match.columns:
                mask = df_match[col].astype(str).str.strip().str.replace(".pdf","",regex=False) == num_factura.replace(".pdf","")
                if mask.any():
                    match_row = df_match[mask].iloc[0]
                    estado = str(match_row.get("estado_matching", match_row.get("estado",""))).strip().upper()
                    datos_extra["importe_po"]         = match_row.get("importe_po", 0)
                    datos_extra["coste_pos_mes"]      = match_row.get("coste_pos_periodo", 0)
                    datos_extra["diferencia_pos_pct"] = match_row.get("diferencia_pos_pct", 0)
                    break

    # Si no se encontró en matching, usar columna estado_matching de la fila principal
    if not estado:
        est_directo = str(row_factura.get("estado_matching", "")).strip().upper()
        if est_directo and est_directo not in ("", "NAN", "NONE", "NO_ENCONTRADO"):
            estado = est_directo

    # 3. Determinar tipo de email
    if estado in ("DISCREPANCIA_PO", "DISCREPANCIA"):
        return "DISCREPANCIA_PO", datos_extra
    elif estado == "SIN_PO":
        return "SIN_PO", datos_extra
    elif estado == "ALERTA_CONSUMO":
        return "ALERTA_CONSUMO", datos_extra
    else:
        return None, None


def procesar_facturas() -> list:
    """Genera todos los emails necesarios y guarda los archivos."""
    print("=" * 60)
    print("GENERADOR DE EMAILS AP — Yve.01")
    print("=" * 60)

    # Cargar datos
    df_facturas = cargar_facturas_con_incidencias()
    if df_facturas.empty:
        print("No hay facturas que procesar.")
        return []

    df_otras = cargar_matching_otras()
    df_fb    = cargar_matching_fb()

    print(f"  Matching OTRAS: {len(df_otras)} registros")
    print(f"  Matching F&B:   {len(df_fb)} registros")
    print()

    emails_generados = []

    for _, row in df_facturas.iterrows():
        num_factura = str(row.get("numero_factura", "")).strip()
        proveedor   = str(row.get("nombre_proveedor", "Proveedor")).strip()

        tipo_email, datos_extra = determinar_email_a_enviar(row, df_otras, df_fb)

        if tipo_email is None:
            continue  # Factura OK, no necesita email

        print(f"→ [{tipo_email}] Factura {num_factura} — {proveedor}")

        # Preparar datos completos para la plantilla
        datos_factura = {
            "numero_factura":    num_factura,
            "nombre_proveedor":  proveedor,
            "email_contacto":    str(row.get("email_contacto", "")),
            "total_factura":     row.get("total_factura", 0),
            "descripcion_concepto": str(row.get("descripcion_concepto", "")),
            **datos_extra,
        }

        # Intentar Claude API primero
        texto_email = None
        if client:
            print(f"   Intentando Claude API...", end=" ")
            texto_email = generar_con_claude(tipo_email, datos_factura)
            if texto_email:
                print("OK")

        # Fallback a template
        if not texto_email:
            if tipo_email == "DISCREPANCIA_PO":
                texto_email = template_discrepancia_po(datos_factura)
            elif tipo_email == "SIN_PO":
                texto_email = template_sin_po(datos_factura)
            elif tipo_email == "ALERTA_CONSUMO":
                texto_email = template_alerta_consumo(datos_factura)
            print(f"   Template aplicado.")

        # Guardar archivo de email
        nombre_seguro = re.sub(r"[^a-zA-Z0-9_\-]", "_", num_factura)
        nombre_archivo = f"email_{tipo_email}_{nombre_seguro}_{HOY}.txt"
        ruta_email = EMAILS_DIR / nombre_archivo

        with open(ruta_email, "w", encoding="utf-8") as f:
            f.write(f"PARA: {datos_factura.get('email_contacto', 'sin_email@ejemplo.com')}\n")
            f.write(f"TIPO: {tipo_email}\n")
            f.write(f"FACTURA: {num_factura}\n")
            f.write(f"PROVEEDOR: {proveedor}\n")
            f.write(f"GENERADO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("─" * 60 + "\n\n")
            f.write(texto_email)

        print(f"   ✓ Guardado: {nombre_archivo}")

        emails_generados.append({
            "tipo":         tipo_email,
            "num_factura":  num_factura,
            "proveedor":    proveedor,
            "archivo":      nombre_archivo,
        })

    # Generar índice
    if emails_generados:
        indice_path = EMAILS_DIR / f"indice_emails_ap_{HOY}.xlsx"
        df_indice = pd.DataFrame(emails_generados)
        df_indice.to_excel(indice_path, index=False)
        print(f"\n✓ Índice guardado: {indice_path.name}")

    print()
    print("=" * 60)
    print(f"RESUMEN: {len(emails_generados)} emails generados")

    conteo = {}
    for e in emails_generados:
        conteo[e["tipo"]] = conteo.get(e["tipo"], 0) + 1
    for tipo, cnt in conteo.items():
        print(f"  {tipo}: {cnt}")

    print(f"Guardados en: {EMAILS_DIR}")
    print("=" * 60)

    return emails_generados


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    procesar_facturas()
