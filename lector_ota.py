"""
lector_ota.py — Yve.01
Lee todas las facturas OTA (PDF) de la carpeta facturas-entrada,
extrae los campos clave y guarda los resultados en un Excel.
"""

import os
import re
from datetime import date
import pdfplumber
import pandas as pd

# ── Rutas ──────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENTRADA_DIR = os.path.join(BASE_DIR, "facturas-entrada")
SALIDA_DIR = os.path.join(BASE_DIR, "facturas-procesadas")
os.makedirs(SALIDA_DIR, exist_ok=True)

FECHA_HOY = date.today().strftime("%Y%m%d")
SALIDA_EXCEL = os.path.join(SALIDA_DIR, f"facturas_procesadas_{FECHA_HOY}.xlsx")

NF = "NO_ENCONTRADO"

# ── Helpers de patrón ─────────────────────────────────────────────────────
# Patrón de importe: "EUR 1,234.56" o "1.234,56" o "1234.56"
_EUR = r"(?:EUR|€|USD|\$)\s*"
_AMT = r"([0-9]{1,3}(?:[.,][0-9]{3})*[.,][0-9]{2}|[0-9]+\.[0-9]{2})"

# ── Patrones de extracción ─────────────────────────────────────────────────
# Cada campo tiene varios patrones ordenados de más específico a más general.
# Se usa el primer match encontrado.

PATRONES = {
    # ── Número de factura ──────────────────────────────────────────────────
    "numero_factura": [
        # "Invoice number: EXP-INV-2024-110034"  /  "Invoice number: 2410088472"
        r"(?:invoice\s*(?:number|no\.?|#|num\.?)|n[uú]mero\s*(?:de\s*)?factura|factura\s*n[uú]m\.?)[:\s#]*([A-Z0-9][A-Z0-9\-\/]{2,30})",
        # Fallback genérico: "invoice" seguido de código
        r"(?:invoice|factura)[:\s]+([A-Z]{0,5}[\-]?[0-9]{4,}[\-\/]?[A-Z0-9]*)",
    ],

    # ── Fecha de emisión ───────────────────────────────────────────────────
    "fecha": [
        # "Invoice Date: 05/11/2024"  /  "Date: 03/11/2024"
        r"(?:invoice\s+date|fecha\s*(?:de\s*)?(?:factura|emisi[oó]n)|date)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        # "Date: 03 November 2024"
        r"(?:invoice\s+date|date)[:\s]*(\d{1,2}\s+(?:january|february|march|april|may|june|july|august|september|october|november|december|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+\d{4})",
    ],

    # ── Nombre OTA emisora ─────────────────────────────────────────────────
    "nombre_ota": [
        r"(Booking\.com|Booking\.es|Expedia|Hotels\.com|Despegar|Airbnb|Agoda|Trip\.com|Trivago|HRS)",
    ],

    # ── Nombre del hotel (destinatario) ───────────────────────────────────
    "nombre_hotel": [
        # "Hotel Arts Barcelona"  /  "Hotel Majestic Barcelona"
        r"(?:hotel|property|establecimiento|cliente|bill\s*to|facturado\s*a)[:\s]+([A-ZÁÉÍÓÚÑ][^\n]{3,60})",
        r"(?:to:|destinatario:)\s*([A-ZÁÉÍÓÚÑ][^\n]{3,60})",
    ],

    # ── Inicio del periodo de facturación ─────────────────────────────────
    "periodo_inicio": [
        # "Period: 01/10/2024 - 31/10/2024"  /  "Billing Period: 01/10/2024 - 31/10/2024"
        r"(?:billing\s+period|period|periodo\s*(?:de\s*facturaci[oó]n)?)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*[-–]",
        # "From: 01/10/2024"
        r"(?:from|desde|check[\s\-]?in)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
    ],

    # ── Fin del periodo de facturación ────────────────────────────────────
    "periodo_fin": [
        # "Period: 01/10/2024 - 31/10/2024" → capturar la SEGUNDA fecha (tras el guión)
        r"(?:billing\s+period|period|periodo)[:\s]*\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\s*[-–]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        # "To: 31/10/2024"  /  "Until: 31/10/2024"
        r"(?:\bto\b|hasta|through|until|check[\s\-]?out|fin\b)[:\s]+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
    ],

    # ── Importe bruto (ventas de habitaciones) ────────────────────────────
    "importe_bruto": [
        # Expedia explícito: "Gross booking revenue: EUR 2,460.00"
        r"(?:gross\s+booking\s+revenue|gross\s+revenue|total\s+room\s+revenue|importe\s+bruto)[:\s]+" + _EUR + _AMT,
        # Booking tabla: "Reservations EUR 12,450.00 EUR 1,867.50" → primer importe
        r"(?:reservations|room\s+sales)\s+" + _EUR + _AMT,
        # Expedia tabla: "TOTAL EUR 2,460.00 EUR 442.80" → primer importe
        r"^TOTAL\s+" + _EUR + _AMT + r"\s+" + _EUR,
        # Genérico: "Total sales / Total revenue"
        r"(?:total\s+(?:sales|revenue))[:\s]+" + _EUR + _AMT,
    ],

    # ── Porcentaje de comisión ────────────────────────────────────────────
    "porcentaje_comision": [
        # "Commission rate: 15%"  /  "Commission rate applied: 18%"
        # [^\n%]{0,40} permite palabras intermedias como "applied:"
        r"commission\s+rate[^\n%]{0,40}?(\d{1,2}(?:[.,]\d{1,2})?)\s*%",
        # "comisión: 15%"  /  "porcentaje: 20%"
        r"(?:comisi[oó]n|porcentaje)[:\s]*(\d{1,2}(?:[.,]\d{1,2})?)\s*%",
        # "15% commission"
        r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*(?:commission|comisi[oó]n)",
    ],

    # ── Importe de comisión ───────────────────────────────────────────────
    "importe_comision": [
        # Expedia explícito: "Total commission amount: EUR 442.80"
        r"(?:total\s+commission\s+amount|commission\s+amount|importe\s+comisi[oó]n)[:\s]+" + _EUR + _AMT,
        # Booking tabla: "Reservations EUR 8,300.00 EUR 1,494.00" → SEGUNDO importe
        r"(?:reservations|TOTAL)\s+" + _EUR + r"[0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2}\s+" + _EUR + _AMT,
        # "Payment charge EUR 24.38" (cargo adicional de Booking)
        r"(?:payment\s+charge|our\s+fee|fee\s+amount)[:\s]+" + _EUR + _AMT,
    ],

    # ── Importe neto ──────────────────────────────────────────────────────
    "importe_neto": [
        # Expedia: "Net amount to be paid to hotel: EUR 2,017.20"
        r"(?:net\s+amount\s+to\s+be\s+paid(?:\s+to\s+hotel)?|net\s+payout|importe\s+neto|total\s+neto)[:\s]+" + _EUR + _AMT,
        # Booking: "Total amount due EUR 1,891.88"
        r"(?:total\s+amount\s+due)\s+" + _EUR + _AMT,
        # Genérico: "Amount due" / "A pagar"
        r"(?:amount\s+due|a\s+(?:abonar|pagar|transferir))[:\s]+" + _EUR + _AMT,
    ],
}


def extraer_texto_pdf(pdf_path: str) -> str:
    """Extrae todo el texto de un PDF usando pdfplumber."""
    texto_completo = []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)
    return "\n".join(texto_completo)


def buscar_campo(texto: str, campo: str) -> str:
    """Aplica los patrones de un campo en el texto y devuelve el primer match."""
    patrones = PATRONES.get(campo, [])
    for patron in patrones:
        match = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
        if match:
            return match.group(1).strip()
    return NF


def procesar_factura(pdf_path: str) -> dict:
    """Procesa un PDF y devuelve un diccionario con todos los campos."""
    nombre_archivo = os.path.basename(pdf_path)
    print(f"  Procesando: {nombre_archivo}")

    try:
        texto = extraer_texto_pdf(pdf_path)
    except Exception as e:
        print(f"    ERROR al leer PDF: {e}")
        return {
            "archivo": nombre_archivo,
            "numero_factura": NF,
            "fecha": NF,
            "nombre_ota": NF,
            "nombre_hotel": NF,
            "periodo_inicio": NF,
            "periodo_fin": NF,
            "importe_bruto": NF,
            "porcentaje_comision": NF,
            "importe_comision": NF,
            "importe_neto": NF,
            "error": str(e),
        }

    resultado = {"archivo": nombre_archivo, "error": ""}
    for campo in PATRONES:
        valor = buscar_campo(texto, campo)
        resultado[campo] = valor
        estado = "✓" if valor != NF else "✗"
        print(f"    [{estado}] {campo}: {valor}")

    return resultado


def guardar_excel(registros: list, ruta_salida: str):
    """Guarda la lista de resultados en un Excel con formato."""
    df = pd.DataFrame(registros)

    columnas_orden = [
        "archivo", "numero_factura", "fecha", "nombre_ota", "nombre_hotel",
        "periodo_inicio", "periodo_fin", "importe_bruto", "porcentaje_comision",
        "importe_comision", "importe_neto", "error",
    ]
    # Solo incluir columnas que existan
    columnas_orden = [c for c in columnas_orden if c in df.columns]
    df = df[columnas_orden]

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas_OTA")
        ws = writer.sheets["Facturas_OTA"]

        # Ajustar anchos de columna automáticamente
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    print(f"\n✅ Excel guardado en: {ruta_salida}")


def main():
    print("=" * 60)
    print("  Yve.01 — Lector de Facturas OTA")
    print("=" * 60)

    # Buscar todos los PDFs en facturas-entrada
    pdfs = [
        os.path.join(ENTRADA_DIR, f)
        for f in os.listdir(ENTRADA_DIR)
        if f.lower().endswith(".pdf")
    ]

    if not pdfs:
        print(f"\n⚠️  No se encontraron PDFs en: {ENTRADA_DIR}")
        print("   Coloca las facturas OTA en PDF en esa carpeta y vuelve a ejecutar.")
        return

    print(f"\nFacturas encontradas: {len(pdfs)}\n")

    registros = []
    for pdf_path in sorted(pdfs):
        registro = procesar_factura(pdf_path)
        registros.append(registro)
        print()

    guardar_excel(registros, SALIDA_EXCEL)
    print(f"\nTotal procesadas: {len(registros)} facturas")
    print("=" * 60)


if __name__ == "__main__":
    main()
