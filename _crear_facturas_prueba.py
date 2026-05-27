"""
Crea 5 facturas PDF de prueba + actualiza datos de referencia.

Escenarios:
  F&B-1  Distribuidora Global Alimentacion SA  172.000 EUR → MATCH_3WAY_OK  (PO 173.000 nuevo)
  F&B-2  Pescados Barcelona SL                  2.200 EUR → DISCREPANCIA_PO (PO 1.850)
  F&B-3  Carnes Premium SL                      2.390 EUR → ALERTA_CONSUMO  (PO 2.400, POS >> factura)
  OTR-1  Endesa Energia SA                     18.500 EUR → MATCH_CORRECTO  (PO 18.500)
  OTR-2  Gas Natural Fenosa SA                  4.200 EUR → SIN_PO          (sin PO registrada)
"""

import os, sys
from pathlib import Path
from datetime import date
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

BASE_DIR    = Path("/sessions/brave-affectionate-hawking/mnt/yve01")
REF_DIR     = BASE_DIR / "datos-referencia"
ENTRADA_DIR = BASE_DIR / "facturas-entrada"
ENTRADA_DIR.mkdir(exist_ok=True)

# ─── 1. Actualizar proveedores.xlsx — añadir Distribuidora Global + Gas Natural ───

prov_path = REF_DIR / "proveedores.xlsx"
df_prov = pd.read_excel(prov_path)
nuevos = pd.DataFrame([
    {
        "nombre_proveedor":   "Distribuidora Global Alimentacion SA",
        "tipo":               "FB",
        "cuenta_contable":    "600",
        "email_contacto":     "facturas@distglobal.es",
        "porcentaje_iva_habitual": 10,
    },
    {
        "nombre_proveedor":   "Gas Natural Fenosa SA",
        "tipo":               "OTRAS",
        "cuenta_contable":    "629",
        "email_contacto":     "facturas@gasnaturalsa.es",
        "porcentaje_iva_habitual": 21,
    },
])
# Añadir solo si no existen
for _, row in nuevos.iterrows():
    if row["nombre_proveedor"] not in df_prov["nombre_proveedor"].values:
        df_prov = pd.concat([df_prov, pd.DataFrame([row])], ignore_index=True)
df_prov.to_excel(prov_path, index=False)
print(f"✓ proveedores.xlsx actualizado ({len(df_prov)} filas)")

# ─── 2. Añadir PO para Distribuidora Global en pos_ordenes.xlsx ───────────────

ordenes_path = REF_DIR / "pos_ordenes.xlsx"
df_ord = pd.read_excel(ordenes_path)
nueva_po = {
    "numero_po":        "PO-2025-0720",
    "proveedor":        "Distribuidora Global Alimentacion SA",
    "descripcion":      "Suministro mensual alimentos y bebidas julio 2025 (contrato marco)",
    "importe_aprobado": 173000,
    "departamento":     "F&B",
    "fecha":            "01/07/2025",
    "estado":           "Aprobada",
}
if "PO-2025-0720" not in df_ord["numero_po"].values:
    df_ord = pd.concat([df_ord, pd.DataFrame([nueva_po])], ignore_index=True)
df_ord.to_excel(ordenes_path, index=False)
print(f"✓ pos_ordenes.xlsx actualizado ({len(df_ord)} POs)")


# ─── 3. Función para crear PDF de factura ─────────────────────────────────────

def crear_factura_pdf(
    nombre_archivo, num_factura, fecha, proveedor, nif_prov,
    descripcion, base_imp, pct_iva, cuota_iva, total,
    email_prov="", dir_prov="", dir_hotel="", tipo="PROVEEDOR"
):
    ruta = ENTRADA_DIR / nombre_archivo
    c = canvas.Canvas(str(ruta), pagesize=A4)
    W, H = A4

    # ── Cabecera proveedor ──
    c.setFont("Helvetica-Bold", 18)
    c.setFillColorRGB(0.13, 0.13, 0.40)
    c.drawString(2*cm, H - 2.5*cm, proveedor)

    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0.3, 0.3, 0.3)
    if dir_prov:
        c.drawString(2*cm, H - 3.3*cm, dir_prov)
    c.drawString(2*cm, H - 3.9*cm, f"NIF: {nif_prov}")
    if email_prov:
        c.drawString(2*cm, H - 4.5*cm, f"Email: {email_prov}")

    # ── Número y fecha ──
    c.setFont("Helvetica-Bold", 11)
    c.setFillColorRGB(0, 0, 0)
    c.drawRightString(W - 2*cm, H - 2.5*cm, f"FACTURA Nº {num_factura}")
    c.setFont("Helvetica", 10)
    c.drawRightString(W - 2*cm, H - 3.2*cm, f"Fecha: {fecha}")
    c.drawRightString(W - 2*cm, H - 3.9*cm, f"Vencimiento: 30 días")

    # ── Línea separadora ──
    c.setStrokeColorRGB(0.2, 0.2, 0.6)
    c.setLineWidth(2)
    c.line(2*cm, H - 5*cm, W - 2*cm, H - 5*cm)

    # ── Cliente (hotel) ──
    c.setFont("Helvetica-Bold", 10)
    c.setFillColorRGB(0.13, 0.13, 0.40)
    c.drawString(2*cm, H - 5.8*cm, "FACTURADO A:")
    c.setFont("Helvetica", 10)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(2*cm, H - 6.5*cm, "Hilton Barcelona SA")
    c.drawString(2*cm, H - 7.1*cm, "Avda. Diagonal 589-591, 08014 Barcelona")
    c.drawString(2*cm, H - 7.7*cm, "NIF: A-08123456")

    # ── Tabla de concepto ──
    y_table = H - 9.5*cm
    c.setFillColorRGB(0.13, 0.13, 0.40)
    c.rect(2*cm, y_table, W - 4*cm, 0.7*cm, fill=1)
    c.setFont("Helvetica-Bold", 9)
    c.setFillColorRGB(1, 1, 1)
    c.drawString(2.3*cm,  y_table + 0.2*cm, "CONCEPTO / DESCRIPCIÓN")
    c.drawRightString(W - 2.3*cm, y_table + 0.2*cm, "IMPORTE (EUR)")

    # Fila de concepto
    y_row = y_table - 1.5*cm
    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(2.3*cm, y_row, descripcion)
    c.drawRightString(W - 2.3*cm, y_row, f"{base_imp:,.2f}")

    # ── Totales ──
    y_totals = y_row - 2.5*cm
    col_lbl = W - 7*cm
    col_val = W - 2*cm

    c.setFont("Helvetica", 10)
    c.drawString(col_lbl, y_totals,         "Base imponible:")
    c.drawRightString(col_val, y_totals,    f"{base_imp:,.2f} EUR")

    c.drawString(col_lbl, y_totals - 0.7*cm, f"IVA ({pct_iva}%):")
    c.drawRightString(col_val, y_totals - 0.7*cm, f"{cuota_iva:,.2f} EUR")

    c.setLineWidth(0.5)
    c.line(col_lbl, y_totals - 1.0*cm, W - 2*cm, y_totals - 1.0*cm)

    c.setFont("Helvetica-Bold", 12)
    c.setFillColorRGB(0.13, 0.13, 0.40)
    c.drawString(col_lbl, y_totals - 1.7*cm, "TOTAL FACTURA:")
    c.drawRightString(col_val, y_totals - 1.7*cm, f"{total:,.2f} EUR")

    # ── Pie ──
    c.setFont("Helvetica", 8)
    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.drawCentredString(W/2, 2.5*cm, "Forma de pago: Transferencia bancaria 30 días fecha factura")
    c.drawCentredString(W/2, 2.0*cm, "Conserve este documento como justificante fiscal")
    c.setLineWidth(0.5)
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.line(2*cm, 2.8*cm, W - 2*cm, 2.8*cm)

    c.save()
    print(f"  ✓ PDF creado: {nombre_archivo}")
    return str(ruta)


# ─── 4. Crear las 5 facturas ──────────────────────────────────────────────────

hoy = "15/07/2025"

facturas = [
    # F&B-1: MATCH_3WAY_OK — Distribuidora Global, PO 173.000
    dict(
        nombre_archivo="FAC-2025-DIST-0001.pdf",
        num_factura="FAC-2025-DIST-0001",
        fecha=hoy,
        proveedor="Distribuidora Global Alimentacion SA",
        nif_prov="A-17234567",
        descripcion="Suministro mensual alimentos y bebidas julio 2025 (contrato marco)",
        base_imp=155818.18,
        pct_iva=10,
        cuota_iva=15581.82,
        total=171400.00,    # dentro del 1% de PO 173.000 y ~0.7% del POS 172.600
        email_prov="facturas@distglobal.es",
        dir_prov="Pol. Ind. Zona Franca, Carrer B num. 45, 08040 Barcelona",
    ),
    # F&B-2: DISCREPANCIA_PO — Pescados Barcelona, PO 1.850, factura 2.200 (+18.9%)
    dict(
        nombre_archivo="FAC-2025-PESC-0027.pdf",
        num_factura="FAC-2025-PESC-0027",
        fecha=hoy,
        proveedor="Pescados Barcelona SL",
        nif_prov="B-08456789",
        descripcion="Pescado fresco y marisco semana 27 — gamba, lubina, dorada, rodaballo",
        base_imp=2000.00,
        pct_iva=10,
        cuota_iva=200.00,
        total=2200.00,      # PO era 1.850 → diff 18.9% > 1% → DISCREPANCIA_PO
        email_prov="admin@pescadosbcn.com",
        dir_prov="Mercabarna, Sector Carnis num. 12, 08040 Barcelona",
    ),
    # F&B-3: ALERTA_CONSUMO — Carnes Premium, PO 2.400, factura 2.390 (<1% OK), POS >>
    dict(
        nombre_archivo="FAC-2025-CARN-0019.pdf",
        num_factura="FAC-2025-CARN-0019",
        fecha=hoy,
        proveedor="Carnes Premium SL",
        nif_prov="B-08567890",
        descripcion="Carnes y embutidos julio primera quincena — solomillo, entrecot, iberico",
        base_imp=2172.73,
        pct_iva=10,
        cuota_iva=217.27,
        total=2390.00,      # PO 2.400 → diff 0.42% ≤1% → PO OK; POS(172.600) vs 2.390 >> 15% → ALERTA
        email_prov="contabilidad@carnespremium.es",
        dir_prov="Mercabarna, Sector Carnis num. 34, 08040 Barcelona",
    ),
    # OTR-1: MATCH_CORRECTO — Endesa, PO 18.500, factura 18.500 (0%)
    dict(
        nombre_archivo="FAC-2025-ENDE-0742.pdf",
        num_factura="FAC-2025-ENDE-0742",
        fecha=hoy,
        proveedor="Endesa Energia SA",
        nif_prov="A-81948077",
        descripcion="Suministro electrico julio 2025 — CUPS: ES0021000012345678ZF",
        base_imp=15289.26,
        pct_iva=21,
        cuota_iva=3210.74,
        total=18500.00,     # = PO-0707 exacto → MATCH_CORRECTO
        email_prov="facturacion@endesa.com",
        dir_prov="Ribera del Loira 60, 28042 Madrid",
    ),
    # OTR-2: SIN_PO — Gas Natural Fenosa, sin PO registrada
    dict(
        nombre_archivo="FAC-2025-GASN-0031.pdf",
        num_factura="FAC-2025-GASN-0031",
        fecha=hoy,
        proveedor="Gas Natural Fenosa SA",
        nif_prov="A-08015497",
        descripcion="Suministro gas natural julio 2025 — Referencia contrato: CN-BCN-28491",
        base_imp=3471.07,
        pct_iva=21,
        cuota_iva=728.93,
        total=4200.00,      # Sin PO → SIN_PO
        email_prov="facturas@gasnaturalsa.es",
        dir_prov="Avda. de San Luis 77, 28033 Madrid",
    ),
]

print("\n=== CREANDO PDFs DE FACTURAS DE PRUEBA ===")
for fac in facturas:
    crear_factura_pdf(**fac)

print(f"\n✓ 5 facturas guardadas en: {ENTRADA_DIR}")
print("\nEscenarios diseñados:")
print("  F&B-1 Distribuidora Global:  172.000 EUR → espera ALERTA_CONSUMO o MATCH_3WAY_OK")
print("  F&B-2 Pescados Barcelona:      2.200 EUR → espera DISCREPANCIA_PO (PO=1.850)")
print("  F&B-3 Carnes Premium:          2.390 EUR → espera ALERTA_CONSUMO (PO OK, POS>>)")
print("  OTR-1 Endesa:                 18.500 EUR → espera MATCH_CORRECTO (PO=18.500)")
print("  OTR-2 Gas Natural Fenosa:      4.200 EUR → espera SIN_PO (sin orden)")
