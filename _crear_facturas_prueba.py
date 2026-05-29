"""
Genera 5 facturas PDF de prueba con layout limpio para pdfplumber.
Los campos clave están en líneas separadas con etiquetas que el regex puede parsear.
"""
from pathlib import Path
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

BASE_DIR    = Path("/sessions/brave-affectionate-hawking/mnt/yve01")
REF_DIR     = BASE_DIR / "datos-referencia"
ENTRADA_DIR = BASE_DIR / "facturas-entrada"
ENTRADA_DIR.mkdir(exist_ok=True)

# Aseguramos que los datos de referencia están actualizados (ya se hizo antes)

def crear_factura_pdf(nombre_archivo, num_factura, fecha, proveedor,
                      nif_prov, descripcion, base_imp, pct_iva, cuota_iva, total,
                      email_prov="", dir_prov=""):
    ruta = ENTRADA_DIR / nombre_archivo
    c = canvas.Canvas(str(ruta), pagesize=A4)
    W, H = A4

    # ── Cabecera proveedor ──────────────────────────────────────────────────
    c.setFont("Helvetica-Bold", 16)
    c.setFillColorRGB(0.10, 0.10, 0.40)
    c.drawString(2*cm, H - 2.5*cm, proveedor)

    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0.3, 0.3, 0.3)
    c.drawString(2*cm, H - 3.2*cm, dir_prov)
    c.drawString(2*cm, H - 3.8*cm, f"NIF proveedor: {nif_prov}")
    c.drawString(2*cm, H - 4.4*cm, f"Email: {email_prov}")

    # ── Número factura y fecha (en columna derecha, sin solapar) ───────────
    c.setFont("Helvetica-Bold", 11)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(W - 8*cm, H - 2.5*cm, "FACTURA")
    c.setFont("Helvetica", 10)
    c.drawString(W - 8*cm, H - 3.1*cm, f"Numero: {num_factura}")
    c.drawString(W - 8*cm, H - 3.7*cm, f"Fecha emision: {fecha}")
    c.drawString(W - 8*cm, H - 4.3*cm, "Vencimiento: 30 dias")

    # ── Línea ───────────────────────────────────────────────────────────────
    c.setStrokeColorRGB(0.2, 0.2, 0.6)
    c.setLineWidth(1.5)
    c.line(2*cm, H - 5*cm, W - 2*cm, H - 5*cm)

    # ── Facturado a ──────────────────────────────────────────────────────────
    c.setFont("Helvetica-Bold", 9)
    c.setFillColorRGB(0.2, 0.2, 0.5)
    c.drawString(2*cm, H - 5.8*cm, "FACTURADO A:")
    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(2*cm, H - 6.5*cm, "Hilton Barcelona SA")
    c.drawString(2*cm, H - 7.1*cm, "Avda. Diagonal 589-591, 08014 Barcelona")
    c.drawString(2*cm, H - 7.7*cm, "NIF cliente: A-08123456")

    # ── Cabecera tabla ───────────────────────────────────────────────────────
    y_t = H - 9.2*cm
    c.setFillColorRGB(0.15, 0.15, 0.45)
    c.rect(2*cm, y_t, W - 4*cm, 0.65*cm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 9)
    c.setFillColorRGB(1, 1, 1)
    c.drawString(2.3*cm, y_t + 0.2*cm, "DESCRIPCION DEL SERVICIO")

    # Fila concepto
    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)
    y_concepto = y_t - 1.0*cm
    c.drawString(2.3*cm, y_concepto, descripcion[:90])

    # ── Totales — cada campo en su propia línea para que pdfplumber lo lea ──
    y_tot = y_concepto - 2.2*cm
    c.setFont("Helvetica", 10)

    # Base
    c.drawString(2*cm,    y_tot,          "Base imponible EUR:")
    c.drawString(9*cm,    y_tot,          f"{base_imp:.2f}")

    # IVA
    c.drawString(2*cm,    y_tot - 0.7*cm, f"Cuota IVA {int(pct_iva)} por ciento EUR:")
    c.drawString(9*cm,    y_tot - 0.7*cm, f"{cuota_iva:.2f}")

    # Separador
    c.setLineWidth(0.5)
    c.line(2*cm, y_tot - 1.0*cm, W - 2*cm, y_tot - 1.0*cm)

    # Total — label en su línea, valor en la siguiente para evitar solapado
    c.setFont("Helvetica-Bold", 12)
    c.setFillColorRGB(0.1, 0.1, 0.4)
    c.drawString(2*cm, y_tot - 1.7*cm, "TOTAL FACTURA EUR:")
    c.drawString(9*cm, y_tot - 1.7*cm, f"{total:.2f}")

    # ── Bloque de datos estructurados (para parser) ─────────────────────────
    y_datos = y_tot - 3.5*cm
    c.setFont("Helvetica", 7)
    c.setFillColorRGB(0.6, 0.6, 0.6)
    c.drawString(2*cm, y_datos,          "DATOS FACTURA SISTEMA:")
    c.setFont("Courier", 7)
    datos = [
        f"NUMERO_FACTURA={num_factura}",
        f"FECHA={fecha}",
        f"PROVEEDOR={proveedor}",
        f"NIF={nif_prov}",
        f"CONCEPTO={descripcion[:60]}",
        f"BASE_IMPONIBLE={base_imp:.2f}",
        f"IVA_PORCENTAJE={int(pct_iva)}",
        f"CUOTA_IVA={cuota_iva:.2f}",
        f"TOTAL={total:.2f}",
    ]
    for i, d in enumerate(datos):
        c.drawString(2*cm, y_datos - (i+1)*0.45*cm, d)

    # ── Pie ──────────────────────────────────────────────────────────────────
    c.setFont("Helvetica", 7)
    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.drawCentredString(W/2, 2.5*cm, "Forma de pago: Transferencia bancaria 30 dias fecha factura")
    c.drawCentredString(W/2, 2.0*cm, "Documento: justificante fiscal valido - IVA incluido segun detalle")

    c.save()
    print(f"  [OK] {nombre_archivo}  |  TOTAL={total:.2f} EUR")


# ─── Definicion de las 5 facturas ─────────────────────────────────────────────

facturas = [
    dict(
        nombre_archivo="FAC-2025-DIST-0001.pdf",
        num_factura="FAC-2025-DIST-0001",
        fecha="15/07/2025",
        proveedor="Distribuidora Global Alimentacion SA",
        nif_prov="A-17234567",
        descripcion="Suministro mensual alimentos y bebidas julio 2025 contrato marco",
        base_imp=155818.18,
        pct_iva=10,
        cuota_iva=15581.82,
        total=171400.00,
        email_prov="facturas@distglobal.es",
        dir_prov="Pol. Ind. Zona Franca, Carrer B 45, 08040 Barcelona",
    ),
    dict(
        nombre_archivo="FAC-2025-PESC-0027.pdf",
        num_factura="FAC-2025-PESC-0027",
        fecha="15/07/2025",
        proveedor="Pescados Barcelona SL",
        nif_prov="B-08456789",
        descripcion="Pescado fresco y marisco semana 27 gamba lubina dorada rodaballo",
        base_imp=2000.00,
        pct_iva=10,
        cuota_iva=200.00,
        total=2200.00,
        email_prov="admin@pescadosbcn.com",
        dir_prov="Mercabarna, Sector Carnis 12, 08040 Barcelona",
    ),
    dict(
        nombre_archivo="FAC-2025-CARN-0019.pdf",
        num_factura="FAC-2025-CARN-0019",
        fecha="15/07/2025",
        proveedor="Carnes Premium SL",
        nif_prov="B-08567890",
        descripcion="Carnes y embutidos julio primera quincena solomillo entrecot iberico",
        base_imp=2172.73,
        pct_iva=10,
        cuota_iva=217.27,
        total=2390.00,
        email_prov="contabilidad@carnespremium.es",
        dir_prov="Mercabarna, Sector Carnis 34, 08040 Barcelona",
    ),
    dict(
        nombre_archivo="FAC-2025-ENDE-0742.pdf",
        num_factura="FAC-2025-ENDE-0742",
        fecha="15/07/2025",
        proveedor="Endesa Energia SA",
        nif_prov="A-81948077",
        descripcion="Suministro electrico julio 2025 CUPS ES0021000012345678ZF",
        base_imp=15289.26,
        pct_iva=21,
        cuota_iva=3210.74,
        total=18500.00,
        email_prov="facturacion@endesa.com",
        dir_prov="Ribera del Loira 60, 28042 Madrid",
    ),
    dict(
        nombre_archivo="FAC-2025-GASN-0031.pdf",
        num_factura="FAC-2025-GASN-0031",
        fecha="15/07/2025",
        proveedor="Gas Natural Fenosa SA",
        nif_prov="A-08015497",
        descripcion="Suministro gas natural julio 2025 Referencia contrato CN-BCN-28491",
        base_imp=3471.07,
        pct_iva=21,
        cuota_iva=728.93,
        total=4200.00,
        email_prov="facturas@gasnaturalsa.es",
        dir_prov="Avda. de San Luis 77, 28033 Madrid",
    ),
]

print("=== REGENERANDO 5 FACTURAS PDF ===")
for f in facturas:
    crear_factura_pdf(**f)
print("LISTO")
