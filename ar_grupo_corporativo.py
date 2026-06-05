
"""
ar_grupo_corporativo.py
Módulo AR Real para grupos corporativos — Hilton style
Procesa: BEO, rooming lists, invoices, conciliación master account
"""

import json
import pandas as pd
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
DATOS = BASE_DIR / "datos-referencia"
REPORTES = BASE_DIR / "reportes"
try:
    REPORTES.mkdir(exist_ok=True)
except (PermissionError, OSError):
    pass

# ─ DATOS HILTON ABBVIE ─
def cargar_grupo():
    """Carga datos del grupo corporativo (Abbvie Ovarian Cancer)"""
    grupo = {
        "master_id": "251527287",
        "grupo_nombre": "AbbVie Ovarian Cancer Educational Forum",
        "hotel": "Hilton Barcelona",
        "fechas": {"check_in": "2025-07-03", "check_out": "2025-07-06"},
        "contactos": {
            "organizer": {"name": "Louisa Worsfold", "email": "louisa.worsfold@stratacreate.com", "tel": "+44 7923 992269"},
            "hotel": {"name": "Gemma Ràfols", "email": "events@hiltonbarcelona.com"}
        },
        "rooming_block": {
            "2025-07-03": {"single": 15, "double": 0, "rate": 210},
            "2025-07-04": {"single": 63, "double": 0, "rate": 210},
            "2025-07-05": {"single": 9, "double": 0, "rate": 210},
        },
        "total_rooms": 87,
        "comisiones": {"comisionables": 10, "rate": "10%"},
        "beos": [
            {"beo_id": "8526", "date": "2025-07-03", "function": "Setup", "room": "Priorat", "cost": 2000},
            {"beo_id": "8527", "date": "2025-07-04", "function": "Meeting + Breaks", "room": "Priorat", "cost": 2500},
            {"beo_id": "8528", "date": "2025-07-05", "function": "Meeting + Lunch + Breaks", "room": "Priorat", "cost": 2250},
        ],
        "sow_total": 27494.96,
        "anticipated_fb": 4500,
        "av_supplier_cost": 8716.61,
    }
    return grupo

def cargar_rooming():
    """Carga y valida lista de rooming (87 pax)"""
    try:
        rooming = pd.read_excel(DATOS / "hilton_abbvie_rooming.xlsx", sheet_name="Rooming")
        print(f"✓ Rooming cargado: {len(rooming)} habitaciones")
        return rooming
    except FileNotFoundError:
        # Mock data
        rooming = pd.DataFrame([
            {"guest": f"Attendee {i}", "checkout": "2025-07-06", "room_type": "S", "nights": 3, "rate": 210, "total": 630}
            for i in range(1, 88)
        ])
        print(f"⚠ Mock rooming: {len(rooming)} pax")
        return rooming

def cargar_beos():
    """Carga Banquet Event Orders (3 BEOs para 3-5 julio)"""
    beos = [
        {
            "beo_id": "8526",
            "date": "2025-07-03",
            "time_start": "19:00",
            "time_end": "23:00",
            "function": "Setup + Pre-Con Meeting",
            "room": "Priorat",
            "setup_cost": 2000,
            "av_cost": 10746.28,
            "pomec_cost": 395,
            "pax": 6,
        },
        {
            "beo_id": "8527",
            "date": "2025-07-04",
            "time_start": "09:00",
            "time_end": "19:00",
            "function": "Meeting + Break + Welcome Dinner",
            "room": "Priorat + Alreves Restaurant",
            "setup_cost": 2500,
            "av_cost": 365,
            "coffee_break_cost": 15 * 67,
            "pax": 80,
            "fb_notes": "Dinner offsite at Alreves (menu Diagonal)"
        },
        {
            "beo_id": "8528",
            "date": "2025-07-05",
            "time_start": "08:00",
            "time_end": "16:00",
            "function": "Full Day Meeting + Breaks + Lunch",
            "room": "Priorat",
            "setup_cost": 2250,
            "av_cost": 365,
            "coffee_break_am": 15 * 67,
            "lunch_buffet": 45 * 67,
            "coffee_break_pm": 15 * 67,
            "pax": 80,
        }
    ]
    return beos

def cargar_invoice_hilton():
    """Carga factura hotel (Invoice 125-01444879)"""
    invoice = {
        "invoice_no": "125-01444879",
        "invoice_date": "2025-08-12",
        "master_id": "251527287",
        "guest_company": "AbbVie / Portugal Group",
        "billing_address": "18581 TELLER AVE, IRVINE CA 92612",
        "period": "2025-07-03 to 2025-07-05",
        "rooms_charged": 5,  # Solo uno pagado en el invoice mostrado (rest individual)
        "room_rate": 210,
        "room_charges": 1050,
        "city_tax": 31.35,
        "subtotal_10_percent_vat": 983.05,
        "vat_10_percent": 98.30,
        "total_invoice": 1081.35,
        "currency": "EUR",
        "payment_method": "Credit Card",
        "status": "Posted",
    }
    return invoice

def reconciliar_rooming_vs_invoice():
    """3-way: Rooming contracted vs Invoice actual vs Master account"""
    grupo = cargar_grupo()
    rooming = cargar_rooming()
    invoice = cargar_invoice_hilton()

    total_room_nights_contracted = sum([
        grupo["rooming_block"]["2025-07-03"]["single"] * 1,
        grupo["rooming_block"]["2025-07-04"]["single"] * 1,
        grupo["rooming_block"]["2025-07-05"]["single"] * 1,
    ])

    total_room_revenue_contracted = sum([
        grupo["rooming_block"]["2025-07-03"]["single"] * grupo["rooming_block"]["2025-07-03"]["rate"],
        grupo["rooming_block"]["2025-07-04"]["single"] * grupo["rooming_block"]["2025-07-04"]["rate"],
        grupo["rooming_block"]["2025-07-05"]["single"] * grupo["rooming_block"]["2025-07-05"]["rate"],
    ])

    reconcil = {
        "master_id": grupo["master_id"],
        "grupo": grupo["grupo_nombre"],
        "contracted_rooms": grupo["total_rooms"],
        "contracted_nights": total_room_nights_contracted,
        "contracted_revenue": total_room_revenue_contracted,
        "invoice_rooms": rooming.shape[0] if not rooming.empty else 0,
        "invoice_total": invoice["total_invoice"],
        "variance": total_room_revenue_contracted - invoice["total_invoice"],
        "status": "Para revisión — invoices individuales pendientes" if invoice["total_invoice"] < total_room_revenue_contracted else "Completo",
    }
    return reconcil

def generar_alerta_ar():
    """Detecta discrepancias en AR"""
    reconcil = reconciliar_rooming_vs_invoice()
    alertas = []

    if abs(reconcil["variance"]) > 100:
        alertas.append({
            "nivel": "AVISO",
            "mensaje": f"Varianza de ingresos: €{reconcil['variance']:.2f} — revisar invoices individuales vs master"
        })

    return {
        "master_id": reconcil["master_id"],
        "grupo": reconcil["grupo"],
        "alertas": alertas,
        "status": "Pendiente de cierre AR"
    }

def ejecutar():
    """Pipeline AR Real — rooming → invoice → reconciliación"""
    print("\n=== MÓDULO AR REAL - GRUPO CORPORATIVO ===\n")
    print(f"Master ID: 251527287 (Abbvie Ovarian Cancer)\n")

    grupo = cargar_grupo()
    print(f"Grupo: {grupo['grupo_nombre']}")
    print(f"Hotel: {grupo['hotel']}")
    print(f"Fechas: {grupo['fechas']['check_in']} to {grupo['fechas']['check_out']}")
    print(f"Total rooms contracted: {grupo['total_rooms']}\n")

    rooming = cargar_rooming()
    print(f"Rooming: {len(rooming)} habitaciones en archivo\n")

    beos = cargar_beos()
    beo_costs = sum([b.get('setup_cost', 0) + b.get('av_cost', 0) for b in beos])
    print(f"BEOs: {len(beos)} órdenes de servicio")
    print(f"Total BEO costs (setup + AV): €{beo_costs:.2f}\n")

    invoice = cargar_invoice_hilton()
    print(f"Invoice: {invoice['invoice_no']}")
    print(f"Monto: €{invoice['total_invoice']:.2f}\n")

    reconcil = reconciliar_rooming_vs_invoice()
    print(f"RECONCILIACIÓN:")
    print(f"  Contracted Revenue: €{reconcil['contracted_revenue']:.2f}")
    print(f"  Invoice Total: €{reconcil['invoice_total']:.2f}")
    print(f"  Variance: €{reconcil['variance']:.2f}")
    print(f"  Status: {reconcil['status']}\n")

    alertas = generar_alerta_ar()
    if alertas["alertas"]:
        print(f"ALERTAS AR:")
        for a in alertas["alertas"]:
            print(f"  {a['nivel']}: {a['mensaje']}")

    return {"grupo": grupo, "reconciliacion": reconcil, "alertas": alertas}

if __name__ == "__main__":
    ejecutar()
