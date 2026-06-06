"""
Demo Mode completo - datos ficticios para presentación
"""
import json
from datetime import datetime, timedelta

def generar_hoteles_demo():
    """3 hoteles ficticios realistas para presentación"""
    return {
        "demo_activo": True,
        "fecha_generacion": datetime.now().isoformat(),
        "hoteles": [
            {
                "id": "DEMO01",
                "nombre": "Barcelona Gran Hotel",
                "categoria": "5 estrellas",
                "ubicacion": "Paseo de Gracia, Barcelona",
                "habitaciones": 250,
                "ocupacion": 89.5,
                "adr": 285,
                "revpar": 255,
                "revenue_hoy": 63750,
                "revenue_mtd": 1875000,
                "ingresos_fb": 450000,
                "ingresos_otros": 125000,
                "gop": 412500,
                "gop_pct": 22.0,
                "ap_pendientes": 3,
                "ar_pendientes": 2,
                "alertas_activas": 1,
                "status": "warning"
            },
            {
                "id": "DEMO02",
                "nombre": "Valencia Beach Resort",
                "categoria": "4 estrellas",
                "ubicacion": "Playa de la Malvarrosa, Valencia",
                "habitaciones": 180,
                "ocupacion": 92.3,
                "adr": 195,
                "revpar": 180,
                "revenue_hoy": 32400,
                "revenue_mtd": 972000,
                "ingresos_fb": 280000,
                "ingresos_otros": 85000,
                "gop": 175000,
                "gop_pct": 18.0,
                "ap_pendientes": 1,
                "ar_pendientes": 0,
                "alertas_activas": 0,
                "status": "ok"
            },
            {
                "id": "DEMO03",
                "nombre": "Sevilla Historic Center",
                "categoria": "4 estrellas",
                "ubicacion": "Barrio Santa Cruz, Sevilla",
                "habitaciones": 95,
                "ocupacion": 87.2,
                "adr": 165,
                "revpar": 144,
                "revenue_hoy": 13680,
                "revenue_mtd": 410400,
                "ingresos_fb": 125000,
                "ingresos_otros": 35000,
                "gop": 82000,
                "gop_pct": 20.0,
                "ap_pendientes": 2,
                "ar_pendientes": 1,
                "alertas_activas": 1,
                "status": "warning"
            }
        ]
    }

def generar_facturas_demo_ar():
    """Facturas AR (OTA) ficticias"""
    return [
        {"id": "BKG-2024-0001", "ota": "Booking.com", "monto": 8500, "comision": 1275, "estado": "procesada", "fecha": "2026-06-06"},
        {"id": "EXP-2024-0002", "ota": "Expedia", "monto": 6200, "comision": 1116, "estado": "procesada", "fecha": "2026-06-06"},
        {"id": "HOT-2024-0003", "ota": "Hotels.com", "monto": 4800, "comision": 864, "estado": "procesada", "fecha": "2026-06-06"},
        {"id": "BKG-2024-0004", "ota": "Booking.com", "monto": 12500, "comision": 1875, "estado": "pendiente_di", "fecha": "2026-06-05"},
        {"id": "EXP-2024-0005", "ota": "Expedia", "monto": 9300, "comision": 1674, "estado": "pendiente_di", "fecha": "2026-06-05"},
    ]

def generar_facturas_demo_ap():
    """Facturas AP (proveedores) ficticias"""
    return [
        {"id": "FAC-2024-001", "proveedor": "Pescados Barcelona SL", "monto": 2850, "tipo": "F&B", "estado": "procesada", "fecha": "2026-06-04"},
        {"id": "FAC-2024-002", "proveedor": "Bebidas Premium SA", "monto": 1600, "tipo": "F&B", "estado": "procesada", "fecha": "2026-06-04"},
        {"id": "FAC-2024-003", "proveedor": "Limpieza Industrial", "monto": 1200, "tipo": "Servicios", "estado": "procesada", "fecha": "2026-06-03"},
        {"id": "FAC-2024-004", "proveedor": "Pescados Barcelona SL", "monto": 3100, "tipo": "F&B", "estado": "pendiente_aprobacion", "fecha": "2026-06-06"},
        {"id": "FAC-2024-005", "proveedor": "Bebidas Premium SA", "monto": 1850, "tipo": "F&B", "estado": "pendiente_aprobacion", "fecha": "2026-06-06"},
    ]

def generar_alertas_demo():
    """Alertas ficticias para dashboard"""
    return [
        {"tipo": "AR DI", "mensaje": "3 facturas sin certificado de doble imposición", "severidad": "warning", "hotel": "DEMO01"},
        {"tipo": "AP Discrepancia", "mensaje": "Discrepancia en PO: Pescados Barcelona", "severidad": "info", "hotel": "DEMO01"},
        {"tipo": "DRR OOB", "mensaje": "Out of Balance: -€245.50 - revisar día 3", "severidad": "error", "hotel": "DEMO02"},
    ]

if __name__ == "__main__":
    print("=== DEMO MODE DATA ===")
    hoteles = generar_hoteles_demo()
    print(f"✓ {len(hoteles['hoteles'])} hoteles demo")
    print(f"✓ {len(generar_facturas_demo_ar())} facturas AR")
    print(f"✓ {len(generar_facturas_demo_ap())} facturas AP")
    print(f"✓ {len(generar_alertas_demo())} alertas")
