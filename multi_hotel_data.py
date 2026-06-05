"""
Generador de datos demo realistas para Multi-Hotel Dashboard
Estructura escalable: desde grupos pequeños (3 hoteles) hasta cadenas grandes (100+)
"""
import json
import os
import random
from datetime import datetime, timedelta

# Estructura de datos: cada hotel tiene métricas, finanzas, alertas
HOTELES_DEMO = [
    # Grupo 1: Cadena pequeña (3 hoteles Calipolis-like)
    {
        "id": "h001",
        "nombre": "Sitges Beach Hotel",
        "grupo": "Mediterranean Hotels Group",
        "ciudad": "Sitges",
        "pais": "España",
        "tier": "4★",
        "habitaciones": 158,
        "ocupacion_pct": 87.3,
        "adr": 142.50,
        "revpar": 124.40,
        "revenue_mtd": 587450,
        "gop_pct": 38.2,
        "fb_pct": 18.6,
        "facturas_pendientes": 23,
        "facturas_importe": 47820,
        "alertas": ["F&B cost 0.6% sobre target"],
        "status": "ok"
    },
    {
        "id": "h002",
        "nombre": "Sitges Promenade Resort",
        "grupo": "Mediterranean Hotels Group",
        "ciudad": "Sitges",
        "pais": "España",
        "tier": "4★",
        "habitaciones": 210,
        "ocupacion_pct": 91.8,
        "adr": 158.20,
        "revpar": 145.20,
        "revenue_mtd": 842900,
        "gop_pct": 41.5,
        "fb_pct": 17.2,
        "facturas_pendientes": 31,
        "facturas_importe": 62100,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h003",
        "nombre": "Sitges Marina Suites",
        "grupo": "Mediterranean Hotels Group",
        "ciudad": "Sitges",
        "pais": "España",
        "tier": "5★",
        "habitaciones": 89,
        "ocupacion_pct": 78.4,
        "adr": 285.00,
        "revpar": 223.42,
        "revenue_mtd": 597230,
        "gop_pct": 35.8,
        "fb_pct": 22.1,
        "facturas_pendientes": 18,
        "facturas_importe": 38500,
        "alertas": ["Ocupación 4.2% bajo budget", "F&B cost crítico: 22.1%"],
        "status": "warning"
    },
    # Grupo 2: Cadena grande internacional (8 hoteles "European Premier")
    {
        "id": "h101",
        "nombre": "Premier Barcelona Diagonal",
        "grupo": "European Premier Hotels",
        "ciudad": "Barcelona",
        "pais": "España",
        "tier": "5★",
        "habitaciones": 412,
        "ocupacion_pct": 90.1,
        "adr": 245.80,
        "revpar": 221.46,
        "revenue_mtd": 2754320,
        "gop_pct": 42.8,
        "fb_pct": 16.5,
        "facturas_pendientes": 87,
        "facturas_importe": 312400,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h102",
        "nombre": "Premier Madrid Recoletos",
        "grupo": "European Premier Hotels",
        "ciudad": "Madrid",
        "pais": "España",
        "tier": "5★",
        "habitaciones": 387,
        "ocupacion_pct": 88.7,
        "adr": 232.40,
        "revpar": 206.14,
        "revenue_mtd": 2412780,
        "gop_pct": 40.2,
        "fb_pct": 17.8,
        "facturas_pendientes": 72,
        "facturas_importe": 285900,
        "alertas": ["3 facturas sin DI cert"],
        "status": "warning"
    },
    {
        "id": "h103",
        "nombre": "Premier Paris Champs",
        "grupo": "European Premier Hotels",
        "ciudad": "Paris",
        "pais": "Francia",
        "tier": "5★",
        "habitaciones": 290,
        "ocupacion_pct": 92.4,
        "adr": 412.50,
        "revpar": 381.15,
        "revenue_mtd": 3289400,
        "gop_pct": 45.1,
        "fb_pct": 19.2,
        "facturas_pendientes": 64,
        "facturas_importe": 421800,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h104",
        "nombre": "Premier London Mayfair",
        "grupo": "European Premier Hotels",
        "ciudad": "London",
        "pais": "Reino Unido",
        "tier": "5★",
        "habitaciones": 245,
        "ocupacion_pct": 89.6,
        "adr": 485.20,
        "revpar": 434.74,
        "revenue_mtd": 3194580,
        "gop_pct": 43.7,
        "fb_pct": 18.9,
        "facturas_pendientes": 91,
        "facturas_importe": 487200,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h105",
        "nombre": "Premier Berlin Mitte",
        "grupo": "European Premier Hotels",
        "ciudad": "Berlin",
        "pais": "Alemania",
        "tier": "4★",
        "habitaciones": 320,
        "ocupacion_pct": 75.2,
        "adr": 178.30,
        "revpar": 134.08,
        "revenue_mtd": 1287400,
        "gop_pct": 32.4,
        "fb_pct": 21.3,
        "facturas_pendientes": 54,
        "facturas_importe": 198400,
        "alertas": ["GOP 8% bajo budget", "F&B cost crítico"],
        "status": "critical"
    },
    {
        "id": "h106",
        "nombre": "Premier Roma Veneto",
        "grupo": "European Premier Hotels",
        "ciudad": "Roma",
        "pais": "Italia",
        "tier": "5★",
        "habitaciones": 195,
        "ocupacion_pct": 86.4,
        "adr": 298.50,
        "revpar": 257.91,
        "revenue_mtd": 1498200,
        "gop_pct": 39.8,
        "fb_pct": 18.4,
        "facturas_pendientes": 41,
        "facturas_importe": 187300,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h107",
        "nombre": "Premier Amsterdam Center",
        "grupo": "European Premier Hotels",
        "ciudad": "Amsterdam",
        "pais": "Países Bajos",
        "tier": "4★",
        "habitaciones": 178,
        "ocupacion_pct": 93.8,
        "adr": 215.40,
        "revpar": 202.04,
        "revenue_mtd": 1115400,
        "gop_pct": 41.2,
        "fb_pct": 16.8,
        "facturas_pendientes": 38,
        "facturas_importe": 142600,
        "alertas": [],
        "status": "ok"
    },
    {
        "id": "h108",
        "nombre": "Premier Lisbon Avenida",
        "grupo": "European Premier Hotels",
        "ciudad": "Lisboa",
        "pais": "Portugal",
        "tier": "4★",
        "habitaciones": 156,
        "ocupacion_pct": 84.7,
        "adr": 168.20,
        "revpar": 142.46,
        "revenue_mtd": 689400,
        "gop_pct": 36.5,
        "fb_pct": 19.6,
        "facturas_pendientes": 28,
        "facturas_importe": 89400,
        "alertas": ["2 facturas duplicadas detectadas"],
        "status": "warning"
    }
]

def get_hoteles():
    """Devuelve lista de hoteles con datos actualizados"""
    return HOTELES_DEMO

def get_grupos():
    """Devuelve grupos hoteleros únicos"""
    return list(set(h["grupo"] for h in HOTELES_DEMO))

def get_kpis_consolidados(grupo=None):
    """Calcula KPIs consolidados del grupo o de todos"""
    hoteles = [h for h in HOTELES_DEMO if not grupo or h["grupo"] == grupo]
    
    if not hoteles:
        return {}
    
    total_revenue = sum(h["revenue_mtd"] for h in hoteles)
    total_rooms = sum(h["habitaciones"] for h in hoteles)
    avg_occupancy = sum(h["ocupacion_pct"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_adr = sum(h["adr"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_revpar = sum(h["revpar"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_gop = sum(h["gop_pct"] * h["revenue_mtd"] for h in hoteles) / total_revenue
    total_facturas = sum(h["facturas_pendientes"] for h in hoteles)
    total_facturas_eur = sum(h["facturas_importe"] for h in hoteles)
    total_alertas = sum(len(h["alertas"]) for h in hoteles)
    hoteles_criticos = sum(1 for h in hoteles if h["status"] == "critical")
    hoteles_warning = sum(1 for h in hoteles if h["status"] == "warning")
    
    return {
        "num_hoteles": len(hoteles),
        "total_rooms": total_rooms,
        "total_revenue_mtd": total_revenue,
        "avg_occupancy": round(avg_occupancy, 1),
        "avg_adr": round(avg_adr, 2),
        "avg_revpar": round(avg_revpar, 2),
        "avg_gop_pct": round(avg_gop, 1),
        "total_facturas_pendientes": total_facturas,
        "total_facturas_importe": total_facturas_eur,
        "total_alertas": total_alertas,
        "hoteles_criticos": hoteles_criticos,
        "hoteles_warning": hoteles_warning,
        "hoteles_ok": len(hoteles) - hoteles_criticos - hoteles_warning
    }

def get_top_performers(grupo=None, metric="revpar", limit=5):
    """Ranking de top hoteles por métrica"""
    hoteles = [h for h in HOTELES_DEMO if not grupo or h["grupo"] == grupo]
    return sorted(hoteles, key=lambda h: h.get(metric, 0), reverse=True)[:limit]

def get_alertas_consolidadas(grupo=None):
    """Lista todas las alertas de hoteles del grupo"""
    hoteles = [h for h in HOTELES_DEMO if not grupo or h["grupo"] == grupo]
    alertas = []
    for h in hoteles:
        for a in h["alertas"]:
            alertas.append({
                "hotel": h["nombre"],
                "ciudad": h["ciudad"],
                "alerta": a,
                "severity": h["status"]
            })
    return alertas

if __name__ == "__main__":
    print("Multi-Hotel Demo Data")
    print("="*60)
    print(f"Total hoteles: {len(HOTELES_DEMO)}")
    print(f"Grupos: {get_grupos()}")
    print()
    print("KPIs Consolidados (todos):")
    kpis = get_kpis_consolidados()
    for k, v in kpis.items():
        print(f"  {k}: {v}")
