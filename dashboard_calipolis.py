"""
Dashboard Calipolis — Vista consolidada de 3 hoteles Sitges
Datos reales de hoteles.json + kpis_hoteles.xlsx
"""
import json
import os
import openpyxl

DATOS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos-referencia")

def cargar_hoteles():
    """Lee hoteles.json"""
    with open(f"{DATOS_PATH}/hoteles.json", "r", encoding="utf-8") as f:
        return json.load(f)

def cargar_kpis():
    """Lee kpis_hoteles.xlsx y retorna dict por hotel_id"""
    wb = openpyxl.load_workbook(f"{DATOS_PATH}/kpis_hoteles.xlsx", data_only=True)
    ws = wb.active
    
    kpis = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        hotel_id = row[0]
        nombre = row[1]
        mes = row[2]
        
        # Columnas: hotel_id, nombre, mes, habitaciones, occ%, adr, revpar, ing_rooms, ing_fb, ing_otros, total_ing, coste_ap, coste_ar, fc%, gop, gop%, ap_pend, ar_pend, alertas, oracle, oob_dias
        if hotel_id not in kpis:
            kpis[hotel_id] = {}
        
        kpis[hotel_id][mes] = {
            "habitaciones": row[3] or 0,
            "ocupacion": row[4] or 0,
            "adr": row[5] or 0,
            "revpar": row[6] or 0,
            "ingresos_rooms": row[7] or 0,
            "ingresos_fb": row[8] or 0,
            "ingresos_otros": row[9] or 0,
            "total_ingresos": row[10] or 0,
            "coste_ap": row[11] or 0,
            "coste_ar": row[12] or 0,
            "food_cost_pct": row[13] or 0,
            "gop": row[14] or 0,
            "gop_pct": row[15] or 0,
            "ap_pendientes": row[16] or 0,
            "ar_pendientes": row[17] or 0,
            "alertas": row[18] or 0,
            "oracle": row[19] or "—",
            "oob_dias": row[20] or 0
        }
    
    return kpis

def get_hoteles_calipolis():
    """Retorna lista de hoteles Calipolis con KPIs"""
    hoteles = cargar_hoteles()
    kpis = cargar_kpis()
    
    result = []
    for h in hoteles:
        kpis_hotel = kpis.get(h["id"], {})
        # Usar último mes disponible (julio)
        meses_disponibles = sorted(kpis_hotel.keys())
        ultimo_mes = meses_disponibles[-1] if meses_disponibles else "2025-07"
        kpi_actual = kpis_hotel.get(ultimo_mes, {})
        
        result.append({
            "id": h["id"],
            "nombre": h["nombre"],
            "ciudad": h["ciudad"],
            "categoria": h["categoria"],
            "habitaciones": h["habitaciones"],
            "contacto": h["contacto"],
            "modulos": h["modulos"],
            "mes_actual": ultimo_mes,
            "ocupacion": kpi_actual.get("ocupacion", 0),
            "adr": kpi_actual.get("adr", 0),
            "revpar": kpi_actual.get("revpar", 0),
            "ingresos_rooms": kpi_actual.get("ingresos_rooms", 0),
            "ingresos_fb": kpi_actual.get("ingresos_fb", 0),
            "total_ingresos": kpi_actual.get("total_ingresos", 0),
            "gop": kpi_actual.get("gop", 0),
            "gop_pct": kpi_actual.get("gop_pct", 0),
            "food_cost_pct": kpi_actual.get("food_cost_pct", 0),
            "ap_pendientes": kpi_actual.get("ap_pendientes", 0),
            "ar_pendientes": kpi_actual.get("ar_pendientes", 0),
            "alertas": kpi_actual.get("alertas", 0),
            "kpis_historicos": kpis_hotel
        })
    
    return result

def get_consolidado():
    """Retorna KPIs consolidados Calipolis"""
    hoteles = get_hoteles_calipolis()
    
    total_rooms = sum(h["habitaciones"] for h in hoteles)
    total_revenue = sum(h["total_ingresos"] for h in hoteles)
    avg_occ = sum(h["ocupacion"] for h in hoteles) / len(hoteles) if hoteles else 0
    avg_adr = sum(h["adr"] for h in hoteles) / len(hoteles) if hoteles else 0
    avg_revpar = sum(h["revpar"] for h in hoteles) / len(hoteles) if hoteles else 0
    total_gop = sum(h["gop"] for h in hoteles)
    avg_gop_pct = sum(h["gop_pct"] for h in hoteles) / len(hoteles) if hoteles else 0
    total_ap_pend = sum(h["ap_pendientes"] for h in hoteles)
    total_ar_pend = sum(h["ar_pendientes"] for h in hoteles)
    total_alertas = sum(h["alertas"] for h in hoteles)
    
    return {
        "grupo": "Calipolis Hotels Group",
        "num_hoteles": len(hoteles),
        "total_rooms": total_rooms,
        "total_revenue_mtd": total_revenue,
        "avg_ocupacion": round(avg_occ, 1),
        "avg_adr": round(avg_adr, 2),
        "avg_revpar": round(avg_revpar, 2),
        "total_gop": round(total_gop, 2),
        "avg_gop_pct": round(avg_gop_pct, 1),
        "total_ap_pendientes": total_ap_pend,
        "total_ar_pendientes": total_ar_pend,
        "total_alertas": total_alertas
    }

def get_hotel_detail(hotel_id):
    """Retorna detalle completo de un hotel"""
    hoteles = get_hoteles_calipolis()
    return next((x for x in hoteles if x["id"] == hotel_id), None)

if __name__ == "__main__":
    print("CALIPOLIS HOTELS GROUP")
    print("=" * 70)
    
    consolidado = get_consolidado()
    print(f"\nGrupo: {consolidado['grupo']}")
    print(f"Hoteles: {consolidado['num_hoteles']}")
    print(f"Habitaciones: {consolidado['total_rooms']}")
    print(f"Revenue MTD: €{consolidado['total_revenue_mtd']:,.0f}")
    print(f"Ocupación Avg: {consolidado['avg_ocupacion']}%")
    print(f"ADR Avg: €{consolidado['avg_adr']:.2f}")
    print(f"RevPAR Avg: €{consolidado['avg_revpar']:.2f}")
    print(f"GOP: €{consolidado['total_gop']:,.0f} ({consolidado['avg_gop_pct']}%)")
    print(f"Facturas AP Pendientes: {consolidado['total_ap_pendientes']}")
    print(f"Facturas AR Pendientes: {consolidado['total_ar_pendientes']}")
    print(f"Alertas Activas: {consolidado['total_alertas']}")
    
    print("\n" + "=" * 70)
    hoteles = get_hoteles_calipolis()
    for h in hoteles:
        print(f"\n{h['nombre']} — {h['categoria']}")
        print(f"  {h['ciudad']} | {h['habitaciones']} rooms")
        print(f"  Revenue: €{h['total_ingresos']:,.0f} (Rooms: €{h['ingresos_rooms']:,.0f} | F&B: €{h['ingresos_fb']:,.0f})")
        print(f"  Occ: {h['ocupacion']}% | ADR: €{h['adr']:.2f} | RevPAR: €{h['revpar']:.2f}")
        print(f"  GOP: €{h['gop']:,.0f} ({h['gop_pct']}%) | F&B Cost: {h['food_cost_pct']}%")
        print(f"  Pendientes: AP={h['ap_pendientes']} | AR={h['ar_pendientes']} | Alertas={h['alertas']}")
