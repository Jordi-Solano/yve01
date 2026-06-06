"""
Blueprint Flask para Dashboard Calipolis
Integra datos reales de kpis_hoteles.xlsx y hoteles.json
"""
from flask import Blueprint, jsonify
from dashboard_calipolis import get_hoteles_calipolis, get_consolidado, get_hotel_detail

calipolis_bp = Blueprint('calipolis', __name__)

@calipolis_bp.route('/api/calipolis/kpis')
def api_calipolis_kpis():
    """Retorna KPIs consolidados Calipolis"""
    consolidado = get_consolidado()
    hoteles = get_hoteles_calipolis()
    
    return jsonify({
        "consolidado": consolidado,
        "hoteles": [
            {
                "id": h["id"],
                "nombre": h["nombre"],
                "categoria": h["categoria"],
                "habitaciones": h["habitaciones"],
                "ciudad": h["ciudad"],
                "ocupacion": h["ocupacion"],
                "adr": h["adr"],
                "revpar": h["revpar"],
                "total_ingresos": h["total_ingresos"],
                "gop": h["gop"],
                "gop_pct": h["gop_pct"],
                "food_cost_pct": h["food_cost_pct"],
                "ap_pendientes": h["ap_pendientes"],
                "ar_pendientes": h["ar_pendientes"],
                "alertas": h["alertas"],
                "status": "ok" if h["alertas"] == 0 else "warning" if h["alertas"] <= 1 else "critical"
            }
            for h in hoteles
        ]
    })

@calipolis_bp.route('/api/calipolis/hotel/<hotel_id>')
def api_calipolis_hotel(hotel_id):
    """Retorna detalle de hotel específico"""
    h = get_hotel_detail(hotel_id)
    if not h:
        return jsonify({"error": "Hotel no encontrado"}), 404
    
    return jsonify({
        "id": h["id"],
        "nombre": h["nombre"],
        "categoria": h["categoria"],
        "ciudad": h["ciudad"],
        "habitaciones": h["habitaciones"],
        "contacto": h["contacto"],
        "modulos": h["modulos"],
        "mes_actual": h["mes_actual"],
        "ocupacion": h["ocupacion"],
        "adr": h["adr"],
        "revpar": h["revpar"],
        "ingresos_rooms": h["ingresos_rooms"],
        "ingresos_fb": h["ingresos_fb"],
        "total_ingresos": h["total_ingresos"],
        "gop": h["gop"],
        "gop_pct": h["gop_pct"],
        "food_cost_pct": h["food_cost_pct"],
        "ap_pendientes": h["ap_pendientes"],
        "ar_pendientes": h["ar_pendientes"],
        "alertas": h["alertas"],
        "historico": h["kpis_historicos"]
    })

@calipolis_bp.route('/api/calipolis/comparativa')
def api_calipolis_comparativa():
    """Retorna comparativa de hoteles"""
    hoteles = get_hoteles_calipolis()
    
    return jsonify({
        "hotels": [
            {
                "nombre": h["nombre"],
                "ingresos": h["total_ingresos"],
                "gop_pct": h["gop_pct"],
                "ocupacion": h["ocupacion"],
                "adr": h["adr"]
            }
            for h in hoteles
        ]
    })
