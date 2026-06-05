"""
Multi-Hotel Dashboard Blueprint
Endpoints para el dashboard multi-hotel
"""
from flask import Blueprint, jsonify, request
from multi_hotel_data import (
    get_hoteles, get_grupos, get_kpis_consolidados,
    get_top_performers, get_alertas_consolidadas
)

multi_hotel_bp = Blueprint('multi_hotel', __name__)

@multi_hotel_bp.route('/api/multi_hotel/overview')
def api_overview():
    """KPIs consolidados + lista de hoteles filtrada por grupo opcional"""
    grupo = request.args.get('grupo')
    
    return jsonify({
        "kpis": get_kpis_consolidados(grupo),
        "hoteles": [h for h in get_hoteles() if not grupo or h["grupo"] == grupo],
        "grupos": get_grupos(),
        "grupo_activo": grupo
    })

@multi_hotel_bp.route('/api/multi_hotel/rankings')
def api_rankings():
    """Top performers por diferentes métricas"""
    grupo = request.args.get('grupo')
    
    return jsonify({
        "top_revenue": get_top_performers(grupo, "revenue_mtd", 5),
        "top_revpar": get_top_performers(grupo, "revpar", 5),
        "top_gop": get_top_performers(grupo, "gop_pct", 5),
        "top_occupancy": get_top_performers(grupo, "ocupacion_pct", 5),
        "bottom_gop": sorted(
            [h for h in get_hoteles() if not grupo or h["grupo"] == grupo],
            key=lambda h: h["gop_pct"]
        )[:5]
    })

@multi_hotel_bp.route('/api/multi_hotel/alertas')
def api_alertas():
    """Alertas consolidadas con severidad"""
    grupo = request.args.get('grupo')
    alertas = get_alertas_consolidadas(grupo)
    
    return jsonify({
        "total": len(alertas),
        "criticas": sum(1 for a in alertas if a["severity"] == "critical"),
        "warnings": sum(1 for a in alertas if a["severity"] == "warning"),
        "lista": alertas
    })

@multi_hotel_bp.route('/api/multi_hotel/hotel/<hotel_id>')
def api_hotel_detail(hotel_id):
    """Detalle completo de un hotel específico"""
    hoteles = get_hoteles()
    hotel = next((h for h in hoteles if h["id"] == hotel_id), None)
    
    if not hotel:
        return jsonify({"error": "Hotel not found"}), 404
    
    return jsonify(hotel)
