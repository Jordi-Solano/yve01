"""
Blueprint para análisis avanzado de Calipolis
"""
from flask import Blueprint, jsonify
from calipolis_analisis import (
    get_tendencias_hotel, get_analisis_fb, get_comparativa_hoteles,
    get_resumen_operacional, calcular_score
)

calipolis_analisis_bp = Blueprint('calipolis_analisis', __name__)

@calipolis_analisis_bp.route('/api/calipolis/analisis/resumen')
def api_calipolis_resumen():
    """Resumen ejecutivo de Calipolis"""
    resumen = get_resumen_operacional()
    return jsonify(resumen)

@calipolis_analisis_bp.route('/api/calipolis/analisis/comparativa')
def api_calipolis_comparativa():
    """Comparativa entre hoteles"""
    comparativa = get_comparativa_hoteles()
    return jsonify({"hoteles": comparativa})

@calipolis_analisis_bp.route('/api/calipolis/analisis/hotel/<hotel_id>')
def api_calipolis_analisis_hotel(hotel_id):
    """Análisis completo de un hotel"""
    tendencias = get_tendencias_hotel(hotel_id)
    fb_analysis = get_analisis_fb(hotel_id)
    score = None
    
    # Obtener score
    from dashboard_calipolis import get_hotel_detail
    h = get_hotel_detail(hotel_id)
    if h:
        score = calcular_score(h)
    
    return jsonify({
        "tendencias": tendencias,
        "fb_analysis": fb_analysis,
        "score_operacional": score
    })
