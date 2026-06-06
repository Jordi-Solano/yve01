"""
Análisis avanzado para Calipolis
Tendencias, comparativas, F&B analysis
"""
from dashboard_calipolis import get_hoteles_calipolis, get_consolidado, get_hotel_detail
import statistics

def get_tendencias_hotel(hotel_id):
    """Retorna tendencias de un hotel (últimos 3 meses)"""
    h = get_hotel_detail(hotel_id)
    if not h or not h.get('kpis_historicos'):
        return None
    
    meses = sorted(h['kpis_historicos'].keys())
    if len(meses) < 2:
        return None
    
    # Últimos 3 meses
    meses_recientes = meses[-3:]
    
    ocupacion_trend = [h['kpis_historicos'][mes]['ocupacion'] for mes in meses_recientes]
    revenue_trend = [h['kpis_historicos'][mes]['total_ingresos'] for mes in meses_recientes]
    gop_trend = [h['kpis_historicos'][mes]['gop_pct'] for mes in meses_recientes]
    
    # Calcular variación
    if len(ocupacion_trend) > 1:
        var_occ = ocupacion_trend[-1] - ocupacion_trend[0]
    else:
        var_occ = 0
    
    if len(revenue_trend) > 1:
        var_rev = revenue_trend[-1] - revenue_trend[0]
    else:
        var_rev = 0
    
    if len(gop_trend) > 1:
        var_gop = gop_trend[-1] - gop_trend[0]
    else:
        var_gop = 0
    
    return {
        "meses": meses_recientes,
        "ocupacion": ocupacion_trend,
        "revenue": revenue_trend,
        "gop_pct": gop_trend,
        "variacion_ocupacion": round(var_occ, 1),
        "variacion_revenue": round(var_rev, 0),
        "variacion_gop": round(var_gop, 1),
        "tendencia_ocupacion": "↑" if var_occ > 0 else "↓" if var_occ < 0 else "→",
        "tendencia_revenue": "↑" if var_rev > 0 else "↓" if var_rev < 0 else "→",
        "tendencia_gop": "↑" if var_gop > 0 else "↓" if var_gop < 0 else "→"
    }

def get_analisis_fb(hotel_id):
    """Análisis F&B del hotel"""
    h = get_hotel_detail(hotel_id)
    if not h:
        return None
    
    ingresos_fb = h.get('ingresos_fb', 0)
    fc_pct = h.get('food_cost_pct', 0)
    total_ingresos = h.get('total_ingresos', 1)
    
    fb_pct_revenue = (ingresos_fb / total_ingresos * 100) if total_ingresos > 0 else 0
    
    # Benchmarking simple
    if fc_pct <= 28:
        evaluacion = "Excelente"
        color = "#1db954"
    elif fc_pct <= 32:
        evaluacion = "Bueno"
        color = "#ff9800"
    else:
        evaluacion = "Revisar"
        color = "#e05252"
    
    return {
        "ingresos_fb": round(ingresos_fb, 2),
        "ingresos_fb_pct_revenue": round(fb_pct_revenue, 1),
        "food_cost_pct": fc_pct,
        "evaluacion": evaluacion,
        "color": color,
        "benchmark_target": 28,
        "diferencia_benchmark": round(fc_pct - 28, 1)
    }

def get_comparativa_hoteles():
    """Comparativa entre los 3 hoteles"""
    hoteles = get_hoteles_calipolis()
    
    comparativa = []
    for h in hoteles:
        comparativa.append({
            "nombre": h['nombre'],
            "categoria": h['categoria'],
            "ocupacion": h['ocupacion'],
            "adr": h['adr'],
            "revpar": h['revpar'],
            "revenue": h['total_ingresos'],
            "gop_pct": h['gop_pct'],
            "fb_cost": h['food_cost_pct'],
            "score_operacional": calcular_score(h)
        })
    
    # Ordenar por revenue
    comparativa.sort(key=lambda x: x['revenue'], reverse=True)
    
    return comparativa

def calcular_score(hotel):
    """Calcula un score operacional (0-100)"""
    score = 0
    
    # Ocupación (max 30 pts)
    score += min(30, (hotel['ocupacion'] / 100) * 30)
    
    # GOP% (max 40 pts)
    score += min(40, (hotel['gop_pct'] / 50) * 40)
    
    # F&B Cost inverso (max 20 pts)
    fb_score = max(0, 20 - (hotel['food_cost_pct'] / 35) * 20)
    score += fb_score
    
    # Penalización por alertas
    # (simplificado, en real sería del hotel)
    
    return round(score, 0)

def get_resumen_operacional():
    """Resumen ejecutivo de Calipolis"""
    hoteles = get_hoteles_calipolis()
    consolidado = get_consolidado()
    
    # Mejores/peores performers
    scores = [(h['nombre'], calcular_score(h)) for h in hoteles]
    scores.sort(key=lambda x: x[1], reverse=True)
    
    mejor = scores[0]
    peor = scores[-1]
    
    # Promedio
    avg_score = statistics.mean([s[1] for s in scores])
    
    return {
        "grupo": consolidado['grupo'],
        "score_grupal": round(avg_score, 0),
        "mejor_hotel": mejor[0],
        "mejor_score": mejor[1],
        "peor_hotel": peor[0],
        "peor_score": peor[1],
        "total_revenue": consolidado['total_revenue_mtd'],
        "avg_gop": consolidado['avg_gop_pct'],
        "hoteles_warning": sum(1 for h in hoteles if calcular_score(h) < 50),
        "hoteles_ok": sum(1 for h in hoteles if calcular_score(h) >= 50)
    }

if __name__ == "__main__":
    print("ANÁLISIS CALIPOLIS")
    print("=" * 70)
    
    comparativa = get_comparativa_hoteles()
    print("\nCOMPARATIVA:")
    for h in comparativa:
        print(f"  {h['nombre']}: Score={h['score_operacional']}, Revenue=€{h['revenue']:,.0f}, GOP={h['gop_pct']}%")
    
    print("\nRESUMEN OPERACIONAL:")
    resumen = get_resumen_operacional()
    print(f"  Score Grupal: {resumen['score_grupal']}")
    print(f"  Mejor: {resumen['mejor_hotel']} ({resumen['mejor_score']})")
    print(f"  Peor: {resumen['peor_hotel']} ({resumen['peor_score']})")
    print(f"  Hoteles OK: {resumen['hoteles_ok']}")
    print(f"  Hoteles Warning: {resumen['hoteles_warning']}")
