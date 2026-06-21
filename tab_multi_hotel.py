"""
tab_multi_hotel.py — Yve.01 Multi-Hotel Dashboard
Provides overview, rankings, and alerts for hotel groups
"""
import os as _os, pandas as _pd
from flask import Blueprint, jsonify

multi_hotel_bp = Blueprint('multi_hotel', __name__)
BASE_DIR = _os.path.dirname(_os.path.abspath(__file__))
DATOS    = _os.path.join(BASE_DIR, 'datos-referencia')

def _load_kpis():
    ruta = _os.path.join(DATOS, 'kpis_hoteles.xlsx')
    if not _os.path.exists(ruta): return _pd.DataFrame()
    return _pd.read_excel(ruta)

def _latest_month(df):
    """Return DataFrame filtered to the most recent month."""
    if df.empty: return df
    latest = df['mes'].max()
    return df[df['mes'] == latest].copy()

@multi_hotel_bp.route('/api/multi_hotel/overview')
def api_multi_overview():
    """Consolidated overview. Accepts ?mes=YYYY-MM&grupo=calipolis."""
    from flask import request as _freq
    df = _load_kpis()
    if df.empty:
        return jsonify({'ok': False, 'error': 'Sin datos KPI'}), 404

    # Filtro de grupo (para Calipolis y futuros grupos)
    grupo_param = _freq.args.get('grupo', '').strip().lower()
    if grupo_param and 'grupo' in df.columns:
        df = df[df['grupo'].str.lower() == grupo_param].copy()
    if df.empty:
        return jsonify({'ok': False, 'error': f'Sin datos para grupo: {grupo_param}'}), 404

    all_months = sorted(df['mes'].unique().tolist())
    mes_param = _freq.args.get('mes', '').strip()
    if mes_param and mes_param in all_months:
        latest = df[df['mes'] == mes_param].copy()
    else:
        latest = _latest_month(df)
    all_months = sorted(df['mes'].unique())
    
    # Consolidado
    total_hab    = int(latest['habitaciones'].sum())
    total_rev    = float(latest['total_ingresos'].sum())
    avg_occ      = round(float(latest['ocupacion_pct'].mean()), 1)
    avg_gop      = round(float(latest['gop_pct'].mean()), 1)
    avg_revpar   = round(float(latest['revpar_eur'].mean()), 2)
    avg_adr      = round(float(latest['adr_eur'].mean()), 2)
    total_gop    = float(latest['gop_eur'].sum())
    total_ap     = int(latest['facturas_ap_pendientes'].sum())
    total_alerts = int(latest['alertas_activas'].sum())
    
    # Revenue trend (last 6 months, all hotels combined)
    rev_trend = []
    for mes in all_months:
        m_df = df[df['mes'] == mes]
        rev_trend.append({
            'mes': mes,
            'revenue': float(m_df['total_ingresos'].sum()),
            'gop': round(float(m_df['gop_pct'].mean()), 1),
            'occ': round(float(m_df['ocupacion_pct'].mean()), 1),
        })
    
    hoteles = []
    for _, row in latest.iterrows():
        # YoY trend (vs first month)
        hist = df[df['hotel_id'] == row['hotel_id']].sort_values('mes')
        prev_rev = float(hist.iloc[-2]['total_ingresos']) if len(hist) >= 2 else float(row['total_ingresos'])
        rev_delta_pct = round((float(row['total_ingresos']) / prev_rev - 1) * 100, 1) if prev_rev > 0 else 0
        
        # Build GOP% trend for sparkline
        h_hist = df[df['hotel_id'] == row['hotel_id']].sort_values('mes')
        gop_trend = [round(float(v), 1) for v in h_hist['gop_pct'].tolist()]
        # Hotel stars from name heuristic
        nombre = str(row['hotel_nombre'])
        stars = '5★' if 'boutique' in nombre.lower() or '5' in nombre else '4★'
        hoteles.append({
            'hotel_id':   str(row['hotel_id']),
            'nombre':     nombre,
            'stars':      stars,
            'habitaciones': int(row['habitaciones']),
            'ocupacion_pct': float(row['ocupacion_pct']),
            'adr_eur':    float(row['adr_eur']),
            'revpar_eur': float(row['revpar_eur']),
            'gop_pct':    float(row['gop_pct']),
            'gop_eur':    float(row.get('gop_eur', 0)),
            'total_ingresos': float(row['total_ingresos']),
            'rev_delta_pct':  rev_delta_pct,
            'facturas_ap':    int(row.get('facturas_ap_pendientes', 0)),
            'alertas':        int(row.get('alertas_activas', 0)),
            'oob_dias':       int(row.get('out_of_balance_dias', 0)),
            'estado_oracle':  str(row.get('estado_oracle', 'SIMULACION')),
            'gop_trend':      gop_trend,
            'avg_gop':        round(float(h_hist['gop_pct'].mean()), 1),
        })
    
    # Sort by GOP%
    hoteles.sort(key=lambda h: h['gop_pct'], reverse=True)
    
    return jsonify({
        'ok': True,
        'mes_actual': latest['mes'].iloc[0] if len(latest) else '',
        'meses_disponibles': all_months,
        'meses_disponibles': all_months,
        'consolidado': {
            'total_habitaciones': total_hab,
            'total_revenue': round(total_rev, 2),
            'avg_occ_pct': avg_occ,
            'avg_gop_pct': avg_gop,
            'avg_revpar': avg_revpar,
            'avg_adr': avg_adr,
            'total_gop': round(total_gop, 2),
            'n_hoteles': len(hoteles),
            'ap_pendientes': total_ap,
            'alertas_activas': total_alerts,
        },
        'hoteles': hoteles,
        'grupos': [{'id': 'CAL', 'nombre': 'Grupo Calipolis', 'n_hoteles': len(hoteles)}],
        'rev_trend': rev_trend,
    })

@multi_hotel_bp.route('/api/multi_hotel/rankings')
def api_multi_rankings():
    """Hotel performance rankings."""
    df = _load_kpis()
    if df.empty: return jsonify({'ok': False, 'error': 'Sin datos'}), 404
    latest = _latest_month(df)
    
    by_gop    = latest.sort_values('gop_pct', ascending=False)[['hotel_nombre','gop_pct']].to_dict('records')
    by_revpar = latest.sort_values('revpar_eur', ascending=False)[['hotel_nombre','revpar_eur']].to_dict('records')
    by_occ    = latest.sort_values('ocupacion_pct', ascending=False)[['hotel_nombre','ocupacion_pct']].to_dict('records')
    by_rev    = latest.sort_values('total_ingresos', ascending=False)[['hotel_nombre','total_ingresos']].to_dict('records')
    
    return jsonify({'ok': True, 'gop': by_gop, 'revpar': by_revpar, 'occ': by_occ, 'revenue': by_rev})

@multi_hotel_bp.route('/api/multi_hotel/alertas')
def api_multi_alertas():
    """Active alerts across all hotels."""
    df = _load_kpis()
    if df.empty: return jsonify({'ok': True, 'alertas': []})
    latest = _latest_month(df)
    
    alertas = []
    for _, row in latest.iterrows():
        nombre = str(row['hotel_nombre'])
        if int(row.get('out_of_balance_dias', 0)) > 0:
            alertas.append({'hotel': nombre, 'tipo': 'OOB', 'msg': f"Out of Balance {row['out_of_balance_dias']}d", 'nivel': 'error'})
        if int(row.get('facturas_ap_pendientes', 0)) > 5:
            alertas.append({'hotel': nombre, 'tipo': 'AP', 'msg': f"{row['facturas_ap_pendientes']} facturas AP pendientes", 'nivel': 'warning'})
        if int(row.get('alertas_activas', 0)) > 0:
            alertas.append({'hotel': nombre, 'tipo': 'AR', 'msg': f"{row['alertas_activas']} alertas AR activas", 'nivel': 'warning'})
    
    return jsonify({'ok': True, 'alertas': alertas, 'total': len(alertas)})

@multi_hotel_bp.route('/api/multi_hotel/hotel/<hotel_id>')
def api_hotel_detail(hotel_id):
    """Detail for a specific hotel: 6-month trend."""
    df = _load_kpis()
    hotel_df = df[df['hotel_id'] == hotel_id].sort_values('mes')
    if hotel_df.empty:
        return jsonify({'ok': False, 'error': f'Hotel {hotel_id} no encontrado'}), 404
    
    latest = hotel_df.iloc[-1]
    trend = hotel_df[['mes','ocupacion_pct','adr_eur','revpar_eur','gop_pct','total_ingresos']].to_dict('records')
    
    return jsonify({
        'ok': True,
        'hotel_id': hotel_id,
        'nombre': str(latest['hotel_nombre']),
        'latest': {k: (float(v) if isinstance(v, (int, float)) else str(v)) for k, v in latest.items()},
        'trend': trend,
    })
