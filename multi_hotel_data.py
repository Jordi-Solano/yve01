"""
multi_hotel_data.py — Multi-Hotel Dashboard data
Primary: kpis_hoteles.xlsx (Calipolis real data)
Fallback: HOTELES_DEMO (extended demo for larger groups)
"""
import json, os
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

BASE_DIR = Path(__file__).parent
from tenant_dirs import datos_dir as _t_ddir

class _TDatos:
    def __truediv__(self, other): return Path(_t_ddir()) / other
    def __str__(self): return _t_ddir()

DATOS = _TDatos()

# ── Load real Calipolis data ──────────────────────────────────────────────
def _load_calipolis():
    try:
        df = pd.read_excel(DATOS / "kpis_hoteles.xlsx")
        hotels = []
        for hid in df['hotel_id'].unique():
            hdf  = df[df['hotel_id'] == hid].sort_values('mes')
            last = hdf.iloc[-1]
            # 6-month trend
            trend_gop = hdf['gop_pct'].tolist()
            trend_occ = hdf['ocupacion_pct'].tolist()
            alertas_list = []
            if float(last.get('alertas_activas', 0)) > 0:
                alertas_list.append(f"{int(last['alertas_activas'])} alerta(s) activa(s)")
            if float(last.get('out_of_balance_dias', 0)) > 0:
                alertas_list.append(f"DRR: {int(last['out_of_balance_dias'])} día(s) OOB")
            if float(last.get('facturas_ap_pendientes', 0)) > 8:
                alertas_list.append(f"AP: {int(last['facturas_ap_pendientes'])} facturas pendientes")
            gop = float(last['gop_pct'])
            status = 'ok' if gop >= 20 else ('warning' if gop >= 15 else 'critical')
            hotels.append({
                "id":                  str(hid),
                "nombre":              str(last['hotel_nombre']),
                "grupo":               "Grupo Calipolis",
                "ciudad":              "Sitges",
                "pais":                "España",
                "tier":                "4★",
                "habitaciones":        int(last['habitaciones']),
                "ocupacion_pct":       float(last['ocupacion_pct']),
                "adr":                 float(last['adr_eur']),
                "revpar":              float(last['revpar_eur']),
                "revenue_mtd":         float(last['total_ingresos']),
                "ingresos_rooms":      float(last['ingresos_rooms']),
                "ingresos_fb":         float(last['ingresos_fb']),
                "gop_pct":             gop,
                "gop_eur":             float(last['gop_eur']),
                "food_cost_pct":       float(last['food_cost_pct']),
                "facturas_pendientes": int(last['facturas_ap_pendientes']) + int(last['facturas_ar_pendientes']),
                "facturas_importe":    float(last['coste_ap_eur']),
                "alertas":             alertas_list,
                "status":              status,
                "trend_gop":           trend_gop,
                "trend_occ":           trend_occ,
                "mes_actual":          str(last['mes']),
            })
        return hotels
    except Exception as e:
        print(f"[multi_hotel_data] Error loading Calipolis: {e}")
        return []

_CAL_CACHE = {"data": None, "ts": 0}

def _get_hotels():
    import time
    now = time.time()
    if _CAL_CACHE["data"] is None or now - _CAL_CACHE["ts"] > 300:
        _CAL_CACHE["data"] = _load_calipolis() or HOTELES_DEMO
        _CAL_CACHE["ts"]   = now
    return _CAL_CACHE["data"]

# ── Fallback demo data ────────────────────────────────────────────────────
HOTELES_DEMO = [
    {"id":"h001","nombre":"Sitges Beach Hotel","grupo":"Mediterranean Hotels Group","ciudad":"Sitges","pais":"España","tier":"4★","habitaciones":158,"ocupacion_pct":87.3,"adr":142.50,"revpar":124.40,"revenue_mtd":587450,"gop_pct":38.2,"fb_pct":18.6,"facturas_pendientes":23,"facturas_importe":47820,"alertas":["F&B cost 0.6% sobre target"],"status":"ok"},
    {"id":"h002","nombre":"Barcelona Diagonal","grupo":"Mediterranean Hotels Group","ciudad":"Barcelona","pais":"España","tier":"5★","habitaciones":245,"ocupacion_pct":82.1,"adr":210.00,"revpar":172.41,"revenue_mtd":1124600,"gop_pct":41.8,"fb_pct":19.2,"facturas_pendientes":18,"facturas_importe":89340,"alertas":[],"status":"ok"},
    {"id":"h003","nombre":"Tarragona Resort","grupo":"Mediterranean Hotels Group","ciudad":"Tarragona","pais":"España","tier":"4★","habitaciones":120,"ocupacion_pct":71.2,"adr":98.50,"revpar":70.13,"revenue_mtd":224400,"gop_pct":28.4,"fb_pct":21.8,"facturas_pendientes":34,"facturas_importe":28740,"alertas":["Ocupación baja — 71%","F&B cost alto"],"status":"warning"},
]

# ── Public API ────────────────────────────────────────────────────────────
def get_hoteles(grupo=None):
    hotels = _get_hotels()
    if grupo:
        hotels = [h for h in hotels if h.get("grupo") == grupo]
    return hotels

def get_grupos():
    seen, groups = set(), []
    for h in _get_hotels():
        g = h.get("grupo", "—")
        if g not in seen:
            seen.add(g)
            groups.append(g)
    return groups

def get_kpis_consolidados(grupo=None):
    hotels = get_hoteles(grupo)
    if not hotels: return {}
    total_rev = sum(h.get("revenue_mtd", 0) for h in hotels)
    avg_occ   = sum(h.get("ocupacion_pct", 0) for h in hotels) / len(hotels)
    avg_gop   = sum(h.get("gop_pct", 0) for h in hotels) / len(hotels)
    total_rooms = sum(h.get("habitaciones", 0) for h in hotels)
    total_pend  = sum(h.get("facturas_pendientes", 0) for h in hotels)
    total_alertas = sum(len(h.get("alertas", [])) for h in hotels)
    return {
        "num_hoteles":              len(hotels),
        "total_rooms":              total_rooms,
        "total_revenue_mtd":        round(total_rev, 0),
        "avg_occupancy":            round(avg_occ, 1),
        "avg_adr":                  round(sum(h.get("adr", 0) for h in hotels) / len(hotels), 2),
        "avg_revpar":               round(sum(h.get("revpar", 0) for h in hotels) / len(hotels), 2),
        "avg_gop_pct":              round(avg_gop, 1),
        "total_facturas_pendientes":total_pend,
        "total_alertas":            total_alertas,
        "hoteles_criticos":         sum(1 for h in hotels if h.get("status") == "critical"),
        "hoteles_warning":          sum(1 for h in hotels if h.get("status") == "warning"),
        "hoteles_ok":               sum(1 for h in hotels if h.get("status") == "ok"),
    }

def get_top_performers(grupo=None, metric="gop_pct", n=5):
    hotels = get_hoteles(grupo)
    return sorted(hotels, key=lambda h: h.get(metric, 0), reverse=True)[:n]

def get_alertas_consolidadas(grupo=None):
    alertas = []
    for h in get_hoteles(grupo):
        for a in h.get("alertas", []):
            alertas.append({
                "hotel":    h["nombre"],
                "hotel_id": h["id"],
                "mensaje":  a,
                "severity": "warning",
            })
    return alertas
