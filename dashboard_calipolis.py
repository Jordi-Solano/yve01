"""
dashboard_calipolis.py — Yve
Datos reales de hoteles.json + kpis_hoteles.xlsx (6 meses 2026)
"""
import pandas as pd, json, os
from pathlib import Path
import time as _t

BASE_DIR  = Path(__file__).parent
DATOS     = BASE_DIR / "datos-referencia"

_CAL_CACHE = {}
_CAL_TTL   = 300

def _cal_excel(path):
    key = str(path)
    now = _t.time()
    if key in _CAL_CACHE:
        df, ts = _CAL_CACHE[key]
        if now - ts < _CAL_TTL: return df
    df = pd.read_excel(path)
    _CAL_CACHE[key] = (df, now)
    return df

def _load():
    with open(DATOS / "hoteles.json") as f:
        hoteles = json.load(f)
    df = _cal_excel(DATOS / "kpis_hoteles.xlsx")
    df["mes"] = df["mes"].astype(str)
    return hoteles, df

def get_hoteles_calipolis():
    hoteles, df = _load()
    # Latest month per hotel
    latest = df.sort_values("mes").groupby("hotel_id").last().reset_index()
    result = []
    for h in hoteles:
        row = latest[latest["hotel_id"] == h["id"]]
        if row.empty:
            continue
        r = row.iloc[0]
        result.append({
            "id": h["id"], "nombre": h["nombre"], "categoria": h["categoria"],
            "habitaciones": h["habitaciones"], "ciudad": h["ciudad"],
            "contacto": h.get("contacto",""), "modulos": h.get("modulos",[]),
            "mes_actual": str(r["mes"]),
            "ocupacion": float(r["ocupacion_pct"]),
            "adr": float(r["adr_eur"]),
            "revpar": float(r["revpar_eur"]),
            "ingresos_rooms": float(r["ingresos_rooms"]),
            "ingresos_fb": float(r["ingresos_fb"]),
            "total_ingresos": float(r["total_ingresos"]),
            "gop": float(r["gop_eur"]),
            "gop_pct": float(r["gop_pct"]),
            "food_cost_pct": float(r["food_cost_pct"]),
            "ap_pendientes": int(r["facturas_ap_pendientes"]),
            "ar_pendientes": int(r["facturas_ar_pendientes"]),
            "alertas": int(r["alertas_activas"]),
            "kpis_historicos": get_historico(h["id"], df),
        })
    return result

def get_historico(hotel_id, df):
    rows = df[df["hotel_id"] == hotel_id].sort_values("mes")
    return [
        {"mes": str(r["mes"]), "ocupacion": float(r["ocupacion_pct"]),
         "adr": float(r["adr_eur"]), "revpar": float(r["revpar_eur"]),
         "total_ingresos": float(r["total_ingresos"]), "gop_pct": float(r["gop_pct"]),
         "food_cost_pct": float(r["food_cost_pct"]),
         "ap_pendientes": int(r["facturas_ap_pendientes"])}
        for _, r in rows.iterrows()
    ]

def get_consolidado():
    hoteles = get_hoteles_calipolis()
    if not hoteles: return {}
    total_rev  = sum(h["total_ingresos"] for h in hoteles)
    total_gop  = sum(h["gop"] for h in hoteles)
    total_rooms = sum(h["habitaciones"] for h in hoteles)
    avg_occ    = sum(h["ocupacion"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_adr    = sum(h["adr"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_revpar = sum(h["revpar"] * h["habitaciones"] for h in hoteles) / total_rooms
    avg_gop    = total_gop / total_rev * 100 if total_rev else 0
    return {
        "num_hoteles": len(hoteles),
        "total_rooms": total_rooms,
        "total_revenue_mtd": round(total_rev),
        "total_gop": round(total_gop),
        "avg_ocupacion": round(avg_occ, 1),
        "avg_adr": round(avg_adr, 2),
        "avg_revpar": round(avg_revpar, 2),
        "avg_gop_pct": round(avg_gop, 1),
        "total_ap_pendientes": sum(h["ap_pendientes"] for h in hoteles),
        "total_alertas": sum(h["alertas"] for h in hoteles),
    }

def get_tendencias():
    """Datos agregados mensuales para gráficos de tendencia."""
    _, df = _load()
    monthly = df.groupby("mes").agg(
        gop_pct_grupo=("gop_pct", "mean"),
        ap_pendientes_total=("facturas_ap_pendientes", "sum"),
        total_revenue=("total_ingresos", "sum"),
        alertas_total=("alertas_activas", "sum"),
    ).reset_index().sort_values("mes")
    return {
        "meses": monthly["mes"].tolist(),
        "gop_pct_grupo": [round(v, 1) for v in monthly["gop_pct_grupo"]],
        "ap_pendientes_total": monthly["ap_pendientes_total"].tolist(),
        "total_revenue": [round(v) for v in monthly["total_revenue"]],
        "alertas_total": monthly["alertas_total"].tolist(),
    }

def get_hotel_detail(hotel_id):
    hoteles = get_hoteles_calipolis()
    return next((h for h in hoteles if h["id"] == hotel_id), None)

if __name__ == "__main__":
    import json
    print("Consolidado:", json.dumps(get_consolidado(), indent=2))
    print("\nTendencias:", json.dumps(get_tendencias(), indent=2))
