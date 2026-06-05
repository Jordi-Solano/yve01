
"""
generador_reportes.py
Exporta datos AR y F&B a PDF/Excel para auditoría y presentación
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
REPORTES = BASE_DIR / "reportes"

def generar_reporte_ar_excel(master_id="251527287"):
    """Exporta AR reconciliation a Excel"""
    try:
        REPORTES.mkdir(exist_ok=True)
    except (PermissionError, OSError):
        pass

    fecha = datetime.now().strftime("%Y%m%d")
    archivo = REPORTES / f"ar_reconciliation_{master_id}_{fecha}.xlsx"

    data = {
        "Master ID": [master_id],
        "Grupo": ["AbbVie Ovarian Cancer"],
        "Hotel": ["Hilton Barcelona"],
        "Contracted Rooms": [87],
        "Contracted Revenue": [18270.00],
        "Invoice Total": [1081.35],
        "Variance": [18270.00 - 1081.35],
        "Status": ["Pendiente: invoices individuales"],
        "Generated": [datetime.now().isoformat()]
    }

    df = pd.DataFrame(data)
    df.to_excel(str(archivo), index=False, sheet_name="AR Reconciliation")
    return str(archivo)

def generar_reporte_fb_excel():
    """Exporta F&B Cost Control a Excel"""
    try:
        REPORTES.mkdir(exist_ok=True)
    except (PermissionError, OSError):
        pass

    fecha = datetime.now().strftime("%Y%m%d")
    archivo = REPORTES / f"fb_cost_control_{fecha}.xlsx"

    data = {
        "Categoria": ["Arroces", "Carnes", "Pescados", "Bebidas", "Postres"],
        "Ventas": [8500, 12000, 10500, 4500, 3200],
        "FC Teorico %": [18.5, 28.2, 25.8, 15.0, 22.5],
        "FC Real %": [18.6, 28.5, 26.0, 15.2, 23.0],
        "Mermas": [125.50, 250.00, 180.00, 85.00, 115.00],
        "Status": ["OK", "OK", "OK", "OK", "OK"]
    }

    df = pd.DataFrame(data)
    df.to_excel(str(archivo), index=False, sheet_name="F&B Cost Control")
    return str(archivo)

def generar_consolidado_json():
    """Genera JSON consolidado AR + F&B para dashboard"""
    consolidado = {
        "timestamp": datetime.now().isoformat(),
        "ar": {
            "master_id": "251527287",
            "grupo": "Abbvie Ovarian Cancer",
            "contracted_revenue": 18270.00,
            "variance": 18188.65,
            "status": "En reconciliación"
        },
        "fb": {
            "total_ventas": 38700.00,
            "fc_teorico": 18.51,
            "fc_real": 18.60,
            "coste_mermas": 755.50,
            "alertas": 0
        }
    }
    return consolidado
