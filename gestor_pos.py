"""
gestor_pos.py — Yve.01 Módulo AP
Crea/actualiza los datos del POS de F&B.
También crea pos_ordenes.xlsx (POs de ejemplo).
Ejecutar: python gestor_pos.py
"""

import os, random
from datetime import date, timedelta
import pandas as pd

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
REFERENCIA_DIR = os.path.join(BASE_DIR, "datos-referencia")
os.makedirs(REFERENCIA_DIR, exist_ok=True)

POS_FILE    = os.path.join(REFERENCIA_DIR, "pos_ventas.xlsx")
ORDENES_FILE= os.path.join(REFERENCIA_DIR, "pos_ordenes.xlsx")

# ── Datos POS de ejemplo — Julio 2025 ─────────────────────────────────────

def generar_pos_ventas():
    random.seed(42)
    registros = []
    inicio = date(2025, 7, 1)
    # Generar ~1 entrada por día para varias categorías
    conceptos = [
        ("Restaurant - Breakfast", "Restaurant", 22.0, 0.30),
        ("Restaurant - Lunch",     "Restaurant", 45.0, 0.32),
        ("Restaurant - Dinner",    "Restaurant", 68.0, 0.35),
        ("Bar",                    "Bar",         18.0, 0.25),
        ("Minibar",                "Minibar",      8.5, 0.40),
        ("Room Service",           "Room Service", 35.0, 0.38),
    ]
    for dia in range(30):
        fecha = inicio + timedelta(days=dia)
        factor = 1.0 + random.uniform(-0.15, 0.20)
        for desc, cat, precio_base, coste_ratio in conceptos:
            # Simular covers/unidades
            covers = random.randint(30, 180) if cat == "Restaurant" else random.randint(15, 80)
            importe = round(covers * precio_base * factor, 2)
            coste   = round(importe * coste_ratio * random.uniform(0.90, 1.10), 2)
            registros.append({
                "fecha":          fecha.strftime("%d/%m/%Y"),
                "descripcion":    desc,
                "categoria":      cat,
                "covers_unidades":covers,
                "importe_venta":  importe,
                "coste_estimado": coste,
            })
    return pd.DataFrame(registros)

# ── Órdenes de compra (POs) de ejemplo ────────────────────────────────────

def generar_pos_ordenes():
    ordenes = [
        {"numero_po":"PO-2025-0701","proveedor":"Makro Cash & Carry SL",
         "descripcion":"Suministro semanal alimentacion seca y congelados",
         "importe_aprobado":3200.00,"departamento":"F&B","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0702","proveedor":"Pescados Barcelona SL",
         "descripcion":"Pescado fresco y marisco semana 27",
         "importe_aprobado":1850.00,"departamento":"F&B","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0703","proveedor":"Carnes Premium SL",
         "descripcion":"Carnes y embutidos julio primera quincena",
         "importe_aprobado":2400.00,"departamento":"F&B","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0704","proveedor":"Frutas Mercabarna SL",
         "descripcion":"Frutas y verduras frescas julio",
         "importe_aprobado":980.00,"departamento":"F&B","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0705","proveedor":"Bodegas Torres SA",
         "descripcion":"Vinos y cavas carta julio-agosto",
         "importe_aprobado":4100.00,"departamento":"F&B","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0706","proveedor":"Telefonica de Espana SAU",
         "descripcion":"Servicio lineas telefonicas y fibra julio 2025",
         "importe_aprobado":1240.00,"departamento":"Administracion","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0707","proveedor":"Endesa Energia SA",
         "descripcion":"Suministro electrico julio 2025",
         "importe_aprobado":18500.00,"departamento":"Administracion","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0708","proveedor":"Limpiezas BCN SL",
         "descripcion":"Servicio limpieza zonas comunes julio",
         "importe_aprobado":6800.00,"departamento":"Housekeeping","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0709","proveedor":"Otis Elevadores SA",
         "descripcion":"Mantenimiento preventivo ascensores julio",
         "importe_aprobado":890.00,"departamento":"Mantenimiento","fecha":"01/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0710","proveedor":"Securitas Direct SL",
         "descripcion":"Servicio vigilancia y alarmas julio 2025",
         "importe_aprobado":3200.00,"departamento":"Seguridad","fecha":"01/07/2025","estado":"ABIERTO"},
        # Segunda quincena
        {"numero_po":"PO-2025-0715","proveedor":"Makro Cash & Carry SL",
         "descripcion":"Suministro semanal alimentacion seca segunda quincena",
         "importe_aprobado":3100.00,"departamento":"F&B","fecha":"15/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0716","proveedor":"Pescados Barcelona SL",
         "descripcion":"Pescado fresco y marisco semana 29",
         "importe_aprobado":1920.00,"departamento":"F&B","fecha":"15/07/2025","estado":"ABIERTO"},
        {"numero_po":"PO-2025-0717","proveedor":"Carnes Premium SL",
         "descripcion":"Carnes y embutidos julio segunda quincena",
         "importe_aprobado":2550.00,"departamento":"F&B","fecha":"15/07/2025","estado":"ABIERTO"},
    ]
    return pd.DataFrame(ordenes)

def main():
    print("="*60)
    print("  Yve.01 — Gestor POS y Órdenes de Compra")
    print("="*60)

    df_pos = generar_pos_ventas()
    with pd.ExcelWriter(POS_FILE, engine="openpyxl") as w:
        df_pos.to_excel(w, index=False, sheet_name="POS_Ventas")
        ws = w.sheets["POS_Ventas"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col)+4, 35)
    print(f"✅ pos_ventas.xlsx creado: {len(df_pos)} líneas de POS")

    df_po = generar_pos_ordenes()
    with pd.ExcelWriter(ORDENES_FILE, engine="openpyxl") as w:
        df_po.to_excel(w, index=False, sheet_name="Ordenes_Compra")
        ws = w.sheets["Ordenes_Compra"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col)+4, 45)
    print(f"✅ pos_ordenes.xlsx creado: {len(df_po)} órdenes de compra")
    print("="*60)

if __name__ == "__main__":
    main()
