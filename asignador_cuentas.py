"""
asignador_cuentas.py — Yve.01 Módulo AP
Asigna cuentas contables a cada factura y genera el asiento contable.
Ejecutar: python asignador_cuentas.py
"""

import os, glob, re
from datetime import date
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
REPORTES_DIR   = os.path.join(BASE_DIR, "reportes")
PROCESADAS_DIR = os.path.join(BASE_DIR, "facturas-procesadas")
REFERENCIA_DIR = os.path.join(BASE_DIR, "datos-referencia")
os.makedirs(REPORTES_DIR, exist_ok=True)

PROV_FILE  = os.path.join(REFERENCIA_DIR, "proveedores.xlsx")
CC_FILE    = os.path.join(REFERENCIA_DIR, "plan_cuentas.xlsx")
FECHA_HOY  = date.today().strftime("%Y%m%d")
SALIDA     = os.path.join(PROCESADAS_DIR, f"facturas_contabilizadas_{FECHA_HOY}.xlsx")
NF         = "NO_ENCONTRADO"

VERDE   = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
AMARILLO= PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
GRIS    = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")

# ── Reglas de asignación ──────────────────────────────────────────────────
# Orden de evaluación: más específico primero

REGLAS_CUENTA_GASTO = [
    # F&B → 600
    ({"tipo_proveedor": "FB"},                                    "600"),
    # Por nombre de proveedor (palabras clave)
    ({"kw_proveedor": ["limpieza","cleaning","housekeeping"]},    "623"),
    ({"kw_proveedor": ["telefon","vodafone","orange","telecom",
                       "movil","internet","fibra"]},              "629"),
    ({"kw_proveedor": ["endesa","iberdrola","gas","electr",
                       "energia","suministro"]},                  "629"),
    ({"kw_proveedor": ["otis","schindler","kone","ascensor",
                       "elevador","mantenimiento","reparacion"]}, "622"),
    ({"kw_proveedor": ["securitas","prosegur","seguridad",
                       "alarma","vigilancia"]},                   "623"),
    ({"kw_proveedor": ["seguro","mapfre","axa","allianz",
                       "zurich","prima"]},                        "625"),
    ({"kw_proveedor": ["arrendamiento","alquiler","renting"]},    "621"),
    # Fallback
    ({"tipo_proveedor": "OTRAS"},                                 "629"),
]

def safe_float(v):
    try:
        if v is None or str(v).strip() in ("", NF, "nan", "None"):
            return None
        s = str(v).replace("EUR","").replace("€","").replace(" ","").strip()
        if "," in s and "." in s:
            s = s.replace(",","") if s.rfind(".") > s.rfind(",") else s.replace(".","").replace(",",".")
        elif "," in s:
            s = s.replace(",",".")
        return float(s)
    except Exception:
        return None

def cargar_proveedores():
    if not os.path.exists(PROV_FILE):
        return {}
    df = pd.read_excel(PROV_FILE)
    return {r["nombre_proveedor"].strip().lower(): r.to_dict() for _, r in df.iterrows()}

def cargar_plan_cuentas():
    if not os.path.exists(CC_FILE):
        return {}
    df = pd.read_excel(CC_FILE)
    return {str(r["codigo_cuenta"]).strip(): r.to_dict() for _, r in df.iterrows()}

def determinar_cuenta_gasto(fila, proveedores):
    tipo      = str(fila.get("tipo_proveedor","")).strip().upper()
    nombre    = str(fila.get("nombre_proveedor","")).strip().lower()
    ya_tiene  = str(fila.get("cuenta_contable","")).strip()

    # Si ya viene asignada desde proveedores.xlsx y no es NF
    if ya_tiene and ya_tiene not in (NF, "nan", "None", ""):
        return ya_tiene, "ASIGNADA_PROVEEDOR"

    for regla, cuenta in REGLAS_CUENTA_GASTO:
        if "tipo_proveedor" in regla and tipo == regla["tipo_proveedor"]:
            return cuenta, "REGLA_TIPO"
        if "kw_proveedor" in regla:
            if any(kw in nombre for kw in regla["kw_proveedor"]):
                return cuenta, "REGLA_KEYWORD"

    return "REVISAR_MANUAL", "SIN_REGLA"

def cuenta_iva(pct_iva):
    try:
        pct = float(str(pct_iva).replace("%","").strip())
        if abs(pct - 21) < 0.5:  return "472",  "H.P. IVA soportado 21%"
        if abs(pct - 10) < 0.5:  return "4720", "H.P. IVA soportado 10%"
        if abs(pct -  4) < 0.5:  return "4721", "H.P. IVA soportado 4%"
    except Exception:
        pass
    return "472", "H.P. IVA soportado"

def generar_asiento(fila, cuenta_gasto, plan_cc):
    num_fac  = fila.get("numero_factura", NF)
    prov     = fila.get("nombre_proveedor", NF)
    base     = safe_float(fila.get("base_imponible"))
    cuota    = safe_float(fila.get("cuota_iva"))
    total    = safe_float(fila.get("total_factura"))
    pct_iva  = fila.get("porcentaje_iva", NF)
    fecha    = fila.get("fecha", NF)

    # Descripción de la cuenta de gasto
    desc_gasto = plan_cc.get(cuenta_gasto, {}).get("descripcion", cuenta_gasto)

    # Cuenta IVA
    c_iva, desc_iva = cuenta_iva(pct_iva)

    # Asiento textual
    lineas = [
        f"FECHA: {fecha}  |  CONCEPTO: Fact. {num_fac} – {prov}",
        f"  DEBE  {cuenta_gasto} {desc_gasto:<40} {base or '?':>12} EUR",
        f"  DEBE  {c_iva:<6} {desc_iva:<40} {cuota or '?':>12} EUR",
        f"  HABER 400    Proveedores ({prov}){'':<20} {total or '?':>12} EUR",
    ]

    return {
        "cuenta_debe_gasto":  cuenta_gasto,
        "cuenta_debe_iva":    c_iva,
        "cuenta_haber":       "400",
        "asiento_contable":   " | ".join(lineas),
    }

def cargar_todas_facturas_ap():
    """Carga facturas_ap + une los resultados de matching si existen."""
    excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_ap_*.xlsx")), reverse=True)
    if not excels:
        raise FileNotFoundError("No hay facturas_ap_*.xlsx. Ejecuta lector_facturas_ap.py")
    df = pd.read_excel(excels[0])
    print(f"  Facturas AP: {os.path.basename(excels[0])}")

    # Intentar unir estado_matching de ambos reportes
    for patron in [f"matching_otras_{FECHA_HOY}.xlsx",
                   f"matching_fb_{FECHA_HOY}.xlsx"]:
        ruta = os.path.join(REPORTES_DIR, patron)
        if os.path.exists(ruta):
            try:
                dm = pd.read_excel(ruta, sheet_name=0)
                if "archivo" in dm.columns and "estado_matching" in dm.columns:
                    df = df.merge(dm[["archivo","estado_matching","alerta_detalle"]]
                                  .rename(columns={"alerta_detalle":"detalle_matching"}),
                                  on="archivo", how="left")
                    print(f"  Matching unido: {patron}")
            except Exception:
                pass
    return df

def aplicar_formato(ws):
    try:
        col_asig = next((i+1 for i, c in enumerate(ws[1]) if c.value == "estado_asignacion"), None)
        if col_asig is None: return
        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            v = ws.cell(ri, col_asig).value
            fill = VERDE if v == "ASIGNADA_PROVEEDOR" else (AMARILLO if v == "REGLA_KEYWORD" else (GRIS if v == "SIN_REGLA" else None))
            if fill:
                for cell in row: cell.fill = fill
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
    except Exception:
        pass

def main():
    print("="*60)
    print("  Yve.01 — Asignador de Cuentas Contables AP")
    print("="*60)
    try:
        df = cargar_todas_facturas_ap()
    except FileNotFoundError as e:
        print(f"\n❌ {e}"); return

    proveedores = cargar_proveedores()
    plan_cc     = cargar_plan_cuentas()
    print(f"  Plan de cuentas: {len(plan_cc)} cuentas\n")

    resultados = []
    manuales   = 0
    for _, fila in df.iterrows():
        cuenta_gasto, metodo = determinar_cuenta_gasto(fila, proveedores)
        asiento_dict = generar_asiento(fila, cuenta_gasto, plan_cc)

        if cuenta_gasto == "REVISAR_MANUAL":
            manuales += 1
            icono = "⚠"
        else:
            icono = "✓"
        print(f"  [{icono}] {fila.get('archivo',NF)} → {cuenta_gasto} [{metodo}]")

        resultados.append({
            **fila.to_dict(),
            **asiento_dict,
            "estado_asignacion": metodo,
        })

    df_res = pd.DataFrame(resultados)

    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        df_res.to_excel(w, index=False, sheet_name="Facturas_Contabilizadas")
        aplicar_formato(w.sheets["Facturas_Contabilizadas"])
        ws = w.sheets["Facturas_Contabilizadas"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col)+4, 50)

    print(f"\n  Asignadas automáticamente: {len(resultados)-manuales}")
    print(f"  Requieren revisión manual:  {manuales}")
    print(f"\n✅ Archivo: {SALIDA}")
    print("="*60)

if __name__ == "__main__":
    main()
