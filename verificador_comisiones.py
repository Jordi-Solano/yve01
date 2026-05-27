"""
verificador_comisiones.py — Yve.01
Cruza Excel de facturas procesadas con comisiones pactadas.
"""
import os, re, glob
from datetime import date
import pandas as pd

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
PROCESADAS_DIR= os.path.join(BASE_DIR, "facturas-procesadas")
REFERENCIA_DIR= os.path.join(BASE_DIR, "datos-referencia")
REPORTES_DIR  = os.path.join(BASE_DIR, "reportes")
os.makedirs(REPORTES_DIR, exist_ok=True)
COMISIONES_FILE = os.path.join(REFERENCIA_DIR, "comisiones_pactadas.xlsx")
FECHA_HOY = date.today().strftime("%Y%m%d")
REPORTE_SALIDA = os.path.join(REPORTES_DIR, f"verificacion_{FECHA_HOY}.xlsx")
TOLERANCIA = 0.5
NF = "NO_ENCONTRADO"

try:
    from openpyxl.styles import PatternFill, Font, Alignment
    VERDE    = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ROJO     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    AMARILLO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    USAR_FORMATO = True
except ImportError:
    USAR_FORMATO = False

def cargar_ultimo_excel_procesadas():
    excels = sorted(glob.glob(os.path.join(PROCESADAS_DIR, "facturas_procesadas_*.xlsx")), reverse=True)
    if not excels:
        raise FileNotFoundError(f"No hay facturas_procesadas_*.xlsx en {PROCESADAS_DIR}")
    ruta = excels[0]
    print(f"  Cargando facturas desde: {os.path.basename(ruta)}")
    return pd.read_excel(ruta), ruta

def cargar_comisiones_pactadas():
    if not os.path.exists(COMISIONES_FILE):
        raise FileNotFoundError(f"No se encontro: {COMISIONES_FILE}")
    df = pd.read_excel(COMISIONES_FILE)
    df["OTA_norm"] = df["OTA"].str.strip().str.lower()
    return df

def convertir_porcentaje(valor):
    if valor is None or str(valor).strip() in ("", NF):
        return None
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return None

def convertir_importe(valor):
    if valor is None or str(valor).strip() in ("", NF):
        return None
    try:
        limpio = str(valor).replace("EUR","").replace("USD","").replace(" ","").strip()
        limpio = limpio.replace("\xa0","")
        # Quitar simbolo euro/dollar si quedara
        limpio = limpio.replace("\u20ac","").strip()
        if "," in limpio and "." in limpio:
            if limpio.rfind(".") > limpio.rfind(","):
                limpio = limpio.replace(",","")         # US: 8,300.00 -> 8300.00
            else:
                limpio = limpio.replace(".","").replace(",",".")  # EU: 8.300,00
        elif "," in limpio:
            if re.search(r",\d{3}$", limpio):
                limpio = limpio.replace(",","")
            else:
                limpio = limpio.replace(",",".")
        return float(limpio)
    except ValueError:
        return None

def verificar_factura(fila, comisiones_df):
    ota_nombre = str(fila.get("nombre_ota", NF))
    ota_norm   = ota_nombre.strip().lower() if ota_nombre not in (NF,"") else ""
    match = comisiones_df[comisiones_df["OTA_norm"] == ota_norm]
    mercado = NF
    comision_pactada = None
    diferencia = None
    importe_discrepancia = None

    if ota_norm == "" or ota_nombre == NF or match.empty:
        estado = "OTA_DESCONOCIDA"
    else:
        comision_pactada = float(match.iloc[0]["Porcentaje_Comision"])
        mercado = match.iloc[0]["Mercado"]
        comision_factura = convertir_porcentaje(fila.get("porcentaje_comision"))
        importe_bruto    = convertir_importe(fila.get("importe_bruto"))
        if comision_factura is None:
            estado = "SIN_PORCENTAJE"
        else:
            diferencia = abs(comision_factura - comision_pactada)
            if diferencia < TOLERANCIA:
                estado = "CORRECTO"
            else:
                estado = "DISCREPANCIA"
                if importe_bruto is not None:
                    importe_discrepancia = round(
                        importe_bruto * (comision_factura - comision_pactada) / 100, 2
                    )
    return {
        "archivo":          fila.get("archivo", NF),
        "numero_factura":   fila.get("numero_factura", NF),
        "fecha":            fila.get("fecha", NF),
        "nombre_ota":       ota_nombre,
        "mercado":          mercado,
        "nombre_hotel":     fila.get("nombre_hotel", NF),
        "periodo_inicio":   fila.get("periodo_inicio", NF),
        "periodo_fin":      fila.get("periodo_fin", NF),
        "importe_bruto":    fila.get("importe_bruto", NF),
        "porcentaje_pactado":       comision_pactada,
        "porcentaje_factura":       fila.get("porcentaje_comision", NF),
        "diferencia_pp":    round(diferencia,2) if diferencia is not None else NF,
        "importe_comision_factura": fila.get("importe_comision", NF),
        "importe_neto":     fila.get("importe_neto", NF),
        "discrepancia_euros": importe_discrepancia if importe_discrepancia is not None else NF,
        "estado":           estado,
    }

def aplicar_formato_excel(ws, df):
    if not USAR_FORMATO:
        return
    col_estado = None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value == "estado":
            col_estado = idx
            break
    if col_estado is None:
        return
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
        estado = ws.cell(row=row_idx, column=col_estado).value
        fill = {"CORRECTO": VERDE, "DISCREPANCIA": ROJO, "OTA_DESCONOCIDA": AMARILLO, "SIN_PORCENTAJE": AMARILLO}.get(estado)
        if fill:
            for cell in row:
                cell.fill = fill
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

def generar_resumen(df_resultado):
    total = len(df_resultado)
    resumen = df_resultado["estado"].value_counts().reset_index()
    resumen.columns = ["Estado", "Cantidad"]
    resumen["Porcentaje"] = (resumen["Cantidad"]/total*100).round(1).astype(str)+"%"
    discrepancias = df_resultado[df_resultado["estado"]=="DISCREPANCIA"]
    montos = [v for v in (convertir_importe(r.get("discrepancia_euros")) for _,r in discrepancias.iterrows()) if v is not None]
    total_disc = sum(montos)
    extra = pd.DataFrame([
        {"Estado":"─"*20,"Cantidad":"","Porcentaje":""},
        {"Estado":"TOTAL FACTURAS","Cantidad":total,"Porcentaje":"100%"},
        {"Estado":"DISCREPANCIA TOTAL (EUR)","Cantidad":f"{total_disc:.2f} EUR","Porcentaje":""},
    ])
    return pd.concat([resumen, extra], ignore_index=True)

def main():
    print("="*60)
    print("  Yve.01 — Verificador de Comisiones OTA")
    print("="*60)
    try:
        df_facturas, _ = cargar_ultimo_excel_procesadas()
        df_comisiones  = cargar_comisiones_pactadas()
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        return
    print(f"  Tabla de referencia: {len(df_comisiones)} OTAs cargadas")
    print(f"  Facturas a verificar: {len(df_facturas)}\n")
    resultados = []
    for _, fila in df_facturas.iterrows():
        resultado = verificar_factura(fila, df_comisiones)
        icono = {"CORRECTO":"v","DISCREPANCIA":"X","OTA_DESCONOCIDA":"?","SIN_PORCENTAJE":"~"}.get(resultado["estado"],".")
        disc_str = f"  ({resultado['discrepancia_euros']} EUR)" if resultado["estado"]=="DISCREPANCIA" and resultado["discrepancia_euros"]!=NF else ""
        print(f"  [{icono}] {resultado['archivo']} -> {resultado['estado']}{disc_str}")
        resultados.append(resultado)
    df_resultado = pd.DataFrame(resultados)
    df_resumen   = generar_resumen(df_resultado)
    with pd.ExcelWriter(REPORTE_SALIDA, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Detalle")
        df_resumen.to_excel(writer, index=False, sheet_name="Resumen")
        aplicar_formato_excel(writer.sheets["Detalle"], df_resultado)
        for sheet in ["Detalle","Resumen"]:
            ws = writer.sheets[sheet]
            for col in ws.columns:
                ancho = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(ancho+4, 40)
    print("\n"+"─"*60)
    print("  RESUMEN VERIFICACION")
    print("─"*60)
    for _, row in df_resumen.iterrows():
        print(f"  {str(row['Estado']):<30} {str(row['Cantidad']):<12} {row['Porcentaje']}")
    print(f"\nReporte guardado en: {REPORTE_SALIDA}")
    print("="*60)

if __name__ == "__main__":
    main()
