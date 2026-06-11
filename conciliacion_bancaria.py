"""
conciliacion_bancaria.py — Yve.01
Conciliación automática: cruza extracto bancario con facturas AP/AR.
Uso: python conciliacion_bancaria.py [ruta_extracto.xlsx]
"""

import os, sys, glob
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd

BASE_DIR       = Path(__file__).parent
REPORTES_DIR   = BASE_DIR / "reportes"
PROCESADAS_DIR = BASE_DIR / "facturas-procesadas"
REFERENCIA_DIR = BASE_DIR / "datos-referencia"
HOY            = datetime.now().strftime("%Y%m%d")

REPORTES_DIR.mkdir(exist_ok=True)


def _sf(v):
    try:
        if v is None or str(v).strip() in ("", "nan", "None", "NO_ENCONTRADO"):
            return 0.0
        return float(str(v).replace(",", "").replace(" EUR", "").replace("EUR", "").strip())
    except Exception:
        return 0.0


def _ultimo_excel(patron, directorio):
    hits = glob.glob(str(directorio / patron))
    if not hits:
        return None
    hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return hits[0]


def cargar_extracto(ruta=None):
    """Carga el extracto bancario."""
    if ruta is None:
        ruta = str(REFERENCIA_DIR / "extracto_banco.xlsx")
    if not os.path.exists(ruta):
        print(f"  No se encontro extracto: {ruta}")
        return pd.DataFrame()
    df = pd.read_excel(ruta)
    df["importe_num"] = df["importe"].apply(_sf)
    return df


def cargar_facturas():
    """Carga todas las facturas AP y AR disponibles para matching."""
    facturas = []

    # AP — facturas proveedores
    for patron in ["facturas_contabilizadas_*.xlsx", "facturas_ap_*.xlsx"]:
        ruta = _ultimo_excel(patron, PROCESADAS_DIR)
        if ruta:
            try:
                df = pd.read_excel(ruta)
                for _, r in df.iterrows():
                    imp = _sf(r.get("total_factura", r.get("importe_total", 0)))
                    if imp > 0:
                        facturas.append({
                            "origen": "AP",
                            "numero": str(r.get("numero_factura", "")),
                            "proveedor": str(r.get("nombre_proveedor", "")),
                            "importe": imp,
                            "fecha": r.get("fecha"),
                            "tipo_mov": "CARGO",
                        })
            except Exception:
                pass
            break

    # AR — facturas OTA (cobros)
    for patron in ["doble_imposicion_*.xlsx", "verificacion_*.xlsx", "facturas_procesadas_*.xlsx"]:
        ruta = _ultimo_excel(patron, REPORTES_DIR)
        if ruta is None:
            ruta = _ultimo_excel(patron, PROCESADAS_DIR)
        if ruta:
            try:
                df = pd.read_excel(ruta)
                for _, r in df.iterrows():
                    imp = _sf(r.get("importe_bruto", 0))
                    if imp > 0:
                        facturas.append({
                            "origen": "AR",
                            "numero": str(r.get("numero_factura", "")),
                            "proveedor": str(r.get("nombre_ota", "")),
                            "importe": imp,
                            "fecha": r.get("fecha"),
                            "tipo_mov": "ABONO",
                        })
            except Exception:
                pass
            break

    return facturas


def conciliar(extracto, facturas, tolerancia_dias=3, tolerancia_importe=0.02):
    """
    Cruza cada movimiento del extracto con facturas.
    Match por: mismo tipo (CARGO/ABONO), importe exacto o cercano, fecha +-tolerancia_dias.
    Devuelve el extracto con columnas adicionales: estado, factura_ref, origen, diferencia.
    """
    usadas = set()
    resultados = []

    for _, mov in extracto.iterrows():
        imp_mov = mov["importe_num"]
        tipo_mov = mov["tipo"]
        fecha_mov = pd.to_datetime(mov["fecha"]) if pd.notna(mov["fecha"]) else None

        mejor_match = None
        mejor_diff = float("inf")

        # Fuzzy proveedor name matching helper
        def _nombre_match(concepto, proveedor):
            c = str(concepto).upper()
            p = str(proveedor).upper().split()[:2]
            return any(word in c for word in p if len(word) > 3)

        for i, fac in enumerate(facturas):
            if i in usadas:
                continue
            if fac["tipo_mov"] != tipo_mov:
                continue

            diff = abs(imp_mov - fac["importe"])
            pct_diff = diff / max(imp_mov, 0.01)

            # Match exacto o con tolerancia de 2%
            if pct_diff <= tolerancia_importe:
                # Check fecha si disponible
                if fecha_mov is not None and fac["fecha"] is not None:
                    try:
                        fecha_fac = pd.to_datetime(fac["fecha"])
                        dias = abs((fecha_mov - fecha_fac).days)
                        if dias > tolerancia_dias * 30:  # Más flexible para facturas
                            continue
                    except Exception:
                        pass

                if diff < mejor_diff:
                    mejor_diff = diff
                    mejor_match = i

        if mejor_match is not None:
            fac = facturas[mejor_match]
            usadas.add(mejor_match)
            diff_real = imp_mov - fac["importe"]
            if abs(diff_real) < 0.01:
                estado = "CONCILIADO"
            else:
                estado = "DIFERENCIA"
            resultados.append({
                "estado": estado,
                "factura_ref": fac["numero"],
                "origen": fac["origen"],
                "match_proveedor": fac["proveedor"],
                "diferencia": round(diff_real, 2),
            })
        else:
            resultados.append({
                "estado": "PENDIENTE",
                "factura_ref": "",
                "origen": "",
                "match_proveedor": "",
                "diferencia": 0.0,
            })

    df_result = extracto.copy()
    df_extra = pd.DataFrame(resultados)
    for col in df_extra.columns:
        df_result[col] = df_extra[col].values

    return df_result


def generar_reporte(df_result):
    """Guarda el reporte de conciliacion en reportes/."""
    nombre = f"conciliacion_{HOY}.xlsx"
    ruta = str(REPORTES_DIR / nombre)

    cols_out = ["fecha", "concepto", "importe", "tipo", "referencia", "saldo",
                "estado", "factura_ref", "origen", "match_proveedor", "diferencia"]
    cols_exist = [c for c in cols_out if c in df_result.columns]
    df_result[cols_exist].to_excel(ruta, index=False)
    return ruta


def main():
    print("=" * 65)
    print("  Yve.01 — Conciliacion Bancaria Automatica")
    print("=" * 65)

    ruta_extracto = sys.argv[1] if len(sys.argv) > 1 else None

    print("\n  [1/3] Cargando extracto bancario...")
    extracto = cargar_extracto(ruta_extracto)
    if extracto.empty:
        print("  No hay extracto. Coloca extracto_banco.xlsx en datos-referencia/")
        return
    print(f"        {len(extracto)} movimientos cargados")

    print("\n  [2/3] Cargando facturas AP/AR...")
    facturas = cargar_facturas()
    print(f"        {len(facturas)} facturas disponibles para matching")
    print(f"        AP: {sum(1 for f in facturas if f['origen']=='AP')}")
    print(f"        AR: {sum(1 for f in facturas if f['origen']=='AR')}")

    print("\n  [3/3] Ejecutando conciliacion automatica...")
    resultado = conciliar(extracto, facturas)

    conciliados = len(resultado[resultado["estado"] == "CONCILIADO"])
    pendientes = len(resultado[resultado["estado"] == "PENDIENTE"])
    diferencias = len(resultado[resultado["estado"] == "DIFERENCIA"])
    imp_pendiente = resultado.loc[resultado["estado"] == "PENDIENTE", "importe_num"].sum()

    ruta = generar_reporte(resultado)

    print(f"\n  {'='*50}")
    print(f"  RESUMEN CONCILIACION")
    print(f"  {'='*50}")
    print(f"  Total movimientos:    {len(resultado)}")
    print(f"  Conciliados:          {conciliados}")
    print(f"  Pendientes:           {pendientes}")
    print(f"  Con diferencia:       {diferencias}")
    print(f"  Importe pendiente:    {imp_pendiente:,.2f} EUR")
    print(f"\n  Reporte: {ruta}")
    print("=" * 65)

    return resultado


if __name__ == "__main__":
    main()
