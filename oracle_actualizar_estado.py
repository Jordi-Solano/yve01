"""
oracle_actualizar_estado.py — Yve.01 Módulo Oracle
Actualiza facturas_contabilizadas_*.xlsx con el Oracle Journal ID
y cambia el estado a CONTABILIZADA (o CONTABILIZADA_SIM en simulación).
"""

import glob
from pathlib import Path
from datetime import datetime

import pandas as pd

from oracle_auth import is_simulation

BASE_DIR       = Path(__file__).parent
PROCESADAS_DIR = BASE_DIR / "facturas-procesadas"
HOY            = datetime.now().strftime("%Y%m%d")


def cargar_facturas() -> tuple:
    """Carga el último facturas_contabilizadas_*.xlsx."""
    archivos = sorted(
        glob.glob(str(PROCESADAS_DIR / "facturas_contabilizadas_*.xlsx")),
        reverse=True,
    )
    if not archivos:
        raise FileNotFoundError("No se encontró facturas_contabilizadas_*.xlsx")
    df = pd.read_excel(archivos[0])
    return df, archivos[0]


def actualizar_estados(resultados: list) -> tuple:
    """
    Actualiza el Excel de facturas con Oracle IDs y estados.
    resultados: lista de dicts con {numero_factura, oracle_id, estado, timestamp}
    Returns: (df_actualizado, ruta_guardada, stats)
    """
    df, ruta_original = cargar_facturas()

    # Asegurar que existen las columnas Oracle
    for col in ("oracle_status", "oracle_id", "fecha_contabilizacion"):
        if col not in df.columns:
            df[col] = ""

    stats = {"actualizadas": 0, "errores": 0, "ya_contabilizadas": 0}

    for res in resultados:
        num_fac = str(res.get("numero_factura", "")).strip()
        mask    = df["numero_factura"].astype(str).str.strip() == num_fac

        if not mask.any():
            print(f"  ⚠  {num_fac} no encontrada en el Excel de facturas")
            continue

        estado_actual = str(df.loc[mask, "oracle_status"].iloc[0]).upper()
        if estado_actual == "CONTABILIZADA":
            print(f"  ⏭  {num_fac} ya estaba CONTABILIZADA — no se modifica")
            stats["ya_contabilizadas"] += 1
            continue

        if res.get("estado") in ("CONTABILIZADA", "CONTABILIZADA_SIM"):
            df.loc[mask, "oracle_status"]         = res["estado"]
            df.loc[mask, "oracle_id"]             = res.get("oracle_id", "")
            df.loc[mask, "fecha_contabilizacion"] = res.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M"))
            print(f"  ✓  {num_fac} → {res['estado']} | ID: {res.get('oracle_id','')}")
            stats["actualizadas"] += 1
        else:
            df.loc[mask, "oracle_status"] = "ERROR_ORACLE"
            df.loc[mask, "oracle_id"]     = f"ERR: {res.get('error','')[:50]}"
            print(f"  ✗  {num_fac} → ERROR: {res.get('error','')[:60]}")
            stats["errores"] += 1

    # Guardar en mismo archivo (sobrescribir)
    df.to_excel(ruta_original, index=False)
    print(f"  💾 Excel actualizado: {Path(ruta_original).name}")

    return df, ruta_original, stats


def mostrar_resumen_oracle(df: pd.DataFrame):
    """Muestra en consola el resumen del estado Oracle de todas las facturas."""
    if "oracle_status" not in df.columns:
        print("  (sin columna oracle_status)")
        return

    conteo = df["oracle_status"].fillna("PENDIENTE").value_counts()
    print("\n  Estado Oracle de facturas:")
    estado_icons = {
        "CONTABILIZADA":     "✅",
        "CONTABILIZADA_SIM": "🟡",
        "PENDIENTE":         "⏳",
        "ERROR_ORACLE":      "❌",
        "":                  "⏳",
    }
    for estado, count in conteo.items():
        icon = estado_icons.get(str(estado).upper(), "•")
        label = estado if estado else "PENDIENTE"
        print(f"    {icon} {label:<22} {count:>3}")

    if "oracle_id" in df.columns and "total_factura" in df.columns:
        mask_cont = df["oracle_status"].isin(["CONTABILIZADA", "CONTABILIZADA_SIM"])
        total_cont = df.loc[mask_cont, "total_factura"].apply(
            lambda v: float(str(v).replace("EUR","").strip() or 0) if v is not None else 0
        ).sum()
        print(f"\n  Importe total contabilizado: {total_cont:,.2f} EUR")


if __name__ == "__main__":
    from oracle_auth import print_status

    print("=" * 65)
    print("  Yve.01 — Oracle Actualizar Estado")
    print("=" * 65)
    print_status()
    print()

    df, ruta = cargar_facturas()
    print(f"  Cargado: {Path(ruta).name} ({len(df)} facturas)")
    mostrar_resumen_oracle(df)
    print("=" * 65)
