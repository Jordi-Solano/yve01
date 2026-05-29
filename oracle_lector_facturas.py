"""
oracle_lector_facturas.py — Yve.01 Módulo Oracle
Lee facturas_contabilizadas_[fecha].xlsx generadas por asignador_cuentas.py
y las convierte al formato de Oracle GL Journal Entry.
Regla crítica: solo procesa facturas con estado APROBADA en el libro de aprobaciones.
"""

import os
import glob
from pathlib import Path
from datetime import datetime

import pandas as pd

from oracle_auth import is_simulation, ORACLE_LEDGER_NAME

# ─── Rutas ────────────────────────────────────────────────────────────────────
BASE_DIR         = Path(__file__).parent
PROCESADAS_DIR   = BASE_DIR / "facturas-procesadas"
APROBACIONES_DIR = BASE_DIR / "aprobaciones"
REFERENCIA_DIR   = BASE_DIR / "datos-referencia"
NF               = "NO_ENCONTRADO"

# Entity por defecto (de CtaCble del DRR — código de hotel Hilton Barcelona)
DEFAULT_ENTITY   = os.getenv("ORACLE_ENTITY", "1662")
DEFAULT_DEPT_ADM = os.getenv("ORACLE_DEPT_ADM", "ADM")


def _sf(v, default=0.0):
    """Safe float."""
    try:
        x = str(v).replace("EUR","").replace("€","").strip()
        if x in ("", NF, "nan", "None", "NaN"):
            return default
        return float(x)
    except Exception:
        return default


def _fecha_oracle(fecha_str):
    """Convierte DD/MM/YYYY → YYYY-MM-DD para Oracle."""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
        try:
            dt = datetime.strptime(str(fecha_str).strip(), fmt)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    return datetime.now().strftime("%Y-%m-%d")


def cargar_aprobaciones_ap():
    """Carga aprobaciones_ap.xlsx y devuelve set de facturas aprobadas."""
    ruta = APROBACIONES_DIR / "aprobaciones_ap.xlsx"
    if not ruta.exists():
        return set(), pd.DataFrame()
    try:
        df = pd.read_excel(ruta)
        if "numero_factura" not in df.columns or "accion" not in df.columns:
            return set(), df
        # Última acción por factura
        ultimas = (df.sort_values("fecha_hora", na_position="first")
                     .groupby("numero_factura").last().reset_index())
        aprobadas = set(
            ultimas.loc[
                ultimas["accion"].str.upper() == "APROBADA", "numero_factura"
            ]
        )
        return aprobadas, ultimas
    except Exception as e:
        print(f"  ⚠  Error leyendo aprobaciones_ap.xlsx: {e}")
        return set(), pd.DataFrame()


def cargar_facturas_contabilizadas():
    """Carga el último facturas_contabilizadas_*.xlsx."""
    archivos = sorted(
        glob.glob(str(PROCESADAS_DIR / "facturas_contabilizadas_*.xlsx")),
        reverse=True,
    )
    if not archivos:
        raise FileNotFoundError(
            "No se encontró facturas_contabilizadas_*.xlsx. "
            "Ejecuta primero asignador_cuentas.py"
        )
    df = pd.read_excel(archivos[0])
    print(f"  Facturas contabilizadas: {Path(archivos[0]).name} ({len(df)} filas)")
    return df, archivos[0]


def construir_journal_lines(row, entity=DEFAULT_ENTITY) -> list:
    """
    Construye las líneas de Journal Entry de Oracle para una factura.
    Formato: 3 líneas por factura:
      1. DEBE  — cuenta de gasto
      2. DEBE  — cuenta de IVA soportado
      3. HABER — cuenta de proveedores (400)
    """
    num_fac   = str(row.get("numero_factura", "")).strip()
    proveedor = str(row.get("nombre_proveedor", "")).strip()
    fecha     = _fecha_oracle(row.get("fecha", ""))
    base      = _sf(row.get("base_imponible", 0))
    iva       = _sf(row.get("cuota_iva", 0))
    total     = _sf(row.get("total_factura", 0))
    cta_gasto = str(row.get("cuenta_debe_gasto", row.get("cuenta_contable", "629"))).strip()
    cta_iva   = str(row.get("cuenta_debe_iva", "472")).strip()
    cta_prov  = str(row.get("cuenta_haber", "400")).strip()
    tipo      = str(row.get("tipo_proveedor", "OTRAS")).strip().upper()

    concepto  = f"Fact. {num_fac} — {proveedor}"

    # Resolver departamento según tipo de proveedor
    dept_gasto = "FB" if tipo == "FB" else DEFAULT_DEPT_ADM

    return [
        {
            "line_number":    1,
            "type":           "DEBE",
            "entity":         entity,
            "department":     dept_gasto,
            "account":        cta_gasto,
            "combination":    f"{entity}.{dept_gasto}.{cta_gasto}",
            "debit":          base,
            "credit":         0.0,
            "description":    f"{concepto} — Base imponible",
            "reference":      num_fac,
            "currency":       "EUR",
            "accounting_date": fecha,
        },
        {
            "line_number":    2,
            "type":           "DEBE",
            "entity":         entity,
            "department":     DEFAULT_DEPT_ADM,
            "account":        cta_iva,
            "combination":    f"{entity}.{DEFAULT_DEPT_ADM}.{cta_iva}",
            "debit":          iva,
            "credit":         0.0,
            "description":    f"{concepto} — IVA soportado",
            "reference":      num_fac,
            "currency":       "EUR",
            "accounting_date": fecha,
        },
        {
            "line_number":    3,
            "type":           "HABER",
            "entity":         entity,
            "department":     DEFAULT_DEPT_ADM,
            "account":        cta_prov,
            "combination":    f"{entity}.{DEFAULT_DEPT_ADM}.{cta_prov}",
            "debit":          0.0,
            "credit":         total,
            "description":    f"Proveedores — {proveedor}",
            "reference":      num_fac,
            "currency":       "EUR",
            "accounting_date": fecha,
        },
    ]


def construir_journal_batch(row) -> dict:
    """Construye el payload completo de un Journal Batch para Oracle."""
    num_fac   = str(row.get("numero_factura", "")).strip()
    proveedor = str(row.get("nombre_proveedor", "")).strip()
    fecha     = _fecha_oracle(row.get("fecha", ""))

    lines = construir_journal_lines(row)

    return {
        "numero_factura":      num_fac,
        "nombre_proveedor":    proveedor,
        "fecha":               fecha,
        "total_factura":       _sf(row.get("total_factura", 0)),
        "ledger_name":         ORACLE_LEDGER_NAME,
        "batch_name":          f"YVE01-AP-{num_fac}",
        "journal_source":      "YVE01",
        "journal_category":    "Purchase Invoices",
        "description":         f"Fact. {num_fac} — {proveedor}",
        "accounting_date":     fecha,
        "currency":            "EUR",
        "journal_lines":       lines,
        # Raw Oracle payload
        "oracle_payload": {
            "JournalBatchName":    f"YVE01-AP-{num_fac}",
            "LedgerName":          ORACLE_LEDGER_NAME,
            "AccountingDate":      fecha,
            "JournalSource":       "YVE01",
            "JournalCategory":     "Purchase Invoices",
            "Description":         f"Fact. {num_fac} — {proveedor}",
            "CurrencyCode":        "EUR",
            "JournalLines": [
                {
                    "LineNumber":           l["line_number"],
                    "AccountCombination":   l["combination"],
                    "EnteredDebitAmount":   l["debit"],
                    "EnteredCreditAmount":  l["credit"],
                    "Description":         l["description"],
                }
                for l in lines
            ],
        },
    }


def preparar_facturas_para_oracle(bypass_aprobacion: bool = False) -> tuple:
    """
    Carga facturas aprobadas y prepara batches para Oracle.
    Returns:
        (batches: list[dict], bloqueadas: list[dict], df_facturas: DataFrame)
    bypass_aprobacion: si True (modo simulación), procesa todas aunque no haya aprobación
    """
    df, archivo = cargar_facturas_contabilizadas()
    aprobadas_set, df_apro = cargar_aprobaciones_ap()

    print(f"  Facturas aprobadas en aprobaciones_ap.xlsx: {len(aprobadas_set)}")
    if is_simulation():
        print("  ℹ  MODO SIMULACIÓN — se procesarán todas las facturas (ignora aprobación)")

    batches   = []
    bloqueadas = []

    # Evitar contabilizar dos veces
    ya_contabilizadas = set()
    if "oracle_status" in df.columns:
        ya_contabilizadas = set(
            df.loc[df["oracle_status"].astype(str).str.upper() == "CONTABILIZADA",
                   "numero_factura"].astype(str)
        )

    for _, row in df.iterrows():
        num_fac = str(row.get("numero_factura", "")).strip()
        estado_asig = str(row.get("estado_asignacion", "")).strip().upper()

        # Saltar ya contabilizadas
        if num_fac in ya_contabilizadas:
            print(f"  ⏭  {num_fac} ya está CONTABILIZADA en Oracle — saltada")
            continue

        # Saltar si la cuenta es REVISAR_MANUAL (sin asignación válida)
        cuenta = str(row.get("cuenta_contable", "")).strip()
        if cuenta.upper() == "REVISAR_MANUAL":
            bloqueadas.append({
                "numero_factura": num_fac,
                "motivo": "Cuenta contable requiere revisión manual (REVISAR_MANUAL)",
            })
            continue

        # Verificar aprobación
        aprobada = (num_fac in aprobadas_set)
        if not aprobada and not bypass_aprobacion and not is_simulation():
            bloqueadas.append({
                "numero_factura": num_fac,
                "motivo": "Factura no aprobada por cabeza de departamento",
            })
            continue

        batch = construir_journal_batch(row)
        batches.append(batch)

    return batches, bloqueadas, df


if __name__ == "__main__":
    from oracle_auth import print_status
    print("=" * 65)
    print("  Yve.01 — Oracle Lector Facturas")
    print("=" * 65)
    print_status()
    print()

    batches, bloqueadas, _ = preparar_facturas_para_oracle()

    print(f"\n  Batches listos para Oracle:  {len(batches)}")
    print(f"  Facturas bloqueadas:         {len(bloqueadas)}")

    for b in batches:
        total = b["total_factura"]
        print(f"\n  ┌─ {b['numero_factura']} — {b['nombre_proveedor']}")
        print(f"  │  Batch: {b['batch_name']}")
        print(f"  │  Fecha: {b['fecha']}  |  Total: {total:,.2f} EUR")
        print(f"  │  Líneas:")
        for l in b["journal_lines"]:
            t = "D" if l["type"] == "DEBE" else "H"
            importe = l["debit"] if l["type"] == "DEBE" else l["credit"]
            print(f"  │    [{t}] {l['combination']:<25} {importe:>12,.2f} EUR  {l['description'][:35]}")

    if bloqueadas:
        print(f"\n  Facturas bloqueadas:")
        for b in bloqueadas:
            print(f"    ✗ {b['numero_factura']}: {b['motivo']}")

    print("\n" + "=" * 65)
