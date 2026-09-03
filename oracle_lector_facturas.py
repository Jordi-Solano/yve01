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
REPORTES_DIR     = BASE_DIR / "reportes"
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
    """Las facturas contabilizadas de TODOS los días, no solo el último fichero.

    BUG 11 — ANTES abría SOLO el `facturas_contabilizadas_*.xlsx` más reciente.
    Si ese fichero salía incompleto (porque `almacen_datos` no pudo abrir uno de
    un día anterior y lo saltó callando), una factura ya CONTABILIZADA reaparecía
    aquí SIN su `oracle_status` y se volvía a contabilizar en el libro mayor.
    Reproducido antes de arreglarlo: 2 de 3 facturas duplicadas.

    Se lee por `almacen_datos.facturas_ap()` —el mismo punto único que usa el
    asignador y, desde ahora, el panel de Aprobaciones—, con las rutas de la
    RAÍZ explícitas para que panel y Oracle miren exactamente lo mismo.

    Se sigue devolviendo el fichero más reciente como `ruta`: es donde
    `oracle_actualizar_estado` escribe el marcador, y eso no cambia.
    """
    import almacen_datos as _alm
    df = _alm.facturas_ap(str(PROCESADAS_DIR), str(REPORTES_DIR))
    _ileg = []
    try:
        _ileg = _alm.ficheros_ilegibles()
    except Exception:
        pass
    if df is None or df.empty:
        # Cero facturas NO es un error del pipeline: es "nada que contabilizar".
        # Antes se lanzaba FileNotFoundError y el boton acababa en
        # "pipeline con errores" con el disco recien vaciado por un deploy.
        print("  Facturas contabilizadas: 0 (no hay facturas procesadas todavia)")
        df = pd.DataFrame()
    archivos = sorted(
        glob.glob(str(PROCESADAS_DIR / "facturas_contabilizadas_*.xlsx")),
        reverse=True,
    )
    ruta = archivos[0] if archivos else str(PROCESADAS_DIR / "facturas_contabilizadas.xlsx")
    if not df.empty:
        print(f"  Facturas contabilizadas: {len(df)} filas de todos los días")
    if _ileg:
        # El aviso que faltaba: un fichero ilegible era un print perdido en el
        # stdout de un subproceso. Aquí importa, porque puede llevarse por
        # delante un `oracle_status`.
        print(f"  ⚠  {len(_ileg)} fichero(s) NO se pudieron leer: {', '.join(_ileg[:5])}")
        print("     Puede faltar algún marcador de contabilizada — el registro "
              "aparte (oracle_contabilizadas.json) es la red de seguridad.")
    return df, ruta


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
    bypass_aprobacion: si True, procesa todas aunque no haya aprobación. SOLO
    para pruebas explícitas: la simulación ya NO lo activa sola (decisión de
    Jordi, 3 sep 2026): la puerta APROBADA se respeta igual que en producción,
    si no la doble firma y el panel de aprobación no se ven hasta tener Oracle real.
    """
    df, archivo = cargar_facturas_contabilizadas()
    aprobadas_set, df_apro = cargar_aprobaciones_ap()

    print(f"  Facturas aprobadas en aprobaciones_ap.xlsx: {len(aprobadas_set)}")
    if is_simulation():
        print("  ℹ  MODO SIMULACIÓN — la puerta APROBADA se respeta igual que en producción")
    if bypass_aprobacion:
        print("  ⚠  bypass_aprobacion=True: se procesan facturas SIN aprobar (solo pruebas)")

    batches   = []
    bloqueadas = []

    # Evitar contabilizar dos veces
    ya_contabilizadas = set()
    if not df.empty and "oracle_status" in df.columns:
        ya_contabilizadas = set(
            df.loc[df["oracle_status"].astype(str).str.upper() == "CONTABILIZADA",
                   "numero_factura"].astype(str)
        )
    # Y el registro aparte, que no depende de ningún Excel (bug 11). Solo
    # SUMA: nunca quita a nadie de la lista, así que solo puede evitar un
    # asiento, jamás provocarlo.
    try:
        from oracle_actualizar_estado import (ya_contabilizadas_registro,
                                              apuntar_en_registro)
        # Primero se SIEMBRA con lo que dice el Excel de hoy. Sin esto, el
        # registro solo tendría lo contabilizado DESPUÉS de este cambio, y una
        # factura marcada de antes se quedaría sin red: justo el caso que se
        # quiere cubrir. Al ser de solo-añadir, sembrar no puede hacer daño.
        if ya_contabilizadas:
            apuntar_en_registro([{"numero_factura": n, "estado": "CONTABILIZADA"}
                                 for n in sorted(ya_contabilizadas)])
        _del_registro = ya_contabilizadas_registro()
        _extra = _del_registro - ya_contabilizadas
        if _extra:
            print(f"  🔒 {len(_extra)} factura(s) ya contabilizadas según el "
                  f"registro, aunque el Excel no lo diga: {', '.join(sorted(_extra)[:5])}")
        ya_contabilizadas |= _del_registro
    except Exception as _er:
        print(f"  ⚠  no se pudo leer el registro de contabilizadas: {str(_er)[:70]}")

    for _, row in (df.iterrows() if not df.empty else []):
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
        if not aprobada and not bypass_aprobacion:
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
