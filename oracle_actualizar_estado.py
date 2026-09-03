"""
oracle_actualizar_estado.py — Yve.01 Módulo Oracle
Actualiza facturas_contabilizadas_*.xlsx con el Oracle Journal ID
y cambia el estado a CONTABILIZADA (o CONTABILIZADA_SIM en simulación).
"""

import glob
import os
from pathlib import Path
from datetime import datetime

import pandas as pd

from oracle_auth import is_simulation

BASE_DIR       = Path(__file__).parent
PROCESADAS_DIR = BASE_DIR / "facturas-procesadas"
HOY            = datetime.now().strftime("%Y%m%d")

# ── El registro de lo ya contabilizado (bug 11) ───────────────────────────────
# El `oracle_status` vive dentro del xlsx. Si ese xlsx no se puede abrir, el
# marcador se pierde y la factura vuelve a parecer nueva: reproducido, Oracle
# montaba otra vez el asiento de dos facturas ya contabilizadas.
#
# Este registro es la copia que NO depende de ningun Excel. Es de SOLO ANADIR:
# nunca quita a nadie de la lista, asi que solo puede EVITAR un asiento, jamas
# provocarlo. Se escribe ANTES de tocar el xlsx, para que un fallo al guardar
# el Excel no deje a Oracle sin memoria de lo que acaba de contabilizar.
#
# Solo entra "CONTABILIZADA". "CONTABILIZADA_SIM" NO, a proposito: es la
# simulacion, y apuntarla dejaria la factura sin contabilizar de verdad para
# siempre. Es la misma regla exacta que ya usa `oracle_lector_facturas`.
REGISTRO_FILE = BASE_DIR / "datos-referencia" / "oracle_contabilizadas.json"


def _leer_registro() -> dict:
    """El registro entero. Un registro ilegible se trata como vacio: este
    fichero solo puede anadir proteccion, nunca quitarla ni romper el flujo."""
    try:
        import json
        with open(REGISTRO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def ya_contabilizadas_registro() -> set:
    """Numeros de factura que Oracle ya contabilizo DE VERDAD."""
    return {str(k) for k, v in _leer_registro().items()
            if str((v or {}).get("estado", "")).upper() == "CONTABILIZADA"}


def apuntar_en_registro(resultados: list) -> int:
    """Apunta en el registro las facturas contabilizadas. Devuelve cuantas."""
    import json
    reg = _leer_registro()
    nuevas = 0
    for res in resultados or []:
        if str(res.get("estado", "")).upper() != "CONTABILIZADA":
            continue
        num = str(res.get("numero_factura", "")).strip()
        if not num or num in reg:
            continue
        reg[num] = {"estado": "CONTABILIZADA",
                    "oracle_id": str(res.get("oracle_id", "")),
                    "fecha": str(res.get("timestamp",
                                         datetime.now().strftime("%Y-%m-%d %H:%M")))}
        nuevas += 1
    if not nuevas:
        return 0
    try:
        os.makedirs(os.path.dirname(str(REGISTRO_FILE)), exist_ok=True)
        tmp = str(REGISTRO_FILE) + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(reg, f, ensure_ascii=False, indent=2, sort_keys=True)
        os.replace(tmp, str(REGISTRO_FILE))
    except Exception as e:
        print(f"  \u26a0  no se pudo escribir el registro de contabilizadas: {e}")
        return 0
    return nuevas


ASIENTOS_FILE = BASE_DIR / "reportes" / "oracle_asientos_producidos.json"


def guardar_asientos_producidos(resultados: list) -> int:
    """Guarda en crudo los asientos que el pipeline acaba de producir.

    Un fichero unico, de solo anadir, con las lineas de cada asiento tal cual
    salieron (simulacion o produccion, con su estado). Es lo que exporta
    `oracle_export_dryrun`: si aqui no hay nada, no se exporta nada.
    """
    import json
    if not resultados:
        return 0
    try:
        with open(ASIENTOS_FILE, "r", encoding="utf-8") as f:
            reg = json.load(f)
        if not isinstance(reg, list):
            reg = []
    except Exception:
        reg = []
    nuevos = 0
    for r in resultados:
        if not r.get("journal_lines"):
            continue
        reg.append({
            "numero_factura":   str(r.get("numero_factura", "")),
            "nombre_proveedor": str(r.get("nombre_proveedor", "")),
            "fecha":            str(r.get("fecha", "")),
            "total_factura":    float(r.get("total_factura") or 0),
            "batch_name":       str(r.get("batch_name", "")),
            "oracle_id":        str(r.get("oracle_id", "")),
            "estado":           str(r.get("estado", "")),
            "modo":             str(r.get("modo", "")),
            "timestamp":        str(r.get("timestamp", "")),
            "journal_lines":    [{k: l.get(k) for k in ("line_number", "type", "entity", "department",
                                                       "account", "combination", "debit", "credit",
                                                       "description", "reference", "accounting_date")}
                                 for l in r.get("journal_lines") or []],
        })
        nuevos += 1
    if not nuevos:
        return 0
    os.makedirs(os.path.dirname(str(ASIENTOS_FILE)), exist_ok=True)
    tmp = str(ASIENTOS_FILE) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=1)
    os.replace(tmp, str(ASIENTOS_FILE))
    return nuevos


def asientos_producidos() -> list:
    """Todo lo que el pipeline ha producido hasta hoy (lista, mas antiguo primero)."""
    import json
    try:
        with open(ASIENTOS_FILE, "r", encoding="utf-8") as f:
            reg = json.load(f)
        return reg if isinstance(reg, list) else []
    except Exception:
        return []


def _ficheros_facturas() -> list:
    """Todos los Excel donde puede vivir una factura, el mas reciente primero.

    Antes solo se miraba `facturas_contabilizadas_*.xlsx` (lo que deja el
    asignador). Las facturas que entran por el lote del panel viven en
    `facturas_ap_*.xlsx` y el paso 4/4 fallaba con "No se encontro ..." aunque
    el paso 2/4 las hubiera leido perfectamente (almacen_datos lee los dos).
    """
    return sorted(
        glob.glob(str(PROCESADAS_DIR / "facturas_contabilizadas_*.xlsx"))
        + glob.glob(str(PROCESADAS_DIR / "facturas_ap_*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )


def cargar_facturas() -> tuple:
    """Carga el Excel de facturas mas reciente. Sin ninguno: (df vacio, '')."""
    archivos = _ficheros_facturas()
    if not archivos:
        return pd.DataFrame(), ""
    df = pd.read_excel(archivos[0])
    return df, archivos[0]


def actualizar_estados(resultados: list) -> tuple:
    """
    Actualiza el Excel de facturas con Oracle IDs y estados.
    resultados: lista de dicts con {numero_factura, oracle_id, estado, timestamp}
    Returns: (df_actualizado, ruta_guardada, stats)
    """
    # Lo PRIMERO, antes de tocar el Excel: si el guardado del xlsx fallara,
    # Oracle seguiria sabiendo que estas facturas ya estan contabilizadas.
    _n_reg = apuntar_en_registro(resultados)
    if _n_reg:
        print(f"  \U0001f4d2 {_n_reg} apuntada(s) en el registro de contabilizadas")

    stats = {"actualizadas": 0, "errores": 0, "ya_contabilizadas": 0, "sin_excel": 0}
    archivos = _ficheros_facturas()
    if not archivos:
        # Sin ningun Excel de facturas (disco recien vaciado, o cero facturas)
        # no hay nada que marcar: el registro de arriba ya guarda lo que se
        # contabilizo de verdad. No es un error del pipeline.
        print("  ℹ  No hay ningún Excel de facturas que marcar (el registro aparte queda actualizado)")
        stats["sin_excel"] = len(resultados or [])
        return pd.DataFrame(), "", stats

    # Cada factura se marca EN EL EXCEL DONDE VIVE (puede haber varios).
    dfs = {}
    for ruta in archivos:
        try:
            dfs[ruta] = pd.read_excel(ruta)
        except Exception as e:
            print(f"  ⚠  no se pudo abrir {Path(ruta).name}: {str(e)[:60]}")
    tocados = set()

    for res in resultados:
        num_fac = str(res.get("numero_factura", "")).strip()
        df = ruta_original = None
        for ruta, d in dfs.items():
            if "numero_factura" in d.columns and (d["numero_factura"].astype(str).str.strip() == num_fac).any():
                df, ruta_original = d, ruta
                break
        if df is None:
            print(f"  ⚠  {num_fac} no encontrada en ningún Excel de facturas")
            stats["sin_excel"] += 1
            continue
        for col in ("oracle_status", "oracle_id", "fecha_contabilizacion"):
            if col not in df.columns:
                df[col] = ""
        mask = df["numero_factura"].astype(str).str.strip() == num_fac
        tocados.add(ruta_original)

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

    # Guardar cada Excel tocado en su mismo sitio (sobrescribir)
    for ruta in sorted(tocados):
        dfs[ruta].to_excel(ruta, index=False)
        print(f"  💾 Excel actualizado: {Path(ruta).name}")

    df_todo = pd.concat(list(dfs.values()), ignore_index=True) if dfs else pd.DataFrame()
    return df_todo, (archivos[0] if archivos else ""), stats


def mostrar_resumen_oracle(df: pd.DataFrame):
    """Muestra en consola el resumen del estado Oracle de todas las facturas."""
    if df is None or df.empty or "oracle_status" not in df.columns:
        print("  (sin facturas con estado Oracle)")
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
