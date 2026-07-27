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
from tenant_dirs import reportes_dir as _t_rdir, procesadas_dir as _t_pdir, datos_dir as _t_ddir

class _TDir:
    """Path por-tenant evaluado en cada uso (no al importar)."""
    def __init__(self, fn): self._fn = fn
    def __truediv__(self, other): return Path(self._fn()) / other
    def __str__(self): return self._fn()
    def mkdir(self, **kw): Path(self._fn()).mkdir(**kw)

REPORTES_DIR   = _TDir(_t_rdir)
PROCESADAS_DIR = _TDir(_t_pdir)
REFERENCIA_DIR = _TDir(_t_ddir)
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


_EXT_COL_MAP = {
    "fecha":      ["fecha", "date", "fecha_operacion", "f. operacion"],
    "concepto":   ["concepto", "descripcion", "description", "detalle"],
    "importe":    ["importe", "cantidad", "amount", "monto"],
    "saldo":      ["saldo", "balance"],
    "referencia": ["referencia", "ref", "reference"],
    "tipo":       ["tipo", "type", "movimiento"],
}


def _normalizar_extracto(df):
    """Renombra columnas alternativas y deriva tipo/importe_num."""
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for canon, alts in _EXT_COL_MAP.items():
        if canon in df.columns:
            continue
        for alt in alts:
            if alt in cols_lower:
                df = df.rename(columns={cols_lower[alt]: canon})
                break
    if "importe" not in df.columns:
        return pd.DataFrame()
    for c in ("concepto", "referencia"):
        if c not in df.columns:
            df[c] = ""
    df["importe_num"] = df["importe"].apply(_sf)
    # tipo: usar columna si trae CARGO/ABONO; si no, derivar del signo
    if "tipo" in df.columns:
        tipos = df["tipo"].astype(str).str.upper().str.strip()
        validos = tipos.isin(["CARGO", "ABONO"])
        df["tipo"] = tipos.where(validos, None)
    else:
        df["tipo"] = None
    df["tipo"] = df.apply(
        lambda r: r["tipo"] if r["tipo"] in ("CARGO", "ABONO")
        else ("CARGO" if r["importe_num"] < 0 else "ABONO"), axis=1)
    return df


def cargar_extracto(ruta=None):
    """Carga el extracto bancario y lo normaliza."""
    if ruta is None:
        ruta = str(REFERENCIA_DIR / "extracto_banco.xlsx")
    if not os.path.exists(ruta):
        print(f"  No se encontro extracto: {ruta}")
        return pd.DataFrame()
    df = pd.read_excel(ruta)
    return _normalizar_extracto(df)


def cargar_facturas():
    """Carga todas las facturas AP y AR disponibles para matching."""
    facturas = []

    # Antes cada bloque hacia su propio "coge el fichero mas reciente" + break,
    # asi que la conciliacion cruzaba los movimientos del banco contra UN SOLO
    # dia de facturas: un pago a una factura de ayer quedaba PENDIENTE para
    # siempre. La lectura y el deduplicado viven en almacen_datos, que es el
    # unico sitio a tocar cuando migremos a persistencia.
    from almacen_datos import facturas_ap as _fap, facturas_ar as _far

    # AP — facturas proveedores (TODOS los dias)
    try:
        for _, r in _fap(PROCESADAS_DIR, REPORTES_DIR).iterrows():
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
    except Exception as e:
        print(f"  [conciliacion] no se pudieron cargar facturas AP: {e}")

    # AR — facturas OTA / cobros (TODOS los dias)
    try:
        for _, r in _far(PROCESADAS_DIR, REPORTES_DIR).iterrows():
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
    except Exception as e:
        print(f"  [conciliacion] no se pudieron cargar facturas AR: {e}")


    return facturas


def conciliar(extracto, facturas, tolerancia_dias=3, tolerancia_importe=0.02):
    """
    Cruza cada movimiento del extracto con facturas.
    Señales de match (de más a menos fuerte):
      1. nº de factura presente en concepto/referencia + importe cuadra
      2. nº de factura presente en concepto/referencia (importe difiere → DIFERENCIA)
      3. nombre de proveedor en concepto + importe cuadra
      4. solo importe cuadra (con fecha dentro de tolerancia si hay fecha)
    Los importes se comparan en valor absoluto (los CARGOs vienen en negativo).
    Devuelve el extracto con columnas: estado, factura_ref, origen, match_proveedor, diferencia.
    """
    if extracto is None or len(extracto) == 0:
        return extracto

    def _nombre_match(concepto, proveedor):
        c = str(concepto).upper()
        palabras = [w for w in str(proveedor).upper().split()[:3] if len(w) > 3]
        return any(w in c for w in palabras)

    def _ref_match(mov_texto, numero):
        n = str(numero).strip()
        if len(n) < 4 or n.lower() in ("nan", "none", ""):
            return False
        return n.upper() in mov_texto

    usadas = set()
    resultados = []

    for _, mov in extracto.iterrows():
        imp_mov   = abs(_sf(mov.get("importe_num", mov.get("importe", 0))))
        tipo_mov  = mov.get("tipo")
        texto_mov = (str(mov.get("concepto", "")) + " " + str(mov.get("referencia", ""))).upper()
        fecha_mov = pd.to_datetime(mov.get("fecha"), errors="coerce")

        mejor = None  # (score, -diff, idx)

        for i, fac in enumerate(facturas):
            if i in usadas:
                continue
            if fac.get("tipo_mov") and tipo_mov and fac["tipo_mov"] != tipo_mov:
                continue

            imp_fac = abs(_sf(fac.get("importe", 0)))
            if imp_fac <= 0:
                continue
            diff = abs(imp_mov - imp_fac)
            importe_ok = (diff / max(imp_mov, 0.01)) <= tolerancia_importe

            ref_ok    = _ref_match(texto_mov, fac.get("numero", ""))
            nombre_ok = _nombre_match(texto_mov, fac.get("proveedor", ""))

            if ref_ok and importe_ok:
                score = 4
            elif ref_ok:
                score = 3
            elif nombre_ok and importe_ok:
                score = 2
            elif importe_ok:
                # sin señal de texto: exigir fecha dentro de tolerancia amplia
                if pd.notna(fecha_mov) and fac.get("fecha") is not None:
                    fecha_fac = pd.to_datetime(fac.get("fecha"), errors="coerce")
                    if pd.notna(fecha_fac) and abs((fecha_mov - fecha_fac).days) > tolerancia_dias * 30:
                        continue
                score = 1
            else:
                continue

            cand = (score, -diff, i)
            if mejor is None or cand > mejor:
                mejor = cand

        if mejor is not None:
            _, _, idx = mejor
            fac = facturas[idx]
            usadas.add(idx)
            diff_real = round(imp_mov - abs(_sf(fac.get("importe", 0))), 2)
            estado = "CONCILIADO" if abs(diff_real) < 0.01 else "DIFERENCIA"
            resultados.append({
                "estado": estado,
                "factura_ref": fac.get("numero", ""),
                "origen": fac.get("origen", "AP"),
                "match_proveedor": fac.get("proveedor", ""),
                "diferencia": diff_real,
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
