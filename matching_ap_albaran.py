"""
matching_ap_albaran.py — Yve.01 Modulo AP
Cruza cada factura de proveedor con los albaranes de entrega que la respaldan.

Responde a la pregunta que mas dinero mueve: **¿te estan facturando algo que
nunca llego?** Y su reverso, que sale gratis del mismo cruce: **¿hay mercancia
entregada que nadie ha facturado?** (pasivo no contabilizado).

DOS NIVELES DE CONFIANZA, en este orden:

  1. REFERENCIA EXPLICITA — el albaran cita el numero de factura
     (`referencia_factura`), o el numero de albaran aparece escrito en el
     concepto de la factura. Es exacto, sin heuristica.
  2. PROVEEDOR + VENTANA DE 45 DIAS — el nombre del proveedor casa y la entrega
     cae dentro de la ventana anterior a la fecha de factura. Una factura
     mensual agrupa VARIOS albaranes, asi que el emparejamiento es 1↔N.

QUE SE COMPARA (y una trampa que costaria caro):
  Un albaran NO lleva IVA y una factura SI. Comparar `total_factura` contra
  `total_albaran` daria una falsa discrepancia del ~21% en TODOS los cruces
  (medido: 55,9% en el caso de prueba). Se compara la **BASE IMPONIBLE** de la
  factura contra la suma de los totales de sus albaranes.

LO QUE ESTE MODULO TODAVIA NO PUEDE HACER:
  Comparar CANTIDADES linea a linea. El albaran tiene lineas; la factura, hoy,
  no — su esquema solo guarda `descripcion_concepto` y los totales. Añadir
  lineas a la factura toca el esquema compartido por los tres caminos de
  entrada, y por eso va en su propia fase (3c). Hasta entonces, la tolerancia
  de cantidad no se puede aplicar: aqui solo se cruzan importes.

Oracle NO interviene: esto genera un informe, no contabiliza nada.
Ejecutar: python matching_ap_albaran.py
"""

import os
import glob
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Multi-tenant desde el primer dia (a diferencia de matching_ap_otras/fb, que
# nacieron con las rutas fijas). tenant_dirs lee YVE_TENANT del entorno cuando
# el dashboard lo lanza por subprocess.
try:
    from tenant_dirs import reportes_dir as _t_rep, procesadas_dir as _t_proc
    REPORTES_DIR = _t_rep()
    PROCESADAS_DIR = _t_proc()
except Exception:
    REPORTES_DIR = os.path.join(BASE_DIR, "reportes")
    PROCESADAS_DIR = os.path.join(BASE_DIR, "facturas-procesadas")
os.makedirs(REPORTES_DIR, exist_ok=True)

FECHA_HOY = date.today().strftime("%Y%m%d")
SALIDA = os.path.join(REPORTES_DIR, f"matching_albaran_{FECHA_HOY}.xlsx")

NF = "NO_ENCONTRADO"

# Ventana hacia atras desde la fecha de factura, en dias. 45 cubre la
# facturacion mensual con holgura (decidido con el usuario).
VENTANA_DIAS = 45

# Tolerancias. Van juntas y se imprimen en el informe para que se vea con que
# criterio se cruzo. La de CANTIDAD es deliberadamente mas laxa: en alimentacion
# el peso servido difiere del pedido por naturaleza, y una alerta que grita por
# 800 gramos de merluza consigue que el AP Manager deje de mirar las alertas.
# OJO: TOL_CANTIDAD no se aplica todavia — hace falta que la factura traiga
# lineas (fase 3c). Se declara aqui para que las dos vivan en el mismo sitio.
TOL_IMPORTE = 0.02    # 2 % sobre la base imponible
TOL_CANTIDAD = 0.10   # 10 % — se usara en la fase 3c

VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ROJO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
AMARILLO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
AZUL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")


# ── utilidades ────────────────────────────────────────────────────────────

def _txt(v):
    """Texto comparable. NaN, NO_ENCONTRADO y demas vacios -> ''.

    En Python plano, nunca con el accesor `.str`: en pandas 3 `astype(str)` deja
    los nulos como NaN y los propaga (ver las reglas del proyecto).
    """
    s = "" if v is None else str(v)
    s = " ".join(s.split()).strip()
    return "" if s.lower() in ("", "nan", "none", "nat", "<na>", "no_encontrado", "null") else s


def _num(v):
    """Float tolerante con '1.234,56', '450 EUR', '€' y los vacios."""
    s = _txt(v)
    if not s:
        return None
    s = s.replace("EUR", "").replace("€", "").replace(" ", "").strip()
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _fecha(v):
    """Fecha -> date. El mismo dato llega como Timestamp desde un fichero y
    como '18/07/2026' desde otro."""
    s = _txt(v)
    if not s:
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date() if hasattr(v, "date") else None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _clave_prov(nombre):
    """Nombre de proveedor comparable: minusculas, sin puntuacion ni forma
    societaria. 'Pescados Rias, S.L.' y 'PESCADOS RIAS SL' son el mismo."""
    s = _txt(nombre).lower()
    for ch in (".", ",", "-", "&", "'", '"'):
        s = s.replace(ch, " ")
    fuera = {"sl", "s", "l", "sa", "slu", "sau", "sccl", "cb", "sociedad",
             "limitada", "anonima", "y", "de", "del", "la", "el", "los", "las"}
    return " ".join(p for p in s.split() if p not in fuera)


def _mismo_proveedor(a, b):
    """Coincidencia parcial en los dos sentidos, como hace buscar_po."""
    ca, cb = _clave_prov(a), _clave_prov(b)
    if not ca or not cb:
        return False
    return ca == cb or (len(ca) > 3 and ca in cb) or (len(cb) > 3 and cb in ca)


def _base_factura(fila):
    """Lo que hay que comparar con el albaran: la base SIN IVA.

    Si no se extrajo la base, se deriva del total y el porcentaje; si tampoco
    hay porcentaje, se usa el total tal cual y se avisa en el detalle, porque
    entonces la comparacion puede llevar el IVA dentro.
    """
    base = _num(fila.get("base_imponible"))
    if base:
        return base, ""
    total = _num(fila.get("total_factura"))
    if not total:
        return None, "sin importe en la factura"
    pct = _num(fila.get("porcentaje_iva"))
    if pct:
        return round(total / (1 + pct / 100.0), 2), ""
    return total, "sin base imponible: se compara el total, que puede llevar IVA"


# ── carga ─────────────────────────────────────────────────────────────────

def cargar_facturas():
    """Todas las facturas AP del tenant, TODOS los dias.

    Por `almacen_datos` desde el principio: la mercancia llega ANTES que la
    factura, asi que el caso normal es cruzar una factura de hoy con un albaran
    de la semana pasada. Con el 'coge el mas reciente' de siempre, la mitad de
    los cruces no encontraria nada.
    """
    import almacen_datos as _alm
    df = _alm.facturas_ap(PROCESADAS_DIR, REPORTES_DIR)
    return df if df is not None else pd.DataFrame()


def cargar_albaranes():
    import almacen_datos as _alm
    df = _alm.albaranes(PROCESADAS_DIR, REPORTES_DIR)
    return df if df is not None else pd.DataFrame()


# ── el cruce ──────────────────────────────────────────────────────────────

def _cita_explicita(fila_f, alb):
    """¿Se citan el uno al otro? Devuelve el motivo, o '' si no."""
    num_f = _txt(fila_f.get("numero_factura"))
    num_a = _txt(alb.get("numero_albaran"))
    ref_f = _txt(alb.get("referencia_factura"))
    if num_f and ref_f and ref_f.lower() == num_f.lower():
        return f"el albarán cita la factura {num_f}"
    if num_a:
        # El esquema de FACTURA no tiene campo para citar un albaran, pero una
        # factura mensual suele listarlos en el concepto ("albaranes 7781,
        # 7782..."). Buscarlo ahi sale gratis y no toca el esquema compartido.
        texto = " ".join(_txt(fila_f.get(c)) for c in
                         ("descripcion_concepto", "numero_factura")).lower()
        if num_a.lower() in texto:
            return f"la factura menciona el albarán {num_a}"
    return ""


def emparejar(df_fact, df_alb):
    """Empareja facturas y albaranes. Devuelve {indice_factura: [indices_alb]}.

    Un albaran solo puede consumirse UNA vez: si dos facturas se lo repartieran,
    la misma entrega estaria justificando dos cobros, que es justo lo que este
    modulo existe para detectar.

    Primero se resuelven TODAS las referencias explicitas y solo despues se
    reparte por proveedor y fecha, para que una coincidencia debil no se lleve
    un albaran que otra factura reclama por su numero.
    """
    asignados = {}          # indice de albaran -> indice de factura
    porque = {}             # (indice factura, indice albaran) -> motivo
    empare = {i: [] for i in df_fact.index}

    # ── nivel 1 · referencia explicita ────────────────────────────────────
    for i_f, fila_f in df_fact.iterrows():
        for i_a, alb in df_alb.iterrows():
            if i_a in asignados:
                continue
            motivo = _cita_explicita(fila_f, alb)
            if motivo:
                asignados[i_a] = i_f
                empare[i_f].append(i_a)
                porque[(i_f, i_a)] = motivo

    # ── nivel 2 · proveedor + ventana ─────────────────────────────────────
    for i_f, fila_f in df_fact.iterrows():
        f_fact = _fecha(fila_f.get("fecha"))
        for i_a, alb in df_alb.iterrows():
            if i_a in asignados:
                continue
            if not _mismo_proveedor(fila_f.get("nombre_proveedor"),
                                    alb.get("nombre_proveedor")):
                continue
            f_ent = _fecha(alb.get("fecha_entrega"))
            if f_fact and f_ent:
                # la entrega tiene que ser ANTERIOR a la factura (o del mismo
                # dia) y caer dentro de la ventana
                if not (f_fact - timedelta(days=VENTANA_DIAS) <= f_ent <= f_fact):
                    continue
                motivo = f"mismo proveedor, entrega del {f_ent.strftime('%d/%m/%Y')}"
            elif f_fact or f_ent:
                # falta una de las dos fechas: se acepta por proveedor, pero se
                # deja dicho, porque la ventana no se ha podido comprobar
                motivo = "mismo proveedor (sin fecha para comprobar la ventana)"
            else:
                motivo = "mismo proveedor (sin fechas)"
            asignados[i_a] = i_f
            empare[i_f].append(i_a)
            porque[(i_f, i_a)] = motivo
    return empare, porque


def analizar_factura(fila_f, indices_alb, df_alb, porque, i_f):
    base, aviso = _base_factura(fila_f)
    albs = [df_alb.loc[i] for i in indices_alb]
    nums = [_txt(a.get("numero_albaran")) or "s/n" for a in albs]
    suma = sum(_num(a.get("total_albaran")) or 0.0 for a in albs)

    if not albs:
        estado = "FACTURA_SIN_ALBARAN"
        detalle = ("no se ha encontrado ninguna entrega que respalde esta factura "
                   f"(mismo proveedor, {VENTANA_DIAS} días antes)")
        diff = dif_pct = NF
    elif base is None:
        estado = "SIN_IMPORTE"
        detalle = f"{len(albs)} albarán(es) encontrado(s), pero {aviso}"
        diff = dif_pct = NF
    elif suma <= 0:
        estado = "SIN_IMPORTE"
        detalle = f"albarán(es) {', '.join(nums)} sin importe extraíble"
        diff = dif_pct = NF
    else:
        diff = round(base - suma, 2)
        dif_pct = abs(diff) / suma
        motivos = "; ".join(porque[(i_f, i)] for i in indices_alb)
        if dif_pct <= TOL_IMPORTE:
            estado = "MATCH_ALBARAN_OK"
            detalle = (f"{len(albs)} albarán(es) {', '.join(nums)} cuadran "
                       f"({dif_pct*100:.2f}% de diferencia) · {motivos}")
        else:
            estado = "DIFERENCIA_IMPORTE"
            signo = "MÁS" if diff > 0 else "menos"
            detalle = (f"la factura cobra {abs(diff):.2f} EUR {signo} de lo entregado: "
                       f"base {base:.2f} vs albaranes {suma:.2f} "
                       f"({', '.join(nums)}) · {motivos}")
        if aviso:
            detalle += f" · OJO: {aviso}"

    return {
        "archivo":            _txt(fila_f.get("archivo")) or NF,
        "numero_factura":     _txt(fila_f.get("numero_factura")) or NF,
        "fecha":              _txt(fila_f.get("fecha")) or NF,
        "nombre_proveedor":   _txt(fila_f.get("nombre_proveedor")) or NF,
        "base_imponible":     base if base is not None else NF,
        "total_factura":      _num(fila_f.get("total_factura")) or NF,
        "n_albaranes":        len(albs),
        "albaranes":          ", ".join(nums) if nums else NF,
        "total_albaranes":    round(suma, 2) if albs else NF,
        "diferencia_importe": diff,
        "diferencia_pct":     f"{dif_pct*100:.2f}%" if isinstance(dif_pct, float) else NF,
        "estado_matching":    estado,
        "detalle_matching":   detalle,
    }


def analizar_albaranes(df_alb, asignados_por_alb, df_fact):
    """El reverso: mercancia entregada que nadie ha facturado."""
    filas = []
    for i_a, alb in df_alb.iterrows():
        i_f = asignados_por_alb.get(i_a)
        if i_f is None:
            estado = "ALBARAN_SIN_FACTURAR"
            detalle = "entregado, pero ninguna factura lo respalda todavía"
            num_f = NF
        else:
            estado = "ALBARAN_FACTURADO"
            num_f = _txt(df_fact.loc[i_f].get("numero_factura")) or NF
            detalle = f"facturado en {num_f}"
        filas.append({
            "clave":            _txt(alb.get("clave")),
            "numero_albaran":   _txt(alb.get("numero_albaran")) or NF,
            "nombre_proveedor": _txt(alb.get("nombre_proveedor")) or NF,
            "fecha_entrega":    _txt(alb.get("fecha_entrega")) or NF,
            "total_albaran":    _num(alb.get("total_albaran")) or NF,
            "n_lineas":         alb.get("n_lineas", NF),
            "numero_factura":   num_f,
            "estado":           estado,
            "detalle":          detalle,
        })
    return filas


# ── informe ───────────────────────────────────────────────────────────────

_COLORES = {
    "MATCH_ALBARAN_OK": VERDE, "ALBARAN_FACTURADO": VERDE,
    "DIFERENCIA_IMPORTE": ROJO,
    "FACTURA_SIN_ALBARAN": AMARILLO, "ALBARAN_SIN_FACTURAR": AMARILLO,
    "SIN_IMPORTE": AZUL,
}


def aplicar_formato(ws, col_estado):
    try:
        idx = next((i + 1 for i, c in enumerate(ws[1]) if c.value == col_estado), None)
        if idx is None:
            return
        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            fill = _COLORES.get(ws.cell(ri, idx).value)
            if fill:
                for cell in row:
                    cell.fill = fill
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
    except Exception:
        pass


def generar_resumen(df_f, df_a):
    filas = []
    if not df_f.empty:
        for est, n in df_f["estado_matching"].value_counts().items():
            filas.append({"Bloque": "Facturas", "Estado": est, "Cantidad": int(n),
                          "Pct": f"{n/len(df_f)*100:.1f}%"})
    if not df_a.empty:
        for est, n in df_a["estado"].value_counts().items():
            filas.append({"Bloque": "Albaranes", "Estado": est, "Cantidad": int(n),
                          "Pct": f"{n/len(df_a)*100:.1f}%"})
    # dejar por escrito CON QUE criterio se cruzo: un informe sin sus umbrales
    # no se puede auditar despues
    filas += [
        {"Bloque": "─" * 12, "Estado": "", "Cantidad": "", "Pct": ""},
        {"Bloque": "Criterio", "Estado": "ventana de emparejamiento",
         "Cantidad": f"{VENTANA_DIAS} días", "Pct": ""},
        {"Bloque": "Criterio", "Estado": "tolerancia de importe",
         "Cantidad": f"{TOL_IMPORTE*100:.0f}%", "Pct": ""},
        {"Bloque": "Criterio", "Estado": "tolerancia de cantidad (aún sin aplicar: la factura no trae líneas)",
         "Cantidad": f"{TOL_CANTIDAD*100:.0f}%", "Pct": ""},
        {"Bloque": "Criterio", "Estado": "se compara la BASE IMPONIBLE (el albarán no lleva IVA)",
         "Cantidad": "", "Pct": ""},
    ]
    return pd.DataFrame(filas)


def main():
    print("=" * 60)
    print("  Yve.01 — Matching AP · Factura ↔ Albarán")
    print("=" * 60)

    df_fact = cargar_facturas()
    df_alb = cargar_albaranes()
    if df_fact.empty:
        print("\n  No hay facturas AP que cruzar.")
        return 0
    if df_alb.empty:
        print("\n  No hay albaranes todavía: nada que cruzar.")
        print("  (Sube los albaranes por Procesar Archivos y vuelve a ejecutarlo.)")
        return 0

    print(f"  Facturas: {len(df_fact)}  |  Albaranes: {len(df_alb)}"
          f"  |  ventana {VENTANA_DIAS} días · tolerancia {TOL_IMPORTE*100:.0f}%\n")

    empare, porque = emparejar(df_fact, df_alb)
    por_alb = {i_a: i_f for i_f, idxs in empare.items() for i_a in idxs}

    res_f = [analizar_factura(df_fact.loc[i_f], idxs, df_alb, porque, i_f)
             for i_f, idxs in empare.items()]
    res_a = analizar_albaranes(df_alb, por_alb, df_fact)

    iconos = {"MATCH_ALBARAN_OK": "✓", "DIFERENCIA_IMPORTE": "✗",
              "FACTURA_SIN_ALBARAN": "?", "SIN_IMPORTE": "~"}
    for r in res_f:
        print(f"  [{iconos.get(r['estado_matching'], '·')}] {r['numero_factura']} → {r['estado_matching']}")
        if r["estado_matching"] != "MATCH_ALBARAN_OK":
            print(f"       {r['detalle_matching']}")
    sin_facturar = [r for r in res_a if r["estado"] == "ALBARAN_SIN_FACTURAR"]
    for r in sin_facturar:
        print(f"  [!] {r['numero_albaran']} ({r['nombre_proveedor']}) → ALBARAN_SIN_FACTURAR")

    df_res_f = pd.DataFrame(res_f)
    df_res_a = pd.DataFrame(res_a)
    df_sum = generar_resumen(df_res_f, df_res_a)

    with pd.ExcelWriter(SALIDA, engine="openpyxl") as w:
        df_res_f.to_excel(w, index=False, sheet_name="Facturas")
        # OJO: esta hoja se llama 'Albaranes' a proposito. almacen_datos lee esa
        # hoja para consolidar los albaranes de todos los dias, asi que el
        # estado del cruce viaja con ellos y no se pierde al cambiar de dia.
        df_res_a.to_excel(w, index=False, sheet_name="Albaranes")
        df_sum.to_excel(w, index=False, sheet_name="Resumen")
        aplicar_formato(w.sheets["Facturas"], "estado_matching")
        aplicar_formato(w.sheets["Albaranes"], "estado")
        for sn in ("Facturas", "Albaranes", "Resumen"):
            ws = w.sheets[sn]
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(len(str(c.value or "")) for c in col) + 4, 60)

    print("\n" + "─" * 60)
    print("  RESUMEN")
    print("─" * 60)
    for _, r in df_sum.iterrows():
        print(f"  {str(r['Bloque']):<12} {str(r['Estado']):<58} {str(r['Cantidad'])}")
    print(f"\n✅ Reporte: {SALIDA}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
