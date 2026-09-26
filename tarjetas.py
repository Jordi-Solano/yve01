# -*- coding: utf-8 -*-
"""tarjetas.py — Yve.01 · liquidacion de tarjetas y su cruce (b109, finanzas 26 sep 2026).

Finanzas: "no hay fichero de ejemplo ni sabemos la pasarela. Haz la pestaña Tarjetas con un
formato de liquidacion generico (CSV/Excel con fecha, bruto, comision, neto, referencia),
plantilla descargable, y el cruce con los abonos del extracto y con lo cobrado con tarjeta del
DRR/TPV. Cuando llegue el primer cliente se le añade el lector de su pasarela."
**Redsys, Adyen, Stripe...: PENDIENTES de un fichero real** (cada pasarela exporta a su manera);
hasta entonces entra el formato generico (y su plantilla).

Formato generico (una fila por operacion, o por remesa diaria):
  fecha        dia de la venta (o de la liquidacion)
  bruto        lo cobrado al cliente
  comision     lo que se queda la pasarela o el banco
  neto         lo que llega al banco (= bruto - comision; si falta uno de los tres, se calcula)
  referencia   numero de operacion o de remesa
  opcionales:  fecha_abono (el dia que llega al banco), tarjeta (Visa, Mastercard...), terminal

Datos: datos-referencia/tarjetas_liquidaciones.xlsx, una fila por operacion con su hotel y el
fichero del que salio. Entra por ⚡ Procesar archivos: el lote lo reconoce por las CABECERAS
(dashboard._CAB_TIPOS['TARJETAS']), no por el nombre.

Cruces (funciones puras):
  1. liquidacion <-> extracto. Lo liquidado se agrupa por dia de abono (fecha_abono, o fecha si
     no viene) y se busca el abono de la pestaña TARJETAS del cuadre de banco: por la referencia
     en el concepto; por el dia que el concepto nombra ("REDSYS LIQ TARJETAS 02/08"); y por
     importe neto (±0,01) entre ese dia y DIAS_MARGEN despues. Lo que queda: dias liquidados sin
     abono y abonos de tarjetas sin liquidacion.
  2. liquidacion <-> ventas con tarjeta. El bruto de cada dia contra lo cobrado con tarjeta segun
     el DRR (cuentas de tarjetas del Trial Balance: "Credit Cards Clearing", Visa...) y el TPV
     (solo si el fichero de ventas trae la forma de pago).
El ASIENTO de la liquidacion (comision...) NO se hace: pendiente de finanzas.
"""
import os
import re
import unicodedata
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd

import candados as _cand

FICHERO = "tarjetas_liquidaciones.xlsx"
COLS = ["fecha", "bruto", "comision", "neto", "referencia", "fecha_abono", "tarjeta", "terminal",
        "hotel_id", "fichero", "cargado"]
OBLIGATORIAS = ("fecha", "bruto", "comision", "neto", "referencia")
DIAS_MARGEN = 4          # un abono de tarjetas llega entre el dia de la venta y 4 dias despues
TOL = 0.01

# cabecera normalizada (minusculas, sin acentos, espacios a '_') -> campo. Nombre COMPLETO,
# nunca subcadena (la cicatriz de _normalize_cols: "id" dentro de "unidades").
SINONIMOS = {
    "fecha": ["fecha", "fecha_operacion", "fecha_venta", "fecha_transaccion", "date", "transaction_date",
              "dia", "fecha_liquidacion"],
    "bruto": ["bruto", "importe_bruto", "gross", "gross_amount", "importe_venta", "total_bruto"],
    "comision": ["comision", "comisiones", "comision_total", "fee", "fees", "tasa", "tasa_descuento", "descuento"],
    "neto": ["neto", "importe_neto", "net", "net_amount", "liquido", "abonado", "importe_abonado"],
    "referencia": ["referencia", "ref", "reference", "remesa", "n_remesa", "num_remesa", "id_liquidacion",
                   "n_operacion", "num_operacion", "numero_operacion", "operacion", "autorizacion", "id"],
    "fecha_abono": ["fecha_abono", "fecha_pago", "fecha_valor", "payout_date", "fecha_liquidacion_banco"],
    "tarjeta": ["tarjeta", "marca", "tipo_tarjeta", "card", "brand", "card_brand"],
    "terminal": ["terminal", "tpv", "datafono", "comercio", "merchant", "fuc"],
}
# cuentas del Trial Balance del DRR que son cobros con tarjeta
CUENTAS_DRR = ["credit card", "credit cards", "card clearing", "cards clearing", "tarjeta", "visa", "mastercard",
               "master card", "amex", "american express", "diners", "jcb", "maestro", "debit card"]
# columnas de ventas del TPV con lo cobrado con tarjeta, o con la forma de pago
TPV_IMPORTE = ["importe_tarjeta", "tarjeta", "cobro_tarjeta", "pago_tarjeta", "total_tarjeta", "card"]
TPV_FORMA = ["forma_pago", "medio_pago", "metodo_pago", "tipo_pago", "payment_method", "pago"]
PALABRAS_TARJETA = ["tarjeta", "card", "visa", "master", "amex", "credito", "debito", "tpv", "datafono"]


# ── utilidades ──────────────────────────────────────────────────────────────
def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _norm(v):
    s = unicodedata.normalize("NFKD", _txt(v))
    s = "".join(c for c in s if not unicodedata.combining(c)).strip().lower()
    for ch in (" ", ".", "-", "/", "\\", "º", "ª"):
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    return s.strip("_")


def _num(v):
    """Importe de una celda: 12.5, '1.234,56', '1,234.56', '12,30 €', '(3,20)'. None si no es numero."""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if (isinstance(v, float) and v != v) else round(float(v), 2)
    s = _txt(v).replace("€", "").replace("EUR", "").replace("eur", "").replace("\xa0", "").replace(" ", "")
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")            # 1.234 = mil doscientos treinta y cuatro
    try:
        f = float(s)
    except ValueError:
        return None
    return round(-f if neg else f, 2)


def _fecha(v):
    """'YYYY-MM-DD' o ''. Acepta fechas de Excel, ISO y dd/mm/aaaa."""
    if v is None:
        return ""
    if isinstance(v, (datetime, pd.Timestamp)):
        return "" if pd.isna(v) else v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    s = _txt(v)
    if not s:
        return ""
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        t = pd.to_datetime(s[:10], errors="coerce")
    else:
        t = pd.to_datetime(s[:10], dayfirst=True, errors="coerce")
    return "" if pd.isna(t) else t.strftime("%Y-%m-%d")


def _ref(v):
    """Referencia como texto: la de Excel llega como 123456.0."""
    if isinstance(v, float) and v == v and v == int(v):
        return str(int(v))
    return _txt(v)


def _eur(v):
    return f"{(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _mes_a_rango(mes):
    m = _txt(mes)[:7]
    if not re.fullmatch(r"\d{4}-\d{2}", m):
        m = date.today().strftime("%Y-%m")
    ini = date(int(m[:4]), int(m[5:7]), 1)
    fin = (pd.Timestamp(ini) + pd.offsets.MonthEnd(0)).date()
    return ini, fin, m


def _ruta(datos_dir=None):
    if datos_dir is None:
        from tenant_dirs import datos_dir as _d
        datos_dir = str(_d())
    return os.path.join(str(datos_dir), FICHERO)


def _dd_candado(dd):
    if dd:
        return str(dd)
    from tenant_dirs import datos_dir as _d
    return str(_d())


# ── la liquidacion ──────────────────────────────────────────────────────────
def columnas_de(cabeceras):
    """{campo: cabecera original} con los sinonimos de SINONIMOS (nombre completo)."""
    usados, out = set(), {}
    normal = [(c, _norm(c)) for c in cabeceras]
    for campo, alias in SINONIMOS.items():
        for c, n in normal:
            if c not in usados and n in alias:
                out[campo] = c; usados.add(c)
                break
    return out


# cabeceras que dicen que NO es una liquidacion de tarjetas sino de una OTA o de reservas
# (una liquidacion de Booking puede traer bruto, comision y neto)
NO_ES = {"huesped", "guest", "guest_name", "cliente", "reserva", "n_reserva", "numero_reserva", "booking",
         "booking_number", "checkin", "check_in", "checkout", "check_out", "llegada", "salida", "entrada",
         "noches", "nights", "ota", "habitacion", "room", "porcentaje", "porcentaje_comision", "comision_pct",
         "pct_comision", "saldo", "concepto"}


def es_liquidacion(cabeceras):
    """True si las cabeceras son las de una liquidacion de tarjetas: fecha, referencia y al
    menos dos de bruto / comision / neto (el tercero se calcula), y ninguna de reservas/OTA."""
    if any(_norm(c) in NO_ES for c in cabeceras):
        return False
    col = columnas_de(cabeceras)
    return "fecha" in col and "referencia" in col and sum(k in col for k in ("bruto", "comision", "neto")) >= 2


def _leer_tabla(fichero, nombre=""):
    """DataFrame de un CSV o Excel (ruta o fichero subido). Del Excel, la hoja de datos: la
    que se llame liquidacion, o la primera que no sea de instrucciones."""
    n = (nombre or getattr(fichero, "filename", "") or (fichero if isinstance(fichero, str) else "")).lower()
    if n.endswith(".csv"):
        if hasattr(fichero, "seek"):
            fichero.seek(0)
        for enc in ("utf-8-sig", "latin-1"):
            try:
                if hasattr(fichero, "seek"):
                    fichero.seek(0)
                return pd.read_csv(fichero, sep=None, engine="python", dtype=str, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("no se puede leer el CSV")
    xl = pd.ExcelFile(fichero)
    hojas = [h for h in xl.sheet_names if not _norm(h).startswith("instruc")]
    hoja = next((h for h in hojas if "liquid" in _norm(h)), hojas[0] if hojas else xl.sheet_names[0])
    return xl.parse(hoja)


def leer(fichero, nombre=""):
    """(DataFrame con fecha, bruto, comision, neto, referencia, fecha_abono, tarjeta, terminal;
    avisos; error). Filas sin fecha o sin importes se saltan (y se avisa)."""
    try:
        df = _leer_tabla(fichero, nombre)
    except Exception as e:
        return pd.DataFrame(columns=COLS[:8]), [], f"no se puede leer el fichero ({str(e)[:80]})"
    col = columnas_de(list(df.columns))
    if not es_liquidacion(list(df.columns)):
        faltan = [c for c in OBLIGATORIAS if c not in col]
        return pd.DataFrame(columns=COLS[:8]), [], "faltan columnas: " + ", ".join(faltan)
    filas, saltadas, descuadre, neg = [], 0, 0, 0
    for _, r in df.iterrows():
        f = _fecha(r.get(col["fecha"]))
        b = _num(r.get(col["bruto"])) if "bruto" in col else None
        c = _num(r.get(col["comision"])) if "comision" in col else None
        n = _num(r.get(col["neto"])) if "neto" in col else None
        if c is not None and c < 0:          # hay pasarelas que dan la comision en negativo
            c = -c; neg += 1
        if b is None and c is not None and n is not None:
            b = round(n + c, 2)
        if n is None and b is not None and c is not None:
            n = round(b - c, 2)
        if c is None and b is not None and n is not None:
            c = round(b - n, 2)
        if not f or b is None or n is None:
            if any(_txt(v) for v in r.values):
                saltadas += 1
            continue
        if abs(round(b - (c or 0) - n, 2)) > 0.02:
            descuadre += 1
        filas.append({"fecha": f, "bruto": b, "comision": c or 0.0, "neto": n,
                      "referencia": _ref(r.get(col["referencia"])),
                      "fecha_abono": _fecha(r.get(col["fecha_abono"])) if "fecha_abono" in col else "",
                      "tarjeta": _txt(r.get(col["tarjeta"])) if "tarjeta" in col else "",
                      "terminal": _ref(r.get(col["terminal"])) if "terminal" in col else ""})
    avisos = []
    if saltadas:
        avisos.append(f"{saltadas} fila(s) sin fecha o sin importes, saltadas")
    if descuadre:
        avisos.append(f"{descuadre} fila(s) con neto distinto de bruto − comisión")
    if neg:
        avisos.append(f"{neg} comisión(es) en negativo: se toman en positivo")
    if not filas:
        return pd.DataFrame(columns=COLS[:8]), avisos, "ninguna fila con fecha e importes"
    return pd.DataFrame(filas, columns=COLS[:8]), avisos, ""


def leer_store(datos_dir=None):
    try:
        df = pd.read_excel(_ruta(datos_dir), dtype={"referencia": str, "terminal": str, "hotel_id": str})
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    return df[COLS]


def _guardar_store(df, datos_dir=None):
    ruta = _ruta(datos_dir)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + f".{os.getpid()}.tmp.xlsx"
    df[COLS].to_excel(tmp, index=False)
    os.replace(tmp, ruta)


def _clave(r):
    return (_txt(r.get("hotel_id")), _fecha(r.get("fecha")), _ref(r.get("referencia")),
            round(float(_num(r.get("bruto")) or 0), 2), round(float(_num(r.get("neto")) or 0), 2))


@_cand.protegido(_dd_candado)       # con 8 hilos, dos ficheros a la vez no se pisan
def importar(fichero, hotel_id="", datos_dir=None, nombre=""):
    """Añade la liquidacion al registro. Una operacion ya cargada (mismo hotel, fecha,
    referencia, bruto y neto) no se duplica. Devuelve un dict con lo que ha pasado."""
    nombre = nombre or os.path.basename(getattr(fichero, "filename", "") or (fichero if isinstance(fichero, str) else ""))
    df, avisos, err = leer(fichero, nombre)
    if err:
        return {"ok": False, "error": err, "avisos": avisos}
    base = leer_store(datos_dir)
    ya = {_clave(r) for _, r in base.iterrows()}
    hid = _txt(hotel_id)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    nuevas = []
    for _, r in df.iterrows():
        d = {**r.to_dict(), "hotel_id": hid, "fichero": nombre, "cargado": ahora}
        k = _clave(d)
        if k in ya:
            continue
        ya.add(k); nuevas.append(d)
    if nuevas:
        _guardar_store(pd.concat([base, pd.DataFrame(nuevas, columns=COLS)], ignore_index=True), datos_dir)
    nb = pd.DataFrame(nuevas, columns=COLS) if nuevas else pd.DataFrame(columns=COLS)
    return {"ok": True, "leidas": len(df), "nuevas": len(nuevas), "repetidas": len(df) - len(nuevas),
            "bruto": round(float(nb["bruto"].sum()) if nuevas else 0.0, 2),
            "comision": round(float(nb["comision"].sum()) if nuevas else 0.0, 2),
            "neto": round(float(nb["neto"].sum()) if nuevas else 0.0, 2),
            "desde": min(df["fecha"]), "hasta": max(df["fecha"]), "avisos": avisos, "fichero": nombre}


@_cand.protegido(_dd_candado)
def quitar_fichero(fichero, hotel_id="", datos_dir=None):
    """Quita del registro las operaciones que entraron con ese fichero (y ese hotel)."""
    base = leer_store(datos_dir)
    if base.empty:
        return 0
    m = (base["fichero"].map(_txt) == _txt(fichero)) & (base["hotel_id"].map(_txt) == _txt(hotel_id))
    n = int(m.sum())
    if n:
        _guardar_store(base[~m], datos_dir)
    return n


def del_hotel(df, hotel_id=None):
    if df is None or df.empty or not _txt(hotel_id):
        return df
    return df[df["hotel_id"].map(_txt) == _txt(hotel_id)]


def plantilla_excel():
    """La plantilla del formato generico: hoja Liquidacion (cabeceras y dos filas de ejemplo
    que hay que borrar) y hoja Instrucciones."""
    buf = BytesIO()
    ej = pd.DataFrame([
        {"fecha": "02/08/2026", "bruto": 1250.00, "comision": 7.50, "neto": 1242.50, "referencia": "EJEMPLO-0001",
         "fecha_abono": "03/08/2026", "tarjeta": "Visa", "terminal": "Recepción"},
        {"fecha": "02/08/2026", "bruto": 380.40, "comision": 2.28, "neto": 378.12, "referencia": "EJEMPLO-0002",
         "fecha_abono": "03/08/2026", "tarjeta": "Mastercard", "terminal": "Restaurante"},
    ], columns=COLS[:8])
    ins = pd.DataFrame([
        ("fecha", "Obligatoria. Día de la venta (o de la liquidación). dd/mm/aaaa."),
        ("bruto", "Obligatoria. Lo cobrado al cliente."),
        ("comision", "Obligatoria. Lo que se queda la pasarela o el banco."),
        ("neto", "Obligatoria. Lo que llega al banco: bruto − comisión (si falta uno de los tres, Yve lo calcula)."),
        ("referencia", "Obligatoria. Número de operación o de remesa; si sale en el extracto, el cruce es directo."),
        ("fecha_abono", "Opcional. Día en que el dinero llega al banco (si no, se busca desde la fecha y 4 días después)."),
        ("tarjeta", "Opcional. Visa, Mastercard, Amex..."),
        ("terminal", "Opcional. TPV o datáfono (recepción, restaurante...)."),
        ("", "Una fila por operación o una por remesa diaria. Borra las dos filas de EJEMPLO."),
        ("", "Súbela con ⚡ Procesar archivos: Yve la reconoce por estas cabeceras."),
    ], columns=["columna", "qué poner"])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        ej.to_excel(w, index=False, sheet_name="Liquidacion")
        ins.to_excel(w, index=False, sheet_name="Instrucciones")
        for hoja, anchos in (("Liquidacion", (12, 10, 10, 10, 16, 12, 12, 14)), ("Instrucciones", (14, 100))):
            ws = w.sheets[hoja]
            for i, a in enumerate(anchos):
                ws.column_dimensions[chr(65 + i)].width = a
    buf.seek(0)
    return buf, "plantilla_liquidacion_tarjetas.xlsx"


# ── lo cobrado con tarjeta segun el DRR y el TPV ────────────────────────────
def drr_por_dia(ruta_drr):
    """{YYYY-MM-DD: importe} de las cuentas de tarjetas del Trial Balance del DRR procesado
    (hoja Trial_Balance_Completo: Débitos de 'Credit Cards Clearing', Visa...). {} si no hay."""
    if not ruta_drr or not os.path.exists(str(ruta_drr)):
        return {}
    try:
        df = pd.read_excel(ruta_drr, sheet_name="Trial_Balance_Completo")
    except Exception:
        return {}
    col = {_norm(c): c for c in df.columns}
    c_cta, c_f, c_deb, c_tot = col.get("cuenta"), col.get("fecha"), col.get("debitos"), col.get("total")
    if not c_cta or not c_f:
        return {}
    out = {}
    for _, r in df.iterrows():
        n = _txt(r.get(c_cta)).lower()
        if not any(k in n for k in CUENTAS_DRR):
            continue
        f = _fecha(r.get(c_f))
        imp = _num(r.get(c_deb)) if c_deb else None
        if not imp and c_tot:
            imp = _num(r.get(c_tot))
        if f and imp:
            out[f] = round(out.get(f, 0.0) + imp, 2)
    return out


def tpv_por_dia(df_ventas):
    """{YYYY-MM-DD: importe} cobrado con tarjeta segun el TPV, o None si el fichero de ventas
    no trae ni importe con tarjeta ni forma de pago (lo normal: el TPV da ventas por plato)."""
    if df_ventas is None or getattr(df_ventas, "empty", True):
        return None
    col = {_norm(c): c for c in df_ventas.columns}
    c_f = col.get("fecha")
    c_imp = next((col[k] for k in TPV_IMPORTE if k in col), None)
    c_forma = next((col[k] for k in TPV_FORMA if k in col), None)
    c_tot = col.get("total_venta") or col.get("total")
    if not c_f or not (c_imp or (c_forma and c_tot)):
        return None
    out = {}
    for _, r in df_ventas.iterrows():
        f = _fecha(r.get(c_f))
        if not f:
            continue
        if c_imp:
            imp = _num(r.get(c_imp))
        else:
            forma = _txt(r.get(c_forma)).lower()
            imp = _num(r.get(c_tot)) if any(k in forma for k in PALABRAS_TARJETA) else None
        if imp:
            out[f] = round(out.get(f, 0.0) + imp, 2)
    return out


# ── los cruces ──────────────────────────────────────────────────────────────
def _grupos(df):
    """Lo liquidado, agrupado por dia de abono esperado (fecha_abono, o fecha)."""
    g = {}
    for _, r in df.iterrows():
        dia = _fecha(r.get("fecha_abono")) or _fecha(r.get("fecha"))
        if not dia:
            continue
        x = g.setdefault(dia, {"dia": dia, "n": 0, "bruto": 0.0, "comision": 0.0, "neto": 0.0, "refs": [],
                               "fechas_venta": set()})
        x["n"] += 1
        x["bruto"] = round(x["bruto"] + (_num(r.get("bruto")) or 0), 2)
        x["comision"] = round(x["comision"] + (_num(r.get("comision")) or 0), 2)
        x["neto"] = round(x["neto"] + (_num(r.get("neto")) or 0), 2)
        ref = _ref(r.get("referencia"))
        if ref and ref not in x["refs"]:
            x["refs"].append(ref)
        x["fechas_venta"].add(_fecha(r.get("fecha")))
    return [g[k] for k in sorted(g)]


def _fecha_del_concepto(concepto, fecha_mov):
    """'REDSYS LIQ TARJETAS 02/08' -> '2026-08-02' (el año, el del movimiento)."""
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b", _txt(concepto))
    if not m or not fecha_mov:
        return ""
    try:
        d, mm = int(m.group(1)), int(m.group(2))
        a = int(m.group(3)) if m.group(3) else int(fecha_mov[:4])
        if a < 100:
            a += 2000
        if not m.group(3) and mm > int(fecha_mov[5:7]):
            a -= 1
        return date(a, mm, d).isoformat()
    except ValueError:
        return ""


def movimientos_tarjetas(df_banco, palabras=None, manual=None, proveedores=None):
    """Los abonos del extracto que el cuadre de banco pone en la pestaña TARJETAS."""
    out = []
    if df_banco is None or getattr(df_banco, "empty", True):
        return out
    import cuadre_banco as CB
    from almacen_datos import clave_movimiento
    pal = palabras or {k: list(v) for k, v in CB.PALABRAS_DEFECTO.items()}
    for _, r in df_banco.iterrows():
        f = _fecha(r.get("fecha"))
        if not f:
            continue
        d = r.to_dict()
        pest, _via = CB.clasificar(d, pal, manual or {}, proveedores or [])
        imp = _num(r.get("importe")) or 0.0
        if pest == "TARJETAS" and imp > 0:
            out.append({"fecha": f, "concepto": _txt(r.get("concepto")), "importe": imp, "clave": clave_movimiento(d)})
    return sorted(out, key=lambda x: (x["fecha"], x["concepto"]))


def cruce_banco(df_liq, movs, ini, fin):
    """Liquidacion <-> abonos de tarjetas del extracto. Devuelve {filas, sin_liquidacion}:
    filas = un dia de abono esperado del mes con su abono (o sin el); sin_liquidacion = abonos
    de tarjetas del mes que no encajan con ninguna liquidacion."""
    ini_s, fin_s = ini.isoformat(), fin.isoformat()
    desde = (ini - timedelta(days=DIAS_MARGEN)).isoformat()
    grupos = [g for g in _grupos(df_liq) if desde <= g["dia"] <= fin_s]
    libres = list(range(len(movs)))
    enl = {}                                    # indice de grupo -> (indice de movimiento, via)

    def tomar(gi, mi, via):
        enl[gi] = (mi, via); libres.remove(mi)

    # 1. por la referencia en el concepto
    for gi, g in enumerate(grupos):
        refs = [_norm(x) for x in g["refs"] if len(_norm(x)) >= 5]
        if not refs:
            continue
        for mi in list(libres):
            c = _norm(movs[mi]["concepto"])
            if any(x in c for x in refs):
                tomar(gi, mi, "referencia"); break
    # 2. por importe neto, entre el dia y DIAS_MARGEN despues
    for gi, g in enumerate(grupos):
        if gi in enl:
            continue
        tope = (date.fromisoformat(g["dia"]) + timedelta(days=DIAS_MARGEN)).isoformat()
        for mi in list(libres):
            m = movs[mi]
            if g["dia"] <= m["fecha"] <= tope and abs(m["importe"] - g["neto"]) <= TOL:
                tomar(gi, mi, "importe"); break
    # 3. por el dia que nombra el concepto ("REDSYS LIQ TARJETAS 02/08"): enlaza aunque no cuadre
    for mi in list(libres):
        m = movs[mi]
        f_c = _fecha_del_concepto(m["concepto"], m["fecha"])
        if not f_c:
            continue
        for gi, g in enumerate(grupos):
            if gi not in enl and (g["dia"] == f_c or f_c in g["fechas_venta"]):
                tomar(gi, mi, "fecha del concepto"); break

    filas = []
    for gi, g in enumerate(grupos):
        mi, via = enl.get(gi, (None, ""))
        m = movs[mi] if mi is not None else None
        f_ref = m["fecha"] if m else g["dia"]
        if not (ini_s <= f_ref <= fin_s):
            continue                        # es del mes de al lado
        dif = round(m["importe"] - g["neto"], 2) if m else None
        filas.append({"dia": g["dia"], "n": g["n"], "bruto": g["bruto"], "comision": g["comision"], "neto": g["neto"],
                      "refs": g["refs"][:6], "abono": m, "via": via, "diferencia": dif,
                      "estado": ("SIN_ABONO" if not m else "CUADRA" if abs(dif) <= TOL else "DIFERENCIA")})
    sin_liq = [movs[mi] for mi in libres if ini_s <= movs[mi]["fecha"] <= fin_s]
    return {"filas": filas, "sin_liquidacion": sin_liq}


def cruce_ventas(df_liq_mes, drr_dia, tpv_dia, ini, fin):
    """Por dia: bruto liquidado contra lo cobrado con tarjeta segun el DRR y el TPV."""
    ini_s, fin_s = ini.isoformat(), fin.isoformat()
    bruto = {}
    for _, r in df_liq_mes.iterrows():
        f = _fecha(r.get("fecha"))
        if ini_s <= f <= fin_s:
            bruto[f] = round(bruto.get(f, 0.0) + (_num(r.get("bruto")) or 0), 2)
    drr = {k: v for k, v in (drr_dia or {}).items() if ini_s <= k <= fin_s}
    tpv = {k: v for k, v in (tpv_dia or {}).items() if ini_s <= k <= fin_s}
    filas = []
    for f in sorted(set(bruto) | set(drr) | set(tpv)):
        b, d, t = bruto.get(f), drr.get(f), tpv.get(f)
        ventas = None if (d is None and t is None) else round((d or 0) + (t or 0), 2)
        dif = round((b or 0) - ventas, 2) if ventas is not None else None
        estado = ("SIN_DATO" if ventas is None else "SIN_LIQUIDACION" if b is None
                  else "CUADRA" if abs(dif) <= TOL else "DIFERENCIA")
        filas.append({"fecha": f, "bruto": b, "drr": d, "tpv": t, "ventas_tarjeta": ventas, "diferencia": dif, "estado": estado})
    return {"filas": filas, "hay_drr": bool(drr_dia), "hay_tpv": tpv_dia is not None}


def resumen_mes(mes, df_store, hotel_id=None, df_banco=None, drr_dia=None, tpv_dia=None,
                palabras=None, manual=None, proveedores=None):
    """Todo lo de la pestaña: operaciones del mes, totales y los dos cruces."""
    ini, fin, mes = _mes_a_rango(mes)
    df = del_hotel(df_store if df_store is not None else pd.DataFrame(columns=COLS), hotel_id)
    ini_s, fin_s = ini.isoformat(), fin.isoformat()
    df_mes = df[df["fecha"].map(_fecha).between(ini_s, fin_s)] if not df.empty else df
    ops = [{"fecha": _fecha(r.get("fecha")), "referencia": _ref(r.get("referencia")), "tarjeta": _txt(r.get("tarjeta")),
            "terminal": _ref(r.get("terminal")), "bruto": _num(r.get("bruto")) or 0.0, "comision": _num(r.get("comision")) or 0.0,
            "neto": _num(r.get("neto")) or 0.0, "fecha_abono": _fecha(r.get("fecha_abono")), "fichero": _txt(r.get("fichero"))}
           for _, r in df_mes.iterrows()]
    ops.sort(key=lambda x: (x["fecha"], x["referencia"]))
    movs = movimientos_tarjetas(df_banco, palabras, manual, proveedores)
    cb = cruce_banco(df, movs, ini, fin)
    cv = cruce_ventas(df_mes, drr_dia, tpv_dia, ini, fin)
    bruto = round(sum(o["bruto"] for o in ops), 2); com = round(sum(o["comision"] for o in ops), 2)
    neto = round(sum(o["neto"] for o in ops), 2)
    ab = [f for f in cb["filas"] if f["abono"]]
    tot = {"n": len(ops), "bruto": bruto, "comision": com, "neto": neto,
           "pct_comision": round(com / bruto * 100, 2) if bruto else None,
           "abonado": round(sum(f["abono"]["importe"] for f in ab), 2), "n_abonos": len(ab),
           "dias_sin_abono": sum(1 for f in cb["filas"] if f["estado"] == "SIN_ABONO"),
           "dias_con_diferencia": sum(1 for f in cb["filas"] if f["estado"] == "DIFERENCIA"),
           "abonos_sin_liquidacion": len(cb["sin_liquidacion"]),
           "importe_sin_liquidacion": round(sum(m["importe"] for m in cb["sin_liquidacion"]), 2),
           "ventas_tarjeta": round(sum(f["ventas_tarjeta"] or 0 for f in cv["filas"]), 2) if (cv["hay_drr"] or cv["hay_tpv"]) else None,
           "dias_ventas_dif": sum(1 for f in cv["filas"] if f["estado"] in ("DIFERENCIA", "SIN_LIQUIDACION"))}
    ficheros = {}
    for _, r in df.iterrows():
        k = _txt(r.get("fichero"))
        if k:
            x = ficheros.setdefault(k, {"fichero": k, "n": 0, "cargado": _txt(r.get("cargado"))})
            x["n"] += 1
    return {"mes": mes, "desde": ini_s, "hasta": fin_s, "operaciones": ops, "totales": tot,
            "banco": cb, "ventas": cv, "ficheros": sorted(ficheros.values(), key=lambda x: x["cargado"], reverse=True),
            "hay_liquidaciones": not df.empty}


def para_cuadre(df_store, df_banco, mes, hotel_id=None, palabras=None, manual=None, proveedores=None):
    """(justificado, n_dias, nota) para la pestaña TARJETAS del cuadre de banco, o None si no
    hay ninguna liquidacion cargada (entonces el cuadre sigue contra el TPV, como antes).
    Justificado = el neto que las liquidaciones dicen que entra en el banco este mes (por el
    dia del abono encontrado, o el esperado)."""
    df = del_hotel(df_store, hotel_id)
    if df is None or df.empty:
        return None
    ini, fin, mes = _mes_a_rango(mes)
    cb = cruce_banco(df, movimientos_tarjetas(df_banco, palabras, manual, proveedores), ini, fin)
    just = round(sum(f["neto"] for f in cb["filas"]), 2)
    n_sin = sum(1 for f in cb["filas"] if f["estado"] == "SIN_ABONO")
    nota = (f"Justificado = neto de las liquidaciones de tarjetas ({len(cb['filas'])} día(s) de abono; "
            f"{n_sin} sin abono en el extracto, {len(cb['sin_liquidacion'])} abono(s) sin liquidación). Detalle en la pestaña Tarjetas.")
    return just, len(cb["filas"]), nota


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(res["operaciones"] or [{"fecha": "sin operaciones"}]).to_excel(w, index=False, sheet_name="Operaciones")
        filas = [{"dia_abono": f["dia"], "operaciones": f["n"], "bruto": f["bruto"], "comision": f["comision"], "neto": f["neto"],
                  "referencias": ", ".join(f["refs"]), "abono_fecha": (f["abono"] or {}).get("fecha", ""),
                  "abono_concepto": (f["abono"] or {}).get("concepto", ""), "abono_importe": (f["abono"] or {}).get("importe"),
                  "via": f["via"], "diferencia": f["diferencia"], "estado": f["estado"]} for f in res["banco"]["filas"]]
        pd.DataFrame(filas or [{"dia_abono": "sin liquidaciones"}]).to_excel(w, index=False, sheet_name="Cruce banco")
        pd.DataFrame(res["banco"]["sin_liquidacion"] or [{"fecha": "ninguno"}]).to_excel(w, index=False, sheet_name="Abonos sin liquidacion")
        pd.DataFrame(res["ventas"]["filas"] or [{"fecha": "sin datos"}]).to_excel(w, index=False, sheet_name="Cruce ventas")
        pd.DataFrame([res["totales"]]).to_excel(w, index=False, sheet_name="Resumen")
    buf.seek(0)
    return buf, f"tarjetas_{res['mes']}.xlsx"
