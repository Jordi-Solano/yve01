# -*- coding: utf-8 -*-
"""caja.py — Yve.01 · arqueo de caja diario y cuadre con el banco (b86, Jordi 23 sep 2026).

Lo que un controller cuadra cada dia y ningun sistema puede saber sin que alguien
cuente: cuanto efectivo hay. Se teclea UNA vez por dia (recepcion o el controller):

  efectivo_sistema  lo que dice el PMS/TPV que se cobro en efectivo ese dia (opcional)
  efectivo_contado  lo contado al cierre del turno / del dia
  diferencia        contado - sistema  (descuadre de caja: +sobra, -falta)

Lo que Yve pone solo: los INGRESOS de efectivo en el banco (los movimientos del
extracto que el cuadre de banco clasifica como CAJA), con lo que sale cuanto
efectivo sigue en la caja fuerte: lo contado acumulado - lo ingresado acumulado.

Datos: datos-referencia/caja.xlsx, una fila por (fecha, hotel). Puro: las funciones
reciben DataFrames y devuelven dicts; el disco solo lo tocan leer()/guardar().

Contabilidad:
  - el ingreso de efectivo en el banco es 572 (banco) contra 570 (caja): entra en
    los asientos del cierre por cada movimiento CAJA del extracto (cierre_mes);
  - b108 (FINANZAS, 26 sep 2026): el descuadre (sobra/falta) SI se contabiliza, en una
    cuenta propia de "overs & shorts" (config_cierre.json -> cuenta_descuadre_caja; 6591
    por defecto, en el plan de cuentas). Al guardar un arqueo con efectivo segun sistema,
    su asiento sale de `asiento_descuadre` (la pestaña Caja lo enseña y el cierre lo
    asienta, origen CAJA_DESCUADRE):
        falta (contado < sistema):  overs&shorts (D)  /  570 Caja (H)
        sobra (contado > sistema):  570 Caja (D)      /  overs&shorts (H)
    Sin efectivo segun sistema no hay descuadre que asentar.
"""
import os
from datetime import date, datetime

import pandas as pd

import candados as _cand

FICHERO = "caja.xlsx"
COLS = ["fecha", "hotel_id", "efectivo_sistema", "efectivo_contado", "diferencia", "nota", "usuario", "actualizado"]


def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _num(v):
    try:
        f = float(v)
        return None if f != f else round(f, 2)
    except (TypeError, ValueError):
        return None


def _iso(v):
    s = _txt(v)
    if not s:
        return ""
    try:
        t = pd.to_datetime(s[:10], dayfirst=("/" in s[:10]), errors="coerce")
        return "" if pd.isna(t) else t.strftime("%Y-%m-%d")
    except Exception:
        return ""


def _ruta(datos_dir=None):
    if datos_dir is None:
        from tenant_dirs import datos_dir as _d
        datos_dir = str(_d())
    return os.path.join(str(datos_dir), FICHERO)


def leer(datos_dir=None):
    try:
        df = pd.read_excel(_ruta(datos_dir))
    except Exception:
        return pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = None
    return df


def guardar(df, datos_dir=None):
    ruta = _ruta(datos_dir)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp.xlsx"
    df[COLS].to_excel(tmp, index=False)
    os.replace(tmp, ruta)


def _dd_candado(dd):
    if dd:
        return str(dd)
    from tenant_dirs import datos_dir as _d
    return str(_d())


def _eur(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def asiento_descuadre(fila, cuenta):
    """El asiento del descuadre de un arqueo (b108), o None si no hay descuadre (o no hay
    efectivo segun sistema). {fecha, concepto, documento, tipo 'sobra'|'falta', importe,
    lineas: [(cuenta, debe, haber), ...]}. `cuenta` = la de overs & shorts."""
    dif = _num((fila or {}).get("diferencia"))
    if dif is None:
        c = _num((fila or {}).get("efectivo_contado")); s = _num((fila or {}).get("efectivo_sistema"))
        dif = round(c - s, 2) if (c is not None and s is not None) else None
    f = _iso((fila or {}).get("fecha"))
    if dif is None or abs(dif) < 0.01 or not f or not _txt(cuenta):
        return None
    imp = round(abs(dif), 2)
    dd = f[8:10] + "/" + f[5:7]
    if dif < 0:
        tipo, lineas = "falta", [(_txt(cuenta), imp, 0.0), ("570", 0.0, imp)]
    else:
        tipo, lineas = "sobra", [("570", imp, 0.0), (_txt(cuenta), 0.0, imp)]
    return {"fecha": f, "tipo": tipo, "importe": imp, "lineas": lineas,
            "concepto": f"Descuadre de caja {dd}: {'faltan' if tipo == 'falta' else 'sobran'} {_eur(imp)} EUR (overs & shorts)",
            "documento": f"ARQUEO-{f}"}


def asiento_texto(a):
    """'6591 D 15,25 / 570 H 15,25' para tablas y Excel."""
    if not a:
        return ""
    return " / ".join(f"{c} {'D' if d else 'H'} {_eur(d or h)}" for c, d, h in a["lineas"])


@_cand.protegido(_dd_candado)      # b108: con 8 hilos dos arqueos a la vez no se pisan
def apuntar_arqueo(fecha, contado, sistema=None, nota="", usuario="", hotel_id="", datos_dir=None):
    """Guarda (o corrige) el arqueo de un dia. Devuelve la fila guardada."""
    f = _iso(fecha)
    if not f:
        raise ValueError("fecha no valida")
    c = _num(contado)
    if c is None or c < 0:
        raise ValueError("el efectivo contado tiene que ser un numero (0 o mas)")
    s = _num(sistema)
    if s is not None and s < 0:
        raise ValueError("el efectivo segun sistema no puede ser negativo")
    df = leer(datos_dir)
    hid = _txt(hotel_id)
    fila = {"fecha": f, "hotel_id": hid, "efectivo_sistema": s, "efectivo_contado": c,
            "diferencia": round(c - s, 2) if s is not None else None, "nota": _txt(nota)[:200], "usuario": _txt(usuario),
            "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M")}
    if not df.empty:
        mismo = (df["fecha"].map(_iso) == f) & (df["hotel_id"].map(_txt) == hid)
        df = df[~mismo]
    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    guardar(df, datos_dir)
    return fila


@_cand.protegido(_dd_candado)
def borrar_arqueo(fecha, hotel_id="", datos_dir=None):
    df = leer(datos_dir)
    f = _iso(fecha); hid = _txt(hotel_id)
    if df.empty:
        return False
    mismo = (df["fecha"].map(_iso) == f) & (df["hotel_id"].map(_txt) == hid)
    if not mismo.any():
        return False
    guardar(df[~mismo], datos_dir)
    return True


def ingresos_banco(df_banco, ini=None, fin=None, palabras=None, manual=None, proveedores=None):
    """Los ingresos de efectivo del extracto (pestaña CAJA del cuadre, con las asignaciones a mano),
    [{fecha, concepto, importe, clave, estado, hotel_id}] ordenados por fecha."""
    out = []
    if df_banco is None or df_banco.empty:
        return out
    import cuadre_banco as CB
    from almacen_datos import clave_movimiento
    pal = palabras or {k: list(v) for k, v in CB.PALABRAS_DEFECTO.items()}
    for _, r in df_banco.iterrows():
        f = _iso(r.get("fecha"))
        if not f:
            continue
        if ini and f < ini:
            continue
        if fin and f > fin:
            continue
        d = r.to_dict()
        pest, _via = CB.clasificar(d, pal, manual or {}, proveedores or [])
        imp = _num(r.get("importe")) or 0.0
        if pest == "CAJA" and imp > 0:
            out.append({"fecha": f, "concepto": _txt(r.get("concepto")), "importe": imp, "clave": clave_movimiento(d),
                        "estado": _txt(r.get("estado")).upper() or "PENDIENTE", "hotel_id": _txt(r.get("hotel_id"))})
    return sorted(out, key=lambda x: x["fecha"])


def ingresos_banco_dir(df_banco, ini=None, fin=None, datos_dir=None):
    """Como ingresos_banco, leyendo palabras y asignaciones a mano de la carpeta de datos."""
    import cuadre_banco as CB
    dd = str(datos_dir) if datos_dir else None
    return ingresos_banco(df_banco, ini, fin, CB.palabras(dd), CB.manuales(dd), CB.proveedores_conocidos(dd))


def contado_mes(df_caja, ini, fin, hotel_id=None):
    """Efectivo contado en los arqueos entre ini y fin (para justificar la pestaña CAJA del cuadre)."""
    if df_caja is None or df_caja.empty:
        return None, 0
    hid = _txt(hotel_id); tot = 0.0; n = 0
    for _, r in df_caja.iterrows():
        if hid and _txt(r.get("hotel_id")) != hid:
            continue
        f = _iso(r.get("fecha"))
        if f and ini <= f <= fin:
            tot += _num(r.get("efectivo_contado")) or 0.0; n += 1
    return (round(tot, 2), n) if n else (None, 0)


def resumen_mes(mes, df_caja, df_banco=None, hotel_id=None, palabras=None, manual=None, proveedores=None):
    """{mes, dias: [...], totales, en_caja, ingresos: [...]} para el panel y el cuadre."""
    ini = f"{mes}-01"
    fin = (pd.Timestamp(ini) + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
    hid = _txt(hotel_id)
    filas = []
    if df_caja is not None and not df_caja.empty:
        for _, r in df_caja.iterrows():
            if hid and _txt(r.get("hotel_id")) != hid:
                continue
            f = _iso(r.get("fecha"))
            if not f:
                continue
            filas.append({"fecha": f, "efectivo_sistema": _num(r.get("efectivo_sistema")), "efectivo_contado": _num(r.get("efectivo_contado")) or 0.0,
                          "diferencia": _num(r.get("diferencia")), "nota": _txt(r.get("nota")), "usuario": _txt(r.get("usuario")),
                          "actualizado": _txt(r.get("actualizado"))})
    filas.sort(key=lambda x: x["fecha"])
    ingresos = ingresos_banco(df_banco, None, fin, palabras, manual, proveedores)
    # acumulados hasta el fin de mes (lo que sigue en caja no se reinicia cada mes)
    contado_total = round(sum(x["efectivo_contado"] for x in filas if x["fecha"] <= fin), 2)
    ingresado_total = round(sum(i["importe"] for i in ingresos), 2)
    dias = [x for x in filas if ini <= x["fecha"] <= fin]
    ing_mes = [i for i in ingresos if i["fecha"] >= ini]
    sist = [x["efectivo_sistema"] for x in dias if x["efectivo_sistema"] is not None]
    difs = [x["diferencia"] for x in dias if x["diferencia"] is not None]
    tot = {
        "n_dias": len(dias),
        "contado": round(sum(x["efectivo_contado"] for x in dias), 2),
        "sistema": round(sum(sist), 2) if sist else None,
        "diferencia": round(sum(difs), 2) if difs else None,
        "dias_con_descuadre": sum(1 for d in difs if abs(d) >= 0.01),
        "ingresado": round(sum(i["importe"] for i in ing_mes), 2),
        "n_ingresos": len(ing_mes),
    }
    return {"mes": mes, "desde": ini, "hasta": fin, "dias": dias, "ingresos": ing_mes, "totales": tot,
            "en_caja": round(contado_total - ingresado_total, 2),
            "nota": ("En caja = todo lo contado hasta fin de mes menos todo lo ingresado en el banco (movimientos "
                     "CAJA del extracto). Solo tiene sentido si se apunta el arqueo cada dia.")}


def exportar_excel(res):
    from io import BytesIO
    buf = BytesIO()
    dias = [{**{k: v for k, v in d.items() if k != "asiento"}, **({"asiento": asiento_texto(d.get("asiento"))} if "asiento" in d else {})}
            for d in (res["dias"] or [])]
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(dias or [{"fecha": "sin arqueos"}]).to_excel(w, index=False, sheet_name="Arqueos")
        pd.DataFrame(res["ingresos"] or [{"fecha": "sin ingresos en banco"}]).to_excel(w, index=False, sheet_name="Ingresos banco")
        pd.DataFrame([{**res["totales"], "en_caja": res["en_caja"]}]).to_excel(w, index=False, sheet_name="Resumen")
    buf.seek(0)
    return buf, f"caja_{res['mes']}.xlsx"
