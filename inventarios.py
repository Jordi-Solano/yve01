# -*- coding: utf-8 -*-
"""inventarios.py — inventarios de cierre: alimentos, bebidas, licores, guest supplies (OLA B · 3).

Con el recuento de fin de mes (`inventario.xlsx`: stock inicial y final por
articulo, coste unitario) Yve saca por FAMILIA:
  valor inicial · compras del mes · valor final · consumo real
  (= inicial + compras - final) · consumo teorico (escandallo × ventas, F&B)
  · desviacion · el asiento de variacion de existencias.

Familias: ALIMENTOS, BEBIDAS, LICORES, GUEST_SUPPLIES, OTROS. Se deciden por
la categoria/nombre del articulo (config_inventarios.json puede fijar la
familia de cada categoria). Lo que no encaja va a OTROS, no se adivina.

Lo que Yve NO sabe no se inventa: las compras se toman de las facturas AP de
proveedores F&B del mes (base imponible); si una familia no tiene compras
identificables su consumo real sale "sin compras" y se dice.

Funciones puras; la unica escritura (el recuento subido) esta en tab_cierre y
pasa por `dashboard._guardar_fb_del_hotel`, la misma puerta que el clasificador.
"""
import json
import os
import unicodedata
from io import BytesIO

import pandas as pd

from provisiones import _fecha, _num, _txt, _mes_a_rango

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FAMILIAS = ("ALIMENTOS", "BEBIDAS", "LICORES", "GUEST_SUPPLIES", "OTROS")
PALABRAS = {
    "LICORES":        ["licor", "whisky", "whiskey", "gin", "ginebra", "ron", "vodka", "brandy", "coñac", "cognac",
                       "tequila", "vermut", "vermouth", "destilado", "orujo", "anis", "espirituoso", "bourbon", "cava", "champagne", "champan"],
    "BEBIDAS":        ["bebida", "vino", "cerveza", "refresco", "agua", "zumo", "cafe", "infusion", "cola",
                       "tonica", "sidra", "soda", "botella"],
    "GUEST_SUPPLIES": ["amenit", "guest", "jabon", "jabón", "champu", "champú", "gel", "papel higienico", "papel higiénico",
                       "gorro", "zapatilla", "kit dental", "cepillo", "peine", "bodymilk", "crema", "toalla", "albornoz",
                       "supplies", "welcome"],
}
CUENTAS = {   # existencias / variacion de existencias (PGC)
    "ALIMENTOS":      ("300", "Mercaderias — alimentos", "610", "Variacion de existencias de mercaderias"),
    "BEBIDAS":        ("300", "Mercaderias — bebidas", "610", "Variacion de existencias de mercaderias"),
    "LICORES":        ("300", "Mercaderias — licores", "610", "Variacion de existencias de mercaderias"),
    "GUEST_SUPPLIES": ("328", "Material diverso — guest supplies", "612", "Variacion de existencias de otros aprovisionamientos"),
    "OTROS":          ("300", "Mercaderias — otros", "610", "Variacion de existencias de mercaderias"),
}


def _norm(x):
    x = unicodedata.normalize("NFKD", _txt(x).lower())
    return "".join(c for c in x if not unicodedata.combining(c))


def config_familias(datos_dir=None):
    """categoria (normalizada) -> familia, desde config_inventarios.json."""
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), "config_inventarios.json")
    out = {}
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh) or {}
        for k, v in (d.get("familias") or d).items():
            if str(v).upper() in FAMILIAS:
                out[_norm(k)] = str(v).upper()
    except Exception:
        pass
    return out


def _encaja(palabra, texto):
    """Por palabras, no por subcadena: 'gel' no debe cazar 'gelatina', pero
    'amenit' si debe cazar 'amenities'."""
    palabra = _norm(palabra).strip()
    for tok in texto.replace("-", " ").replace("/", " ").split():
        if tok == palabra or (len(palabra) >= 5 and tok.startswith(palabra)):
            return True
    return len(palabra.split()) > 1 and palabra in texto


def familia(categoria, nombre="", cfg=None):
    cat = _norm(categoria); nom = _norm(nombre)
    if cfg and cat in cfg:
        return cfg[cat]
    for fam in ("LICORES", "GUEST_SUPPLIES", "BEBIDAS"):
        if any(_encaja(p, cat) or _encaja(p, nom) for p in PALABRAS[fam]):
            return fam
    if cat in ("carnes", "pescados", "mariscos", "verduras", "lacteos", "secos", "frutas", "panaderia",
               "congelados", "charcuteria", "conservas", "alimentos", "food", "cocina"):
        return "ALIMENTOS"
    if cat:
        return "ALIMENTOS" if any(p in cat for p in ("aliment", "comida", "fresc")) else "OTROS"
    return "OTROS"


def valorar(mes, df_inv, df_ap=None, coste_teorico_fb=None, cfg=None):
    """Devuelve {familias: [...], articulos: [...], asientos: [...], revisar: [...], resumen}."""
    ini, fin, mes = _mes_a_rango(mes)
    cfg = cfg or {}
    arts = []
    if df_inv is not None and not df_inv.empty:
        for _, r in df_inv.iterrows():
            nombre = _txt(r.get("ingrediente")) or _txt(r.get("producto"))
            if not nombre:
                continue
            coste = _num(r.get("coste_unitario"))
            fin_q = _num(r.get("stock_actual_kg_l")) if _txt(r.get("stock_actual_kg_l")) else None
            ini_q = _num(r.get("stock_inicial_kg_l")) if _txt(r.get("stock_inicial_kg_l")) else None
            fam = familia(r.get("categoria"), nombre, cfg)
            arts.append({
                "articulo": nombre, "categoria": _txt(r.get("categoria")), "familia": fam,
                "unidad": _txt(r.get("unidad")), "coste_unitario": coste,
                "stock_inicial": ini_q, "stock_final": fin_q,
                "valor_inicial": round((ini_q or 0) * coste, 2), "valor_final": round((fin_q or 0) * coste, 2),
                "proveedor": _txt(r.get("proveedor")), "hotel_id": _txt(r.get("hotel_id")),
                "revisar": (not coste) or fin_q is None or (fin_q is not None and fin_q < 0),
                "motivo": ("sin coste unitario" if not coste else "sin recuento final" if fin_q is None
                           else "stock negativo" if fin_q < 0 else ""),
            })
    # compras del mes: facturas AP de proveedores F&B (base imponible)
    compras = 0.0; n_compras = 0
    if df_ap is not None and not df_ap.empty:
        for _, r in df_ap.iterrows():
            if _txt(r.get("tipo_proveedor")).upper() != "FB":
                continue
            fecha = r.get("fecha_factura") if _txt(r.get("fecha_factura")) else r.get("fecha")
            f = _fecha(fecha)
            if f and ini <= f <= fin:
                base = _num(r.get("base_imponible"))
                if not base:
                    tot = _num(r.get("total_factura")); base = round(tot / 1.10, 2) if tot else 0.0
                compras = round(compras + base, 2); n_compras += 1

    fams = []
    asientos = []
    for fam in FAMILIAS:
        xs = [a for a in arts if a["familia"] == fam]
        if not xs:
            continue
        v_ini = round(sum(a["valor_inicial"] for a in xs), 2)
        v_fin = round(sum(a["valor_final"] for a in xs), 2)
        delta = round(v_fin - v_ini, 2)
        cta_ex, d_ex, cta_var, d_var = CUENTAS[fam]
        if abs(delta) >= 0.01:
            if delta > 0:
                asientos.append({"fecha": fin.isoformat(), "cuenta": cta_ex, "desc_cuenta": d_ex, "concepto": f"Variacion existencias {mes} — {fam}", "debe": delta, "haber": 0.0, "familia": fam})
                asientos.append({"fecha": fin.isoformat(), "cuenta": cta_var, "desc_cuenta": d_var, "concepto": f"Variacion existencias {mes} — {fam}", "debe": 0.0, "haber": delta, "familia": fam})
            else:
                asientos.append({"fecha": fin.isoformat(), "cuenta": cta_var, "desc_cuenta": d_var, "concepto": f"Variacion existencias {mes} — {fam}", "debe": -delta, "haber": 0.0, "familia": fam})
                asientos.append({"fecha": fin.isoformat(), "cuenta": cta_ex, "desc_cuenta": d_ex, "concepto": f"Variacion existencias {mes} — {fam}", "debe": 0.0, "haber": -delta, "familia": fam})
        fams.append({"familia": fam, "n": len(xs), "valor_inicial": v_ini, "valor_final": v_fin, "variacion": delta,
                     "revisar": sum(1 for a in xs if a["revisar"])})

    # consumo real F&B = inicial + compras - final (alimentos + bebidas + licores)
    fb = [f for f in fams if f["familia"] in ("ALIMENTOS", "BEBIDAS", "LICORES")]
    ini_fb = round(sum(f["valor_inicial"] for f in fb), 2); fin_fb = round(sum(f["valor_final"] for f in fb), 2)
    consumo_real = round(ini_fb + compras - fin_fb, 2) if fb else None
    desv = None
    if consumo_real is not None and coste_teorico_fb is not None:
        desv = round(consumo_real - coste_teorico_fb, 2)
    resumen = {
        "mes": mes, "n_articulos": len(arts), "n_revisar": sum(1 for a in arts if a["revisar"]),
        "valor_inicial": round(sum(f["valor_inicial"] for f in fams), 2),
        "valor_final": round(sum(f["valor_final"] for f in fams), 2),
        "compras_fb": compras, "n_facturas_fb": n_compras,
        "consumo_real_fb": consumo_real, "consumo_teorico_fb": coste_teorico_fb,
        "desviacion_fb": desv,
        "desviacion_pct": round(desv / coste_teorico_fb * 100, 1) if desv is not None and coste_teorico_fb else None,
        "nota": ("Consumo real = existencias iniciales + compras F&B del mes - existencias finales. "
                 "Teorico = escandallo × unidades vendidas. La diferencia son mermas, regalos, robos o recuento mal hecho."),
    }
    return {"mes": mes, "familias": fams, "articulos": arts, "asientos": asientos,
            "revisar": [a for a in arts if a["revisar"]], "resumen": resumen}


def hoja_recuento(df_inv, mes, cfg=None):
    """Excel para contar: lo que dice el sistema y una columna vacia 'recuento'."""
    filas = []
    if df_inv is not None and not df_inv.empty:
        for _, r in df_inv.iterrows():
            nombre = _txt(r.get("ingrediente")) or _txt(r.get("producto"))
            if not nombre:
                continue
            filas.append({"ingrediente": nombre, "categoria": _txt(r.get("categoria")),
                          "familia": familia(r.get("categoria"), nombre, cfg), "unidad": _txt(r.get("unidad")),
                          "stock_sistema": _num(r.get("stock_actual_kg_l")), "recuento": None,
                          "coste_unitario": _num(r.get("coste_unitario")), "observaciones": ""})
    filas.sort(key=lambda x: (x["familia"], x["ingrediente"]))
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(filas or [{"ingrediente": "", "recuento": None}]).to_excel(w, index=False, sheet_name="Recuento")
        ws = w.sheets["Recuento"]
        ws.freeze_panes = "A2"
        for col, ancho in zip("ABCDEFGH", (28, 14, 16, 8, 14, 12, 14, 30)):
            ws.column_dimensions[col].width = ancho
    buf.seek(0)
    return buf, f"recuento_inventario_{mes}.xlsx"


def leer_recuento(fichero):
    """Lee la hoja de recuento rellenada -> DataFrame para inventario.xlsx.

    Solo se toman las filas con `recuento` informado; el resto no se toca.
    `recuento` pasa a stock_actual_kg_l. Devuelve (df, n_contadas, n_saltadas).
    """
    df = pd.read_excel(fichero)
    cols = {c.lower().strip(): c for c in df.columns}
    c_nom = cols.get("ingrediente") or cols.get("producto") or cols.get("articulo")
    c_rec = cols.get("recuento") or cols.get("stock_final") or cols.get("stock_actual")
    if not c_nom or not c_rec:
        raise ValueError("La hoja tiene que traer las columnas 'ingrediente' y 'recuento'")
    filas = []; saltadas = 0
    for _, r in df.iterrows():
        nombre = _txt(r.get(c_nom))
        rec = r.get(c_rec)
        if not nombre or not _txt(rec):
            saltadas += 1
            continue
        try:
            q = float(str(rec).replace(",", "."))
        except Exception:
            saltadas += 1
            continue
        fila = {"ingrediente": nombre, "stock_actual_kg_l": q}
        for k in ("categoria", "unidad", "coste_unitario", "proveedor"):
            if k in cols and _txt(r.get(cols[k])):
                fila[k] = r.get(cols[k])
        filas.append(fila)
    return pd.DataFrame(filas), len(filas), saltadas


def exportar_excel(res):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([res["resumen"]]).to_excel(w, index=False, sheet_name="Resumen")
        pd.DataFrame(res["familias"] or [{}]).to_excel(w, index=False, sheet_name="Familias")
        pd.DataFrame(res["articulos"] or [{}]).to_excel(w, index=False, sheet_name="Articulos")
        pd.DataFrame(res["asientos"] or [{}]).to_excel(w, index=False, sheet_name="Asientos")
        pd.DataFrame(res["revisar"] or [{}]).to_excel(w, index=False, sheet_name="Revisar")
    buf.seek(0)
    return buf, f"inventarios_{res['mes']}.xlsx"
