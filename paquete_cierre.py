# -*- coding: utf-8 -*-
"""paquete_cierre.py — el archivo de fin de mes para la central (OLA B · bloque 5).

Lo ultimo que hace el controller: un solo paquete con todo el cierre y sus
comentarios. Yve lo monta desde los bloques ya hechos (asientos, reconciliacion,
banco, provisiones, inventarios, inmovilizado, aging) y le añade:

  · un CHECKLIST del cierre: cada bloque con su estado (OK / PENDIENTE / SIN
    DATO) y la cifra que lo resume; nunca un OK sin dato detras;
  · la cuenta de resultados del mes segun lo asentado (ingresos 7xx, gastos
    6xx, resultado) — con la advertencia de que solo recoge lo que entra por
    documento;
  · los COMENTARIOS del controller por seccion (comentarios_cierre.json), que
    son lo que la central lee primero.

`montar` es pura (recibe los resultados de cada bloque); `comentarios` /
`guardar_comentario` son la unica lectura/escritura propia.
"""
import json
import os
from datetime import datetime
from io import BytesIO

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FICHERO_COMENTARIOS = "comentarios_cierre.json"
SECCIONES = ("resumen", "asientos", "reconciliacion", "banco", "provisiones", "inventarios", "inmovilizado", "aging", "fiscal")


def _r(x):
    return round(float(x or 0), 2)


def comentarios(mes, datos_dir=None):
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), FICHERO_COMENTARIOS)
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh) or {}
        return d.get(mes, {}) if isinstance(d, dict) else {}
    except Exception:
        return {}


def guardar_comentario(mes, seccion, texto, usuario="", datos_dir=None):
    if seccion not in SECCIONES:
        raise ValueError("seccion desconocida")
    ruta = os.path.join(datos_dir or os.path.join(BASE_DIR, "datos-referencia"), FICHERO_COMENTARIOS)
    try:
        with open(ruta, encoding="utf-8") as fh:
            d = json.load(fh) or {}
    except Exception:
        d = {}
    if not isinstance(d, dict):
        d = {}
    m = d.setdefault(mes, {})
    texto = (texto or "").strip()
    if texto:
        m[seccion] = {"texto": texto, "usuario": usuario, "fecha": datetime.now().strftime("%Y-%m-%d %H:%M")}
    else:
        m.pop(seccion, None)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)
    return m


def cuenta_resultados(mayor):
    """Ingresos (7xx) y gastos (6xx) del Diario del mes -> resultado."""
    ing = [m for m in mayor if str(m["cuenta"]).startswith("7")]
    gas = [m for m in mayor if str(m["cuenta"]).startswith("6")]
    ingresos = _r(sum(m["haber"] - m["debe"] for m in ing))
    gastos = _r(sum(m["debe"] - m["haber"] for m in gas))
    return {"ingresos": ingresos, "gastos": gastos, "resultado": _r(ingresos - gastos),
            "por_cuenta": [{"cuenta": m["cuenta"], "descripcion": m["descripcion"],
                            "importe": _r(m["haber"] - m["debe"]) if str(m["cuenta"]).startswith("7") else _r(m["debe"] - m["haber"]),
                            "tipo": "ingreso" if str(m["cuenta"]).startswith("7") else "gasto"} for m in ing + gas]}


def _eur(v, dec=2):
    """1234.5 -> '1.234,50 €' (mismo formato que el panel: ronda de pruebas, punto 6)."""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return "—"
    if x != x:
        return "—"
    s = f"{abs(x):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if x < 0 else "") + s + " €"


def montar(mes, asientos=None, reconciliacion=None, banco=None, provisiones=None, inventarios=None,
           inmovilizado=None, aging=None, fiscal=None, comentarios_mes=None, drr=None):
    """Checklist + resumen. Cada bloque puede venir None (no calculado / sin datos)."""
    chk = []

    def item(clave, titulo, estado, cifra, detalle=""):
        chk.append({"clave": clave, "titulo": titulo, "estado": estado, "cifra": cifra, "detalle": detalle,
                    "comentario": (comentarios_mes or {}).get(clave, {}).get("texto", "")})

    if asientos:
        item("asientos", "Asientos del mes", "OK" if asientos.get("cuadra") and asientos.get("n_asientos") else ("SIN_DATO" if not asientos.get("n_asientos") else "PENDIENTE"),
             f"{asientos.get('n_asientos', 0)} asientos · {_eur(asientos.get('debe', 0))}",
             "; ".join(asientos.get("avisos") or []))
    else:
        item("asientos", "Asientos del mes", "SIN_DATO", "", "")
    hay_asientos = bool(asientos and asientos.get("n_asientos"))
    if reconciliacion and not hay_asientos:
        # Sin asientos no hay nada que reconciliar: "7 cuadran" con todo a
        # cero no es un OK, es que no hay datos del mes (ronda de pruebas, fase 6).
        item("reconciliacion", "Reconciliacion de cuentas", "SIN_DATO", "sin asientos del mes", "")
    elif reconciliacion:
        n = reconciliacion.get("resumen", {})
        item("reconciliacion", "Reconciliacion de cuentas", "OK" if reconciliacion.get("ok") else "PENDIENTE",
             f"{n.get('CUADRA', 0)} cuadran · {n.get('DIFERENCIA', 0)} con diferencia · {n.get('PENDIENTE', 0)} pendientes · {n.get('SIN_DATO', 0)} sin dato",
             "; ".join(c["concepto"] for c in reconciliacion.get("checks", []) if c["estado"] in ("DIFERENCIA", "REVISAR")))
    else:
        item("reconciliacion", "Reconciliacion de cuentas", "SIN_DATO", "", "")
    if banco and banco.get("n"):
        p = banco.get("pestanas", {})
        item("banco", "Cuadre de banco por pestañas", "OK" if banco.get("ok") else "PENDIENTE",
             f"{banco.get('n', 0)} movimientos · {banco.get('sin_clasificar', 0)} sin clasificar · {banco.get('sin_conciliar', 0)} sin conciliar"
             + (f" · saldo {_eur(banco.get('saldo_final'))}" if banco.get("saldo_final") is not None else ""),
             " · ".join(f"{k}: {_eur(v.get('total', 0))}" for k, v in p.items() if v.get("n")))
    else:
        item("banco", "Cuadre de banco por pestañas", "SIN_DATO", "sin extracto del mes", "")
    if provisiones:
        alb = provisiones[0] if len(provisiones) > 0 else {}
        com = provisiones[1] if len(provisiones) > 1 else {}
        tot = _r((alb or {}).get("total", 0) + (com or {}).get("total", 0))
        item("provisiones", "Provisiones (albaran sin factura, comisiones)", "OK" if provisiones else "SIN_DATO",
             f"{_eur(tot)} · {(alb or {}).get('n', 0)} albaranes · {(com or {}).get('n', 0)} liquidaciones", "")
    else:
        item("provisiones", "Provisiones", "SIN_DATO", "", "")
    if inventarios and inventarios.get("resumen", {}).get("n_articulos"):
        s = inventarios["resumen"]
        item("inventarios", "Inventarios de cierre", "PENDIENTE" if s.get("n_revisar") else "OK",
             f"existencias {_eur(s.get('valor_final', 0))} · consumo real {_eur(s.get('consumo_real_fb')) if s.get('consumo_real_fb') is not None else '—'}"
             + (f" · desviacion {s.get('desviacion_pct')} %" if s.get("desviacion_pct") is not None else ""),
             f"{s.get('n_revisar', 0)} articulos a revisar" if s.get("n_revisar") else "")
    else:
        _otros = ((inventarios or {}).get("resumen") or {}).get("otros_meses") or []
        item("inventarios", "Inventarios de cierre", "SIN_DATO",
             f"sin recuento de {mes} (el inventario guardado es de {', '.join(_otros)})" if _otros else "sin recuento", "")
    if inmovilizado and inmovilizado.get("resumen", {}).get("n_activos"):
        s = inmovilizado["resumen"]
        item("inmovilizado", "Inmovilizado y amortizaciones", "PENDIENTE" if (s.get("n_error") or s.get("altas_pendientes")) else "OK",
             f"{s.get('n_activos', 0)} activos · cuota {_eur(s.get('cuota_mes', 0))} · VNC {_eur(s.get('vnc_total', 0))}",
             (f"{s.get('altas_pendientes')} posibles altas sin registrar" if s.get("altas_pendientes") else "") + (f" · {s.get('n_error')} con error" if s.get("n_error") else ""))
    else:
        item("inmovilizado", "Inmovilizado y amortizaciones", "SIN_DATO", "sin activos registrados", "")
    if aging and aging.get("n"):
        item("aging", "Aging AP (deuda con proveedores)", "PENDIENTE" if aging.get("mas_de_60") else "OK",
             f"{_eur(aging.get('total', 0))} pendientes · {_eur(aging.get('mas_de_60', 0))} a mas de 60 dias", "")
    else:
        item("aging", "Aging AP", "SIN_DATO", "nada pendiente", "")
    if fiscal:
        item("fiscal", "Fiscal (303 / 349 / SII)", fiscal.get("estado", "SIN_DATO"), fiscal.get("cifra", ""), fiscal.get("detalle", ""))
    else:
        item("fiscal", "Fiscal (303 / 349 / SII)", "SIN_DATO", "pendiente del bloque fiscal", "")

    resultado = cuenta_resultados((reconciliacion or {}).get("mayor") or []) if reconciliacion else {"ingresos": 0.0, "gastos": 0.0, "resultado": 0.0, "por_cuenta": []}
    if drr and drr.get("rooms_revenue_mtd") is not None:
        resultado["drr_rooms_revenue"] = drr["rooms_revenue_mtd"]
    n = {k: sum(1 for c in chk if c["estado"] == k) for k in ("OK", "PENDIENTE", "SIN_DATO")}
    return {
        "mes": mes, "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        # "Listo" exige que haya asientos del mes: un mes vacio (todo SIN_DATO)
        # no esta listo para la central, esta sin datos.
        "checklist": chk, "resumen_checklist": n, "listo": n["PENDIENTE"] == 0 and hay_asientos,
        "sin_datos": not hay_asientos,
        "resultado": resultado,
        "comentario_general": (comentarios_mes or {}).get("resumen", {}).get("texto", ""),
        "nota": ("La cuenta de resultados recoge SOLO lo que entra por documento en Yve (facturas, TPV, extracto, "
                 "provisiones). Los ingresos de habitaciones cobrados por el PMS no estan hasta tener conector: "
                 "compara con el DRR."),
    }


def exportar_excel(paq, asientos=None, reconciliacion=None, banco=None, provisiones=None, inventarios=None,
                   inmovilizado=None, aging=None, fiscal=None):
    """Un Excel con todo: Portada, Checklist, Resultados, y una hoja por bloque."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        portada = [{"campo": "Mes", "valor": paq["mes"]}, {"campo": "Generado", "valor": paq["generado"]},
                   {"campo": "Estado", "valor": "LISTO PARA LA CENTRAL" if paq["listo"] else ("SIN DATOS DEL MES" if paq.get("sin_datos") else f"{paq['resumen_checklist']['PENDIENTE']} bloque(s) pendiente(s)")},
                   {"campo": "Comentario general", "valor": paq.get("comentario_general", "")},
                   {"campo": "Ingresos asentados", "valor": paq["resultado"]["ingresos"]},
                   {"campo": "Gastos asentados", "valor": paq["resultado"]["gastos"]},
                   {"campo": "Resultado del mes (segun documentos)", "valor": paq["resultado"]["resultado"]},
                   {"campo": "Nota", "valor": paq["nota"]}]
        pd.DataFrame(portada).to_excel(w, index=False, sheet_name="Portada")
        pd.DataFrame(paq["checklist"]).to_excel(w, index=False, sheet_name="Checklist")
        pd.DataFrame(paq["resultado"]["por_cuenta"] or [{}]).to_excel(w, index=False, sheet_name="Resultados")
        if asientos:
            pd.DataFrame(asientos.get("asientos") or [{}]).to_excel(w, index=False, sheet_name="Libro Diario")
        if reconciliacion:
            pd.DataFrame(reconciliacion.get("mayor") or [{}]).to_excel(w, index=False, sheet_name="Mayor")
            pd.DataFrame(reconciliacion.get("checks") or [{}]).to_excel(w, index=False, sheet_name="Reconciliacion")
        if banco:
            pd.DataFrame([{"pestana": k, **v} for k, v in banco.get("pestanas", {}).items()]).to_excel(w, index=False, sheet_name="Banco pestañas")
            pd.DataFrame(banco.get("movimientos") or [{}]).to_excel(w, index=False, sheet_name="Banco movimientos")
        if provisiones:
            filas = []
            for b in provisiones:
                filas += (b or {}).get("asientos") or []
            pd.DataFrame(filas or [{}]).to_excel(w, index=False, sheet_name="Provisiones")
        if inventarios:
            pd.DataFrame(inventarios.get("familias") or [{}]).to_excel(w, index=False, sheet_name="Inventarios")
        if inmovilizado:
            pd.DataFrame(inmovilizado.get("activos") or [{}]).to_excel(w, index=False, sheet_name="Inmovilizado")
        if aging:
            pd.DataFrame(aging.get("por_acreedor") or [{}]).to_excel(w, index=False, sheet_name="Aging AP")
        if fiscal and fiscal.get("hojas"):
            for nombre, filas in fiscal["hojas"].items():
                pd.DataFrame(filas or [{}]).to_excel(w, index=False, sheet_name=nombre[:31])
        for ws in w.sheets.values():
            ws.freeze_panes = "A2"
    buf.seek(0)
    return buf, f"cierre_{paq['mes']}_paquete_central.xlsx"
