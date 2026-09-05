# -*- coding: utf-8 -*-
"""control_importes.py — ¿cuadran los importes de la factura? (Jordi, sep 2026)

Al leer una factura AP se comprueba, sin tocar el clasificador:
  base imponible + cuota de IVA = total          (tolerancia 2 centimos)
  cuota de IVA = base × porcentaje / 100         (si viene el porcentaje)
Si no cuadra se AVISA (columna `aviso_importes` en la factura guardada, badge
en el panel, linea en el log). No se corrige nada: el que no cuadre es
informacion — o el proveedor se ha equivocado, o la IA ha leido mal un numero.
Ambas cosas hay que verlas antes de aprobar.
"""
TOL = 0.02


def _num(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if v != v else float(v)
    s = str(v).strip().replace("€", "").replace("EUR", "").replace(" ", "")
    if not s or s.upper() in ("NO_ENCONTRADO", "NAN", "NONE"):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _es(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"


def comprobar(fila):
    """Devuelve (cuadra, aviso). cuadra=None si faltan datos para comprobar."""
    base = _num(fila.get("base_imponible"))
    cuota = _num(fila.get("cuota_iva"))
    total = _num(fila.get("total_factura"))
    pct = _num(fila.get("porcentaje_iva"))
    avisos = []
    if base is None or total is None:
        return None, ""
    if cuota is None and pct is not None:
        cuota = round(base * pct / 100, 2)
    if cuota is None:
        # sin cuota ni porcentaje solo se puede mirar que el total no sea menor que la base
        if total + TOL < base:
            return False, f"el total ({_es(total)}) es menor que la base ({_es(base)})"
        return None, ""
    suma = round(base + cuota, 2)
    if abs(suma - total) > TOL:
        avisos.append(f"base {_es(base)} + IVA {_es(cuota)} = {_es(suma)}, pero el total dice {_es(total)} (diferencia {_es(round(total - suma, 2))})")
    if pct is not None and pct > 0:
        esperada = round(base * pct / 100, 2)
        if abs(esperada - cuota) > TOL:
            avisos.append(f"IVA al {pct:g} % sobre {_es(base)} serian {_es(esperada)}, no {_es(cuota)}")
    return (not avisos), " · ".join(avisos)


def anotar(filas):
    """Escribe `importes_cuadran` y `aviso_importes` en cada dict. Devuelve las que no cuadran."""
    mal = []
    for f in filas:
        if not isinstance(f, dict):
            continue
        ok, aviso = comprobar(f)
        f["importes_cuadran"] = "" if ok is None else ("SI" if ok else "NO")
        f["aviso_importes"] = aviso
        if ok is False:
            mal.append(f)
    return mal
