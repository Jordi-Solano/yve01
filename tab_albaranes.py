# -*- coding: utf-8 -*-
"""tab_albaranes.py — pantalla de albaranes (ronda de pruebas de Jordi, punto 5).

Hasta ahora el albaran se guardaba "para cruces internos" y no habia donde
verlo. Esto lo lista con sus lineas y con que factura ha cruzado. Solo lectura:
lee de almacen_datos (cabeceras + lineas, ya consolidadas por hotel) y no
escribe nada.
"""
from io import BytesIO

from flask import Blueprint, jsonify, request, send_file

from tenant_dirs import procesadas_dir as _t_pdir, reportes_dir as _t_rdir

albaranes_bp = Blueprint('albaranes', __name__)


def _txt(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat", "no_encontrado") else s


def _num(v):
    try:
        import pandas as pd
        x = pd.to_numeric(v, errors="coerce")
        return None if x != x else round(float(x), 2)
    except Exception:
        return None


def listar(hotel=None):
    import almacen_datos as ALM
    p, r = str(_t_pdir()), str(_t_rdir())
    try:
        cab = ALM.albaranes(p, r, hotel=hotel or None)
    except Exception:
        cab = None
    try:
        lin = ALM.lineas_albaran(p, r, hotel=hotel or None)
    except Exception:
        lin = None
    por_clave = {}
    if lin is not None and not lin.empty and "clave" in lin.columns:
        for _, l in lin.iterrows():
            por_clave.setdefault(_txt(l.get("clave")), []).append({
                "n": int(_num(l.get("n_linea")) or 0), "descripcion": _txt(l.get("descripcion")),
                "cantidad": _num(l.get("cantidad")), "unidad": _txt(l.get("unidad")),
                "precio_unitario": _num(l.get("precio_unitario")), "importe": _num(l.get("importe")),
            })
    out = []
    if cab is not None and not cab.empty:
        for _, a in cab.iterrows():
            clave = _txt(a.get("clave"))
            lineas = sorted(por_clave.get(clave, []), key=lambda x: x["n"])
            estado = _txt(a.get("estado")).upper() or "SIN_CRUZAR"
            out.append({
                "clave": clave, "numero_albaran": _txt(a.get("numero_albaran")) or "s/n",
                "proveedor": _txt(a.get("nombre_proveedor")), "fecha_entrega": _txt(a.get("fecha_entrega")),
                "total": _num(a.get("total_albaran")), "referencia_pedido": _txt(a.get("referencia_pedido")),
                "referencia_factura": _txt(a.get("referencia_factura")),
                "estado": estado, "numero_factura": _txt(a.get("numero_factura")),
                "detalle": _txt(a.get("detalle")), "archivo": _txt(a.get("archivo")),
                "hotel_id": _txt(a.get("hotel_id")), "n_lineas": len(lineas), "lineas": lineas,
            })
    out.sort(key=lambda x: (x["fecha_entrega"] == "", x["fecha_entrega"], x["numero_albaran"]), reverse=False)
    resumen = {"n": len(out),
               "facturados": sum(1 for x in out if x["estado"] == "ALBARAN_FACTURADO"),
               "sin_facturar": sum(1 for x in out if x["estado"] == "ALBARAN_SIN_FACTURAR"),
               "sin_cruzar": sum(1 for x in out if x["estado"] not in ("ALBARAN_FACTURADO", "ALBARAN_SIN_FACTURAR")),
               "total": round(sum(x["total"] or 0 for x in out), 2)}
    return {"albaranes": out, "resumen": resumen}


@albaranes_bp.route('/api/albaranes')
def api_albaranes():
    try:
        import censo_hoteles
        hotel = censo_hoteles.activo() or None
    except Exception:
        hotel = None
    try:
        return jsonify({'ok': True, 'hotel': hotel or '', **listar(hotel)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500


@albaranes_bp.route('/api/exportar/albaranes')
def api_exportar_albaranes():
    import pandas as pd
    try:
        import censo_hoteles
        hotel = censo_hoteles.activo() or None
    except Exception:
        hotel = None
    d = listar(hotel)
    cab = [{k: v for k, v in a.items() if k != "lineas"} for a in d["albaranes"]]
    lin = [{"numero_albaran": a["numero_albaran"], "proveedor": a["proveedor"], **l} for a in d["albaranes"] for l in a["lineas"]]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(cab or [{}]).to_excel(w, index=False, sheet_name="Albaranes")
        pd.DataFrame(lin or [{}]).to_excel(w, index=False, sheet_name="Lineas")
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="albaranes.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
