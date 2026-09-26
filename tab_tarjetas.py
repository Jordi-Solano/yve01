"""
tab_tarjetas.py — Yve.01 · pestaña Tarjetas (b109, finanzas 26 sep 2026).

La liquidacion de tarjetas entra por ⚡ Procesar archivos (formato generico: fecha, bruto,
comision, neto, referencia; ver tarjetas.py). Aqui solo se consulta, se descarga y se quita.

  GET    /api/tarjetas?mes=YYYY-MM       operaciones del mes, totales, cruce con el extracto y
                                         con lo cobrado con tarjeta segun DRR/TPV
  GET    /api/tarjetas/plantilla         Excel del formato generico (con instrucciones)
  GET    /api/exportar/tarjetas?mes=     Excel del cruce
  POST   /api/tarjetas/quitar            {fichero}  quita las operaciones que entraron con ese fichero
"""
import os

import pandas as pd
from flask import Blueprint, jsonify, request, send_file
from flask_login import login_required, current_user

from tenant_dirs import datos_dir as _t_ddir, reportes_dir as _t_rdir

tarjetas_bp = Blueprint('tarjetas', __name__)


def _usuario():
    try:
        return getattr(current_user, "username", None) or getattr(current_user, "id", "") or ""
    except Exception:
        return ""


def _hotel():
    try:
        import censo_hoteles
        return censo_hoteles.activo() or ""
    except Exception:
        return ""


def _mes():
    m = (request.args.get('mes') or '').strip()[:7]
    if len(m) == 7 and m[4] == '-':
        return m
    from datetime import date
    return date.today().strftime('%Y-%m')


def _banco(hotel):
    import almacen_datos as ALM
    try:
        bk, _ = ALM.movimientos_banco(datos_dir=str(_t_ddir()), reportes_dir=str(_t_rdir()))
        return ALM.banco_del_hotel(bk, hotel or None)
    except Exception:
        return pd.DataFrame()


def _ventas(hotel):
    import almacen_datos as ALM
    try:
        df = pd.read_excel(os.path.join(str(_t_ddir()), 'ventas_fb_diarias.xlsx'))
        return ALM._filtrar_hotel(df, hotel) if hotel else df
    except Exception:
        return pd.DataFrame()


def _drr_dia():
    import tarjetas as TJ
    try:
        import dashboard as D
        return TJ.drr_por_dia(D.drr_del_hotel(str(_t_rdir())))
    except Exception:
        return {}


def resumen(mes, hotel=None):
    import tarjetas as TJ
    import cuadre_banco as CB
    dd = str(_t_ddir())
    hotel = hotel if hotel is not None else _hotel()
    return TJ.resumen_mes(mes, TJ.leer_store(dd), hotel, _banco(hotel), _drr_dia(), TJ.tpv_por_dia(_ventas(hotel)),
                          CB.palabras(dd), CB.manuales(dd), CB.proveedores_conocidos(dd))


@tarjetas_bp.route('/api/tarjetas')
@login_required
def api_tarjetas():
    try:
        res = resumen(_mes())
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    return jsonify({'ok': True, 'hotel': _hotel(), **res})


@tarjetas_bp.route('/api/tarjetas/plantilla')
@login_required
def api_tarjetas_plantilla():
    import tarjetas as TJ
    buf, nombre = TJ.plantilla_excel()
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@tarjetas_bp.route('/api/exportar/tarjetas')
@login_required
def api_exportar_tarjetas():
    import tarjetas as TJ
    try:
        buf, nombre = TJ.exportar_excel(resumen(_mes()))
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@tarjetas_bp.route('/api/tarjetas/quitar', methods=['POST'])
@login_required
def api_tarjetas_quitar():
    import tarjetas as TJ
    data = request.get_json(force=True, silent=True) or {}
    fichero = str(data.get('fichero') or '').strip()
    if not fichero:
        return jsonify({'ok': False, 'error': 'falta el fichero'}), 400
    try:
        n = TJ.quitar_fichero(fichero, _hotel(), str(_t_ddir()))
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    if not n:
        return jsonify({'ok': False, 'error': 'no hay operaciones de ese fichero'}), 404
    try:
        from dashboard import _audit
        _audit('TARJETAS_QUITAR', f"{fichero}: {n} operaciones", _usuario() or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True, 'quitadas': n})
