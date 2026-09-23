"""
tab_caja.py — Yve.01 · pestaña Caja (b86, Jordi 23 sep 2026).

El arqueo de caja se apunta a mano (no hay fichero que lo traiga); los ingresos
de efectivo en el banco los pone Yve leyendo el extracto (pestaña CAJA del cuadre).

  GET    /api/caja?mes=YYYY-MM        resumen del mes: arqueos, ingresos en banco, en caja
  POST   /api/caja/arqueo             {fecha, contado, sistema?, nota?}  guarda o corrige un dia
  POST   /api/caja/arqueo/borrar      {fecha}
  GET    /api/exportar/caja?mes=      Excel (arqueos, ingresos banco, resumen)
"""
import pandas as pd
from flask import Blueprint, jsonify, request, send_file
from flask_login import login_required, current_user

from tenant_dirs import datos_dir as _t_ddir, reportes_dir as _t_rdir

caja_bp = Blueprint('caja', __name__)


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


def resumen(mes, hotel=None):
    import caja as CJ
    import cuadre_banco as CB
    dd = str(_t_ddir())
    hotel = hotel if hotel is not None else _hotel()
    return CJ.resumen_mes(mes, CJ.leer(dd), _banco(hotel), hotel, CB.palabras(dd), CB.manuales(dd), CB.proveedores_conocidos(dd))


@caja_bp.route('/api/caja')
@login_required
def api_caja():
    try:
        res = resumen(_mes())
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    return jsonify({'ok': True, 'hotel': _hotel(), **res})


@caja_bp.route('/api/caja/arqueo', methods=['POST'])
@login_required
def api_caja_arqueo():
    import caja as CJ
    data = request.get_json(force=True, silent=True) or {}
    try:
        fila = CJ.apuntar_arqueo(data.get('fecha'), data.get('contado'), data.get('sistema') if data.get('sistema') not in ('', None) else None,
                                 data.get('nota') or '', _usuario(), _hotel(), str(_t_ddir()))
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    try:
        from dashboard import _audit
        _audit('CAJA_ARQUEO', f"{fila['fecha']} contado {fila['efectivo_contado']} sistema {fila['efectivo_sistema']}", _usuario() or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True, 'arqueo': fila})


@caja_bp.route('/api/caja/arqueo/borrar', methods=['POST'])
@login_required
def api_caja_borrar():
    import caja as CJ
    data = request.get_json(force=True, silent=True) or {}
    try:
        hecho = CJ.borrar_arqueo(data.get('fecha'), _hotel(), str(_t_ddir()))
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    if not hecho:
        return jsonify({'ok': False, 'error': 'no hay arqueo ese dia'}), 404
    return jsonify({'ok': True})


@caja_bp.route('/api/exportar/caja')
@login_required
def api_exportar_caja():
    import caja as CJ
    try:
        buf, nombre = CJ.exportar_excel(resumen(_mes()))
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
