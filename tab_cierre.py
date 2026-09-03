"""
tab_cierre.py — Yve.01 · Cierre de mes (OLA B)
Endpoints del apartado "Cierre": asientos del mes, reconciliacion de cuentas
y su Excel. Solo lectura: nada de aqui escribe en los datos.
"""
import os
from flask import Blueprint, jsonify, request, send_file

from tenant_dirs import datos_dir as _t_ddir, procesadas_dir as _t_pdir, reportes_dir as _t_rdir

cierre_bp = Blueprint('cierre', __name__)


def _args():
    mes = (request.args.get('mes') or '').strip()[:7] or None
    try:
        import censo_hoteles
        hotel = censo_hoteles.activo() or None
    except Exception:
        hotel = None
    return mes, hotel


def _dirs():
    return {"procesadas_dir": str(_t_pdir()), "reportes_dir": str(_t_rdir()), "datos_dir": str(_t_ddir())}


@cierre_bp.route('/api/cierre/asientos')
def api_cierre_asientos():
    import cierre_mes as CM
    mes, hotel = _args()
    try:
        res, rec = CM.cierre_completo(mes, hotel, **_dirs())
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    limite = int(request.args.get('limite') or 400)
    return jsonify({'ok': True, 'hotel': hotel or '', **{k: v for k, v in res.items() if k != 'asientos'},
                    'asientos': res['asientos'][:limite], 'truncado': len(res['asientos']) > limite,
                    'mayor': rec['mayor'], 'reconciliacion': {k: v for k, v in rec.items() if k != 'mayor'}})


@cierre_bp.route('/api/exportar/cierre')
def api_exportar_cierre():
    import cierre_mes as CM
    mes, hotel = _args()
    res, rec = CM.cierre_completo(mes, hotel, **_dirs())
    buf, nombre = CM.exportar_excel(res, rec)
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Cuadre de banco por pestañas (Ola B · bloque 2) ─────────────────────────
def _cuadre(mes, hotel):
    import cuadre_banco as CB
    import almacen_datos as ALM
    import pandas as pd
    dd = str(_t_ddir())
    try:
        bk, _ = ALM.movimientos_banco(datos_dir=dd, reportes_dir=str(_t_rdir()))
        if hotel:
            bk = ALM._filtrar_hotel(bk, hotel)
    except Exception:
        bk = pd.DataFrame()
    try:
        vf = pd.read_excel(os.path.join(dd, 'ventas_fb_diarias.xlsx'))
        if hotel:
            vf = ALM._filtrar_hotel(vf, hotel)
    except Exception:
        vf = pd.DataFrame()
    return CB.cuadrar(mes, bk, vf, CB.palabras(dd), CB.manuales(dd), CB.proveedores_conocidos(dd))


@cierre_bp.route('/api/cuadre_banco')
def api_cuadre_banco():
    mes, hotel = _args()
    try:
        res = _cuadre(mes, hotel)
    except Exception as e:
        return jsonify({'ok_api': False, 'error': str(e)[:200]}), 500
    return jsonify({'ok_api': True, 'hotel': hotel or '', **res})


@cierre_bp.route('/api/cuadre_banco/asignar', methods=['POST'])
def api_cuadre_banco_asignar():
    """Pone un movimiento en una pestaña a mano (clave = clave_movimiento)."""
    import cuadre_banco as CB
    data = request.get_json(force=True, silent=True) or {}
    clave = str(data.get('clave') or '').strip()
    pestana = str(data.get('pestana') or '').strip().upper()
    if not clave:
        return jsonify({'ok': False, 'error': 'falta clave'}), 400
    if pestana and pestana not in CB.PESTANAS:
        return jsonify({'ok': False, 'error': 'pestaña desconocida'}), 400
    CB.guardar_manual(clave, pestana if pestana != 'SIN_CLASIFICAR' else '', str(_t_ddir()))
    try:
        from dashboard import _audit
        from flask_login import current_user
        _audit('CUADRE_BANCO_ASIGNAR', f'{clave} -> {pestana or "(quitar)"}', getattr(current_user, 'username', None) or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True})


@cierre_bp.route('/api/exportar/cuadre_banco')
def api_exportar_cuadre_banco():
    import cuadre_banco as CB
    mes, hotel = _args()
    buf, nombre = CB.exportar_excel(_cuadre(mes, hotel))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
