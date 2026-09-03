"""
tab_cierre.py — Yve.01 · Cierre de mes (OLA B)
Endpoints del apartado "Cierre": asientos del mes, reconciliacion de cuentas
y su Excel. Solo lectura: nada de aqui escribe en los datos.
"""
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
