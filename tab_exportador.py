"""
Blueprint Flask para descarga de reportes
"""
from flask import Blueprint, send_file, jsonify
from exportador_reportes import exportar_excel, crear_reporte_calipolis_excel
from ar_real_completo import procesar_ar_real_completo

exportador_bp = Blueprint('exportador', __name__)

@exportador_bp.route('/api/exportar/<tipo>')
def api_exportar(tipo):
    """Descarga reporte en Excel"""
    valid_tipos = ['ar', 'ap', 'drr', 'multihotel', 'banco', 'fb', 'ar_real']
    
    if tipo not in valid_tipos and tipo != 'ar_real' and tipo != 'calipolis':
        return jsonify({"error": "Tipo de reporte inválido"}), 400
    
    try:
        if tipo == 'ar_real':
            output_file = procesar_ar_real_completo()
            if not output_file:
                return jsonify({"error": "No se pudo procesar AR Real"}), 500
            output = open(output_file, 'rb').read()
            from io import BytesIO
            output = BytesIO(output)
            import os
            filename = os.path.basename(output_file)
            result = (output, filename)
        elif tipo == 'banco':
            import glob as _g, os as _os
            from datetime import datetime as _dt
            from tenant_dirs import reportes_dir as _t_rdir
            REPORTES_DIR = _t_rdir()
            hits = _g.glob(_os.path.join(REPORTES_DIR, 'conciliacion_*.xlsx'))
            if not hits: return jsonify({'error': 'No hay datos bancarios'}), 404
            hits.sort(key=lambda p: _os.path.getmtime(p), reverse=True)
            from io import BytesIO
            result = (open(hits[0],'rb'), _os.path.basename(hits[0]))
        elif tipo == 'ar_real':
            import os as _os
            ruta = __import__('tenant_dirs').datos_dir() + '/reservas_credito.xlsx'
            if not _os.path.exists(ruta): return jsonify({'error': 'Sin datos AR Real'}), 404
            from io import BytesIO as _BIO
            with open(ruta,'rb') as fh: data = fh.read()
            from datetime import datetime as _dt
            result = (_BIO(data), f'ar_real_facturas_{_dt.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'fb':
            import os as _os
            ruta = __import__('tenant_dirs').datos_dir() + '/ventas_fb_diarias.xlsx'
            if not _os.path.exists(ruta): return jsonify({'error': 'Sin datos F&B'}), 404
            from io import BytesIO
            with open(ruta,'rb') as fh: data = fh.read()
            from datetime import datetime as _dt
            result = (BytesIO(data), f'fb_ventas_{_dt.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'calipolis':
            wb = crear_reporte_calipolis_excel()
            from io import BytesIO
            from datetime import datetime
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            filename = f"Calipolis_Consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            result = (output, filename)
        else:
            result = exportar_excel(tipo)
        if not result:
            return jsonify({"error": "Error generando reporte"}), 500
        
        output, filename = result
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
