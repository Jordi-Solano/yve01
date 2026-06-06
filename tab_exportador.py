"""
Blueprint Flask para descarga de reportes
"""
from flask import Blueprint, send_file, jsonify
from exportador_reportes import exportar_excel, crear_reporte_calipolis_excel
from ar_real_completo import exportar_excel as exportar_ar_real

exportador_bp = Blueprint('exportador', __name__)

@exportador_bp.route('/api/exportar/<tipo>')
def api_exportar(tipo):
    """Descarga reporte en Excel"""
    valid_tipos = ['ar', 'ap', 'drr', 'multihotel']
    
    if tipo not in valid_tipos and tipo != 'ar_real' and tipo != 'calipolis':
        return jsonify({"error": "Tipo de reporte inválido"}), 400
    
    try:
        if tipo == 'ar_real':
            result = exportar_ar_real(tipo)
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
