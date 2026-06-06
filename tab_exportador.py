"""
Blueprint Flask para descarga de reportes
"""
from flask import Blueprint, send_file, jsonify
from exportador_reportes import exportar_excel

exportador_bp = Blueprint('exportador', __name__)

@exportador_bp.route('/api/exportar/<tipo>')
def api_exportar(tipo):
    """Descarga reporte en Excel"""
    valid_tipos = ['ar', 'ap', 'drr', 'multihotel']
    
    if tipo not in valid_tipos:
        return jsonify({"error": "Tipo de reporte inválido"}), 400
    
    try:
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
