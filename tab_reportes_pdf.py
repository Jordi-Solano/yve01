"""
Blueprint para reportes PDF automáticos
"""
from flask import Blueprint, jsonify, send_file
from generador_reportes_pdf import generar_reporte_diario, generar_reporte_semanal, generar_reporte_mensual
import os

reportes_pdf_bp = Blueprint('reportes_pdf', __name__)

@reportes_pdf_bp.route('/api/reportes/diario')
def api_reporte_diario():
    """Genera y descarga reporte diario"""
    try:
        pdf_file = generar_reporte_diario()
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=os.path.basename(pdf_file)
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@reportes_pdf_bp.route('/api/reportes/semanal')
def api_reporte_semanal():
    """Genera y descarga reporte semanal"""
    try:
        pdf_file = generar_reporte_semanal()
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=os.path.basename(pdf_file)
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@reportes_pdf_bp.route('/api/reportes/mensual')
def api_reporte_mensual():
    """Genera y descarga reporte mensual"""
    try:
        pdf_file = generar_reporte_mensual()
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=os.path.basename(pdf_file)
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
