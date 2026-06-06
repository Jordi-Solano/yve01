"""
Blueprint Demo Mode — Datos ficticios para presentaciones
"""
from flask import Blueprint, jsonify, session
from demo_mode import generar_datos_demo, exportar_demo_excel

demo_bp = Blueprint('demo', __name__)

@demo_bp.route('/api/demo/status')
def api_demo_status():
    """Retorna si demo mode está activo"""
    demo_active = session.get('demo_mode', False)
    return jsonify({"demo_mode": demo_active})

@demo_bp.route('/api/demo/toggle')
def api_demo_toggle():
    """Activa/desactiva demo mode"""
    current = session.get('demo_mode', False)
    session['demo_mode'] = not current
    return jsonify({"demo_mode": session['demo_mode'], "message": "Demo mode " + ("activado" if not current else "desactivado")})

@demo_bp.route('/api/demo/data')
def api_demo_data():
    """Retorna datos demo"""
    demo_active = session.get('demo_mode', False)
    if not demo_active:
        return jsonify({"error": "Demo mode no activo"}), 403
    
    data = generar_datos_demo()
    return jsonify(data)

@demo_bp.route('/api/demo/export')
def api_demo_export():
    """Exporta datos demo a Excel"""
    try:
        excel_file = exportar_demo_excel()
        with open(excel_file, 'rb') as f:
            from io import BytesIO
            output = BytesIO(f.read())
        
        import os
        filename = os.path.basename(excel_file)
        
        from flask import send_file
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
