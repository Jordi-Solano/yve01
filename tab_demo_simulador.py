"""
Blueprint para exponer el simulador de demo en tiempo real
"""
from flask import Blueprint, jsonify, session
from demo_simulador import get_simulador

demo_sim_bp = Blueprint('demo_sim', __name__)

@demo_sim_bp.route('/api/demo/simular/paso')
def api_demo_simular_paso():
    """Simula un paso en el tiempo - genera cambios realistas"""
    if not session.get('demo_mode'):
        return jsonify({"error": "Demo mode no activo"}), 403
    
    sim = get_simulador()
    evento = sim.simular_paso()
    
    return jsonify({
        "evento": evento,
        "data_actual": sim.get_data(),
        "iteracion": sim.iteration
    })

@demo_sim_bp.route('/api/demo/simular/auto')
def api_demo_simular_auto():
    """Inicia simulación automática (SSE)"""
    if not session.get('demo_mode'):
        return jsonify({"error": "Demo mode no activo"}), 403
    
    def generate():
        import time
        sim = get_simulador()
        for i in range(30):  # 30 pasos
            evento = sim.simular_paso()
            yield f"data: {__import__('json').dumps({
                'evento': evento,
                'consolidado': sim.data['consolidado'],
                'iteracion': sim.iteration
            })}\n\n"
            time.sleep(2)  # 2 segundos entre pasos
    
    from flask import Response
    return Response(generate(), mimetype='text/event-stream')

@demo_sim_bp.route('/api/demo/simular/eventos-recientes')
def api_demo_eventos_recientes():
    """Retorna últimos eventos simulados"""
    if not session.get('demo_mode'):
        return jsonify({"error": "Demo mode no activo"}), 403
    
    sim = get_simulador()
    return jsonify({
        "eventos": sim.get_eventos_recientes(15),
        "total_eventos": len(sim.eventos)
    })

@demo_sim_bp.route('/api/demo/simular/reset')
def api_demo_reset():
    """Reinicia la simulación"""
    sim = get_simulador()
    sim.reset()
    return jsonify({"message": "Simulación reiniciada"})
