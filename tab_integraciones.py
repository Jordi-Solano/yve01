"""
Blueprint para integraciones (Slack, WhatsApp, Email)
"""
from flask import Blueprint, jsonify, request
from integraciones_externas import AutomatizadorAlertas

integraciones_bp = Blueprint('integraciones', __name__)

automatizador = AutomatizadorAlertas()

@integraciones_bp.route('/api/integraciones/status')
def api_status_integraciones():
    """Retorna estado de integraciones disponibles"""
    return jsonify({
        "slack_disponible": automatizador.slack.enabled,
        "whatsapp_disponible": automatizador.whatsapp.enabled,
        "email_disponible": automatizador.email.enabled,
        "mensaje": "Configura variables de entorno para habilitar integraciones"
    })

@integraciones_bp.route('/api/integraciones/enviar-alerta-critica', methods=['POST'])
def api_enviar_alerta_critica():
    """Envía alerta crítica a todos los canales"""
    data = request.json
    resultados = automatizador.enviar_alerta_critica(
        data.get("titulo", "Alerta crítica"),
        data.get("detalles", {})
    )
    return jsonify({"resultados": resultados})

@integraciones_bp.route('/api/integraciones/test-slack', methods=['POST'])
def api_test_slack():
    """Test de conexión a Slack"""
    if not automatizador.slack.enabled:
        return jsonify({"error": "Slack no configurado"}), 403
    
    result = automatizador.slack.enviar_alerta(
        "Test YVE",
        "Esto es un mensaje de prueba desde YVE",
        "info",
        "Test"
    )
    return jsonify({"success": result})
