"""
Integraciones externas: Slack, WhatsApp, Email
Envía alertas automáticamente
"""
import os
from datetime import datetime

class IntegradorSlack:
    """Envía mensajes a Slack"""
    def __init__(self, webhook_url=None):
        self.webhook_url = webhook_url or os.getenv('SLACK_WEBHOOK_URL')
        self.enabled = bool(self.webhook_url)
    
    def enviar_alerta(self, titulo, mensaje, severidad="info", hotel=""):
        """Envía alerta a Slack"""
        if not self.enabled:
            return False
        
        import requests
        import json
        
        color_map = {"info": "#1a73e8", "warning": "#ff9800", "critical": "#e05252"}
        
        payload = {
            "text": f"🚨 {titulo}",
            "attachments": [{
                "color": color_map.get(severidad, "#1a73e8"),
                "title": titulo,
                "text": mensaje,
                "fields": [
                    {"title": "Hotel", "value": hotel, "short": True},
                    {"title": "Severidad", "value": severidad.upper(), "short": True},
                    {"title": "Hora", "value": datetime.now().strftime("%H:%M:%S"), "short": True}
                ]
            }]
        }
        
        try:
            response = requests.post(self.webhook_url, json=payload)
            return response.status_code == 200
        except:
            return False

class IntegradorWhatsApp:
    """Envía mensajes a WhatsApp (Twilio)"""
    def __init__(self, account_sid=None, auth_token=None, from_number=None):
        self.account_sid = account_sid or os.getenv('TWILIO_ACCOUNT_SID')
        self.auth_token = auth_token or os.getenv('TWILIO_AUTH_TOKEN')
        self.from_number = from_number or os.getenv('TWILIO_WHATSAPP_NUMBER')
        self.enabled = all([self.account_sid, self.auth_token, self.from_number])
    
    def enviar_alerta(self, numero, titulo, mensaje):
        """Envía alerta a WhatsApp"""
        if not self.enabled:
            return False
        
        try:
            from twilio.rest import Client
            client = Client(self.account_sid, self.auth_token)
            msg = client.messages.create(
                from_=f"whatsapp:{self.from_number}",
                to=f"whatsapp:{numero}",
                body=f"🚨 {titulo}\n{mensaje}\n\n{datetime.now().strftime('%H:%M')}"
            )
            return msg.sid is not None
        except:
            return False

class IntegradorEmail:
    """Envía emails (ya existe en notificaciones.py)"""
    def __init__(self, smtp_server=None, smtp_user=None, smtp_password=None):
        self.smtp_server = smtp_server or os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_user = smtp_user or os.getenv('SMTP_USER')
        self.smtp_password = smtp_password or os.getenv('SMTP_PASSWORD')
        self.enabled = bool(self.smtp_user and self.smtp_password)
    
    def enviar_alerta(self, destinatario, asunto, cuerpo):
        """Envía email de alerta"""
        if not self.enabled:
            return False
        
        import smtplib
        from email.mime.text import MIMEText
        
        try:
            mensaje = MIMEText(cuerpo, 'html')
            mensaje['Subject'] = asunto
            mensaje['From'] = self.smtp_user
            mensaje['To'] = destinatario
            
            server = smtplib.SMTP(self.smtp_server, 587)
            server.starttls()
            server.login(self.smtp_user, self.smtp_password)
            server.send_message(mensaje)
            server.quit()
            return True
        except:
            return False

class AutomatizadorAlertas:
    """Orquesta alertas a múltiples canales"""
    def __init__(self):
        self.slack = IntegradorSlack()
        self.whatsapp = IntegradorWhatsApp()
        self.email = IntegradorEmail()
    
    def enviar_alerta_critica(self, titulo, detalles):
        """Envía alerta crítica a todos los canales"""
        resultados = {
            "slack": self.slack.enviar_alerta(titulo, detalles["mensaje"], "critical", detalles.get("hotel", "")),
            "email": self.email.enviar_alerta(
                detalles.get("email", "admin@hotel.es"),
                f"🚨 ALERTA CRÍTICA: {titulo}",
                f"<h2>{titulo}</h2><p>{detalles['mensaje']}</p>"
            )
        }
        return resultados
    
    def enviar_alerta_warning(self, titulo, detalles):
        """Envía alerta de warning"""
        resultados = {
            "slack": self.slack.enviar_alerta(titulo, detalles["mensaje"], "warning", detalles.get("hotel", "")),
            "email": self.email.enviar_alerta(
                detalles.get("email", "auditor@hotel.es"),
                f"⚠️ {titulo}",
                f"<p>{detalles['mensaje']}</p>"
            )
        }
        return resultados
    
    def enviar_reporte_diario(self, reporte_html, destino):
        """Envía reporte diario por email"""
        return self.email.enviar_alerta(
            destino,
            f"📊 Reporte Diario YVE — {datetime.now().strftime('%d/%m/%Y')}",
            reporte_html
        )

if __name__ == "__main__":
    print("INTEGRACIONES EXTERNAS")
    print("=" * 70)
    
    automatizador = AutomatizadorAlertas()
    
    print(f"\nSlack disponible: {automatizador.slack.enabled}")
    print(f"WhatsApp disponible: {automatizador.whatsapp.enabled}")
    print(f"Email disponible: {automatizador.email.enabled}")
    
    print("\n(Configura SLACK_WEBHOOK_URL, TWILIO_ACCOUNT_SID, etc en .env)")
