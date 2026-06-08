"""
legal.py — Yve.01 Legal Pages
Términos de uso, política de privacidad, cookies
"""
from flask import Blueprint, Response

legal_bp = Blueprint("legal", __name__)

_HEAD = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="canonical" href="https://yve01.onrender.com{path}">
<title>{title} | Yve.01</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#0f172a;color:#f1f5f9;font-family:-apple-system,'Inter',sans-serif;line-height:1.7}
nav{background:#0f172a;border-bottom:1px solid #1e293b;padding:16px 5%;display:flex;align-items:center;justify-content:space-between}
.logo{display:flex;align-items:center;gap:8px;text-decoration:none}
.dot{width:9px;height:9px;border-radius:50%;background:#3b82f6;box-shadow:0 0 8px #3b82f6}
.logo-name{font-size:18px;font-weight:800;color:#f1f5f9}
.logo-name span{color:#60a5fa}
.content{max-width:720px;margin:0 auto;padding:60px 24px}
h1{font-size:32px;font-weight:800;margin-bottom:8px;background:linear-gradient(135deg,#f1f5f9,#94a3b8);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.meta{font-size:13px;color:#64748b;margin-bottom:40px}
h2{font-size:18px;font-weight:700;color:#e2e8f0;margin:32px 0 12px;padding-left:12px;border-left:3px solid #3b82f6}
p{color:#94a3b8;margin-bottom:16px;font-size:15px}
ul{color:#94a3b8;padding-left:20px;margin-bottom:16px}
li{margin-bottom:6px;font-size:15px}
a{color:#60a5fa}
.back{display:inline-flex;align-items:center;gap:6px;color:#64748b;text-decoration:none;font-size:13px;margin-top:40px;padding:8px 14px;border:1px solid #334155;border-radius:8px;transition:.15s}
.back:hover{border-color:#475569;color:#94a3b8}
footer{text-align:center;padding:32px;font-size:12px;color:#475569;border-top:1px solid #1e293b;margin-top:40px}
</style>
</head>
<body>
<nav>
  <a href="/" class="logo"><div class="dot"></div><span class="logo-name">Yve<span>.01</span></span></a>
  <a href="/" style="font-size:13px;color:#64748b;text-decoration:none">← Volver</a>
</nav>
"""

@legal_bp.route("/terminos")
def terminos():
    html = _HEAD.format(title="Términos de Uso", path="/terminos") + """
<div class="content">
  <h1>Términos de Uso</h1>
  <p class="meta">Última actualización: 7 de junio de 2026</p>

  <h2>1. Aceptación de los términos</h2>
  <p>Al acceder y usar Yve.01 ("el Servicio"), aceptas estos Términos de Uso. Si no estás de acuerdo, no uses el Servicio.</p>

  <h2>2. Descripción del servicio</h2>
  <p>Yve.01 es una plataforma de gestión financiera para establecimientos hoteleros. Ofrece herramientas para automatizar la gestión de cuentas a pagar (AP), cuentas a cobrar (AR), informes de ingresos diarios (DRR) y control de costes F&B.</p>

  <h2>3. Cuenta de usuario</h2>
  <p>Eres responsable de mantener la confidencialidad de tu contraseña y de todas las actividades realizadas bajo tu cuenta. Notifícanos inmediatamente cualquier uso no autorizado en <a href="mailto:jordi@yve01.com">jordi@yve01.com</a>.</p>

  <h2>4. Datos financieros</h2>
  <ul>
    <li>Los datos que introduces en Yve.01 son de tu exclusiva propiedad.</li>
    <li>Yve.01 no vende, cede ni comparte tus datos financieros con terceros.</li>
    <li>Los datos se transmiten cifrados mediante HTTPS/TLS.</li>
    <li>Recomendamos no subir datos que contengan información personal de huéspedes.</li>
  </ul>

  <h2>5. Uso aceptable</h2>
  <p>Te comprometes a no:</p>
  <ul>
    <li>Usar el servicio para actividades ilegales o fraudulentas</li>
    <li>Intentar acceder a cuentas de otros usuarios</li>
    <li>Realizar ingeniería inversa del software</li>
    <li>Sobrecargar intencionalmente el servicio</li>
  </ul>

  <h2>6. Facturación</h2>
  <p>Las suscripciones se facturan mensualmente mediante Stripe. Puedes cancelar en cualquier momento desde el panel de administración. No hay permanencia mínima. El reembolso de periodos ya pagados queda a criterio de Yve.01.</p>

  <h2>7. Disponibilidad del servicio</h2>
  <p>Nos esforzamos por mantener una disponibilidad del 99.5%. No garantizamos un servicio ininterrumpido. Programaremos mantenimientos preferentemente en horario de baja actividad.</p>

  <h2>8. Limitación de responsabilidad</h2>
  <p>Yve.01 es una herramienta de asistencia. Las decisiones financieras son responsabilidad del usuario. No nos responsabilizamos de pérdidas derivadas del uso o imposibilidad de uso del servicio.</p>

  <h2>9. Modificaciones</h2>
  <p>Podemos actualizar estos términos. Te notificaremos por email con 15 días de antelación ante cambios significativos.</p>

  <h2>10. Ley aplicable</h2>
  <p>Estos términos se rigen por la legislación española. Los conflictos se someterán a los juzgados de Barcelona.</p>

  <a href="/" class="back">← Volver a inicio</a>
</div>
<footer>© 2026 Yve.01 · <a href="/privacidad">Privacidad</a> · <a href="/terminos">Términos</a> · <a href="/cookies">Cookies</a></footer>
</body></html>"""
    return Response(html, mimetype="text/html")


@legal_bp.route("/privacidad")
def privacidad():
    html = _HEAD.format(title="Política de Privacidad", path="/privacidad") + """
<div class="content">
  <h1>Política de Privacidad</h1>
  <p class="meta">Última actualización: 7 de junio de 2026 · Conforme al RGPD (UE 2016/679)</p>

  <h2>1. Responsable del tratamiento</h2>
  <p><strong>Yve.01</strong> · Barcelona, España · Contacto DPO: <a href="mailto:privacidad@yve01.com">privacidad@yve01.com</a></p>

  <h2>2. Datos que recopilamos</h2>
  <ul>
    <li><strong>Datos de cuenta:</strong> nombre, email, nombre del hotel, rol de usuario</li>
    <li><strong>Datos financieros:</strong> facturas, extractos bancarios, informes DRR que subes voluntariamente</li>
    <li><strong>Datos técnicos:</strong> IP de acceso, navegador, timestamps de login (logs de seguridad)</li>
    <li><strong>Datos de pago:</strong> gestionados exclusivamente por Stripe. Yve.01 no almacena datos de tarjetas.</li>
  </ul>

  <h2>3. Finalidad y base legal</h2>
  <ul>
    <li>Prestación del servicio contratado (Art. 6.1.b RGPD — ejecución de contrato)</li>
    <li>Cumplimiento de obligaciones legales contables y fiscales (Art. 6.1.c)</li>
    <li>Interés legítimo en la seguridad del servicio (Art. 6.1.f)</li>
    <li>Comunicaciones de marketing solo con consentimiento explícito (Art. 6.1.a)</li>
  </ul>

  <h2>4. Conservación de datos</h2>
  <p>Los datos de cuenta se conservan mientras tengas contrato activo + 3 años. Los datos financieros pueden eliminarse en cualquier momento desde el panel. Los logs de seguridad se conservan 12 meses.</p>

  <h2>5. Destinatarios</h2>
  <p>Tus datos pueden compartirse con:</p>
  <ul>
    <li><strong>Stripe</strong> (procesamiento de pagos) — EEUU, transferencia bajo cláusulas estándar UE</li>
    <li><strong>Render</strong> (alojamiento del servidor) — EEUU, transferencia bajo cláusulas estándar UE</li>
    <li><strong>Autoridades</strong> cuando sea legalmente obligatorio</li>
  </ul>
  <p>No vendemos ni cedemos datos a terceros para fines comerciales.</p>

  <h2>6. Tus derechos</h2>
  <p>Tienes derecho a acceder, rectificar, suprimir, oponerte al tratamiento y solicitar la portabilidad de tus datos. Contacta: <a href="mailto:privacidad@yve01.com">privacidad@yve01.com</a>. Puedes reclamar ante la AEPD (aepd.es).</p>

  <h2>7. Seguridad</h2>
  <p>Aplicamos cifrado TLS en tránsito, contraseñas hasheadas (bcrypt), y acceso restringido a datos por rol. Realizamos revisiones de seguridad periódicas.</p>

  <a href="/" class="back">← Volver a inicio</a>
</div>
<footer>© 2026 Yve.01 · <a href="/privacidad">Privacidad</a> · <a href="/terminos">Términos</a> · <a href="/cookies">Cookies</a></footer>
</body></html>"""
    return Response(html, mimetype="text/html")


@legal_bp.route("/cookies")
def cookies():
    html = _HEAD.format(title="Política de Cookies", path="/cookies") + """
<div class="content">
  <h1>Política de Cookies</h1>
  <p class="meta">Última actualización: 7 de junio de 2026</p>

  <h2>Qué son las cookies</h2>
  <p>Las cookies son pequeños archivos de texto que los sitios web guardan en tu navegador para recordar preferencias y mejorar la experiencia.</p>

  <h2>Cookies que usamos</h2>
  <ul>
    <li><strong>session</strong> — Cookie de sesión Flask. Imprescindible para el login. Desaparece al cerrar el navegador.</li>
    <li><strong>yve_lang</strong> — Guarda tu preferencia de idioma (localStorage, no cookie). Persiste entre sesiones.</li>
  </ul>

  <h2>Cookies de terceros</h2>
  <ul>
    <li><strong>Stripe</strong> — Si usas el checkout de pago, Stripe puede instalar cookies para prevención de fraude. Consulta su <a href="https://stripe.com/es/privacy" target="_blank">política de privacidad</a>.</li>
  </ul>

  <h2>Gestión de cookies</h2>
  <p>Puedes eliminar las cookies en cualquier momento desde la configuración de tu navegador. Desactivar la cookie de sesión impedirá el acceso al panel.</p>

  <p>Yve.01 <strong>no usa cookies de seguimiento, analítica ni publicidad</strong>.</p>

  <a href="/" class="back">← Volver a inicio</a>
</div>
<footer>© 2026 Yve.01 · <a href="/privacidad">Privacidad</a> · <a href="/terminos">Términos</a> · <a href="/cookies">Cookies</a></footer>
</body></html>"""
    return Response(html, mimetype="text/html")
