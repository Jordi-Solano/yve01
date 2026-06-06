"""
blog.py — Yve Blog SEO
Rutas: /blog, /blog/<slug>
Artículos en memoria, optimizados para SEO hotelero en España
"""
from flask import Blueprint, Response

blog_bp = Blueprint('blog', __name__)

POSTS = [
  {
    "slug": "software-gestion-financiera-hoteles-espana",
    "title": "Software de gestión financiera para hoteles en España: guía 2026",
    "desc": "Comparativa completa de las mejores herramientas para automatizar AP, AR y conciliación bancaria en hoteles independientes y cadenas medianas.",
    "date": "2026-05-15",
    "cat": "Gestión hotelera",
    "read": "8 min",
    "body": """
<h2>¿Por qué los hoteles necesitan software financiero específico?</h2>
<p>Los hoteles tienen una estructura financiera única: ingresos de múltiples canales (PMS, POS de restaurante, OTAs), proveedores de todo tipo y un volumen de facturas que puede superar las 150 al mes para un hotel de 100 habitaciones. Los ERPs genéricos no están diseñados para esto.</p>
<h2>Los 5 procesos que más tiempo consumen</h2>
<p><strong>1. Verificación de comisiones OTA.</strong> Cada factura de Booking.com o Expedia tiene que verificarse manualmente contra las tarifas pactadas. Un error en el porcentaje de comisión puede costar miles de euros al año sin que nadie lo detecte.</p>
<p><strong>2. 3-way matching en F&B.</strong> Para los hoteles con restaurante, cada compra requiere cruzar la factura del proveedor, el albarán físico y los datos del POS. Hacerlo manualmente son 3-4 horas diarias.</p>
<p><strong>3. Contabilización en Oracle.</strong> Introducir asientos contables en Oracle Fusion manualmente es lento y propenso a errores. La automatización puede reducir este tiempo un 90%.</p>
<p><strong>4. Daily Revenue Report.</strong> El Income Auditor dedica entre 1 y 2 horas cada mañana a revisar el DRR. Un sistema que detecte automáticamente los Out of Balance libera ese tiempo.</p>
<p><strong>5. Conciliación bancaria.</strong> Cruzar el extracto bancario con las facturas pagadas a mano puede llevar medio día a la semana.</p>
<h2>Qué buscar en un software financiero hotelero</h2>
<p>Las características clave son: integración con Oracle GL, lectura automática de facturas PDF, soporte para el Plan General Contable español, y capacidad multi-hotel para grupos.</p>
<h2>Yve.01: la alternativa AI-first para hoteles europeos</h2>
<p>A diferencia de las soluciones legacy como M3 o Aptech, Yve está diseñado desde cero con IA en el núcleo. El setup tarda 15 minutos y el precio comienza en 400€/mes, accesible para hoteles independientes.</p>
""",
  },
  {
    "slug": "automatizar-cuentas-pagar-hotel",
    "title": "Cómo automatizar las cuentas a pagar (AP) en un hotel",
    "desc": "Guía práctica para automatizar el proceso AP hotelero: desde la recepción de facturas hasta la contabilización en Oracle, con ejemplos reales.",
    "date": "2026-05-28",
    "cat": "Cuentas a pagar",
    "read": "6 min",
    "body": """
<h2>El flujo AP en un hotel: el problema</h2>
<p>Un hotel de tamaño medio recibe facturas por tres canales: email, correo físico y portales de proveedores. Sin un sistema centralizado, el equipo financiero pierde horas buscando qué facturas faltan y cuáles están duplicadas.</p>
<h2>El proceso AP estándar en hotelería</h2>
<p>El camino correcto es: recepción → verificación → matching con PO → aprobación del jefe de departamento → contabilización en Oracle → pago.</p>
<p>Para facturas F&B se añade un paso crítico: el <strong>3-way matching</strong>, que cruza la factura con el albarán físico sellado y los datos del POS. Este proceso manual puede llevar 2-3 horas al día en un hotel con restaurante activo.</p>
<h2>La automatización con IA</h2>
<p>Los sistemas modernos como Yve utilizan OCR e IA para leer PDFs de facturas, extraer los campos clave (NIF, importe, fecha, número de factura) y cruzarlos automáticamente con las órdenes de compra. Las discrepancias se detectan y se generan emails de reclamación automáticamente.</p>
<h2>Integración con Oracle GL</h2>
<p>El último paso, contabilizar en Oracle, puede automatizarse completamente via la API REST de Oracle Fusion Cloud. Cada factura aprobada genera automáticamente el asiento con las cuentas correctas del PGC español.</p>
""",
  },
  {
    "slug": "revenue-management-hoteles-independientes",
    "title": "Revenue management para hoteles independientes: lo que nadie te explica",
    "desc": "Los hoteles independientes pueden competir con las grandes cadenas en revenue management sin invertir en software caro. Aquí está cómo.",
    "date": "2026-06-01",
    "cat": "Revenue Management",
    "read": "7 min",
    "body": """
<h2>El mito del revenue management caro</h2>
<p>Muchos hoteleros independientes asumen que el revenue management avanzado es solo para grandes cadenas con presupuestos millonarios. No es verdad. Las métricas clave — ADR, RevPAR, Occupancy% — son accesibles y medibles para cualquier hotel.</p>
<h2>Las métricas que realmente importan</h2>
<p><strong>ADR (Average Daily Rate)</strong>: el precio medio por habitación ocupada. Es tu palanca más directa de revenue.</p>
<p><strong>RevPAR (Revenue Per Available Room)</strong>: combina ocupación y ADR. Es la métrica que usan los inversores.</p>
<p><strong>GOP% (Gross Operating Profit)</strong>: el margen operativo real del hotel. Un hotel 4★ bien gestionado debería estar entre el 30-45%.</p>
<h2>El Daily Revenue Report: tu brújula diaria</h2>
<p>El DRR es el informe que resume todos los ingresos del día anterior por departamento. Un Income Auditor lo revisa cada mañana para detectar desviaciones. Automatizar esta revisión es el primer paso hacia un revenue management eficiente.</p>
<h2>Temporada alta en la Costa Dorada: datos reales</h2>
<p>En junio 2026, los hoteles de Sitges están viendo ocupaciones del 83-95% con ADR entre 180€ y 315€ según categoría. El RevPAR medio del grupo Calipolis supera los 170€, por encima de la media provincial.</p>
""",
  },
]

def post_html(post):
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="{post['desc']}">
<meta property="og:title" content="{post['title']} | Yve Blog">
<meta property="og:description" content="{post['desc']}">
<meta property="og:type" content="article">
<title>{post['title']} | Yve Blog</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;-webkit-font-smoothing:antialiased}}
.nav{{background:rgba(15,23,42,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--s2);padding:0 5%;height:60px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}}
.nav-logo{{font-size:18px;font-weight:800;color:#fff}}.nav-logo span{{color:var(--acc2)}}
.nav a{{font-size:14px;color:var(--mut);transition:color .15s}}.nav a:hover{{color:var(--tx)}}
.article{{max-width:720px;margin:0 auto;padding:64px 24px 96px}}
.breadcrumb{{font-size:13px;color:var(--dim);margin-bottom:28px}}
.breadcrumb a{{color:var(--acc2);text-decoration:none}}
.cat-badge{{display:inline-block;background:rgba(59,130,246,.1);color:var(--acc2);border:1px solid rgba(59,130,246,.2);border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;margin-bottom:16px}}
h1{{font-size:clamp(26px,4vw,42px);font-weight:800;letter-spacing:-1px;line-height:1.15;margin-bottom:16px}}
.meta{{font-size:13px;color:var(--dim);margin-bottom:40px;padding-bottom:32px;border-bottom:1px solid var(--s2)}}
.content h2{{font-size:22px;font-weight:700;margin:36px 0 14px;color:var(--tx)}}
.content p{{font-size:17px;color:var(--mut);line-height:1.8;margin-bottom:18px}}
.content strong{{color:var(--tx);font-weight:600}}
.cta-box{{background:var(--s1);border:1px solid rgba(59,130,246,.2);border-radius:16px;padding:32px;text-align:center;margin-top:56px}}
.cta-box h3{{font-size:22px;font-weight:700;margin-bottom:10px}}
.cta-box p{{color:var(--mut);margin-bottom:24px}}
.btn{{display:inline-block;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:12px 28px;border-radius:10px;font-size:15px;font-weight:700;text-decoration:none;box-shadow:0 4px 20px rgba(59,130,246,.35)}}
footer{{background:var(--s1);border-top:1px solid var(--s2);padding:32px 5%;text-align:center;font-size:13px;color:var(--dim);margin-top:64px}}
</style>
</head>
<body>
<nav class="nav">
  <a href="/" class="nav-logo" style="text-decoration:none">Yve<span>.01</span></a>
  <div style="display:flex;gap:24px">
    <a href="/blog">Blog</a>
    <a href="/#pricing">Precios</a>
    <a href="/login" style="color:var(--acc2)">Acceder →</a>
  </div>
</nav>
<article class="article">
  <div class="breadcrumb"><a href="/">Inicio</a> / <a href="/blog">Blog</a> / {post['cat']}</div>
  <span class="cat-badge">{post['cat']}</span>
  <h1>{post['title']}</h1>
  <div class="meta">{post['date']} &nbsp;·&nbsp; {post['read']} de lectura</div>
  <div class="content">{post['body']}</div>
  <div class="cta-box">
    <h3>¿Quieres automatizar tu hotel?</h3>
    <p>14 días gratis. Sin tarjeta. Setup en 15 minutos.</p>
    <a href="/login" class="btn">Empezar gratis →</a>
  </div>
</article>
<footer>© 2026 Yve.01 · Barcelona · <a href="/blog" style="color:var(--acc2)">Blog</a> · <a href="/" style="color:var(--acc2)">Inicio</a></footer>
</body></html>"""

def index_html(posts):
    cards = ""
    for p in posts:
        cards += f"""
<a href="/blog/{p['slug']}" style="text-decoration:none;display:block;background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px;transition:border-color .2s" onmouseover="this.style.borderColor='rgba(59,130,246,.4)'" onmouseout="this.style.borderColor='var(--s2)'">
  <span style="display:inline-block;background:rgba(59,130,246,.1);color:var(--acc2);border:1px solid rgba(59,130,246,.2);border-radius:20px;padding:3px 10px;font-size:11px;font-weight:600;margin-bottom:12px">{p['cat']}</span>
  <h2 style="font-size:20px;font-weight:700;color:var(--tx);margin-bottom:10px;line-height:1.3">{p['title']}</h2>
  <p style="font-size:14px;color:var(--mut);line-height:1.6;margin-bottom:14px">{p['desc']}</p>
  <span style="font-size:12px;color:var(--dim)">{p['date']} · {p['read']}</span>
</a>"""
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Blog de Yve.01 — artículos sobre gestión financiera hotelera, automatización AP/AR, revenue management y tecnología para hoteles en España.">
<title>Blog | Yve.01 — Gestión financiera para hoteles</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;-webkit-font-smoothing:antialiased}}
.nav{{background:rgba(15,23,42,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--s2);padding:0 5%;height:60px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}}
.nav-logo{{font-size:18px;font-weight:800;color:#fff;text-decoration:none}}.nav-logo span{{color:var(--acc2)}}
.nav a{{font-size:14px;color:var(--mut);text-decoration:none;transition:color .15s}}.nav a:hover{{color:var(--tx)}}
.hero{{padding:72px 5% 56px;text-align:center}}
.hero h1{{font-size:clamp(28px,4vw,44px);font-weight:800;letter-spacing:-1px;margin-bottom:14px}}
.hero p{{font-size:17px;color:var(--mut);max-width:540px;margin:0 auto}}
.posts{{max-width:860px;margin:0 auto;padding:0 5% 96px;display:grid;grid-template-columns:1fr;gap:16px}}
footer{{background:var(--s1);border-top:1px solid var(--s2);padding:32px 5%;text-align:center;font-size:13px;color:var(--dim)}}
</style>
</head>
<body>
<nav class="nav">
  <a href="/" class="nav-logo">Yve<span>.01</span></a>
  <div style="display:flex;gap:24px">
    <a href="/blog" style="color:var(--tx)">Blog</a>
    <a href="/#pricing">Precios</a>
    <a href="/login" style="color:var(--acc2)">Acceder →</a>
  </div>
</nav>
<div class="hero">
  <h1>Blog de Yve</h1>
  <p>Gestión financiera hotelera, automatización AP/AR y tecnología para hoteles en España y Europa.</p>
</div>
<div class="posts">{cards}</div>
<footer>© 2026 Yve.01 · <a href="/" style="color:var(--acc2)">Inicio</a></footer>
</body></html>"""

@blog_bp.route('/blog')
def blog_index():
    return Response(index_html(POSTS), mimetype='text/html')

@blog_bp.route('/blog/<slug>')
def blog_post(slug):
    post = next((p for p in POSTS if p['slug'] == slug), None)
    if not post:
        return Response('<h1>Post no encontrado</h1>', status=404, mimetype='text/html')
    return Response(post_html(post), mimetype='text/html')
