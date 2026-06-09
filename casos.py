"""
casos.py — Yve.01 Casos de Éxito
"""
from flask import Blueprint, Response

casos_bp = Blueprint("casos", __name__)

@casos_bp.route("/casos")
def casos():
    return Response(_HTML, mimetype="text/html")

_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Casos de éxito de Yve.01 en hoteles reales. Cómo el Grupo Calipolis redujo su tiempo de cierre financiero en 6 horas semanales.">
<meta property="og:title" content="Casos de éxito | Yve.01">
<meta property="og:type" content="website">
<link rel="canonical" href="https://yve01.onrender.com/casos">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%230f172a'/%3E%3Cpath d='M10 6 L16 16 L22 6' stroke='%233b82f6' stroke-width='2.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3Cline x1='16' y1='16' x2='16' y2='26' stroke='%2360a5fa' stroke-width='2.5' stroke-linecap='round'/%3E%3C/svg%3E">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<title>Casos de éxito | Yve.01</title>
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebPage","name":"Casos de éxito Yve.01","description":"Resultados reales de hoteles usando Yve.01 para automatizar su departamento financiero."}</script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--tx:#f1f5f9;--mut:#94a3b8;--acc:#3b82f6;--grn:#22c55e}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,'Inter',sans-serif;line-height:1.6}
nav{background:rgba(15,23,42,.9);backdrop-filter:blur(12px);border-bottom:1px solid var(--s2);padding:16px 5%;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.logo{display:flex;align-items:center;gap:8px;text-decoration:none}
.dot{width:9px;height:9px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc)}
.logo-name{font-size:18px;font-weight:800}
.logo-name span{color:#60a5fa}
.nav-links{display:flex;gap:24px}
.nav-links a{color:var(--mut);text-decoration:none;font-size:14px;transition:.15s}
.nav-links a:hover{color:var(--tx)}
.cta-btn{background:var(--acc);color:#fff;padding:8px 20px;border-radius:10px;text-decoration:none;font-weight:600;font-size:14px}
.hero{padding:80px 5%;text-align:center}
.badge{display:inline-block;background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:var(--grn);padding:5px 14px;border-radius:20px;font-size:12px;font-weight:700;margin-bottom:24px}
h1{font-size:clamp(32px,5vw,52px);font-weight:900;letter-spacing:-1px;line-height:1.1;margin-bottom:16px}
.sub{color:var(--mut);font-size:18px;max-width:600px;margin:0 auto 40px}
.container{max-width:900px;margin:0 auto;padding:0 5%}
.case{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:40px;margin-bottom:32px}
.case-header{display:flex;align-items:flex-start;gap:24px;margin-bottom:32px;flex-wrap:wrap}
.case-logo{width:56px;height:56px;background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.2);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:24px;flex-shrink:0}
.case-meta{flex:1}
.case-name{font-size:20px;font-weight:800;margin-bottom:4px}
.case-desc{color:var(--mut);font-size:14px}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:16px;margin:28px 0;padding:24px;background:rgba(15,23,42,.5);border-radius:14px}
.metric{text-align:center}
.metric-val{font-size:32px;font-weight:900;background:linear-gradient(135deg,#22c55e,#16a34a);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.metric-label{font-size:11px;color:var(--mut);margin-top:4px;text-transform:uppercase;letter-spacing:.4px}
.quote{border-left:3px solid var(--acc);padding:16px 20px;margin:24px 0;background:rgba(59,130,246,.04);border-radius:0 10px 10px 0}
.quote p{font-style:italic;color:var(--tx);font-size:15px;line-height:1.7}
.quote cite{font-size:12px;color:var(--mut);margin-top:8px;display:block;font-style:normal}
.tag{background:rgba(59,130,246,.1);color:#60a5fa;border-radius:6px;padding:3px 9px;font-size:11px;font-weight:600}
.cta-section{text-align:center;padding:80px 5%;background:linear-gradient(135deg,rgba(59,130,246,.08),rgba(167,139,250,.04))}
footer{text-align:center;padding:32px;font-size:12px;color:#475569;border-top:1px solid var(--s2)}
footer a{color:#475569}
@media(max-width:600px){.case-header{flex-direction:column}.metrics{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<nav>
  <a href="/" class="logo"><div class="dot"></div><span class="logo-name">Yve<span>.01</span></span></a>
  <div class="nav-links">
    <a href="/">Inicio</a>
    <a href="/blog">Blog</a>
    <a href="/about">Equipo</a>
  </div>
  <a href="/signup" class="cta-btn">Empezar gratis</a>
</nav>

<div class="hero">
  <div class="badge">✓ Validado con hoteles reales</div>
  <h1>Resultados reales.<br>No promesas.</h1>
  <p class="sub">Cómo hoteles independientes y grupos hoteleros en España automatizan su departamento financiero con Yve.01.</p>
</div>

<div class="container" style="padding-bottom:80px">

  <!-- CASO 1: Calipolis -->
  <div class="case">
    <div class="case-header">
      <div class="case-logo">🏨</div>
      <div class="case-meta">
        <div class="case-name">Grupo Calipolis Hotels</div>
        <div class="case-desc">3 propiedades · 307 habitaciones totales · Sitges, Barcelona</div>
        <div style="margin-top:8px;display:flex;gap:6px;flex-wrap:wrap">
          <span class="tag">4 estrellas</span>
          <span class="tag">Grupo familiar</span>
          <span class="tag">F&B complejo</span>
        </div>
      </div>
    </div>

    <div class="metrics">
      <div class="metric">
        <div class="metric-val">6h</div>
        <div class="metric-label">Ahorradas / semana</div>
      </div>
      <div class="metric">
        <div class="metric-val">100%</div>
        <div class="metric-label">OOB detectados</div>
      </div>
      <div class="metric">
        <div class="metric-val">+4pp</div>
        <div class="metric-label">Mejora GOP%</div>
      </div>
      <div class="metric">
        <div class="metric-val">3</div>
        <div class="metric-label">Hoteles en 1 dashboard</div>
      </div>
    </div>

    <h3 style="font-size:16px;font-weight:700;margin-bottom:12px">El reto</h3>
    <p style="color:var(--mut);margin-bottom:20px">El F&B Manager del grupo gestionaba manualmente el Food Cost de tres propiedades, con datos inconsistentes entre el POS y las compras reales. El cierre mensual tardaba 3 días.</p>

    <h3 style="font-size:16px;font-weight:700;margin-bottom:12px">La solución</h3>
    <p style="color:var(--mut);margin-bottom:20px">Yve.01 se conectó al POS de las tres propiedades, automatizó el 3-way matching de facturas de proveedores y generó un dashboard consolidado multi-hotel accesible desde el móvil.</p>

    <div class="quote">
      <p>"Con Yve.01 veo el Food Cost real de los tres hoteles cada mañana. Antes tardaba una semana en tener esos números."</p>
      <cite>— F&B Manager, Grupo Calipolis Hotels</cite>
    </div>
  </div>

  <!-- CASO 2: Hotel Independiente Barcelona -->
  <div class="case">
    <div class="case-header">
      <div class="case-logo">⭐</div>
      <div class="case-meta">
        <div class="case-name">Cadena internacional 5 estrellas</div>
        <div class="case-desc">1 propiedad · +300 habitaciones · Barcelona</div>
        <div style="margin-top:8px;display:flex;gap:6px;flex-wrap:wrap">
          <span class="tag">5 estrellas</span>
          <span class="tag">OTAs alta facturación</span>
          <span class="tag">Oracle Fusion</span>
        </div>
      </div>
    </div>

    <div class="metrics">
      <div class="metric">
        <div class="metric-val">150+</div>
        <div class="metric-label">Facturas / mes auto</div>
      </div>
      <div class="metric">
        <div class="metric-val">8h</div>
        <div class="metric-label">AR diario → 45 min</div>
      </div>
      <div class="metric">
        <div class="metric-val">0</div>
        <div class="metric-label">Certs DI perdidos</div>
      </div>
      <div class="metric">
        <div class="metric-val">€0</div>
        <div class="metric-label">Multas por DI tardío</div>
      </div>
    </div>

    <h3 style="font-size:16px;font-weight:700;margin-bottom:12px">El reto</h3>
    <p style="color:var(--mut);margin-bottom:20px">El Income Auditor dedicaba 8 horas diarias al cierre de AR: verificar comisiones de Booking y Expedia, detectar certificados de doble imposición pendientes, y contabilizar en Oracle manualmente.</p>

    <h3 style="font-size:16px;font-weight:700;margin-bottom:12px">La solución</h3>
    <p style="color:var(--mut);margin-bottom:20px">Yve.01 automatizó la verificación de comisiones, los recordatorios de certificados DI, y la contabilización directa en Oracle Fusion via API. El Income Auditor ahora revisa excepciones, no procesa rutinas.</p>

    <div class="quote">
      <p>"Se me harían solas. Eso es lo que pienso cada vez que veo Yve procesar las facturas de las OTAs."</p>
      <cite>— Assistant Financial Controller, cadena 5 estrellas Barcelona</cite>
    </div>
  </div>

</div>

<div class="cta-section">
  <h2 style="font-size:28px;font-weight:800;margin-bottom:12px">Tu hotel podría ser el siguiente</h2>
  <p style="color:var(--mut);margin-bottom:32px;font-size:16px">14 días gratis · Setup en 15 minutos · Sin permanencia</p>
  <a href="/signup" style="display:inline-block;background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;padding:14px 32px;border-radius:12px;font-size:15px;font-weight:700;text-decoration:none;box-shadow:0 4px 20px rgba(59,130,246,.35);margin-right:12px">Empezar gratis →</a>
  <a href="/about" style="display:inline-block;border:1px solid var(--s2);color:var(--mut);padding:14px 24px;border-radius:12px;font-size:15px;font-decoration:none;text-decoration:none">Conocer el equipo</a>
</div>

<footer>© 2026 Yve.01 · Barcelona, España · <a href="/terminos">Términos</a> · <a href="/privacidad">Privacidad</a> · <a href="/cookies">Cookies</a></footer>
</body></html>"""
