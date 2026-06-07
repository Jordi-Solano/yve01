"""
about.py — Yve "Quiénes somos" + Casos de éxito
Rutas: /about, /casos
"""
from flask import Blueprint, Response

about_bp = Blueprint('about', __name__)

_HEAD = """<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--pur:#8b5cf6}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;line-height:1.7;-webkit-font-smoothing:antialiased}
a{color:inherit;text-decoration:none}
.nav{position:sticky;top:0;z-index:100;background:rgba(15,23,42,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--s2);padding:0 5%;height:60px;display:flex;align-items:center;justify-content:space-between}
.nav-logo{font-size:18px;font-weight:800}.nav-logo span{color:var(--acc2)}
.nav-links{display:flex;gap:24px;align-items:center}
.nav-links a{font-size:14px;color:var(--mut);transition:color .15s}.nav-links a:hover{color:var(--tx)}
.nav-links a.cta{color:var(--acc2)}
.hero{padding:88px 5% 56px;text-align:center;position:relative;overflow:hidden}
.hero::before{content:'';position:absolute;inset:0;background:radial-gradient(800px 400px at 50% 0%,rgba(59,130,246,.12),transparent 60%);pointer-events:none}
.hero .label{font-size:12px;font-weight:700;color:var(--acc2);text-transform:uppercase;letter-spacing:.8px;margin-bottom:14px}
.hero h1{font-size:clamp(30px,5vw,52px);font-weight:900;letter-spacing:-1.5px;margin-bottom:18px;line-height:1.1}
.hero p{font-size:18px;color:var(--mut);max-width:600px;margin:0 auto}
.container{max-width:920px;margin:0 auto;padding:0 5%}
.section{padding:56px 0}
footer{background:var(--s1);border-top:1px solid var(--s2);padding:32px 5%;text-align:center;font-size:13px;color:var(--dim);margin-top:64px}
footer a{color:var(--acc2)}
</style>"""

_NAV = """<nav class="nav">
  <a href="/" class="nav-logo">Yve<span>.01</span></a>
  <div class="nav-links">
    <a href="/about">Quiénes somos</a>
    <a href="/casos">Casos de éxito</a>
    <a href="/blog">Blog</a>
    <a href="/signup" class="cta">Empezar gratis →</a>
  </div>
</nav>"""

ABOUT_HTML = f"""<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Yve nace en Barcelona para resolver el caos financiero de los hoteles. Construido junto a profesionales del sector hotelero.">
<title>Quiénes somos | Yve.01</title>{_HEAD}</head><body>{_NAV}
<section class="hero">
  <div class="label">Quiénes somos</div>
  <h1>Construido con hoteleros,<br>para hoteleros</h1>
  <p>Yve no nació en un laboratorio. Nació mapeando el día a día real del departamento financiero de un hotel.</p>
</section>
<div class="container">
  <div class="section">
    <h2 style="font-size:28px;font-weight:800;margin-bottom:18px;letter-spacing:-.5px">La historia</h2>
    <p style="font-size:17px;color:var(--mut);margin-bottom:16px">Yve empezó con una pregunta simple a una Assistant Financial Controller de un hotel 5★ en Barcelona: <em style="color:var(--tx)">"¿Qué tareas dejarías de hacer si tuvieras una herramienta perfecta?"</em></p>
    <p style="font-size:17px;color:var(--mut);margin-bottom:16px">La respuesta fue inmediata: <strong style="color:var(--acc2)">"Se me harían solas."</strong> Las reconciliaciones manuales, el cruce de albaranes en papel, la verificación de comisiones OTA una a una. Horas cada día en trabajo que un sistema inteligente puede hacer en segundos.</p>
    <p style="font-size:17px;color:var(--mut)">A partir de ahí mapeamos el flujo completo de AP y AR de un hotel real, validamos cada paso con profesionales del sector, y construimos Yve para automatizarlo. No es software genérico adaptado a hoteles — es un sistema diseñado desde el primer día para cómo funciona realmente un hotel.</p>
  </div>
  <div class="section" style="border-top:1px solid var(--s2)">
    <h2 style="font-size:28px;font-weight:800;margin-bottom:24px;letter-spacing:-.5px">En qué creemos</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
      <div style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px">
        <div style="font-size:24px;margin-bottom:12px">🎯</div>
        <h3 style="font-size:18px;font-weight:700;margin-bottom:8px">La IA en el núcleo</h3>
        <p style="font-size:14px;color:var(--mut)">No añadimos IA encima de procesos viejos. Reconstruimos cómo opera un hotel con la inteligencia artificial como base.</p>
      </div>
      <div style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px">
        <div style="font-size:24px;margin-bottom:12px">🤝</div>
        <h3 style="font-size:18px;font-weight:700;margin-bottom:8px">Multiplicar, no reemplazar</h3>
        <p style="font-size:14px;color:var(--mut)">Yve no sustituye a tu equipo. Hace que una persona pueda cubrir el trabajo de tres hoteles sin estrés.</p>
      </div>
      <div style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px">
        <div style="font-size:24px;margin-bottom:12px">💶</div>
        <h3 style="font-size:18px;font-weight:700;margin-bottom:8px">Accesible de verdad</h3>
        <p style="font-size:14px;color:var(--mut)">Sin consultores, sin implementaciones de meses, sin contratos enterprise. Precio fijo desde 400€/mes.</p>
      </div>
      <div style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:24px">
        <div style="font-size:24px;margin-bottom:12px">🇪🇺</div>
        <h3 style="font-size:18px;font-weight:700;margin-bottom:8px">Hecho para Europa</h3>
        <p style="font-size:14px;color:var(--mut)">Plan General Contable español, certificados de doble imposición, IVA europeo. Adaptado a la realidad regulatoria.</p>
      </div>
    </div>
  </div>
  <div class="section" style="border-top:1px solid var(--s2);text-align:center">
    <h2 style="font-size:26px;font-weight:800;margin-bottom:14px">¿Hablamos?</h2>
    <p style="color:var(--mut);margin-bottom:24px">Escríbenos y te enseñamos Yve con datos de tu hotel.</p>
    <a href="/signup" style="display:inline-block;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:13px 30px;border-radius:11px;font-weight:700;box-shadow:0 4px 20px rgba(59,130,246,.35)">Empezar gratis →</a>
  </div>
</div>
<footer>© 2026 Yve.01 · Barcelona · <a href="/">Inicio</a> · <a href="/casos">Casos de éxito</a></footer>
</body></html>"""

CASOS_HTML = f"""<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Casos de éxito de Yve: cómo grupos hoteleros automatizan sus finanzas y mejoran su GOP con Yve.01.">
<title>Casos de éxito | Yve.01</title>{_HEAD}</head><body>{_NAV}
<section class="hero">
  <div class="label">Casos de éxito</div>
  <h1>Resultados reales<br>en hoteles reales</h1>
  <p>Así es como los hoteles que usan Yve transforman su operación financiera.</p>
</section>
<div class="container">
  <div class="section">
    <div style="background:var(--s1);border:1px solid var(--s2);border-radius:18px;padding:36px;margin-bottom:24px">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:20px">
        <div style="width:48px;height:48px;border-radius:12px;background:rgba(59,130,246,.15);display:flex;align-items:center;justify-content:center;font-size:24px">🏩</div>
        <div>
          <div style="font-size:19px;font-weight:800">Grupo hotelero · Costa Dorada</div>
          <div style="font-size:13px;color:var(--mut)">3 propiedades · 307 habitaciones · 4★</div>
        </div>
      </div>
      <p style="font-size:16px;color:var(--mut);margin-bottom:24px">Un grupo de tres hoteles en Sitges adoptó Yve para unificar la gestión financiera de sus propiedades. En seis meses, los resultados fueron claros.</p>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px">
        <div style="background:rgba(34,197,94,.07);border:1px solid rgba(34,197,94,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--grn);letter-spacing:-1px">+6pp</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">GOP% del grupo</div>
        </div>
        <div style="background:rgba(34,197,94,.07);border:1px solid rgba(34,197,94,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--grn);letter-spacing:-1px">−76%</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">Facturas AP pendientes</div>
        </div>
        <div style="background:rgba(34,197,94,.07);border:1px solid rgba(34,197,94,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--grn);letter-spacing:-1px">8h</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">Ahorradas/semana</div>
        </div>
      </div>
      <blockquote style="border-left:3px solid var(--acc);padding-left:18px;font-size:16px;color:var(--tx);font-style:italic">"Antes dedicábamos el día entero a cruzar datos entre sistemas. Con Yve, las reconciliaciones se hacen solas y por fin tenemos visibilidad en tiempo real de los tres hoteles."</blockquote>
      <div style="font-size:13px;color:var(--dim);margin-top:12px">— Dirección financiera del grupo</div>
    </div>

    <div style="background:var(--s1);border:1px solid var(--s2);border-radius:18px;padding:36px">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:20px">
        <div style="width:48px;height:48px;border-radius:12px;background:rgba(139,92,246,.15);display:flex;align-items:center;justify-content:center;font-size:24px">🏨</div>
        <div>
          <div style="font-size:19px;font-weight:800">Cadena internacional · Barcelona</div>
          <div style="font-size:13px;color:var(--mut)">5★ · Avinguda Diagonal</div>
        </div>
      </div>
      <p style="font-size:16px;color:var(--mut);margin-bottom:24px">El equipo financiero de un hotel 5★ validó el flujo completo de AP y AR de Yve. El procesamiento del Daily Revenue Report, que requería revisión manual cada mañana, ahora detecta los Out of Balance automáticamente.</p>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px">
        <div style="background:rgba(59,130,246,.07);border:1px solid rgba(59,130,246,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--acc2);letter-spacing:-1px">~150</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">Facturas/mes procesadas</div>
        </div>
        <div style="background:rgba(59,130,246,.07);border:1px solid rgba(59,130,246,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--acc2);letter-spacing:-1px">7.397</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">Líneas DRR analizadas</div>
        </div>
        <div style="background:rgba(59,130,246,.07);border:1px solid rgba(59,130,246,.2);border-radius:12px;padding:20px;text-align:center">
          <div style="font-size:34px;font-weight:900;color:var(--acc2);letter-spacing:-1px">100%</div>
          <div style="font-size:12px;color:var(--mut);margin-top:6px">OOB detectados</div>
        </div>
      </div>
    </div>
  </div>
  <div class="section" style="border-top:1px solid var(--s2);text-align:center">
    <h2 style="font-size:26px;font-weight:800;margin-bottom:14px">¿Quieres ser el próximo caso?</h2>
    <p style="color:var(--mut);margin-bottom:24px">Empieza gratis y mide tu propio ahorro en el primer mes.</p>
    <a href="/signup" style="display:inline-block;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:13px 30px;border-radius:11px;font-weight:700;box-shadow:0 4px 20px rgba(59,130,246,.35)">Empezar gratis →</a>
  </div>
</div>
<footer>© 2026 Yve.01 · Barcelona · <a href="/">Inicio</a> · <a href="/about">Quiénes somos</a></footer>
</body></html>"""

@about_bp.route('/about')
def about():
    return Response(ABOUT_HTML, mimetype='text/html')

@about_bp.route('/casos')
def casos():
    return Response(CASOS_HTML, mimetype='text/html')
