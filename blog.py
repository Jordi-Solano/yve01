"""blog.py — Yve Blog SEO · 3 artículos optimizados para hoteleros España"""
from flask import Blueprint, Response

blog_bp = Blueprint('blog', __name__)

POSTS = [
  {
    "slug": "software-gestion-financiera-hoteles-espana",
    "title": "Software de gestión financiera para hoteles en España: guía 2026",
    "desc": "Comparativa completa de las mejores herramientas para automatizar AP, AR y conciliación bancaria en hoteles independientes y cadenas medianas.",
    "date": "2026-05-15", "cat": "Gestión hotelera", "read": "8 min",
    "body": """<h2>¿Por qué los hoteles necesitan software financiero específico?</h2>
<p>Los hoteles tienen una estructura financiera única: ingresos de múltiples canales (PMS, POS de restaurante, OTAs), proveedores de todo tipo y un volumen de facturas que puede superar las 150 al mes para un hotel de 100 habitaciones. Los ERPs genéricos no están diseñados para esto.</p>
<h2>Los 5 procesos que más tiempo consumen</h2>
<p><strong>1. Verificación de comisiones OTA.</strong> Cada factura de Booking.com o Expedia tiene que verificarse manualmente contra las tarifas pactadas. Un error en el porcentaje de comisión puede costar miles de euros al año sin que nadie lo detecte.</p>
<p><strong>2. 3-way matching en F&B.</strong> Para los hoteles con restaurante, cada compra requiere cruzar la factura del proveedor, el albarán físico y los datos del POS. Hacerlo manualmente son 3-4 horas diarias.</p>
<p><strong>3. Contabilización en Oracle.</strong> Introducir asientos contables en Oracle Fusion manualmente es lento y propenso a errores. La automatización puede reducir este tiempo un 90%.</p>
<p><strong>4. Daily Revenue Report.</strong> El Income Auditor dedica entre 1 y 2 horas cada mañana a revisar el DRR. Un sistema que detecte automáticamente los Out of Balance libera ese tiempo.</p>
<p><strong>5. Conciliación bancaria.</strong> Cruzar el extracto bancario con las facturas pagadas a mano puede llevar medio día a la semana.</p>
<h2>Yve.01: la alternativa AI-first para hoteles europeos</h2>
<p>A diferencia de las soluciones legacy como M3 o Aptech, Yve está diseñado desde cero con IA en el núcleo. El setup tarda 15 minutos y el precio comienza en 400€/mes.</p>""",
  },
  {
    "slug": "automatizar-cuentas-pagar-hotel",
    "title": "Cómo automatizar las cuentas a pagar (AP) en un hotel",
    "desc": "Guía práctica para automatizar el proceso AP hotelero: desde la recepción de facturas hasta la contabilización en Oracle, con ejemplos reales.",
    "date": "2026-05-28", "cat": "Cuentas a pagar", "read": "6 min",
    "body": """<h2>El flujo AP en un hotel: el problema</h2>
<p>Un hotel de tamaño medio recibe facturas por tres canales: email, correo físico y portales de proveedores. Sin un sistema centralizado, el equipo financiero pierde horas buscando qué facturas faltan y cuáles están duplicadas.</p>
<h2>El proceso AP estándar en hotelería</h2>
<p>El camino correcto es: recepción → verificación → matching con PO → aprobación del jefe de departamento → contabilización en Oracle → pago. Para facturas F&B se añade el <strong>3-way matching</strong>: factura + albarán físico sellado + POS.</p>
<h2>La automatización con IA</h2>
<p>Los sistemas modernos como Yve utilizan OCR e IA para leer PDFs de facturas, extraer los campos clave y cruzarlos automáticamente con las órdenes de compra. Las discrepancias se detectan y se generan emails de reclamación automáticamente.</p>
<h2>Integración con Oracle GL</h2>
<p>El último paso, contabilizar en Oracle, puede automatizarse completamente via la API REST de Oracle Fusion Cloud. Cada factura aprobada genera automáticamente el asiento con las cuentas correctas del PGC español.</p>""",
  },
  {
    "slug": "revenue-management-hoteles-independientes",
    "title": "Revenue management para hoteles independientes: lo que nadie te explica",
    "desc": "Los hoteles independientes pueden competir con las grandes cadenas en revenue management sin invertir en software caro. Aquí está cómo.",
    "date": "2026-06-01", "cat": "Revenue Management", "read": "7 min",
    "body": """<h2>El mito del revenue management caro</h2>
<p>Muchos hoteleros independientes asumen que el revenue management avanzado es solo para grandes cadenas. No es verdad. Las métricas clave — ADR, RevPAR, Occupancy% — son accesibles y medibles para cualquier hotel.</p>
<h2>Las métricas que realmente importan</h2>
<p><strong>ADR (Average Daily Rate)</strong>: el precio medio por habitación ocupada. Es tu palanca más directa de revenue.</p>
<p><strong>RevPAR (Revenue Per Available Room)</strong>: combina ocupación y ADR. Es la métrica que usan los inversores.</p>
<p><strong>GOP% (Gross Operating Profit)</strong>: el margen operativo real del hotel. Un hotel 4★ bien gestionado debería estar entre el 30-45%.</p>
<h2>El Daily Revenue Report: tu brújula diaria</h2>
<p>El DRR es el informe que resume todos los ingresos del día anterior por departamento. Automatizar su revisión y la detección de Out of Balance es el primer paso hacia un revenue management eficiente.</p>""",
  },
  {
    "slug": "food-cost-hotel-restaurante-como-calcularlo",
    "title": "Food Cost en hoteles con restaurante: cómo calcularlo y reducirlo",
    "desc": "El Food Cost real vs teórico es la métrica más importante para el F&B Manager. Te explicamos cómo calcularlo automáticamente desde los datos del POS.",
    "date": "2026-06-03", "cat": "F&B Management", "read": "6 min",
    "body": """<h2>¿Qué es el Food Cost y por qué importa?</h2>
<p>El Food Cost es el porcentaje de los ingresos de F&B que se destina al coste de los ingredientes. Un FC del 18-25% es saludable para un restaurante de hotel; por encima del 30% hay un problema.</p>
<p>La ecuación básica: <strong>Food Cost % = (Coste de ingredientes / Ingresos F&B) × 100</strong></p>
<h2>FC Teórico vs FC Real: la diferencia que nadie mide</h2>
<p>El <strong>FC Teórico</strong> es lo que debería costar producir lo que has vendido, según tus recetas. El <strong>FC Real</strong> es lo que realmente has gastado. La diferencia entre ambos revela mermas, robos, errores de porcionado o recetas incorrectas.</p>
<p>Si tu FC Teórico es 20% y tu FC Real es 28%, estás perdiendo un 8% de tus ingresos en algún lugar de la cadena. En un restaurante de hotel con 60.000€/mes de ventas, eso son 4.800€/mes que se evaporan.</p>
<h2>Cómo calcularlo automáticamente</h2>
<p>El proceso manual requiere exportar datos del POS, cruzarlos con las recetas estándar, calcular el consumo teórico por ingrediente y compararlo con las compras reales. Son horas de trabajo en Excel.</p>
<p>Con un sistema como Yve.01, este cálculo se actualiza diariamente de forma automática: los datos del POS alimentan el FC Real, las recetas calculan el FC Teórico, y las desviaciones se muestran por categoría y por plato.</p>
<h2>Reducir el Food Cost: las 5 palancas</h2>
<p><strong>1. Estandarizar recetas:</strong> El porcionado inconsistente puede añadir un 3-5% al FC Real.</p>
<p><strong>2. Control de mermas:</strong> Registrar y analizar las mermas permite identificar qué ingredientes se pierden más y por qué.</p>
<p><strong>3. Análisis por plato:</strong> No todos los platos tienen el mismo margen. Identificar los platos con FC% alto y rediseñarlos o retirarlos del menú.</p>
<p><strong>4. Revisión periódica de precios de compra:</strong> Los precios de proveedores cambian. Si no actualizas tu tabla de costes, tu FC Teórico será incorrecto.</p>
<p><strong>5. Gestión de inventario en tiempo real:</strong> Saber qué tienes en stock evita compras innecesarias y caducidades.</p>""",
  },
  {
    "slug": "conciliacion-bancaria-hotel-guia-completa",
    "title": "Conciliación bancaria en hoteles: guía completa para el Financial Controller",
    "desc": "La conciliación bancaria es uno de los procesos más tediosos del departamento financiero hotelero. Automatizarla puede ahorrar medio día de trabajo a la semana.",
    "date": "2026-06-05", "cat": "Contabilidad hotelera", "read": "7 min",
    "body": """<h2>Qué es la conciliación bancaria en un hotel</h2>
<p>La conciliación bancaria consiste en comparar los movimientos del extracto bancario con las facturas registradas en el sistema contable, identificando qué pagos han sido procesados y cuáles siguen pendientes.</p>
<p>Para un hotel mediano, esto implica cruzar 150-300 movimientos mensuales: pagos a proveedores, liquidaciones de OTAs (Booking, Expedia), cargos TPV, transferencias de grupos corporativos y comisiones bancarias.</p>
<h2>Los problemas típicos</h2>
<p><strong>Facturas sin movimiento bancario:</strong> una factura está contabilizada pero el pago aún no ha aparecido en el extracto. ¿Es un error o simplemente tarda en procesarse?</p>
<p><strong>Movimientos sin factura:</strong> hay un cargo en el banco que nadie reconoce. Puede ser una comisión oculta, un cargo duplicado o una factura que no ha llegado al sistema.</p>
<p><strong>Diferencias de importe:</strong> el banco muestra 12.450€ pagados a Booking.com, pero en el sistema solo consta una factura de 12.200€. Los 250€ de diferencia hay que justificarlos.</p>
<h2>El proceso manual vs automatizado</h2>
<p>El proceso manual requiere exportar el extracto en Excel, ordenar por proveedor e importe, y cruzarlo manualmente con las facturas del sistema. Con 200 movimientos mensuales, esto son 3-4 horas semanales de trabajo.</p>
<p>La automatización aplica algoritmos de matching: primero por importe exacto y referencia, luego por importe aproximado y proveedor, finalmente detecta los movimientos sin match como alertas. El proceso que tomaba horas se completa en segundos.</p>
<h2>Integración con Oracle GL</h2>
<p>Cuando la conciliación está integrada con Oracle Fusion, los movimientos conciliados se contabilizan automáticamente en el libro mayor. Los movimientos pendientes generan asientos provisionales que se cierran cuando el pago se confirma.</p>""",
  },
  {
    "slug": "out-of-balance-drr-como-detectarlo",
    "title": "Out of Balance en el DRR: qué es, por qué ocurre y cómo detectarlo automáticamente",
    "desc": "El Out of Balance es el error más común en el Daily Revenue Report de los hoteles. Te explicamos sus causas y cómo detectarlo automáticamente cada mañana.",
    "date": "2026-06-07", "cat": "Daily Revenue Report", "read": "5 min",
    "body": """<h2>¿Qué es un Out of Balance?</h2>
<p>Un Out of Balance (OOB) ocurre cuando la suma de todos los ingresos y cargos del día en el DRR no cuadra a cero. Es decir, hay una diferencia entre el debe y el haber del informe diario.</p>
<p>Ejemplo real: el DRR reporta 59.600€ en ingresos de habitaciones, 8.400€ en F&B, 1.200€ en spa. Pero la suma de los departamentos da 69.350€ mientras el total consolidado marca 69.100€. La diferencia de 250€ es el Out of Balance.</p>
<h2>Las causas más frecuentes</h2>
<ul>
<li><strong>Posting manual incorrecto:</strong> Un recepcionista ha introducido manualmente un cargo y se ha equivocado en 1 céntimo. Multiplicado por 200 reservas, puede sumar una diferencia significativa.</li>
<li><strong>Rounding en tipos de cambio:</strong> Los hoteles que trabajan con monedas extranjeras tienen OOB frecuentes por el redondeo en la conversión.</li>
<li><strong>Interfaces entre sistemas:</strong> Cuando el PMS y el POS no están perfectamente sincronizados, una venta en el restaurante puede no llegar al DRR.</li>
<li><strong>Correcciones de días anteriores:</strong> Un ajuste contable de un día anterior puede crear un OOB en el día actual si no se procesa correctamente.</li>
</ul>
<h2>Cómo detectarlo cada mañana automáticamente</h2>
<p>El Income Auditor dedica típicamente 30-60 minutos cada mañana a revisar el DRR y detectar OOB. Con Yve.01, el sistema lee el DRR automáticamente, calcula el balance de cada Trial Balance y genera una alerta inmediata si detecta una discrepancia.</p>
<p>El sistema muestra exactamente en qué línea contable está la diferencia, en qué departamento y en qué momento del día ocurrió, acelerando la resolución de días que normalmente llevarían horas.</p>""",
  },

]

_CSS = """
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e;--pur:#8b5cf6}
*{box-sizing:border-box;margin:0;padding:0}html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;line-height:1.65;-webkit-font-smoothing:antialiased}
a{text-decoration:none;color:inherit}
.pub-nav{position:sticky;top:0;z-index:100;background:rgba(15,23,42,.9);backdrop-filter:blur(16px);border-bottom:1px solid rgba(51,65,85,.6);padding:0 5%;height:62px;display:flex;align-items:center;justify-content:space-between;gap:20px}
.pub-nav-logo{display:flex;align-items:center;gap:9px;flex-shrink:0}
.pub-nav-logo .dot{width:9px;height:9px;border-radius:50%;background:var(--acc);box-shadow:0 0 10px var(--acc)}
.pub-nav-logo .name{font-size:19px;font-weight:800;letter-spacing:-.5px;color:#fff}
.pub-nav-logo .name span{color:var(--acc2)}
.pub-nav-logo .badge{font-size:10px;font-weight:600;color:var(--acc2);background:rgba(59,130,246,.12);border:1px solid rgba(59,130,246,.25);border-radius:12px;padding:2px 8px;margin-left:4px}
.pub-nav-links{display:flex;gap:24px}
.pub-nav-links a{font-size:14px;color:var(--mut);font-weight:500;transition:color .15s}
.pub-nav-links a:hover,.pub-nav-links a.active{color:var(--tx)}
.pub-nav-cta{display:flex;gap:10px;flex-shrink:0}
.pub-btn-ghost{font-size:13px;font-weight:600;color:var(--mut);padding:7px 16px;border-radius:9px;border:1px solid var(--s2);transition:.15s}
.pub-btn-ghost:hover{border-color:var(--acc2);color:var(--acc2)}
.pub-btn-cta{font-size:13px;font-weight:700;color:#fff;padding:8px 18px;border-radius:9px;background:linear-gradient(135deg,var(--acc),#2563eb);box-shadow:0 0 16px rgba(59,130,246,.3);transition:.15s}
.pub-btn-cta:hover{box-shadow:0 0 24px rgba(59,130,246,.5);transform:translateY(-1px)}
@media(max-width:768px){.pub-nav-links{display:none}.pub-nav-logo .badge{display:none}}
@media(max-width:420px){.pub-btn-ghost{display:none}}
.pub-footer{background:var(--s1);border-top:1px solid var(--s2);padding:48px 5% 28px;margin-top:72px}
.pub-footer-grid{max-width:1100px;margin:0 auto;display:grid;grid-template-columns:2fr 1fr 1fr;gap:40px;margin-bottom:36px}
.pub-footer-brand .logo-row{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.pub-footer-brand .dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 6px var(--acc)}
.pub-footer-brand .name{font-size:17px;font-weight:800}.pub-footer-brand .name span{color:var(--acc2)}
.pub-footer-brand p{font-size:13px;color:var(--dim);line-height:1.7;max-width:260px}
.pub-footer-col h4{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--mut);margin-bottom:14px}
.pub-footer-col a{display:block;font-size:13px;color:var(--dim);margin-bottom:8px;transition:color .15s}
.pub-footer-col a:hover{color:var(--tx)}
.pub-footer-bottom{max-width:1100px;margin:0 auto;border-top:1px solid var(--s2);padding-top:18px;display:flex;justify-content:space-between;font-size:12px;color:var(--dim);flex-wrap:wrap;gap:8px}
@media(max-width:640px){.pub-footer-grid{grid-template-columns:1fr}}
#read-progress{position:fixed;top:0;left:0;height:2px;background:linear-gradient(90deg,var(--acc),var(--pur));z-index:9999;width:0%;transition:width .1s}
.fade-up{opacity:0;transform:translateY(20px);transition:opacity .55s ease,transform .55s ease}
.fade-up.visible{opacity:1;transform:none}
.fade-up-delay-1{transition-delay:.1s}.fade-up-delay-2{transition-delay:.2s}.fade-up-delay-3{transition-delay:.3s}
.post-cat{display:inline-block;background:rgba(59,130,246,.1);color:var(--acc2);border:1px solid rgba(59,130,246,.2);border-radius:20px;padding:3px 11px;font-size:11px;font-weight:700;margin-bottom:12px}
.post-title{font-size:20px;font-weight:700;color:var(--tx);margin-bottom:9px;line-height:1.35}
.post-desc{font-size:14px;color:var(--mut);line-height:1.65;margin-bottom:12px}
.article-content h2{font-size:22px;font-weight:700;margin:32px 0 13px;color:var(--tx)}
.article-content p{font-size:17px;color:var(--mut);line-height:1.8;margin-bottom:18px}
.article-content strong{color:var(--tx);font-weight:600}
"""

_HEAD = lambda title,desc: f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="{desc}">
<meta property="og:title" content="{title}">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<title>{title}</title><style>{{}}</style></head>"""

_NAV = """<nav class="pub-nav">
  <a href="/" class="pub-nav-logo"><div class="dot"></div><span class="name">Yve<span>.01</span></span><span class="badge">Beta</span></a>
  <div class="pub-nav-links">
    <a href="/#features">Funciones</a><a href="/#pricing">Precios</a>
    <a href="/blog" class="active">Blog</a><a href="/casos">Casos</a><a href="/about">Nosotros</a>
  </div>
  <div class="pub-nav-cta">
    <a href="/login" class="pub-btn-ghost">Acceder</a>
    <a href="/signup" class="pub-btn-cta">Empezar gratis →</a>
  </div>
</nav>"""

_FOOTER = """<footer class="pub-footer">
  <div class="pub-footer-grid">
    <div class="pub-footer-brand">
      <div class="logo-row"><div class="dot"></div><span class="name">Yve<span>.01</span></span></div>
      <p>Sistema operativo AI-first para la gestión financiera hotelera. Hecho en Barcelona para hoteles europeos.</p>
    </div>
    <div class="pub-footer-col"><h4>Producto</h4>
      <a href="/#features">Funciones</a><a href="/#pricing">Precios</a>
      <a href="/casos">Casos de éxito</a><a href="/about">Quiénes somos</a><a href="/blog">Blog</a>
    </div>
    <div class="pub-footer-col"><h4>Contacto</h4>
      <a href="mailto:jordi@yve01.com">jordi@yve01.com</a>
      <a href="https://github.com/Jordi-Solano/yve01">GitHub</a>
      <a href="/login">Panel</a><a href="/signup">Crear cuenta</a>
    </div>
  </div>
  <div class="pub-footer-bottom"><span>© 2026 Yve.01 · Barcelona</span><span>Hecho con IA · Validado con hoteleros reales</span></div>
</footer>"""

_SCROLL = """<script>
const _obs=new IntersectionObserver(e=>{e.forEach(x=>{if(x.isIntersecting){x.target.classList.add('visible');_obs.unobserve(x.target);}});},{threshold:.12});
document.querySelectorAll('.fade-up').forEach(el=>_obs.observe(el));
const _p=document.getElementById('read-progress');
if(_p)document.addEventListener('scroll',()=>{const h=document.documentElement;_p.style.width=Math.min(h.scrollTop/(h.scrollHeight-h.clientHeight)*100,100)+'%';});
</script>"""

def post_html(post):
    return f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="{post['desc']}">
<meta property="og:title" content="{post['title']} | Yve Blog">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<title>{post['title']} | Yve Blog</title>
<style>{_CSS}</style></head>
<body>
<div id="read-progress"></div>
{_NAV}
<div style="max-width:720px;margin:0 auto;padding:48px 5% 80px">
  <div class="fade-up" style="font-size:13px;color:var(--dim);margin-bottom:22px"><a href="/" style="color:var(--acc2)">Inicio</a> / <a href="/blog" style="color:var(--acc2)">Blog</a> / {post['cat']}</div>
  <span class="post-cat fade-up fade-up-delay-1">{post['cat']}</span>
  <h1 class="fade-up fade-up-delay-1" style="font-size:clamp(24px,4vw,40px);font-weight:900;letter-spacing:-1px;line-height:1.12;margin:12px 0">{post['title']}</h1>
  <div class="fade-up fade-up-delay-2" style="font-size:13px;color:var(--dim);margin-bottom:36px;padding-bottom:26px;border-bottom:1px solid var(--s2)">{post['date']} &nbsp;·&nbsp; {post['read']} de lectura</div>
  <div class="article-content fade-up fade-up-delay-2">{post['body']}</div>
  <div class="fade-up" style="background:var(--s1);border:1px solid rgba(59,130,246,.2);border-radius:16px;padding:32px;text-align:center;margin-top:52px">
    <h3 style="font-size:21px;font-weight:700;margin-bottom:9px">¿Quieres automatizar tu hotel?</h3>
    <p style="color:var(--mut);margin-bottom:22px">14 días gratis · Sin tarjeta · Setup en 15 min</p>
    <a href="/signup" style="display:inline-block;background:linear-gradient(135deg,var(--acc),#2563eb);color:#fff;padding:12px 28px;border-radius:11px;font-size:15px;font-weight:700;box-shadow:0 4px 20px rgba(59,130,246,.35)">Empezar gratis →</a>
  </div>
</div>
{_FOOTER}
{_SCROLL}
</body></html>"""

def index_html(posts):
    cards = "".join(
        f'<a href="/blog/{p["slug"]}" style="background:var(--s1);border:1px solid var(--s2);border-radius:14px;padding:26px;display:block;transition:border-color .18s,transform .18s" class="fade-up fade-up-delay-{min(i+1,3)}" onmouseover="this.style.borderColor=\'rgba(59,130,246,.4)\';this.style.transform=\'translateY(-2px)\'" onmouseout="this.style.borderColor=\'var(--s2)\';this.style.transform=\'\'"><span class="post-cat">{p["cat"]}</span><h2 class="post-title">{p["title"]}</h2><p class="post-desc">{p["desc"]}</p><span style="font-size:12px;color:var(--dim)">{p["date"]} · {p["read"]}</span></a>'
        for i, p in enumerate(posts)
    )
    return f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Artículos sobre gestión financiera hotelera, automatización AP/AR y revenue management para hoteles en España.">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<title>Blog | Yve.01</title><style>{_CSS}</style></head>
<body>
{_NAV}
<div style="padding:64px 5% 40px;text-align:center;position:relative;overflow:hidden">
  <div style="position:absolute;inset:0;background:radial-gradient(600px 250px at 50% 0%,rgba(59,130,246,.1),transparent 60%);pointer-events:none"></div>
  <div style="font-size:11px;font-weight:700;color:var(--acc2);text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px" class="fade-up">Recursos</div>
  <h1 style="font-size:clamp(28px,4vw,44px);font-weight:900;letter-spacing:-1px;margin-bottom:14px" class="fade-up fade-up-delay-1">Blog de Yve</h1>
  <p style="font-size:17px;color:var(--mut);max-width:500px;margin:0 auto" class="fade-up fade-up-delay-2">Gestión financiera hotelera, automatización y tecnología para hoteles en España y Europa.</p>
</div>
<div style="max-width:860px;margin:0 auto;padding:0 5% 80px;display:grid;gap:18px">{cards}</div>
{_FOOTER}
{_SCROLL}
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
