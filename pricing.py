"""pricing.py — Yve.01 Pricing Page"""
from flask import Blueprint, Response
pricing_bp = Blueprint("pricing", __name__)

PRICING_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5">
<title>Precios | Yve.01 — Software financiero para hoteles</title>
<meta name="description" content="Planes y precios de Yve.01. Sin contratos largos. Sin letra pequeña. El sistema operativo AI para finanzas hoteleras.">
<style>
:root{--bg:#0a0f1e;--s1:#0f172a;--s2:#1e293b;--s3:#334155;--acc:#3b82f6;--acc2:#60a5fa;--grn:#22c55e;--ora:#f59e0b;--pur:#a78bfa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:Inter,system-ui,sans-serif;min-height:100vh}
.nav{display:flex;align-items:center;justify-content:space-between;padding:16px 5%;border-bottom:1px solid rgba(255,255,255,.06);position:sticky;top:0;background:rgba(10,15,30,.95);backdrop-filter:blur(12px);z-index:100}
.logo{font-size:20px;font-weight:800;color:#fff;text-decoration:none}.logo span{color:var(--acc2)}
.nav-links{display:flex;gap:28px;font-size:14px;color:var(--mut)}
.nav-links a{color:var(--mut);text-decoration:none;transition:.15s}.nav-links a:hover{color:var(--tx)}
.btn-demo{background:var(--acc);color:#fff;padding:8px 18px;border-radius:10px;font-size:13px;font-weight:600;text-decoration:none;transition:.2s}
.btn-demo:hover{background:#2563eb}

/* Hero */
.hero{text-align:center;padding:80px 5% 60px;max-width:700px;margin:0 auto}
.hero h1{font-size:42px;font-weight:900;line-height:1.15;margin-bottom:16px}
.hero h1 span{background:linear-gradient(135deg,var(--acc2),var(--pur));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hero p{font-size:17px;color:var(--mut);line-height:1.6;margin-bottom:12px}
.guarantee{display:inline-flex;align-items:center;gap:6px;background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);padding:8px 16px;border-radius:20px;font-size:12px;color:var(--grn);font-weight:600;margin-top:8px}

/* Toggle annual/monthly */
.billing-toggle{display:flex;align-items:center;gap:12px;justify-content:center;margin-bottom:48px;font-size:13px;color:var(--mut)}
.toggle-pill{width:44px;height:24px;background:var(--acc);border-radius:12px;cursor:pointer;position:relative;transition:.2s}
.toggle-pill::after{content:'';position:absolute;width:18px;height:18px;background:#fff;border-radius:50%;top:3px;left:3px;transition:.2s}
.toggle-pill.monthly::after{left:23px}
.badge-save{background:rgba(34,197,94,.15);color:var(--grn);font-size:10px;font-weight:700;padding:2px 7px;border-radius:6px;margin-left:4px}

/* Plans grid */
.plans{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:20px;max-width:1000px;margin:0 auto;padding:0 5% 80px}
.plan{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:32px;position:relative;transition:.3s}
.plan:hover{border-color:var(--s3);transform:translateY(-2px)}
.plan.featured{border-color:var(--acc);background:linear-gradient(135deg,rgba(59,130,246,.06),rgba(167,139,250,.04))}
.plan.featured::before{content:"MÁS POPULAR";position:absolute;top:-12px;left:50%;transform:translateX(-50%);background:var(--acc);color:#fff;font-size:10px;font-weight:800;letter-spacing:.5px;padding:4px 14px;border-radius:20px}
.plan-name{font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);margin-bottom:8px}
.plan-price{font-size:40px;font-weight:900;line-height:1;margin-bottom:4px}
.plan-price span{font-size:16px;font-weight:400;color:var(--mut)}
.plan-sub{font-size:12px;color:var(--dim);margin-bottom:24px}
.plan-features{list-style:none;margin-bottom:28px;display:flex;flex-direction:column;gap:10px}
.plan-features li{font-size:13px;color:var(--mut);display:flex;align-items:flex-start;gap:8px}
.plan-features li::before{content:"✓";color:var(--grn);font-weight:700;flex-shrink:0;margin-top:1px}
.plan-features li.no::before{content:"—";color:var(--s3)}
.plan-features li.no{opacity:.5}
.btn-plan{display:block;text-align:center;padding:13px;border-radius:12px;font-size:14px;font-weight:700;text-decoration:none;transition:.2s;cursor:pointer;border:none;width:100%}
.btn-plan.primary{background:var(--acc);color:#fff}.btn-plan.primary:hover{background:#2563eb}
.btn-plan.outline{background:none;border:1px solid var(--s3);color:var(--tx)}.btn-plan.outline:hover{border-color:var(--acc2)}

/* FAQ */
.faq{max-width:640px;margin:0 auto;padding:0 5% 80px}
.faq h2{text-align:center;font-size:28px;font-weight:800;margin-bottom:40px}
.faq-item{border-bottom:1px solid var(--s2);padding:20px 0;cursor:pointer}
.faq-item summary{font-size:15px;font-weight:600;list-style:none;display:flex;justify-content:space-between;align-items:center}
.faq-item summary::after{content:"+";font-size:20px;color:var(--dim)}
.faq-item[open] summary::after{content:"−"}
.faq-item p{font-size:13px;color:var(--mut);margin-top:12px;line-height:1.7}

/* Enterprise */
.enterprise{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:40px;text-align:center;max-width:700px;margin:0 auto 80px;display:flex;flex-direction:column;align-items:center;gap:12px}
.enterprise h3{font-size:22px;font-weight:800}
.enterprise p{color:var(--mut);font-size:14px}
.enterprise a{background:var(--s2);border:1px solid var(--s3);color:var(--tx);padding:12px 24px;border-radius:12px;text-decoration:none;font-weight:600;font-size:14px;transition:.2s}.enterprise a:hover{border-color:var(--acc2)}

@media(max-width:600px){.hero h1{font-size:28px}.plans{grid-template-columns:1fr}.nav-links{display:none}}
</style>
</head>
<body>
<nav class="nav">
  <a href="/" class="logo">Yve<span>.01</span></a>
  <div class="nav-links">
    <a href="/nosotros">Nosotros</a>
    <a href="/blog">Blog</a>
    <a href="/casos">Casos de éxito</a>
  </div>
  <a href="/signup" class="btn-demo">Empezar gratis</a>
</nav>

<div class="hero">
  <h1>Precios <span>claros</span> para hoteles que quieren crecer</h1>
  <p>Sin contratos anuales obligatorios. Sin costes ocultos. Cancela cuando quieras.</p>
  <div class="guarantee">🔒 14 días de prueba gratis · Sin tarjeta de crédito</div>
</div>

<div class="billing-toggle" id="toggle-wrap">
  <span id="lbl-annual" style="font-weight:700;color:var(--tx)">Anual</span>
  <div class="toggle-pill monthly" id="billing-toggle" onclick="toggleBilling()"></div>
  <span id="lbl-monthly">Mensual</span>
  <span class="badge-save" id="save-badge">Ahorra 20%</span>
</div>

<div class="plans">
  <!-- Starter -->
  <div class="plan">
    <div class="plan-name">Starter</div>
    <div class="plan-price"><span>€</span><span id="p-start">199</span><span>/mes</span></div>
    <div class="plan-sub" id="ps-start">Hasta 80 habitaciones · Pago mensual</div>
    <ul class="plan-features">
      <li>AR OTA — Booking, Expedia, 2 más</li>
      <li>AP — 3-way matching básico</li>
      <li>DRR — carga manual .xlsm</li>
      <li>1 usuario + 1 FC</li>
      <li>Exportación Excel</li>
      <li>Email de soporte</li>
      <li class="no">Multi-hotel</li>
      <li class="no">Oracle Fusion</li>
      <li class="no">API acceso</li>
    </ul>
    <a href="/checkout/starter" class="btn-plan outline">Empezar — €400/mes</a>
<p style="font-size:11px;color:#64748b;text-align:center;margin-top:6px">14 días de prueba · Sin tarjeta hasta confirmar</p>
  </div>
  
  <!-- Pro (featured) -->
  <div class="plan featured">
    <div class="plan-name">Pro</div>
    <div class="plan-price"><span>€</span><span id="p-pro">349</span><span>/mes</span></div>
    <div class="plan-sub" id="ps-pro">Hasta 200 habitaciones · Ideal para 4★</div>
    <ul class="plan-features">
      <li>Todo lo de Starter</li>
      <li>Multi-hotel hasta 3 propiedades</li>
      <li>Oracle GL dry-run + export</li>
      <li>F&B Cost Control avanzado</li>
      <li>Conciliación bancaria automática</li>
      <li>5 usuarios</li>
      <li>Telegram + email alerts</li>
      <li>AR Real — grupos corporativos</li>
      <li>PDF facturas + recordatorios</li>
    </ul>
    <a href="/checkout/pro" class="btn-plan primary">Empezar — €600/mes</a>
<p style="font-size:11px;color:#64748b;text-align:center;margin-top:6px">14 días de prueba · Sin tarjeta hasta confirmar</p>
  </div>

  <!-- Business -->
  <div class="plan">
    <div class="plan-name">Business</div>
    <div class="plan-price"><span>€</span><span id="p-biz">599</span><span>/mes</span></div>
    <div class="plan-sub" id="ps-biz">Habitaciones ilimitadas · Grupos</div>
    <ul class="plan-features">
      <li>Todo lo de Pro</li>
      <li>Multi-hotel ilimitado</li>
      <li>Oracle Fusion integración real</li>
      <li>API acceso completo</li>
      <li>Usuarios ilimitados</li>
      <li>Onboarding dedicado</li>
      <li>WhatsApp + Slack alerts</li>
      <li>SLA respuesta 4h</li>
      <li>Custom reporting</li>
    </ul>
    <a href="/signup?plan=business" class="btn-plan outline">Contactar</a>
  </div>
</div>

<div style="max-width:1000px;margin:0 auto;padding:0 5% 80px">
<div class="enterprise">
  <div style="font-size:40px">🏨</div>
  <h3>¿Cadena con más de 10 hoteles?</h3>
  <p>Tenemos precios especiales por volumen para grupos hoteleros. Descuentos desde el 30% con facturación consolidada.</p>
  <a href="mailto:jordi@yve01.com?subject=Enterprise Yve.01">Hablar con el equipo →</a>
</div>
</div>

<div class="faq">
  <h2>Preguntas frecuentes</h2>
  <details class="faq-item"><summary>¿Puedo probarlo sin tarjeta de crédito?</summary>
    <p>Sí. Los 14 días de prueba son completamente gratis y no necesitas introducir ningún dato de pago. Si decides no continuar, no se te cobra nada.</p></details>
  <details class="faq-item"><summary>¿Qué pasa si supero el límite de habitaciones?</summary>
    <p>Te avisamos cuando te acercas al límite y puedes cambiar de plan en cualquier momento. Nunca bloqueamos el acceso de golpe.</p></details>
  <details class="faq-item"><summary>¿Funciona con Oracle Fusion?</summary>
    <p>En el plan Starter y Pro exportamos los asientos en formato Oracle GL Interface (Excel) para importación manual. En el plan Business ofrecemos integración directa via REST API.</p></details>
  <details class="faq-item"><summary>¿Puedo importar datos de mi sistema actual?</summary>
    <p>Sí. Aceptamos Excel, CSV, y PDF. El onboarding incluye una sesión de migración de datos históricos.</p></details>
  <details class="faq-item"><summary>¿Dónde están alojados mis datos?</summary>
    <p>En servidores en Europa (Frankfurt). Cumplimos con el RGPD. No vendemos datos ni los compartimos con terceros.</p></details>
  <details class="faq-item"><summary>¿Hay contrato de permanencia?</summary>
    <p>No. Pago mensual sin permanencia. Si pagas anual obtienes un 20% de descuento pero también puedes cancelar y te devolvemos los meses no usados.</p></details>
</div>

<script>
var annual = true;
var prices = {starter:[199,159],pro:[349,279],biz:[599,479]};

function toggleBilling() {
  annual = !annual;
  var pill = document.getElementById('billing-toggle');
  var saveBadge = document.getElementById('save-badge');
  var lblA = document.getElementById('lbl-annual');
  var lblM = document.getElementById('lbl-monthly');
  pill.classList.toggle('monthly', !annual);
  saveBadge.style.display = annual ? 'inline' : 'none';
  lblA.style.fontWeight = annual ? '700' : '400';
  lblA.style.color = annual ? 'var(--tx)' : 'var(--mut)';
  lblM.style.fontWeight = !annual ? '700' : '400';
  lblM.style.color = !annual ? 'var(--tx)' : 'var(--mut)';
  var idx = annual ? 0 : 1;
  document.getElementById('p-start').textContent = prices.starter[idx];
  document.getElementById('p-pro').textContent   = prices.pro[idx];
  document.getElementById('p-biz').textContent   = prices.biz[idx];
  var freq = annual ? 'Pago mensual (factura anual)' : 'Pago mensual';
  document.getElementById('ps-start').textContent = 'Hasta 80 habitaciones · ' + freq;
  document.getElementById('ps-pro').textContent   = 'Hasta 200 habitaciones · ' + freq;
  document.getElementById('ps-biz').textContent   = 'Habitaciones ilimitadas · ' + freq;
}
</script>
</body>
</html>"""

@pricing_bp.route('/precios')
@pricing_bp.route('/pricing')
def pricing_page():
    return Response(PRICING_HTML, mimetype='text/html')
