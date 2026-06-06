"""
landing.py — Yve.01 Public Landing Page
Ruta: / (pública si no hay sesión, dashboard si la hay)
"""

LANDING_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Yve — Sistema operativo AI-first para hoteles. Automatiza AP, AR, DRR y conciliación bancaria. Para hoteles independientes y grupos hoteleros en Europa.">
<title>Yve.01 — El sistema operativo para hoteles</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,400&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#0f172a;--s1:#1e293b;--s2:#334155;--s3:#475569;
  --acc:#3b82f6;--acc2:#60a5fa;--acc3:#93c5fd;
  --tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;
  --grn:#22c55e;--red:#ef4444;--ora:#f97316;--yel:#eab308;--pur:#8b5cf6;
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;line-height:1.6;-webkit-font-smoothing:antialiased}
a{color:inherit;text-decoration:none}
/* NAV */
.nav{position:sticky;top:0;z-index:100;background:rgba(15,23,42,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--s2);padding:0 5%;height:64px;display:flex;align-items:center;justify-content:space-between}
.nav-logo{display:flex;align-items:baseline;gap:8px}
.nav-logo .dot{width:8px;height:8px;border-radius:50%;background:var(--acc);box-shadow:0 0 8px var(--acc);margin-bottom:2px}
.nav-logo span{font-size:20px;font-weight:800;color:#fff;letter-spacing:-.5px}
.nav-logo .sub{font-size:11px;color:var(--mut);font-weight:400}
.nav-links{display:flex;align-items:center;gap:28px}
.nav-links a{font-size:14px;color:var(--mut);font-weight:500;transition:color .15s}
.nav-links a:hover{color:var(--tx)}
.nav-cta{display:flex;align-items:center;gap:12px}
.btn-outline{border:1px solid var(--s2);color:var(--tx);padding:8px 18px;border-radius:9px;font-size:14px;font-weight:500;transition:.15s;cursor:pointer;background:none}
.btn-outline:hover{border-color:var(--acc);color:var(--acc2)}
.btn-primary{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:9px 20px;border-radius:9px;font-size:14px;font-weight:700;box-shadow:0 0 20px rgba(59,130,246,.35);transition:.15s}
.btn-primary:hover{box-shadow:0 0 28px rgba(59,130,246,.55);transform:translateY(-1px)}
@media(max-width:768px){.nav-links{display:none}.nav-logo .sub{display:none}}
/* HERO */
.hero{padding:120px 5% 100px;text-align:center;position:relative;overflow:hidden}
.hero::before{content:'';position:absolute;inset:0;background:radial-gradient(900px 600px at 50% 0%,rgba(59,130,246,.15),transparent 60%),radial-gradient(600px 400px at 80% 80%,rgba(139,92,246,.08),transparent 55%);pointer-events:none}
.hero-badge{display:inline-flex;align-items:center;gap:8px;background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.25);border-radius:20px;padding:6px 16px;font-size:13px;color:var(--acc3);margin-bottom:28px}
.hero-badge .dot{width:6px;height:6px;border-radius:50%;background:var(--grn);box-shadow:0 0 6px var(--grn);animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.hero h1{font-size:clamp(36px,6vw,72px);font-weight:900;line-height:1.08;letter-spacing:-2px;margin-bottom:24px;max-width:900px;margin-left:auto;margin-right:auto}
.hero h1 .accent{background:linear-gradient(135deg,var(--acc2),var(--pur));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hero-sub{font-size:clamp(16px,2vw,21px);color:var(--mut);max-width:620px;margin:0 auto 40px;line-height:1.7;font-weight:400}
.hero-btns{display:flex;gap:14px;justify-content:center;flex-wrap:wrap}
.btn-hero-primary{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:14px 32px;border-radius:12px;font-size:16px;font-weight:700;box-shadow:0 4px 24px rgba(59,130,246,.45);transition:.2s}
.btn-hero-primary:hover{box-shadow:0 8px 36px rgba(59,130,246,.6);transform:translateY(-2px)}
.btn-hero-outline{border:1px solid var(--s2);color:var(--tx);padding:14px 32px;border-radius:12px;font-size:16px;font-weight:600;transition:.15s}
.btn-hero-outline:hover{border-color:var(--acc2);color:var(--acc2)}
/* STATS BAR */
.stats-bar{background:var(--s1);border-top:1px solid var(--s2);border-bottom:1px solid var(--s2);padding:32px 5%;display:flex;justify-content:center;gap:0}
.stat-item{text-align:center;padding:0 48px;border-right:1px solid var(--s2)}
.stat-item:last-child{border-right:none}
.stat-val{font-size:36px;font-weight:800;letter-spacing:-1px;color:var(--acc2)}
.stat-lbl{font-size:13px;color:var(--mut);margin-top:4px}
@media(max-width:768px){.stats-bar{flex-wrap:wrap;gap:24px}.stat-item{border-right:none;padding:0 24px}}
/* SECTIONS */
.section{padding:96px 5%}
.section-alt{background:rgba(30,41,59,.3)}
.section-label{font-size:12px;font-weight:700;color:var(--acc2);text-transform:uppercase;letter-spacing:.8px;margin-bottom:14px}
.section-title{font-size:clamp(28px,4vw,44px);font-weight:800;letter-spacing:-1px;margin-bottom:16px;line-height:1.15}
.section-sub{font-size:17px;color:var(--mut);max-width:560px;line-height:1.7}
.container{max-width:1200px;margin:0 auto}
/* PROBLEM */
.problem-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-top:56px}
.prob-card{background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px}
.prob-icon{font-size:28px;margin-bottom:16px}
.prob-title{font-size:17px;font-weight:700;margin-bottom:10px}
.prob-desc{font-size:14px;color:var(--mut);line-height:1.7}
@media(max-width:768px){.problem-grid{grid-template-columns:1fr}}
/* FEATURES */
.features-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:20px;margin-top:56px}
.feat-card{background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;transition:border-color .2s,transform .2s}
.feat-card:hover{border-color:rgba(59,130,246,.4);transform:translateY(-2px)}
.feat-tag{display:inline-block;background:rgba(59,130,246,.1);color:var(--acc2);border:1px solid rgba(59,130,246,.2);border-radius:6px;padding:3px 10px;font-size:11px;font-weight:700;letter-spacing:.4px;margin-bottom:14px}
.feat-title{font-size:20px;font-weight:700;margin-bottom:10px}
.feat-desc{font-size:14px;color:var(--mut);line-height:1.7;margin-bottom:16px}
.feat-list{list-style:none;display:flex;flex-direction:column;gap:7px}
.feat-list li{font-size:13px;color:var(--dim);display:flex;align-items:center;gap:8px}
.feat-list li::before{content:'✓';color:var(--grn);font-weight:700;flex-shrink:0}
@media(max-width:768px){.features-grid{grid-template-columns:1fr}}
/* ROI CALCULATOR */
.roi-wrap{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:40px;margin-top:56px;display:grid;grid-template-columns:1fr 1fr;gap:40px;align-items:center}
.roi-inputs{display:flex;flex-direction:column;gap:28px}
.roi-label{font-size:13px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
.roi-slider{width:100%;accent-color:var(--acc);height:4px}
.roi-val{font-size:22px;font-weight:800;color:var(--tx);margin-top:4px}
.roi-result{background:linear-gradient(135deg,rgba(59,130,246,.12),rgba(139,92,246,.08));border:1px solid rgba(59,130,246,.2);border-radius:16px;padding:32px;text-align:center}
.roi-saving{font-size:52px;font-weight:900;color:var(--grn);letter-spacing:-2px;line-height:1}
.roi-saving-lbl{font-size:16px;color:var(--mut);margin-top:8px;margin-bottom:24px}
.roi-breakdown{display:flex;flex-direction:column;gap:10px;text-align:left;margin-top:24px;padding-top:24px;border-top:1px solid var(--s2)}
.roi-row{display:flex;justify-content:space-between;font-size:14px}
.roi-row .k{color:var(--mut)}.roi-row .v{font-weight:700;color:var(--tx)}
@media(max-width:768px){.roi-wrap{grid-template-columns:1fr;padding:24px}.roi-saving{font-size:40px}}
/* PRICING */
.pricing-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-top:56px}
.price-card{background:var(--s1);border:1px solid var(--s2);border-radius:16px;padding:28px;position:relative}
.price-card.featured{border-color:var(--acc);background:linear-gradient(135deg,rgba(59,130,246,.08),rgba(30,41,59,1))}
.price-badge{position:absolute;top:-12px;left:50%;transform:translateX(-50%);background:var(--acc);color:#fff;font-size:11px;font-weight:700;padding:4px 14px;border-radius:20px;white-space:nowrap}
.price-name{font-size:14px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px}
.price-amount{font-size:44px;font-weight:900;letter-spacing:-2px;color:var(--tx);line-height:1}
.price-period{font-size:14px;color:var(--mut);margin-bottom:24px}
.price-desc{font-size:14px;color:var(--mut);margin-bottom:20px;padding-bottom:20px;border-bottom:1px solid var(--s2)}
.price-features{list-style:none;display:flex;flex-direction:column;gap:10px;margin-bottom:28px}
.price-features li{font-size:14px;color:var(--dim);display:flex;gap:9px}
.price-features li::before{content:'✓';color:var(--grn);font-weight:700;flex-shrink:0}
.price-btn{display:block;text-align:center;padding:12px;border-radius:10px;font-size:15px;font-weight:700;transition:.15s}
.price-btn.primary{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;box-shadow:0 4px 20px rgba(59,130,246,.3)}
.price-btn.primary:hover{box-shadow:0 6px 28px rgba(59,130,246,.5);transform:translateY(-1px)}
.price-btn.outline{border:1px solid var(--s2);color:var(--tx)}
.price-btn.outline:hover{border-color:var(--acc);color:var(--acc2)}
@media(max-width:900px){.pricing-grid{grid-template-columns:1fr;max-width:420px;margin-left:auto;margin-right:auto}}
/* COMPARISON */
.comp-table{width:100%;border-collapse:collapse;margin-top:56px;font-size:14px}
.comp-table th{background:var(--s1);padding:14px 20px;font-size:11px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;text-align:left;border-bottom:2px solid var(--s2)}
.comp-table th:first-child{width:200px}
.comp-table td{padding:14px 20px;border-bottom:1px solid rgba(51,65,85,.4);vertical-align:middle}
.comp-table tr:last-child td{border-bottom:none}
.comp-table tr:hover td{background:rgba(255,255,255,.02)}
.comp-table td:first-child{font-weight:600;color:var(--tx)}
.comp-table .yve{color:var(--grn);font-weight:700}
.comp-table .no{color:var(--dim)}
/* CTA BANNER */
.cta-banner{background:linear-gradient(135deg,rgba(59,130,246,.15),rgba(139,92,246,.1));border:1px solid rgba(59,130,246,.2);border-radius:24px;padding:64px 48px;text-align:center;margin:0 5%}
.cta-banner h2{font-size:clamp(28px,4vw,44px);font-weight:800;letter-spacing:-1px;margin-bottom:16px}
.cta-banner p{font-size:18px;color:var(--mut);margin-bottom:36px}
/* FOOTER */
footer{background:var(--s1);border-top:1px solid var(--s2);padding:48px 5% 32px;margin-top:96px}
.footer-grid{display:grid;grid-template-columns:2fr 1fr 1fr;gap:40px;margin-bottom:40px}
.footer-brand p{font-size:14px;color:var(--mut);margin-top:12px;max-width:300px;line-height:1.7}
.footer-col h4{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);margin-bottom:14px}
.footer-col a{display:block;font-size:14px;color:var(--dim);margin-bottom:9px;transition:color .15s}
.footer-col a:hover{color:var(--tx)}
.footer-bottom{border-top:1px solid var(--s2);padding-top:24px;display:flex;justify-content:space-between;align-items:center;font-size:13px;color:var(--dim)}
@media(max-width:768px){.footer-grid{grid-template-columns:1fr}}
</style>
</head>
<body>

<!-- NAV -->
<nav class="nav">
  <div class="nav-logo">
    <div class="dot"></div>
    <span>Yve<span style="color:var(--acc2)">.01</span></span>
    <span class="sub">Beta</span>
  </div>
  <div class="nav-links">
    <a href="#features">Funciones</a>
    <a href="#roi">ROI</a>
    <a href="#pricing">Precios</a>
    <a href="#comparativa">Comparativa</a>
  </div>
  <div class="nav-cta">
    <a href="/login" class="btn-outline">Iniciar sesión</a>
    <a href="/login" class="btn-primary">Ver demo →</a>
  </div>
</nav>

<!-- HERO -->
<section class="hero">
  <div class="hero-badge">
    <div class="dot"></div>
    Nuevo · Yve.01 Beta disponible para hoteles independientes
  </div>
  <h1>El sistema operativo<br><span class="accent">AI-first para hoteles</span></h1>
  <p class="hero-sub">Automatiza AP, AR, DRR y conciliación bancaria. Tu equipo financiero multiplica su capacidad sin ampliar plantilla.</p>
  <div class="hero-btns">
    <a href="/login" class="btn-hero-primary">Empezar gratis 14 días</a>
    <a href="#roi" class="btn-hero-outline">Calcular mi ahorro →</a>
  </div>
</section>

<!-- STATS BAR -->
<div class="stats-bar">
  <div class="stat-item"><div class="stat-val">~150</div><div class="stat-lbl">Facturas/mes automatizadas</div></div>
  <div class="stat-item"><div class="stat-val">8h</div><div class="stat-lbl">Ahorradas a la semana</div></div>
  <div class="stat-item"><div class="stat-val">3→1</div><div class="stat-lbl">Sistemas reemplazados</div></div>
  <div class="stat-item"><div class="stat-val">0€</div><div class="stat-lbl">Setup. Sin consultores.</div></div>
</div>

<!-- PROBLEM -->
<section class="section">
  <div class="container">
    <div class="section-label">El problema</div>
    <h2 class="section-title">Tu equipo financiero<br>trabaja en silos</h2>
    <p class="section-sub">PMS, POS, OTAs, extracto bancario y Oracle no se comunican entre sí. Alguien dedica el día a cruzar datos manualmente.</p>
    <div class="problem-grid">
      <div class="prob-card">
        <div class="prob-icon">📧</div>
        <div class="prob-title">Facturas en 3 canales distintos</div>
        <div class="prob-desc">Email, correo físico, portales de proveedores. Cada una llega diferente y hay que adjuntarla manualmente al expediente.</div>
      </div>
      <div class="prob-card">
        <div class="prob-icon">📄</div>
        <div class="prob-title">Albaranes en papel</div>
        <div class="prob-desc">El matching de factura + PO + albarán físico se hace a mano. Un error y se paga lo que no corresponde.</div>
      </div>
      <div class="prob-card">
        <div class="prob-icon">🔄</div>
        <div class="prob-title">Nada habla con Oracle</div>
        <div class="prob-desc">Contabilizar una factura en el ERP requiere entrar datos dos veces. Nadie tiene visibilidad en tiempo real.</div>
      </div>
    </div>
  </div>
</section>

<!-- FEATURES -->
<section class="section section-alt" id="features">
  <div class="container">
    <div class="section-label">La solución</div>
    <h2 class="section-title">Un dashboard.<br>Todo conectado.</h2>
    <p class="section-sub">Yve integra todos los flujos financieros en una sola plataforma con IA en el núcleo, no añadida encima.</p>
    <div class="features-grid">
      <div class="feat-card">
        <span class="feat-tag">AP — PROVEEDORES</span>
        <h3 class="feat-title">3-way matching automático</h3>
        <p class="feat-desc">Lee facturas PDF via IA, cruza automáticamente con PO y POS del restaurante, detecta discrepancias y genera el email de reclamación.</p>
        <ul class="feat-list">
          <li>Lectura de PDF escaneados y digitales</li>
          <li>Matching F&B: factura + albarán + POS</li>
          <li>Contabilización directa en Oracle GL</li>
          <li>Flujo de aprobación por departamento</li>
        </ul>
      </div>
      <div class="feat-card">
        <span class="feat-tag">AR — OTAs</span>
        <h3 class="feat-title">Verificación de comisiones OTA</h3>
        <p class="feat-desc">Comprueba que Booking, Expedia y el resto cobren exactamente las tarifas pactadas. Detecta facturas extranjeras que necesitan certificado de doble imposición.</p>
        <ul class="feat-list">
          <li>Verificación vs tabla de comisiones pactadas</li>
          <li>Detección automática de certificados DI</li>
          <li>Generación de emails de reclamación</li>
          <li>Clientes corporativos y grupos</li>
        </ul>
      </div>
      <div class="feat-card">
        <span class="feat-tag">DRR — DAILY REVENUE REPORT</span>
        <h3 class="feat-title">Análisis de revenue diario</h3>
        <p class="feat-desc">Procesa tu DRR en formato .xlsm directamente. Extrae KPIs, detecta días Out of Balance y visualiza la tendencia de revenue mes a mes.</p>
        <ul class="feat-list">
          <li>Occupancy, ADR, RevPAR, GOP en tiempo real</li>
          <li>Detección automática de Out of Balance</li>
          <li>Gráfico de revenue diario con tendencia</li>
          <li>Mapeo al Plan de Cuentas Oracle</li>
        </ul>
      </div>
      <div class="feat-card">
        <span class="feat-tag">MULTI-HOTEL</span>
        <h3 class="feat-title">Vista consolidada del grupo</h3>
        <p class="feat-desc">Para grupos de 2 a 100 hoteles: KPIs consolidados, ranking de performance, alertas activas y GOP% comparativo entre propiedades.</p>
        <ul class="feat-list">
          <li>Dashboard consolidado multi-propiedad</li>
          <li>F&B Cost % real vs teórico por hotel</li>
          <li>Rankings y benchmarking interno</li>
          <li>Conciliación bancaria automática</li>
        </ul>
      </div>
    </div>
  </div>
</section>

<!-- ROI CALCULATOR -->
<section class="section" id="roi">
  <div class="container">
    <div class="section-label">Calculadora de ROI</div>
    <h2 class="section-title">¿Cuánto ahorras<br>con Yve?</h2>
    <p class="section-sub">Ajusta los sliders con los datos de tu hotel y calcula el retorno mensual.</p>
    <div class="roi-wrap">
      <div class="roi-inputs">
        <div>
          <div class="roi-label">Hoteles del grupo</div>
          <input type="range" class="roi-slider" id="sl-hotels" min="1" max="20" value="1" oninput="calcROI()">
          <div class="roi-val" id="v-hotels">1 hotel</div>
        </div>
        <div>
          <div class="roi-label">Facturas AP al mes</div>
          <input type="range" class="roi-slider" id="sl-inv" min="30" max="400" step="10" value="100" oninput="calcROI()">
          <div class="roi-val" id="v-inv">100 facturas</div>
        </div>
        <div>
          <div class="roi-label">Horas semanales en tareas manuales</div>
          <input type="range" class="roi-slider" id="sl-hours" min="2" max="40" value="10" oninput="calcROI()">
          <div class="roi-val" id="v-hours">10 horas/semana</div>
        </div>
        <div>
          <div class="roi-label">Coste hora del equipo financiero</div>
          <input type="range" class="roi-slider" id="sl-cost" min="15" max="60" value="25" oninput="calcROI()">
          <div class="roi-val" id="v-cost">25 €/hora</div>
        </div>
      </div>
      <div class="roi-result">
        <div class="roi-saving" id="roi-saving">€800</div>
        <div class="roi-saving-lbl">ahorro neto al mes</div>
        <div style="font-size:30px;font-weight:800;color:var(--acc2);margin-top:8px" id="roi-pct">+167%</div>
        <div style="font-size:13px;color:var(--mut)">ROI sobre el precio de Yve</div>
        <div class="roi-breakdown">
          <div class="roi-row"><span class="k">Ahorro en horas</span><span class="v" id="roi-save">€1.200/mes</span></div>
          <div class="roi-row"><span class="k">Precio Yve</span><span class="v" id="roi-price">€400/mes</span></div>
          <div class="roi-row"><span class="k">Payback</span><span class="v" style="color:var(--grn)">Inmediato</span></div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- PRICING -->
<section class="section section-alt" id="pricing">
  <div class="container">
    <div class="section-label">Precios</div>
    <h2 class="section-title">Sin sorpresas.<br>Sin consultores.</h2>
    <p class="section-sub">Precio fijo mensual por hotel. Sin costes de implementación ni formación.</p>
    <div class="pricing-grid">
      <div class="price-card">
        <div class="price-name">Starter</div>
        <div class="price-amount">€400</div>
        <div class="price-period">/mes · 1 hotel</div>
        <div class="price-desc">Para hoteles independientes que quieren empezar a automatizar finanzas.</div>
        <ul class="price-features">
          <li>Módulo AP — Proveedores</li>
          <li>Módulo AR — OTAs</li>
          <li>DRR & Conciliación bancaria</li>
          <li>Oracle GL (simulación)</li>
          <li>Soporte email</li>
        </ul>
        <a href="/login" class="price-btn outline">Empezar gratis</a>
      </div>
      <div class="price-card featured">
        <div class="price-badge">Más popular</div>
        <div class="price-name">Pro</div>
        <div class="price-amount">€600</div>
        <div class="price-period">/mes · 1 hotel</div>
        <div class="price-desc">Para hoteles con restaurante o que quieren integración completa Oracle.</div>
        <ul class="price-features">
          <li>Todo lo de Starter</li>
          <li>F&B Cost Control avanzado</li>
          <li>Oracle GL API en producción</li>
          <li>AR Real — Grupos corporativos</li>
          <li>Notificaciones (email, WhatsApp)</li>
          <li>Soporte prioritario</li>
        </ul>
        <a href="/login" class="price-btn primary">Empezar gratis</a>
      </div>
      <div class="price-card">
        <div class="price-name">Multi-Hotel</div>
        <div class="price-amount">€400</div>
        <div class="price-period">/mes por hotel · mín. 2</div>
        <div class="price-desc">Para grupos hoteleros. Dashboard consolidado incluido sin coste adicional.</div>
        <ul class="price-features">
          <li>Todo lo de Pro en cada hotel</li>
          <li>Dashboard Multi-Hotel consolidado</li>
          <li>Benchmarking entre propiedades</li>
          <li>20% dto pago anual</li>
          <li>Gestor de cuenta dedicado</li>
        </ul>
        <a href="/login" class="price-btn outline">Hablar con ventas</a>
      </div>
    </div>
    <p style="text-align:center;color:var(--dim);font-size:13px;margin-top:24px">14 días gratis. Sin tarjeta de crédito. Cancela cuando quieras.</p>
  </div>
</section>

<!-- COMPARATIVA -->
<section class="section" id="comparativa">
  <div class="container">
    <div class="section-label">Comparativa</div>
    <h2 class="section-title">Por qué Yve<br>es diferente</h2>
    <p class="section-sub">Los competidores existen, pero ninguno integra todo con IA a precio accesible para hoteles independientes.</p>
    <div style="overflow-x:auto;margin-top:56px">
    <table class="comp-table">
      <thead>
        <tr>
          <th>Funcionalidad</th>
          <th class="yve">Yve.01</th>
          <th>Oracle Hospitality</th>
          <th>M3 / Aptech</th>
          <th>Rillion</th>
          <th>BlackLine</th>
        </tr>
      </thead>
      <tbody>
        <tr><td>AP automatizado con IA</td><td class="yve">✓</td><td class="no">—</td><td class="no">—</td><td>Parcial</td><td class="no">—</td></tr>
        <tr><td>3-way matching F&B</td><td class="yve">✓</td><td class="no">—</td><td class="no">—</td><td class="no">—</td><td class="no">—</td></tr>
        <tr><td>Verificación comisiones OTA</td><td class="yve">✓</td><td class="no">—</td><td class="no">—</td><td class="no">—</td><td class="no">—</td></tr>
        <tr><td>DRR procesamiento .xlsm</td><td class="yve">✓</td><td class="no">—</td><td>Parcial</td><td class="no">—</td><td class="no">—</td></tr>
        <tr><td>Integración Oracle GL</td><td class="yve">✓</td><td class="yve">✓</td><td>Parcial</td><td class="no">—</td><td>✓</td></tr>
        <tr><td>Multi-hotel dashboard</td><td class="yve">✓</td><td>✓</td><td>✓</td><td class="no">—</td><td>✓</td></tr>
        <tr><td>Setup en minutos</td><td class="yve">✓</td><td class="no">Meses</td><td class="no">Semanas</td><td class="no">Semanas</td><td class="no">Meses</td></tr>
        <tr><td>Precio hotel independiente</td><td class="yve">€400/mes</td><td class="no">€15K+/año</td><td class="no">€8K+/año</td><td class="no">€6K+/año</td><td class="no">Enterprise</td></tr>
        <tr><td>Disponible en España</td><td class="yve">✓</td><td>✓</td><td class="no">—</td><td class="no">—</td><td>Parcial</td></tr>
      </tbody>
    </table>
    </div>
  </div>
</section>

<!-- CTA BANNER -->
<section class="section">
  <div class="container">
    <div class="cta-banner">
      <h2>¿Listo para automatizar<br>tus finanzas hoteleras?</h2>
      <p>14 días gratis. Sin tarjeta. Setup en 15 minutos.</p>
      <div style="display:flex;gap:14px;justify-content:center;flex-wrap:wrap">
        <a href="/login" class="btn-hero-primary">Empezar gratis →</a>
        <a href="mailto:jordi@yve01.com" class="btn-hero-outline">Hablar con el equipo</a>
      </div>
    </div>
  </div>
</section>

<!-- FOOTER -->
<footer>
  <div class="container">
    <div class="footer-grid">
      <div class="footer-brand">
        <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:4px">
          <div style="width:7px;height:7px;border-radius:50%;background:var(--acc);box-shadow:0 0 6px var(--acc)"></div>
          <span style="font-size:18px;font-weight:800">Yve<span style="color:var(--acc2)">.01</span></span>
        </div>
        <p>Sistema operativo AI-first para la gestión financiera hotelera. Construido en Barcelona para hoteles europeos.</p>
      </div>
      <div class="footer-col">
        <h4>Producto</h4>
        <a href="#features">Funciones</a>
        <a href="#pricing">Precios</a>
        <a href="#roi">Calculadora ROI</a>
        <a href="#comparativa">Comparativa</a>
        <a href="/login">Demo</a>
      </div>
      <div class="footer-col">
        <h4>Contacto</h4>
        <a href="mailto:jordi@yve01.com">jordi@yve01.com</a>
        <a href="https://github.com/Jordi-Solano/yve01">GitHub</a>
        <a href="/login">Panel de acceso</a>
      </div>
    </div>
    <div class="footer-bottom">
      <span>© 2026 Yve.01 — Barcelona, España</span>
      <span>Hecho con IA · Validado con hoteles reales</span>
    </div>
  </div>
</footer>

<script>
function calcROI() {
  const hotels = parseInt(document.getElementById('sl-hotels').value);
  const inv    = parseInt(document.getElementById('sl-inv').value);
  const hours  = parseInt(document.getElementById('sl-hours').value);
  const cost   = parseInt(document.getElementById('sl-cost').value);
  document.getElementById('v-hotels').textContent = hotels + (hotels === 1 ? ' hotel' : ' hoteles');
  document.getElementById('v-inv').textContent    = inv + ' facturas';
  document.getElementById('v-hours').textContent  = hours + ' horas/semana';
  document.getElementById('v-cost').textContent   = cost + ' €/hora';
  const save = hours * 4 * cost;
  const price = hotels <= 1 ? 400 : hotels <= 5 ? hotels * 400 : hotels * 350;
  const net = save - price;
  const roi = price > 0 ? Math.round((net / price) * 100) : 0;
  document.getElementById('roi-saving').textContent = '€' + net.toLocaleString('es-ES');
  document.getElementById('roi-pct').textContent    = (roi >= 0 ? '+' : '') + roi + '%';
  document.getElementById('roi-save').textContent   = '€' + save.toLocaleString('es-ES') + '/mes';
  document.getElementById('roi-price').textContent  = '€' + price.toLocaleString('es-ES') + '/mes';
  document.getElementById('roi-saving').style.color = net >= 0 ? 'var(--grn)' : 'var(--red)';
}
calcROI();
</script>
</body>
</html>"""
