"""
billing.py — Yve Stripe Billing
Rutas: /pricing, /checkout/<plan>, /billing/success, /billing/cancel, /billing/webhook
Planes: starter (400€), pro (600€), multi (400€/hotel)
"""
import os, json
from flask import Blueprint, request, redirect, Response, jsonify, session

billing_bp = Blueprint('billing', __name__)

# Stripe price IDs — se configuran en .env
# Crear en Stripe Dashboard → Products → Add product
PLANS = {
    'starter': {
        'name': 'Starter',
        'price_eur': 400,
        'stripe_price_id': os.environ.get('STRIPE_PRICE_STARTER', ''),
        'desc': '1 hotel · AP + AR + DRR + Banco',
    },
    'pro': {
        'name': 'Pro',
        'price_eur': 600,
        'stripe_price_id': os.environ.get('STRIPE_PRICE_PRO', ''),
        'desc': '1 hotel · Todo incluido + Oracle',
    },
    'multi': {
        'name': 'Multi-Hotel',
        'price_eur': 400,
        'stripe_price_id': os.environ.get('STRIPE_PRICE_MULTI', ''),
        'desc': 'Por hotel · Mín. 2 hoteles',
    },
}

def _stripe_client():
    """Devuelve cliente Stripe o None si no hay clave."""
    key = os.environ.get('STRIPE_SECRET_KEY', '')
    if not key:
        return None
    try:
        import stripe
        stripe.api_key = key
        return stripe
    except ImportError:
        return None

@billing_bp.route('/checkout/<plan>')
def checkout(plan):
    if plan not in PLANS:
        return redirect('/#pricing')

    stripe = _stripe_client()
    price_id = PLANS[plan]['stripe_price_id']

    # Modo simulación si no hay Stripe configurado
    if not stripe or not price_id:
        return Response(CHECKOUT_SIM_HTML.format(
            plan_name=PLANS[plan]['name'],
            plan_price=PLANS[plan]['price_eur'],
            plan_desc=PLANS[plan]['desc'],
        ), mimetype='text/html')

    # Stripe Checkout real
    try:
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{'price': price_id, 'quantity': 1}],
            mode='subscription',
            success_url=request.host_url + 'billing/success?session_id={CHECKOUT_SESSION_ID}',
            cancel_url=request.host_url + 'billing/cancel',
            locale='es',
            metadata={'plan': plan},
        )
        return redirect(checkout_session.url, code=303)
    except Exception as e:
        return Response(f'<p>Error: {e}</p>', status=500, mimetype='text/html')


@billing_bp.route('/billing/success')
def success():
    return Response(SUCCESS_HTML, mimetype='text/html')


@billing_bp.route('/billing/cancel')
def cancel():
    return redirect('/#pricing')


@billing_bp.route('/billing/webhook', methods=['POST'])
def webhook():
    stripe = _stripe_client()
    webhook_secret = os.environ.get('STRIPE_WEBHOOK_SECRET', '')
    payload = request.get_data()
    sig = request.headers.get('Stripe-Signature', '')

    if not stripe or not webhook_secret:
        return jsonify({'received': True})

    try:
        event = stripe.Webhook.construct_event(payload, sig, webhook_secret)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    if event['type'] == 'checkout.session.completed':
        session_obj = event['data']['object']
        plan = session_obj.get('metadata', {}).get('plan', 'unknown')
        customer_email = session_obj.get('customer_details', {}).get('email', '')
        print(f"[BILLING] New subscription: {plan} — {customer_email}")
        # TODO: activar cuenta en usuarios.json

    return jsonify({'received': True})


# ── Checkout simulación (sin Stripe configurado) ─────────────────────────────
CHECKOUT_SIM_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Checkout — Yve.01</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;-webkit-font-smoothing:antialiased}}
.card{{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:40px;max-width:440px;width:100%;text-align:center}}
.logo{{font-size:22px;font-weight:800;margin-bottom:24px}}
.logo span{{color:var(--acc2)}}
.plan-name{{font-size:28px;font-weight:800;margin-bottom:6px}}
.plan-price{{font-size:48px;font-weight:900;color:var(--acc2);letter-spacing:-2px;margin-bottom:4px}}
.plan-period{{font-size:14px;color:var(--mut);margin-bottom:16px}}
.plan-desc{{font-size:14px;color:var(--mut);margin-bottom:32px;padding-bottom:28px;border-bottom:1px solid var(--s2)}}
.notice{{background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.2);border-radius:10px;padding:16px;font-size:13px;color:var(--acc2);margin-bottom:28px;line-height:1.6}}
.btn{{display:block;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:14px;border-radius:12px;font-size:16px;font-weight:700;text-decoration:none;margin-bottom:12px;box-shadow:0 4px 20px rgba(59,130,246,.35)}}
.btn-out{{display:block;border:1px solid var(--s2);color:var(--mut);padding:13px;border-radius:12px;font-size:14px;text-decoration:none}}
</style>
</head>
<body>
<div class="card">
  <div class="logo">Yve<span>.01</span></div>
  <div class="plan-name">Plan {plan_name}</div>
  <div class="plan-price">€{plan_price}</div>
  <div class="plan-period">/mes · sin permanencia</div>
  <div class="plan-desc">{plan_desc}</div>
  <div class="notice">
    <strong>Pago en vivo próximamente.</strong><br>
    Stripe está configurándose. Para activar tu cuenta ahora,
    escríbenos a <strong>jordi@yve01.com</strong> y lo gestionamos en 24h.
  </div>
  <a href="mailto:jordi@yve01.com?subject=Quiero%20el%20plan%20{plan_name}&body=Hola%2C%20quiero%20contratar%20Yve%20plan%20{{plan_name}}%20(€{plan_price}%2Fmes)." class="btn">Contactar para activar →</a>
  <a href="/#pricing" class="btn-out">← Volver a precios</a>
</div>
</body></html>"""

SUCCESS_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>¡Bienvenido a Yve! — Suscripción activada</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700;800&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--grn:#22c55e}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;-webkit-font-smoothing:antialiased}
.card{background:var(--s1);border:1px solid var(--s2);border-radius:20px;padding:48px 40px;max-width:440px;width:100%;text-align:center}
.check{width:64px;height:64px;background:rgba(34,197,94,.15);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 24px;font-size:28px}
h1{font-size:28px;font-weight:800;margin-bottom:12px}
p{font-size:15px;color:var(--mut);line-height:1.7;margin-bottom:28px}
.btn{display:block;background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;padding:14px;border-radius:12px;font-size:16px;font-weight:700;text-decoration:none;box-shadow:0 4px 20px rgba(59,130,246,.35)}
</style>
</head>
<body>
<div class="card">
  <div class="check">✓</div>
  <h1>¡Suscripción activada!</h1>
  <p>Bienvenido a Yve. Recibirás un email con tus credenciales de acceso en los próximos minutos.</p>
  <a href="/login" class="btn">Acceder al panel →</a>
</div>
</body></html>"""
