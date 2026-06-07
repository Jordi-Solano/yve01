"""
billing.py — Yve Stripe Billing
Rutas: /checkout/<plan>, /billing/success, /billing/cancel, /billing/webhook
Modo simulación cuando no hay STRIPE_SECRET_KEY configurado.

Para activar Stripe real, añade en Render → Environment:
  STRIPE_SECRET_KEY      = sk_live_...  (Dashboard Stripe → Developers → API Keys)
  STRIPE_PRICE_STARTER   = price_...   (Dashboard Stripe → Products → Starter → Price ID)
  STRIPE_PRICE_PRO       = price_...
  STRIPE_PRICE_MULTI     = price_...
  STRIPE_WEBHOOK_SECRET  = whsec_...   (Dashboard Stripe → Webhooks → Signing secret)
  STRIPE_MODE            = live        (o "test" para pruebas)
"""
import os, json
from pathlib import Path
from flask import Blueprint, request, redirect, Response, jsonify

billing_bp = Blueprint('billing', __name__)

# Precio base: ~2€/habitación/mes (100 hab ≈ €200, 300 hab ≈ €600)
# Para simplificar el checkout usamos 3 tiers flat.
# En el futuro: precio dinámico según habitaciones registradas en onboarding.
PLANS = {
    'starter': {'name':'Starter',     'price_eur':400, 'stripe_price_id':os.environ.get('STRIPE_PRICE_STARTER',''),
                'desc':'Hasta 150 hab · AP + AR + DRR + Banco',
                'features':['Módulo AP — Proveedores','Módulo AR — OTAs','DRR & Conciliación bancaria',
                            'Hasta 150 habitaciones','Soporte por email']},
    'pro':     {'name':'Pro',         'price_eur':600, 'stripe_price_id':os.environ.get('STRIPE_PRICE_PRO',''),
                'desc':'Hasta 400 hab · Todo incluido + Oracle',
                'features':['Todo lo de Starter','F&B Cost Control','Oracle GL API producción',
                            'Hasta 400 habitaciones','Notificaciones WhatsApp/Slack','Soporte prioritario']},
    'multi':   {'name':'Multi-Hotel', 'price_eur':400, 'stripe_price_id':os.environ.get('STRIPE_PRICE_MULTI',''),
                'desc':'Por hotel · Mín. 2 hoteles · Dashboard consolidado',
                'features':['Todo lo de Pro en cada hotel','Sin límite de habitaciones',
                            'Dashboard Multi-Hotel','Benchmarking propiedades','Gestor de cuenta dedicado']},
}

def _stripe():
    key = os.environ.get('STRIPE_SECRET_KEY','')
    if not key: return None
    try:
        import stripe as _s; _s.api_key = key; return _s
    except ImportError: return None

def _is_test_mode():
    key = os.environ.get('STRIPE_SECRET_KEY','')
    return key.startswith('sk_test_')


@billing_bp.route('/checkout/<plan>')
def checkout(plan):
    if plan not in PLANS: return redirect('/#pricing')
    stripe = _stripe()
    price_id = PLANS[plan]['stripe_price_id']

    if not stripe or not price_id:
        return Response(_sim_html(plan), mimetype='text/html')

    try:
        sess = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{'price': price_id, 'quantity':1}],
            mode='subscription',
            success_url=request.host_url.rstrip('/') + '/billing/success?session_id={CHECKOUT_SESSION_ID}',
            cancel_url=request.host_url.rstrip('/') + '/billing/cancel',
            locale='es',
            metadata={'plan': plan},
            allow_promotion_codes=True,
        )
        return redirect(sess.url, code=303)
    except Exception as e:
        return Response(_err_html(str(e)), mimetype='text/html')


@billing_bp.route('/billing/success')
def success():
    session_id = request.args.get('session_id','')
    # Try to get customer email for personalization
    email = ''
    stripe = _stripe()
    if stripe and session_id:
        try:
            sess  = stripe.checkout.Session.retrieve(session_id)
            email = sess.get('customer_details',{}).get('email','')
        except Exception: pass
    return Response(_success_html(email), mimetype='text/html')


@billing_bp.route('/billing/cancel')
def cancel():
    return redirect('/#pricing')


@billing_bp.route('/billing/webhook', methods=['POST'])
def webhook():
    stripe = _stripe()
    secret = os.environ.get('STRIPE_WEBHOOK_SECRET','')
    payload = request.get_data()
    sig = request.headers.get('Stripe-Signature','')
    if not stripe or not secret:
        return jsonify({'received': True})
    try:
        event = stripe.Webhook.construct_event(payload, sig, secret)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    if event['type'] == 'checkout.session.completed':
        s    = event['data']['object']
        plan = s.get('metadata',{}).get('plan','unknown')
        mail = s.get('customer_details',{}).get('email','')
        sub  = s.get('subscription','')
        print(f"[BILLING] ✅ New {plan} subscription — {mail} — {sub}")
        _activate_account(mail, plan)
    elif event['type'] == 'customer.subscription.deleted':
        sub_id = event['data']['object']['id']
        print(f"[BILLING] ❌ Subscription cancelled: {sub_id}")

    return jsonify({'received': True})


def _activate_account(email, plan):
    """Activa o actualiza el plan del usuario con este email."""
    try:
        users_path = Path(__file__).parent / 'datos-referencia' / 'usuarios.json'
        users = json.loads(users_path.read_text())
        for u in users:
            if u.get('email','').lower() == email.lower():
                u['plan'] = plan
                u['activo'] = True
                users_path.write_text(json.dumps(users, indent=2, ensure_ascii=False))
                print(f"[BILLING] ✅ Account {email} activated with plan {plan}")
                return
    except Exception as e:
        print(f"[BILLING] Warning: could not activate account: {e}")


# ── HTML helpers ──────────────────────────────────────────────────────────

_HEAD = """<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle cx='16' cy='16' r='9' fill='%233b82f6'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--acc:#3b82f6;--acc2:#60a5fa;--tx:#f1f5f9;--mut:#94a3b8;--dim:#64748b;--grn:#22c55e}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;position:relative;overflow-x:hidden;-webkit-font-smoothing:antialiased}
body::before{content:"";position:fixed;inset:0;z-index:0;pointer-events:none;background:radial-gradient(800px 400px at 80% -15%,rgba(59,130,246,.12),transparent 55%)}
.card{position:relative;z-index:1;background:linear-gradient(160deg,rgba(30,41,59,.96),rgba(10,18,35,.97));border:1px solid var(--s2);border-radius:22px;padding:40px 36px;max-width:460px;width:100%;box-shadow:0 28px 80px rgba(0,0,0,.55);animation:rise .4s cubic-bezier(.2,.8,.2,1)}
@keyframes rise{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:none}}
.logo{display:flex;align-items:center;gap:8px;margin-bottom:28px}
.logo .dot{width:9px;height:9px;border-radius:50%;background:var(--acc);box-shadow:0 0 10px var(--acc)}
.logo .name{font-size:20px;font-weight:800;letter-spacing:-.5px}.logo .name span{color:var(--acc2)}
.plan-name{font-size:15px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.price{font-size:52px;font-weight:900;letter-spacing:-2px;color:var(--tx);line-height:1}
.period{font-size:14px;color:var(--mut);margin:6px 0 18px}
.desc{font-size:14px;color:var(--mut);margin-bottom:22px;padding-bottom:20px;border-bottom:1px solid var(--s2)}
.features{list-style:none;display:flex;flex-direction:column;gap:9px;margin-bottom:28px}
.features li{font-size:14px;color:var(--dim);display:flex;gap:9px;align-items:flex-start}
.features li::before{content:"✓";color:var(--grn);font-weight:700;flex-shrink:0;margin-top:1px}
.btn{display:block;text-align:center;padding:14px;border-radius:12px;font-size:15px;font-weight:700;transition:.15s;text-decoration:none;margin-bottom:10px}
.btn-primary{background:linear-gradient(135deg,var(--acc),#1d4ed8);color:#fff;box-shadow:0 4px 22px rgba(59,130,246,.4)}
.btn-primary:hover{box-shadow:0 8px 32px rgba(59,130,246,.6);transform:translateY(-1px)}
.btn-outline{border:1px solid var(--s2);color:var(--mut)}
.btn-outline:hover{border-color:var(--acc2);color:var(--acc2)}
.notice{background:rgba(59,130,246,.07);border:1px solid rgba(59,130,246,.18);border-radius:10px;padding:14px 16px;font-size:13px;color:var(--acc2);margin-bottom:20px;line-height:1.6}
.test-badge{display:inline-block;background:rgba(251,191,36,.1);border:1px solid rgba(251,191,36,.3);color:#fbbf24;border-radius:8px;padding:3px 10px;font-size:11px;font-weight:700;margin-bottom:16px}
</style></head>"""

def _sim_html(plan):
    p = PLANS[plan]
    test = _is_test_mode()
    features = ''.join(f'<li>{f}</li>' for f in p['features'])
    mail_subject = f"Quiero%20el%20plan%20{p['name']}"
    mail_body = f"Hola%2C%20quiero%20contratar%20Yve%20plan%20{p['name']}%20(€{p['price_eur']}%2Fmes)."
    return _HEAD + f"""
<body><div class="card">
  <div class="logo"><div class="dot"></div><span class="name">Yve<span>.01</span></span></div>
  {"<div class='test-badge'>🧪 Test Mode</div>" if test else ""}
  <div class="plan-name">Plan {p['name']}</div>
  <div class="price">€{p['price_eur']}</div>
  <div class="period">/mes · sin permanencia</div>
  <div class="desc">{p['desc']}</div>
  <ul class="features">{features}</ul>
  <div class="notice">
    <strong>Pago en vivo próximamente.</strong><br>
    Para activar tu cuenta ahora, escríbenos a <strong>jordi@yve01.com</strong> — te respondemos en 24h.
  </div>
  <a href="mailto:jordi@yve01.com?subject={mail_subject}&body={mail_body}" class="btn btn-primary">Contactar para activar →</a>
  <a href="/#pricing" class="btn btn-outline">← Ver todos los planes</a>
</div></body></html>"""

def _success_html(email=''):
    greeting = f'Bienvenido, <strong>{email}</strong>.' if email else 'Tu suscripción está activa.'
    return _HEAD + f"""
<body><div class="card" style="text-align:center">
  <div style="width:64px;height:64px;background:rgba(34,197,94,.15);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:28px;margin:0 auto 22px">✓</div>
  <h1 style="font-size:26px;font-weight:800;margin-bottom:10px;letter-spacing:-.5px">¡Suscripción activada!</h1>
  <p style="color:var(--mut);font-size:15px;line-height:1.7;margin-bottom:28px">{greeting}<br>Recibirás un email de confirmación en breve.</p>
  <a href="/login" class="btn btn-primary" style="margin-bottom:0">Acceder al panel →</a>
</div></body></html>"""

def _err_html(msg):
    return _HEAD + f"""
<body><div class="card" style="text-align:center">
  <div style="font-size:36px;margin-bottom:16px">⚠️</div>
  <h2 style="font-size:20px;font-weight:700;margin-bottom:10px">Error en el pago</h2>
  <p style="color:var(--mut);font-size:14px;margin-bottom:24px">{msg[:180]}</p>
  <a href="/#pricing" class="btn btn-outline">← Volver a precios</a>
</div></body></html>"""
