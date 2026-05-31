"""Parche para añadir el chat AI al dashboard.py"""
import re, os

ruta = "/sessions/brave-affectionate-hawking/mnt/yve01/dashboard.py"
with open(ruta, "r", encoding="utf-8") as f:
    c = f.read()

# ════════════════════════════════════════════════════════════════════════
# BLOQUE 1: Ruta Flask /api/chat — insertar antes de @app.route("/")
# ════════════════════════════════════════════════════════════════════════

CHAT_ROUTE = r'''
# ── Chat AI — Yve Copilot ──────────────────────────────────────────────

def _cargar_contexto_chat():
    """Construye el contexto financiero actual para el system prompt del chat."""
    try:
        df_ar, _  = cargar_datos()
        stats_ar  = calcular_stats(df_ar) if not df_ar.empty else {}
        stats_ap  = calcular_stats_ap(cargar_datos_ap())

        # Detalles AR
        facturas_ar = df_a_lista(df_ar)[:50]
        disc_ar  = [f for f in facturas_ar if str(f.get("estado","")).upper()
                    not in ("CORRECTO","APROBADA","")]
        ota_data = calcular_chart(df_ar)
        otas_str = "; ".join(f"{o['ota']}: {o['n']} facturas" for o in ota_data[:8]) if ota_data else "sin datos"

        # Detalles AP
        df_ap    = cargar_datos_ap()
        lista_ap = df_ap_a_lista(df_ap) if not df_ap.empty else []
        pend_ap  = [f for f in lista_ap if not f.get("accion")]
        disc_ap  = [f for f in lista_ap if f.get("estado") in
                    ("DISCREPANCIA_PO","SIN_PO","ALERTA_CONSUMO","DISCREPANCIA")]

        # Top proveedores con más errores
        from collections import Counter
        prov_err = Counter(f.get("proveedor","") for f in disc_ap)
        top_err  = "; ".join(f"{p}: {n}" for p,n in prov_err.most_common(5)) or "ninguno"

        ctx = f"""ESTADO FINANCIERO ACTUAL DEL HOTEL — Yve.01

=== MÓDULO AR (Facturas OTA) ===
Total facturas AR procesadas hoy: {stats_ar.get('total_facturas', 0)}
Importe total AR: {stats_ar.get('importe_total', 0):,.2f} EUR
Facturas correctas: {stats_ar.get('correctas', 0)}
Discrepancias AR: {stats_ar.get('discrepancias', 0)} — importe reclamable: {stats_ar.get('importe_discrepancias', 0):,.2f} EUR
DI pendientes: {stats_ar.get('di_pendientes', 0)}
Aprobadas: {stats_ar.get('aprobadas', 0)} | Rechazadas: {stats_ar.get('rechazadas', 0)}
OTAs y volumen: {otas_str}
Facturas con discrepancias AR: {'; '.join(f"{f.get('ota','?')} {f.get('importe','?')}€" for f in disc_ar[:5]) or 'ninguna'}

=== MÓDULO AP (Facturas Proveedores) ===
Total facturas AP: {stats_ap.get('total', 0)}
Importe total AP: {stats_ap.get('importe', 0):,.2f} EUR
Matches correctos (F&B+OTRAS): {stats_ap.get('matches', 0)}
Discrepancias PO: {stats_ap.get('discrepancias', 0)}
Sin Orden de Compra: {stats_ap.get('sin_po', 0)}
Alertas consumo F&B: {stats_ap.get('alertas_consumo', 0)}
Pendientes asignación manual: {stats_ap.get('manual', 0)}
Aprobadas AP: {stats_ap.get('aprobadas', 0)} | Rechazadas AP: {stats_ap.get('rechazadas', 0)}
Facturas AP pendientes de aprobar: {len(pend_ap)}
Facturas AP con discrepancias: {'; '.join(f"{f.get('proveedor','?')} ({f.get('estado','?')}) {f.get('total',0):,.0f}EUR" for f in disc_ap[:5]) or 'ninguna'}
Proveedores con más errores: {top_err}"""
        return ctx
    except Exception as e:
        return f"Error cargando contexto financiero: {e}"


@app.route("/api/chat", methods=["POST"])
def api_chat():
    """Endpoint del chat AI — llama a Claude con contexto de datos reales."""
    import json as _json

    data     = request.get_json(force=True, silent=True) or {}
    messages = data.get("messages", [])
    if not messages:
        return jsonify({"error": "No messages provided"}), 400

    # Cargar API key
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        env_path = os.path.join(BASE_DIR, ".env")
        if os.path.exists(env_path):
            for line in open(env_path).readlines():
                if line.startswith("ANTHROPIC_API_KEY="):
                    api_key = line.split("=",1)[1].strip().strip('"').strip("'")
                    break

    if not api_key:
        return jsonify({"reply":
            "⚠️ API key de Anthropic no configurada. Añade ANTHROPIC_API_KEY en el archivo .env."}), 200

    contexto = _cargar_contexto_chat()

    system_prompt = f"""Eres Yve, el asistente financiero inteligente del hotel integrado en el dashboard Yve.01.
Tienes acceso en tiempo real a todos los datos financieros del hotel: facturas AR (OTAs), facturas AP (proveedores), aprobaciones, discrepancias, importes reclamables y estados de contabilización Oracle.

DATOS ACTUALES DEL HOTEL:
{contexto}

INSTRUCCIONES:
- Responde SIEMPRE en español, con tono profesional pero cercano
- Sé directo y específico — da números reales de los datos de arriba
- Si te preguntan por discrepancias, menciona los importes exactos y las OTAs/proveedores concretos
- Si no tienes el dato exacto, dilo claramente y sugiere qué módulo revisar
- Usa emojis con moderación para hacer las respuestas más legibles
- Nunca inventes datos que no aparezcan en el contexto financiero anterior
- Si te preguntan quién eres: "Soy Yve, tu copiloto financiero de Yve.01. Tengo acceso a todos los datos del dashboard en tiempo real."
- Para preguntas sobre facturas concretas, busca en los datos del contexto
- Respuestas concisas: máximo 4-5 líneas salvo que se pida un análisis detallado"""

    try:
        import anthropic
        client  = anthropic.Anthropic(api_key=api_key)
        # Filtrar solo mensajes user/assistant
        msgs = [{"role": m["role"], "content": m["content"]}
                for m in messages if m.get("role") in ("user","assistant")]
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            system=system_prompt,
            messages=msgs,
        )
        reply = resp.content[0].text.strip()
        return jsonify({"reply": reply})
    except Exception as e:
        return jsonify({"reply": f"⚠️ Error al llamar a Claude: {str(e)[:120]}"}), 200

'''

marker = '@app.route("/")\ndef index():'
c = c.replace(marker, CHAT_ROUTE + marker, 1)
print("✓ Ruta /api/chat insertada")

# ════════════════════════════════════════════════════════════════════════
# BLOQUE 2: CSS del chat — insertar antes del cierre </style>
# ════════════════════════════════════════════════════════════════════════

CHAT_CSS = """
/* ── Chat AI — Yve Copilot ─────────────────────────────── */
#chat-fab{position:fixed;bottom:28px;right:28px;z-index:1000;
  display:flex;align-items:center;gap:10px;
  background:linear-gradient(135deg,#7c3aed,#3b82f6);
  color:#fff;border:none;border-radius:50px;padding:14px 22px 14px 18px;
  cursor:pointer;font-size:.95rem;font-weight:700;letter-spacing:.02em;
  box-shadow:0 4px 24px rgba(124,58,237,.5);transition:.2s;white-space:nowrap}
#chat-fab:hover{transform:translateY(-2px);box-shadow:0 8px 32px rgba(124,58,237,.6)}
#chat-fab .fab-dot{width:9px;height:9px;border-radius:50%;
  background:#22c55e;box-shadow:0 0 6px #22c55e;animation:pulse-dot 2s infinite}
@keyframes pulse-dot{0%,100%{opacity:1}50%{opacity:.4}}

#chat-panel{position:fixed;bottom:0;right:0;width:420px;height:100vh;
  background:#0f172a;border-left:1px solid #1e293b;z-index:999;
  display:flex;flex-direction:column;transform:translateX(100%);
  transition:transform .3s cubic-bezier(.4,0,.2,1);
  box-shadow:-8px 0 40px rgba(0,0,0,.5)}
#chat-panel.open{transform:translateX(0)}
@media(max-width:480px){#chat-panel{width:100vw}}

#chat-header{padding:18px 20px;background:#1e293b;border-bottom:1px solid #334155;
  display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
#chat-header .chat-title{display:flex;align-items:center;gap:10px}
#chat-header .chat-title span:first-child{font-size:1.5rem}
#chat-header h3{margin:0;font-size:1rem;font-weight:700;color:#f1f5f9}
#chat-header p{margin:0;font-size:.75rem;color:#60a5fa}
#chat-close{background:none;border:none;color:#64748b;cursor:pointer;
  font-size:1.4rem;padding:4px 8px;border-radius:6px;transition:.15s}
#chat-close:hover{background:#334155;color:#f1f5f9}

#chat-msgs{flex:1;overflow-y:auto;padding:20px;display:flex;
  flex-direction:column;gap:14px;scroll-behavior:smooth}
#chat-msgs::-webkit-scrollbar{width:4px}
#chat-msgs::-webkit-scrollbar-track{background:transparent}
#chat-msgs::-webkit-scrollbar-thumb{background:#334155;border-radius:2px}

.msg{max-width:90%;padding:12px 16px;border-radius:16px;font-size:.88rem;
  line-height:1.55;animation:msgIn .2s ease}
@keyframes msgIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.msg.user{align-self:flex-end;background:linear-gradient(135deg,#3b82f6,#2563eb);
  color:#fff;border-bottom-right-radius:4px}
.msg.bot{align-self:flex-start;background:#1e293b;color:#e2e8f0;
  border:1px solid #334155;border-bottom-left-radius:4px}
.msg.bot.thinking{color:#64748b;font-style:italic;border-style:dashed}

#chat-suggestions{padding:0 16px 12px;display:flex;flex-wrap:wrap;gap:7px;flex-shrink:0}
.sug{background:#1e293b;border:1px solid #334155;color:#94a3b8;
  border-radius:20px;padding:6px 13px;font-size:.78rem;cursor:pointer;
  transition:.15s;white-space:nowrap}
.sug:hover{border-color:#60a5fa;color:#60a5fa;background:#1e3a5f}

#chat-input-row{padding:14px 16px;border-top:1px solid #1e293b;
  display:flex;gap:10px;align-items:center;flex-shrink:0;background:#0f172a}
#chat-input{flex:1;background:#1e293b;border:1px solid #334155;color:#f1f5f9;
  border-radius:24px;padding:11px 18px;font-size:.88rem;outline:none;
  resize:none;font-family:inherit;transition:.15s;max-height:120px}
#chat-input:focus{border-color:#60a5fa;box-shadow:0 0 0 2px rgba(96,165,250,.15)}
#chat-input::placeholder{color:#475569}
#chat-send{background:linear-gradient(135deg,#7c3aed,#3b82f6);border:none;
  color:#fff;border-radius:50%;width:42px;height:42px;cursor:pointer;
  font-size:1.1rem;flex-shrink:0;transition:.15s;display:flex;
  align-items:center;justify-content:center}
#chat-send:hover{transform:scale(1.08)}
#chat-send:disabled{opacity:.4;cursor:not-allowed;transform:none}
"""

c = c.replace("</style>", CHAT_CSS + "\n</style>", 1)
print("✓ CSS chat insertado")

# ════════════════════════════════════════════════════════════════════════
# BLOQUE 3: HTML del chat — insertar antes de </body> en el HTML string
# ════════════════════════════════════════════════════════════════════════

CHAT_HTML = r"""
<!-- ── Chat AI — Yve Copilot ─────────────────────────────── -->

<!-- Botón flotante -->
<button id="chat-fab" onclick="toggleChat()">
  <span style="font-size:1.3rem">💬</span>
  <span>Pregunta a Yve</span>
  <div class="fab-dot"></div>
</button>

<!-- Panel lateral -->
<div id="chat-panel">
  <div id="chat-header">
    <div class="chat-title">
      <span>🤖</span>
      <div>
        <h3>Yve — Copiloto Financiero</h3>
        <p>Acceso en tiempo real a los datos del hotel</p>
      </div>
    </div>
    <button id="chat-close" onclick="toggleChat()">✕</button>
  </div>
  <div id="chat-msgs"></div>
  <div id="chat-suggestions">
    <button class="sug" onclick="askSug(this)">¿Cuánto llevamos facturado?</button>
    <button class="sug" onclick="askSug(this)">¿Qué facturas tienen discrepancias?</button>
    <button class="sug" onclick="askSug(this)">¿Qué proveedor tiene más errores?</button>
    <button class="sug" onclick="askSug(this)">¿Cuánto podemos reclamar a Booking?</button>
    <button class="sug" onclick="askSug(this)">¿Qué facturas están pendientes de aprobar?</button>
  </div>
  <div id="chat-input-row">
    <textarea id="chat-input" rows="1" placeholder="Pregunta sobre el estado financiero del hotel…"
      onkeydown="chatKeydown(event)" oninput="autoResize(this)"></textarea>
    <button id="chat-send" onclick="sendChat()">➤</button>
  </div>
</div>
"""

c = c.replace("</body>\\n</html>", CHAT_HTML + "\n</body>\\n</html>", 1)
print("✓ HTML chat insertado")

# ════════════════════════════════════════════════════════════════════════
# BLOQUE 4: JS del chat — insertar antes de </script>
# ════════════════════════════════════════════════════════════════════════

CHAT_JS = r"""
// ══════════════════════════════════════════════════════════════
// CHAT AI — Yve Copilot
// ══════════════════════════════════════════════════════════════

let chatHistory  = [];
let chatOpen     = false;
let chatGreeted  = false;

function toggleChat() {
  chatOpen = !chatOpen;
  const panel = document.getElementById('chat-panel');
  const fab   = document.getElementById('chat-fab');
  panel.classList.toggle('open', chatOpen);
  fab.style.display = chatOpen ? 'none' : 'flex';
  if (chatOpen && !chatGreeted) {
    chatGreeted = true;
    addMsg('bot', '¡Hola! Soy Yve, tu copiloto financiero 👋\nTengo acceso en tiempo real a todos los datos del hotel. ¿En qué puedo ayudarte?');
  }
  if (chatOpen) setTimeout(() => document.getElementById('chat-input').focus(), 300);
}

function addMsg(role, text) {
  const msgs = document.getElementById('chat-msgs');
  const div  = document.createElement('div');
  div.className = `msg ${role}`;
  div.textContent = text;
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
  return div;
}

async function sendChat() {
  const input = document.getElementById('chat-input');
  const send  = document.getElementById('chat-send');
  const text  = input.value.trim();
  if (!text || send.disabled) return;

  addMsg('user', text);
  chatHistory.push({ role: 'user', content: text });
  input.value = '';
  input.style.height = 'auto';
  send.disabled = true;

  const thinkDiv = addMsg('bot', 'Consultando datos del hotel…');
  thinkDiv.classList.add('thinking');

  try {
    const resp = await fetch('/api/chat', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ messages: chatHistory }),
    });
    const data = await resp.json();
    const reply = data.reply || '⚠️ Sin respuesta del servidor.';

    thinkDiv.textContent = reply;
    thinkDiv.classList.remove('thinking');
    chatHistory.push({ role: 'assistant', content: reply });

    // Mantener historial manejable (últimas 20 interacciones)
    if (chatHistory.length > 40) chatHistory = chatHistory.slice(-40);
  } catch(e) {
    thinkDiv.textContent = '⚠️ Error de conexión con el servidor.';
    thinkDiv.classList.remove('thinking');
  } finally {
    send.disabled = false;
    document.getElementById('chat-input').focus();
  }
}

function askSug(btn) {
  const input = document.getElementById('chat-input');
  input.value = btn.textContent;
  sendChat();
}

function chatKeydown(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendChat();
  }
}

function autoResize(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight, 120) + 'px';
}
"""

c = c.replace(
    "// Cargar datos AP al iniciar\nloadAP();",
    CHAT_JS + "\n// Cargar datos AP al iniciar\nloadAP();",
    1
)
print("✓ JS chat insertado")

# También necesita `request` importado en Flask
if "from flask import" in c and "request" not in c.split("from flask import")[1].split("\n")[0]:
    c = c.replace(
        "from flask import Flask, Response, jsonify, stream_with_context",
        "from flask import Flask, Response, jsonify, request, stream_with_context"
    )
    print("✓ request añadido a imports Flask")
else:
    print("  (request ya importado o import pattern diferente)")

# Guardar
with open(ruta, "w", encoding="utf-8") as f:
    f.write(c)
print("\n✅ dashboard.py modificado")

import py_compile
py_compile.compile(ruta, doraise=True)
print("SINTAXIS OK")
