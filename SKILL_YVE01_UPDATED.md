---
name: yve01
description: >
  Full context skill for Yve.01 — an AI-first hotel operations automation startup
  targeting the hospitality finance sector. Use this skill whenever the user mentions
  Yve.01, hotel finance automation, AP/AR workflows, the Hilton project, Oracle hotel
  integration, OTA commissions, albaranes, or anything related to building this startup.
  Also trigger for any coding task related to the project (invoice processing, OCR,
  3-way matching, Flask dashboards, Oracle API). This skill gives Claude complete
  context so the user never has to re-explain the project.
---

# Yve.01 — Project Context Skill

## What is Yve.01

Yve.01 is an AI-first hotel operations automation startup targeting the hospitality
finance sector. The founder is 16 years old, based in Barcelona. No programming
experience — uses Claude as the primary development tool.

The entry point is automating the Finance department of hotels, starting with
AP (Accounts Payable) and AR (Accounts Receivable) workflows. The long-term
vision is full hotel operations automation across all departments.

Validated through direct conversations with the Assistant Financial Controller of
Hilton Barcelona (Avinguda Diagonal), who acts as advisor and lives with the founder.

**Production URL:** https://yve01.onrender.com (AR dashboard live)

---

## Current Build State

### ✅ Phase 0 — Setup (Complete)
Python, VS Code, Claude installed. Project at `C:\Users\Jo\yve01`.

### ✅ Phase 1 — AR Module (Complete, in production)
5 files built, tested end-to-end, running at https://yve01.onrender.com.

### ✅ Phase 2 — AP Module (Complete)
10 files built, tested with 5 real PDF invoices:
- F&B: MATCH_3WAY_OK ✅ | DISCREPANCIA_PO ❌ | ALERTA_CONSUMO 🔵
- OTRAS: MATCH_CORRECTO ✅ | SIN_PO 🟡

### 🔜 Phase 3 — Oracle Integration (Next)
Auto-contabilización via Oracle REST API.
Plan documented in `C:\Users\Jo\yve01\ORACLE_INTEGRATION.md`.

### 🔜 Phase 4 — First Client
Target: independent 4–5★ hotels in Barcelona.

---

## Project Files — Complete Inventory

All files live in `C:\Users\Jo\yve01\`.

### Módulo AR — Accounts Receivable (OTA invoices)

| File | Purpose | Port |
|---|---|---|
| `lector_ota.py` | Reads OTA PDFs, extracts fields via Claude API + regex fallback | — |
| `verificador_comisiones.py` | Verifies commission rates against negotiated table | — |
| `detector_doble_imposicion.py` | Detects foreign OTA invoices needing DI certificate | — |
| `generador_emails.py` | Generates professional Spanish emails for AR discrepancies | — |
| `app_aprobacion.py` | Flask approval app — department heads approve/reject | 5000 |

### Módulo AP — Accounts Payable (supplier invoices)

| File | Purpose | Port |
|---|---|---|
| `datos-referencia/proveedores.xlsx` | 12 suppliers (F&B + OTRAS) with account codes + email | — |
| `datos-referencia/plan_cuentas.xlsx` | Spanish PGC chart of accounts (600–629, 400, 410, 472) | — |
| `lector_facturas_ap.py` | Reads supplier PDFs; Claude API + structured regex fallback | — |
| `gestor_pos.py` | POS data (180 rows) + Purchase Orders (14 POs) | — |
| `matching_ap_otras.py` | OTRAS invoices vs POs — 1% tolerance | — |
| `matching_ap_fb.py` | 3-way matching: Factura + PO + POS (1% / 15% tolerances) | — |
| `asignador_cuentas.py` | Rule-based account assignment + full double-entry record | — |
| `app_aprobacion_ap.py` | Flask approval app — per-department, mandatory comment | 5002 |
| `generador_emails_ap.py` | Emails for DISCREPANCIA_PO / SIN_PO / ALERTA_CONSUMO | — |

### Dashboard

| File | Purpose | Port |
|---|---|---|
| `dashboard.py` | Main dashboard — AR + AP tabs, pipeline SSE buttons, Chart.js | 5001 |

### Folder structure

```
datos-referencia/
  proveedores.xlsx        — 12 suppliers
  plan_cuentas.xlsx       — 12 account codes
  pos_ventas.xlsx         — 180 POS rows (July 2025)
  pos_ordenes.xlsx        — 14 POs
  cta_cble.xlsx           — Oracle account map (Phase 3, one per hotel)

facturas-entrada/         — drop PDFs here (OTA + supplier)
facturas-procesadas/      — facturas_ap_*.xlsx, facturas_contabilizadas_*.xlsx
aprobaciones/             — aprobaciones.xlsx, aprobaciones_ap.xlsx
reportes/
  verificacion_*.xlsx
  doble_imposicion_*.xlsx
  matching_otras_*.xlsx
  matching_fb_*.xlsx
  emails_pendientes/
  emails_pendientes_ap/
```

---

## Running the System

```bash
# 3 terminals
python dashboard.py          # http://localhost:5001
python app_aprobacion.py     # http://localhost:5000
python app_aprobacion_ap.py  # http://localhost:5002

# AR pipeline
python lector_ota.py && python verificador_comisiones.py && \
python detector_doble_imposicion.py && python generador_emails.py

# AP pipeline
python lector_facturas_ap.py && python matching_ap_otras.py && \
python matching_ap_fb.py && python asignador_cuentas.py && \
python generador_emails_ap.py
```

---

## Critical Technical Patterns

### MANDATORY: Use bash heredoc for files >100 lines
The Write/Edit tools silently truncate at ~256 lines. Always use:
```bash
cat > archivo.py << 'ENDOFFILE'
...code...
ENDOFFILE
python3 -c "import py_compile; py_compile.compile('archivo.py', doraise=True); print('SINTAXIS OK')"
```

### Claude API always needs regex fallback (sandbox has no internet)
```python
try:
    resp = client.messages.create(model="claude-sonnet-4-6", ...)
    return json.loads(resp.content[0].text)
except Exception:
    return extraer_con_regex(texto)
```

### Test PDF format — structured block for reliable regex
```
DATOS FACTURA SISTEMA:
NUMERO_FACTURA=FAC-2025-ENDE-0742
FECHA=15/07/2025
PROVEEDOR=Endesa Energia SA
NIF=A-81948077
BASE_IMPONIBLE=15289.26
IVA_PORCENTAJE=21
CUOTA_IVA=3210.74
TOTAL=18500.00
```

### Flask SSE pipeline pattern
```python
_pipeline_lock = threading.Lock()

@app.route("/api/procesar")
def api_procesar():
    def generar():
        with _pipeline_lock:
            if _running: yield "data: Proceso en curso\n\n"; return
        try:
            for script, label in scripts:
                res = subprocess.run([sys.executable, script], ...)
                for l in res.stdout.splitlines():
                    yield f"data: {l}\n\n"
            yield "data: PIPELINE_COMPLETO\n\n"
        finally:
            pass
    return Response(stream_with_context(generar()), mimetype="text/event-stream",
                    headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})
```

### Account assignment rules (PGC español)
```
F&B suppliers   → 600  |  limpieza/prof → 623  |  telecom/internet → 629
energía         → 629  |  mantenimiento → 622  |  seguro  → 625
arrendamiento   → 621  |  OTRAS fallback → 629  |  unknown → REVISAR_MANUAL
IVA 21% → 472  |  IVA 10% → 4720  |  IVA 4% → 4721
Proveedores → 400  |  Acreedores → 410
```

### Matching states
- MATCH_CORRECTO / MATCH_3WAY_OK → facturas OK, no email
- DISCREPANCIA_PO → email soliciting factura rectificativa
- SIN_PO → email informing no approved PO exists
- ALERTA_CONSUMO → email requesting F&B consumption explanation

---

## AP/AR Workflow (validated with Hilton Barcelona)

### CAMINO AR — Accounts Receivable
OTA invoices → verify rates → sign → attach group info →
DI certificate if foreign OTA → dept head sign → passes to AP

### CAMINO AP — Accounts Payable
**F&B:** Attach + verify albaranes sellados + POS + PO → dept head sign
**OTRAS:** Verify + attach PO only → dept head sign
Both paths → CONTABILIZAR EN ORACLE → PAGAR → attach bank receipt

---

## Hotel Systems

| System | Role | Integration |
|---|---|---|
| PEP | Hilton's new cloud PMS (replaces OnQ) | Proprietary — Corporate approval needed |
| Oracle (OPERA/MICROS) | Accounting ERP — contabilización | **Phase 3 — public REST APIs** |
| Opera Cloud | Oracle PMS (IHG, Accor, Hyatt) | Open APIs — for non-Hilton clients |

---

## DRR Structure (real Hilton file: DailyHilton BCNJUL_2025NT.xlsm)

**Layer 1 — 31 daily sheets:** Trial Balance per day (ACCOUNT NAME / DEBITS / CREDITS)
**Layer 2 — Summaries:** DAILY_MASTER, DHBCN (KPIs: Today/MTD/Budget/LY), F&B, EXP_INC
**Layer 3 — Analysis:** KPIs (time series), Nat.Mix, Datas, **CtaCble**

### CtaCble sheet — Oracle account mapping

| Column | Oracle field | Example |
|---|---|---|
| Entity | Legal entity / ledger | `HILBCN` |
| Department | Cost centre | `FB`, `ADM`, `MAINT`, `HK`, `SEC` |
| Account | GL account code | `629`, `600`, `472` |
| Line Description | Human label | "Suministro Eléctrico" |
| Interco | Intercompany flag | Usually blank |

Oracle account combination: `Entity.Department.Account`
Example: `HILBCN.ADM.629` = Hilton Barcelona / Administration / Otros Servicios

Key DRR KPIs: Rooms, Occupancy%, ADR, RevPAR, Rooms Revenue, F&B Revenue,
Total Revenue, SpendPAR, GOP, GOP%, MTD, OTB

---

## Oracle Integration — 5-Step Plan (Phase 3)

Full code in `ORACLE_INTEGRATION.md`. Summary:

| Step | File | What it does |
|---|---|---|
| 1 | `oracle_auth.py` | OAuth 2.0 token + connection test |
| 2 | `oracle_lector_facturas.py` | Read facturas_contabilizadas → Oracle line dicts |
| 3 | `oracle_crear_asientos.py` | POST journal batches to Oracle GL |
| 4 | `oracle_actualizar_estado.py` | Write Oracle ID back → dashboard "CONTABILIZADA" |
| 5 | `oracle_pipeline.py` | Full daily automated run with error handling |

Key rule: only post invoices with `accion == "APROBADA"` — legal requirement in Spain.

---

## Key Vocabulary

| Term | Definition |
|---|---|
| Albarán | Physical delivery note. Signed on receipt. Must match invoice. |
| PO (Orden de Compra) | Purchase authorization. Required before payment. |
| 3-way matching | PO + albarán/POS + factura cross-check |
| Asiento contable | Double-entry accounting record (DEBE / HABER) |
| CtaCble | DRR sheet mapping P&L lines to Oracle GL codes |
| Doble imposición | Double taxation on foreign OTA invoices |
| Certificado DI | Certificate proving tax paid in country of origin |
| Cabeza de departamento | Department head — approves all invoices |
| PGC | Plan General Contable — Spanish chart of accounts |
| USALI | Uniform System of Accounts for the Lodging Industry |
| Night Audit | Daily financial close |
| DRR | Daily Revenue Report |
| RevPAR | Revenue Per Available Room |
| ADR | Average Daily Rate |
| GOP | Gross Operating Profit |
| OTA | Online Travel Agency |

---

## Competitors

| Company | Gap |
|---|---|
| Phacet | France, large chains only |
| Nimble Property | Generic, not enterprise |
| Rillion | US-focused, no Spain |
| M3 / Aptech | Legacy, slow, costly |
| BlackLine | Not hospitality-specific |

**Yve.01's gap:** No competitor serves mid-size independent hotels in Spain/Europe
with fast, affordable, AI-native AP/AR automation that works out of the box.

---

## Business Model

| Plan | Price | Target |
|---|---|---|
| Starter | 300€/month | Independent hotel, 1 module |
| Pro | 600€/month | Independent hotel, all modules |
| Multi | 400€/month/hotel | Groups of 2–5 hotels |

**NEVER position as replacing staff.** Always: multiplying team capacity.
ROI: 1 auditor covers 1 hotel manually → with Yve.01 covers 3–4 hotels.

---

## Advisor Constraints

- No Hilton brand in marketing (reference as "international 5★ chain")
- DRR needs human review before being reliable
- Hilton = Corporate decisions — start with independent hotels
- Spanish severance is significant — never sell as headcount reduction

---

## How to Use This Skill

1. Use full context without asking user to re-explain anything
2. Correct workflow: AR feeds AP, both feed Oracle
3. Correct vocabulary: albarán, asiento, CtaCble, cabeza de departamento
4. Correct constraints: no Hilton brand, no replacement narrative
5. Phase order: Oracle integration (Phase 3) is next
6. Files >100 lines: ALWAYS bash heredoc
7. After every file: verify syntax with py_compile
8. Claude API calls: ALWAYS include regex fallback
9. Ports: 5000 (AR approvals), 5001 (dashboard), 5002 (AP approvals)
