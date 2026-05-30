---
name: yve01
description: >
  Full context skill for Yve.01 — an AI-first hotel operations automation startup
  targeting the hospitality finance sector. Use this skill whenever the user mentions
  Yve.01, hotel finance automation, AP/AR workflows, the Hilton project, Oracle hotel
  integration, OTA commissions, albaranes, DRR, lector_drr, or anything related to
  building this startup. Also trigger for any coding task related to the project
  (invoice processing, OCR, 3-way matching, Flask dashboards, Oracle API, Trial Balance).
  This skill gives Claude complete context so the user never has to re-explain the project.
---

# Yve.01 — Project Context Skill

## What is Yve.01

Yve.01 is an AI-first hotel operations automation startup targeting the hospitality
finance sector. The founder is 16 years old, based in Barcelona. No programming
experience — uses Claude as the primary development tool.

Validated through direct conversations with the Assistant Financial Controller of
Hilton Barcelona (Avinguda Diagonal), who acts as advisor and lives with the founder.

**Production URL:** https://yve01.onrender.com
**GitHub:** https://github.com/Jordi-Solano/yve01
**Project folder:** `C:\Users\Jo\yve01`

---

## Current Build State

### ✅ Phase 0 — Setup (Complete)
Python, VS Code, Claude installed. Stack: Python, Flask, pdfplumber, anthropic SDK,
openpyxl, pandas, reportlab, gunicorn. Deployed on Render.

### ✅ Phase 1 — AR Module (Complete, live in production)
5 files. OTA invoice processing running at https://yve01.onrender.com.

### ✅ Phase 2 — AP Module (Complete)
9 files. Supplier invoice processing with 3-way matching, account assignment,
email generation. Tested end-to-end with 5 real PDF invoices.

### ✅ Phase 3 — Oracle Integration (Complete, simulation mode)
5 files. Full Oracle GL journal entry pipeline. Simulation mode active until
hotel provides real Oracle Cloud credentials. Simulation tested: 5 facturas,
198.690 EUR, IDs SIM-20260529-001 to 005.

### ✅ Lector DRR (Complete)
1 file (`lector_drr.py`). Processes real Hilton .xlsm DRR files.
Tested on `DailyHilton BCNJUL_2025NT.xlsm` (July 2025, 45 sheets):
- 31 days processed, 7.397 Trial Balance lines extracted
- 239 accounts per day, 105 CtaCble Oracle mappings
- 1 Out of Balance day detected (Day 1: -2.185,01 EUR)

### 🔜 Phase 4 — First Client
Target: independent 4–5★ hotels in Barcelona.
Reference: "validated with finance team of international 5★ chain hotel in Barcelona."

---

## Complete File Inventory

All files in `C:\Users\Jo\yve01\`

### Módulo AR — Accounts Receivable (OTA invoices)

| File | Purpose | Port |
|---|---|---|
| `lector_ota.py` | Reads OTA PDFs, extracts fields — Claude API + regex fallback | — |
| `verificador_comisiones.py` | Verifies commission rates vs negotiated table | — |
| `detector_doble_imposicion.py` | Detects foreign OTA invoices needing DI certificate | — |
| `generador_emails.py` | Generates Spanish emails for AR discrepancies | — |
| `app_aprobacion.py` | Flask approval app — department head approve/reject | 5000 |

### Módulo AP — Accounts Payable (supplier invoices)

| File | Purpose | Port |
|---|---|---|
| `lector_facturas_ap.py` | Reads supplier PDFs — Claude API + structured regex fallback | — |
| `gestor_pos.py` | POS data (180 rows) + Purchase Orders (14 POs) | — |
| `matching_ap_otras.py` | OTRAS invoices vs POs — 1% tolerance | — |
| `matching_ap_fb.py` | 3-way matching: Factura + PO + POS (1% / 15%) | — |
| `asignador_cuentas.py` | Rule-based account assignment + double-entry record | — |
| `app_aprobacion_ap.py` | Flask approval — per-department, mandatory comment | 5002 |
| `generador_emails_ap.py` | Emails for DISCREPANCIA_PO / SIN_PO / ALERTA_CONSUMO | — |

### Módulo Oracle — GL Journal Entries

| File | Purpose |
|---|---|
| `oracle_auth.py` | OAuth 2.0 token mgmt. Auto-detects simulation vs production. |
| `oracle_lector_facturas.py` | Reads `facturas_contabilizadas` → Oracle journal line dicts. 3 lines/invoice: DEBE gasto + DEBE IVA + HABER proveedores. |
| `oracle_crear_asientos.py` | POSTs journal batches to Oracle GL REST API. In simulation: saves `reportes/oracle_simulacion_[fecha].xlsx`. |
| `oracle_actualizar_estado.py` | Writes Oracle Journal ID back to Excel. Sets status CONTABILIZADA / CONTABILIZADA_SIM. Never re-processes already contabilizadas. |
| `oracle_pipeline.py` | Orchestrates all 4 above. Critical rule: blocks any invoice without APROBADA status (bypassed in simulation). Full log + final summary. |

### Dashboard + DRR

| File | Purpose | Port |
|---|---|---|
| `dashboard.py` | Main dashboard: AR + AP tabs, Chart.js, pipeline SSE buttons, Oracle button | 5001 |
| `lector_drr.py` | Reads .xlsm DRR files: DAILY_MASTER KPIs, 31-day Trial Balance, CtaCble mapping, Out-of-Balance detection | — |

### Reference data & outputs

```
datos-referencia/
  proveedores.xlsx        — 12 suppliers (F&B + OTRAS)
  plan_cuentas.xlsx       — 12 PGC account codes
  pos_ventas.xlsx         — 180 POS rows (July 2025)
  pos_ordenes.xlsx        — 14 Purchase Orders
  cta_cble.xlsx           — Oracle CtaCble map (one per hotel client)

facturas-entrada/         — drop PDFs here (OTA + supplier)
facturas-procesadas/
  facturas_ap_*.xlsx      — extracted supplier invoices
  facturas_contabilizadas_*.xlsx  — with account assignments + oracle_status + oracle_id
aprobaciones/
  aprobaciones.xlsx       — AR approvals
  aprobaciones_ap.xlsx    — AP approvals (created by app_aprobacion_ap.py)
reportes/
  verificacion_*.xlsx
  doble_imposicion_*.xlsx
  matching_otras_*.xlsx
  matching_fb_*.xlsx
  drr_procesado_*.xlsx    — 4 sheets: Resumen, Trial_Balance_Completo, Alertas, CtaCble_Mapping
  oracle_simulacion_*.xlsx — Oracle journal batches (simulation mode)
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

# Oracle pipeline (simulation until credentials added)
python oracle_pipeline.py

# DRR reader
python lector_drr.py "path/to/DailyHilton.xlsm"
```

---

## Oracle Integration Details

### Activating Production Mode
Add to `.env`:
```
ORACLE_BASE_URL=https://your-hotel.oraclecloud.com
ORACLE_CLIENT_ID=your_client_id
ORACLE_CLIENT_SECRET=your_client_secret
ORACLE_LEDGER_NAME=Hilton Barcelona
ORACLE_ENTITY=1662
```
Without these variables → simulation mode activates automatically.

### Oracle Journal Entry Structure (per invoice)
```
DEBE  {Entity}.{Dept}.{cta_gasto}   base_imponible EUR   (e.g. 1662.FB.600)
DEBE  {Entity}.ADM.{cta_iva}        cuota_iva EUR         (e.g. 1662.ADM.4720)
HABER {Entity}.ADM.400              total_factura EUR     (e.g. 1662.ADM.400)
```

### Oracle API Endpoints Used
```
POST /oauth/token                                           — auth
GET  /fscmRestApi/resources/11.13.18.05/ledgers            — test connection
POST /fscmRestApi/resources/11.13.18.05/journalBatches     — create journal
POST /fscmRestApi/resources/11.13.18.05/journalBatches/{id}/action/post — post to GL
```

### Critical Rule
**Never post to Oracle without `accion == "APROBADA"`** in `aprobaciones_ap.xlsx`.
This is a legal requirement in Spain. The pipeline enforces this gate.

---

## DRR Structure (real Hilton file: DailyHilton BCNJUL_2025NT.xlsm, 45 sheets)

### Sheet inventory
- `DAILY_MASTER` — KPI summary (rows 10–32): Today / MTD / Rest of Month / Full Month Forecast / Budget / LY
- `DHBCN`, `FCST`, `F&B`, `LY`, `EXP_INC`, `LM`, `CONTROL` — summary sheets
- `1`–`31` — daily Trial Balance sheets
- `Nat.Mix`, `Datas`, `KPIs` — analysis sheets
- **`CtaCble`** — Oracle account mapping (263 rows, 105 unique Entity+Dept+Account combos)

### Daily sheet structure (cols, 0-indexed)
- A (0): section header — ASSETS / LIABILITIES / EXPENSES / INCOME
- C (2): ACCOUNT NAME
- H (7): DEBITS
- I (8): CREDITS
- J (9): Total
- R01 col D: "Out of Balance" or "OK"
- R01 col E: imbalance amount
- R03 col D: date of the day

### CtaCble columns
| Column | Oracle field | Example value |
|---|---|---|
| Entity | Legal entity / ledger | `1662` |
| Department | Cost centre | `0`, `FB`, `ADM`, `000` |
| Account | GL account code | `10560`, `12100`, `40000` |
| Interco | Intercompany flag | `0` |
| Project | Project code | `0` or `351102` |
| Future | Future use | `0` |
| Line Description | Human label | "BANK DEPOSITS" |

Oracle combination: `{Entity}.{Department}.{Account}` → e.g. `1662.000.10560`

### Key KPIs extracted from DAILY_MASTER
Occupancy %, ADR, Revenue PAR, Rooms Revenue, Food Revenue, Beverage Revenue,
F&B Revenue Total, Total Revenue, Spend PAR, GOP, GOP % — all with Today/MTD/Forecast/Budget

---

## Critical Technical Patterns

### MANDATORY: bash heredoc for files >100 lines
Write/Edit tools silently truncate at ~256 lines. Always:
```bash
cat > archivo.py << 'ENDOFFILE'
...code...
ENDOFFILE
python3 -c "import py_compile; py_compile.compile('archivo.py', doraise=True); print('SINTAXIS OK')"
```

### Claude API + regex fallback (sandbox/CI has no internet)
```python
try:
    resp = client.messages.create(model="claude-sonnet-4-6", max_tokens=512, ...)
    return json.loads(resp.content[0].text)
except Exception:
    return extraer_con_regex(texto)   # always works offline
```

### Test PDF format — structured block for reliable parsing
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

### Flask SSE pipeline — shared pattern for all 3 pipeline buttons
```python
_lock = threading.Lock()
_running = False

@app.route("/api/procesar_X")
def api_procesar_X():
    def generar():
        global _running
        with _lock:
            if _running: yield "data: Proceso en curso\n\n"; return
            _running = True
        try:
            res = subprocess.run([sys.executable, "script.py"], ...)
            for l in res.stdout.splitlines():
                yield f"data: {l}\n\n"
            yield "data: PIPELINE_COMPLETO\n\n"
        finally:
            _running = False
    return Response(stream_with_context(generar()), mimetype="text/event-stream",
                    headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})
```

### Matching states
- `MATCH_CORRECTO` / `MATCH_3WAY_OK` → invoice OK, no email needed
- `DISCREPANCIA_PO` → email requesting factura rectificativa
- `SIN_PO` → email: no approved PO exists, cannot pay
- `ALERTA_CONSUMO` → email: F&B consumption difference >15% vs POS

### Account assignment (PGC español)
```
F&B suppliers   → 600  |  limpieza/prof  → 623  |  telecom → 629
energía         → 629  |  mantenimiento  → 622  |  seguro  → 625
arrendamiento   → 621  |  OTRAS fallback → 629  |  unknown → REVISAR_MANUAL
IVA 21% → 472  |  IVA 10% → 4720  |  IVA 4% → 4721
Proveedores → 400  |  Acreedores → 410
```

### Port map
- 5000 → AR approvals (`app_aprobacion.py`)
- 5001 → Main dashboard (`dashboard.py`)
- 5002 → AP approvals (`app_aprobacion_ap.py`)

---

## AP/AR Workflow (validated with Hilton Barcelona)

### CAMINO AR — Accounts Receivable
OTA invoices → verify rates → sign → attach group info →
DI certificate if foreign OTA → dept head sign → passes to AP

### CAMINO AP — Accounts Payable
**F&B:** Attach + verify albaranes sellados + POS data + PO → dept head sign
**OTRAS:** Verify + attach PO only → dept head sign
Both paths → **CONTABILIZAR EN ORACLE** → PAGAR → attach bank receipt

---

## Hotel Systems

| System | Role | Integration status |
|---|---|---|
| PEP | Hilton's new cloud PMS (replaces OnQ) | Proprietary — Corporate approval needed |
| Oracle Fusion Cloud Finance | Accounting ERP — contabilización | **Built — needs credentials** |
| Opera Cloud | Oracle PMS (IHG, Accor, Hyatt) | Open APIs — for non-Hilton clients |

---

## Key Vocabulary

| Term | Definition |
|---|---|
| Albarán | Physical delivery note. Signed on receipt. Must match invoice. |
| PO (Orden de Compra) | Purchase authorization. Required before payment. |
| 3-way matching | PO + albarán/POS + factura cross-check |
| Asiento contable | Double-entry accounting record (DEBE / HABER) |
| Contabilizar | Post a transaction to the General Ledger in Oracle |
| CtaCble | DRR sheet mapping each P&L line to Oracle GL code (Entity+Dept+Account) |
| DRR | Daily Revenue Report — the hotel's daily .xlsm financial file |
| Trial Balance | Daily list of all account movements (DEBITS / CREDITS / Total) |
| Out of Balance | When total debits ≠ total credits in a Trial Balance day |
| Journal Batch | Oracle GL unit: one batch per invoice, 3 lines each |
| Doble imposición | Double taxation on foreign OTA invoices |
| Certificado DI | Certificate proving tax paid in country of origin |
| Cabeza de departamento | Department head — approves all invoices |
| PGC | Plan General Contable — Spanish chart of accounts |
| USALI | Uniform System of Accounts for the Lodging Industry |
| Night Audit | Daily financial close |
| RevPAR | Revenue Per Available Room |
| ADR | Average Daily Rate |
| GOP | Gross Operating Profit |
| SpendPAR | Spend Per Available Room |
| OTA | Online Travel Agency (Booking, Expedia, etc.) |

---

## Competitors

| Company | Gap |
|---|---|
| Phacet | France, large chains only |
| Nimble Property | Generic, not enterprise-grade |
| Rillion | US-focused, no Spain |
| M3 / Aptech | Legacy, slow, costly |
| BlackLine | Not hospitality-specific |

**Yve.01's gap:** No competitor offers fast, affordable, AI-native AP/AR + Oracle automation
for mid-size independent hotels in Spain/Europe that works out of the box as a service.

---

## Business Model

| Plan | Price | Target |
|---|---|---|
| Starter | 300€/month | Independent hotel, 1 module |
| Pro | 600€/month | Independent hotel, all modules |
| Multi | 400€/month/hotel | Groups of 2–5 hotels |

**NEVER position as replacing staff** — always position as multiplying team capacity.
ROI: 1 Income Auditor covers 1 hotel manually → with Yve.01 covers 3–4 hotels.

---

## Advisor Constraints

- No Hilton brand in marketing (reference as "international 5★ chain in Barcelona")
- DRR requires human review before being reliable (data quality — confirmed by lector_drr.py)
- Hilton decisions at Corporate level — start with independent hotels
- Spanish severance is significant — never sell as headcount reduction

---

## How to Use This Skill

1. Use full context — never ask the user to re-explain anything
2. Correct workflow: AR feeds AP → AP feeds Oracle → Oracle posts to GL
3. Correct vocabulary: albarán, asiento, CtaCble, cabeza de departamento, contabilizar
4. Correct constraints: no Hilton brand, no replacement narrative
5. **Files >100 lines: ALWAYS bash heredoc** — Write/Edit tool truncates silently
6. After every file: verify syntax with `py_compile.compile()`
7. Claude API calls: ALWAYS include regex/template fallback
8. Oracle: simulation mode is automatic without credentials, production is automatic with them
9. Ports: 5000 AR / 5001 dashboard / 5002 AP
10. Next priority: first paying client (independent hotel Barcelona)
