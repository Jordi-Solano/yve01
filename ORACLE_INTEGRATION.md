# Oracle Integration Plan — Yve.01

## Overview

This document describes how to build the Oracle REST API integration (Phase 3 of Yve.01).
The goal: take invoices already processed by the AP module (`facturas_contabilizadas_*.xlsx`)
and create journal entries (asientos contables) automatically in Oracle, eliminating manual
data entry by the accounting team.

---

## Part 1 — The CtaCble Sheet as Oracle Account Map

### What is CtaCble

The `CtaCble` sheet is found in the hotel's Daily Revenue Report (DRR) — the `.xlsm` file
produced by the hotel's accounting system each day. At Hilton Barcelona this is
`DailyHilton BCNJUL_2025NT.xlsm`.

It is the bridge between the hotel's operational P&L and Oracle's General Ledger.
Every line of revenue and cost in the hotel maps to a specific Oracle account code
via this sheet.

### CtaCble Column Structure

| Column | Oracle Field | Description |
|---|---|---|
| **Entity** | Legal entity code | Identifies the legal company. E.g. `HILBCN` or `ESP001`. Used in Oracle as the Ledger/BU identifier. |
| **Department** | Cost centre / department | Maps to Oracle department segment. E.g. `FB` (Food & Beverage), `ROOMS`, `ADM` (Administration), `MAINT` (Maintenance), `HK` (Housekeeping), `SEC` (Security). |
| **Account** | Natural account code | The 3–6 digit GL account number. In Spanish PGC: 600 (F&B purchases), 621 (rent), 622 (maintenance), 623 (professional services), 629 (other services), 400 (suppliers), 472 (IVA soportado). |
| **Line Description** | Account description | Human-readable label for the account line. E.g. "Compras de Alimentación", "Suministro Eléctrico", "Servicios de Limpieza". Used in Oracle's journal line description. |
| **Interco** | Intercompany flag | Optional. Marks intercompany transactions for consolidation. Usually blank for independent hotels. |

### How Yve.01 Uses CtaCble

When `asignador_cuentas.py` assigns account `629` to a Gas Natural invoice,
the CtaCble sheet tells us:
- **Entity:** `HILBCN`
- **Department:** `ADM`
- **Account:** `629`
- **Line Description:** "Otros Servicios Exteriores"

These four fields compose the Oracle account segment string:
```
HILBCN.ADM.629
```
That string is what goes into the Oracle Journal Entry API call as the `account` field.

### Practical Note

The CtaCble sheet is hotel-specific. Before deploying Yve.01 to a new hotel,
the onboarding process must:
1. Export their CtaCble sheet (or equivalent account mapping)
2. Load it into `datos-referencia/cta_cble.xlsx`
3. The integration script reads this file to resolve Entity/Department for each account code

---

## Part 2 — Oracle REST API Integration Plan (5 Steps)

### Prerequisites

- Oracle Fusion Cloud Finance or Oracle OPERA Cloud access
- API credentials: Client ID + Client Secret (OAuth 2.0)
- Base URL: `https://{hostname}/fscmRestApi/resources/11.13.18.05/`
- The hotel's CtaCble mapping loaded in `datos-referencia/cta_cble.xlsx`

---

### Step 1 — Authentication & Connection Test
**File to build:** `oracle_auth.py`

Implement OAuth 2.0 token retrieval and test the connection:

```python
import requests

ORACLE_BASE_URL = os.getenv("ORACLE_BASE_URL")   # e.g. https://hotel.oraclecloud.com
CLIENT_ID       = os.getenv("ORACLE_CLIENT_ID")
CLIENT_SECRET   = os.getenv("ORACLE_CLIENT_SECRET")

def get_token():
    resp = requests.post(
        f"{ORACLE_BASE_URL}/oauth/token",
        data={"grant_type": "client_credentials"},
        auth=(CLIENT_ID, CLIENT_SECRET)
    )
    return resp.json()["access_token"]

def test_connection(token):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.get(
        f"{ORACLE_BASE_URL}/fscmRestApi/resources/11.13.18.05/ledgers",
        headers=headers
    )
    return resp.status_code == 200
```

**Expected output:** Token + list of available ledgers (confirms access).

---

### Step 2 — Read Processed Invoices from Yve.01
**File to build:** `oracle_lector_facturas.py`

Read `facturas_contabilizadas_*.xlsx` and transform each row into an Oracle-ready
journal line dict:

```python
import pandas as pd, glob

def cargar_facturas_para_oracle():
    archivos = sorted(glob.glob("facturas-procesadas/facturas_contabilizadas_*.xlsx"), reverse=True)
    df = pd.read_excel(archivos[0])
    
    # Only load invoices approved by department head and not yet posted to Oracle
    df = df[df["accion"].str.upper() == "APROBADA"]
    df = df[df.get("oracle_status", "PENDIENTE") == "PENDIENTE"]
    return df

def construir_lineas_oracle(row, cta_cble_df):
    """Convert one invoice row into Oracle journal line pairs (DEBE + HABER)."""
    cuenta_gasto = str(row["cuenta_contable"])
    cuenta_iva   = str(row.get("cuenta_debe_iva", "472"))
    cuenta_prov  = "400"  # Proveedores
    
    # Resolve Entity + Department from CtaCble
    def resolver(account_code):
        match = cta_cble_df[cta_cble_df["Account"].astype(str) == account_code]
        if not match.empty:
            return match.iloc[0]["Entity"], match.iloc[0]["Department"]
        return "DEFAULT", "ADM"
    
    entity, dept = resolver(cuenta_gasto)
    
    base = float(row.get("base_imponible", 0) or 0)
    iva  = float(row.get("cuota_iva", 0) or 0)
    total = float(row.get("total_factura", 0) or 0)
    
    return [
        # DEBE — Gasto
        {"accountCombination": f"{entity}.{dept}.{cuenta_gasto}",
         "creditAmount": 0, "debitAmount": base,
         "description": f"Fact. {row['numero_factura']} — {row['nombre_proveedor']}"},
        # DEBE — IVA soportado
        {"accountCombination": f"{entity}.ADM.{cuenta_iva}",
         "creditAmount": 0, "debitAmount": iva,
         "description": f"IVA s/ Fact. {row['numero_factura']}"},
        # HABER — Proveedores
        {"accountCombination": f"{entity}.ADM.{cuenta_prov}",
         "creditAmount": total, "debitAmount": 0,
         "description": f"Proveedores — {row['nombre_proveedor']}"},
    ]
```

**Expected output:** List of journal line dicts ready for Oracle API.

---

### Step 3 — Create Journal Batches in Oracle
**File to build:** `oracle_crear_asientos.py`

Use the Oracle Financials REST API to create journal entries:

```python
def crear_asiento_oracle(token, factura_row, lineas):
    """POST a journal batch to Oracle General Ledger."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "JournalBatchName":        f"YVE01-AP-{factura_row['numero_factura']}",
        "LedgerName":              os.getenv("ORACLE_LEDGER_NAME", "Hilton Barcelona"),
        "AccountingDate":          factura_row["fecha"],  # DD/MM/YYYY → YYYY-MM-DD
        "JournalSource":           "YVE01",
        "JournalCategory":         "Purchase Invoices",
        "Description":             f"Factura {factura_row['numero_factura']} — {factura_row['nombre_proveedor']}",
        "JournalLines": [
            {
                "LineNumber":         i + 1,
                "AccountCombination": l["accountCombination"],
                "CreditAmount":       l["creditAmount"],
                "EnteredDebitAmount": l["debitAmount"],
                "Description":        l["description"],
            }
            for i, l in enumerate(lineas)
        ]
    }
    
    resp = requests.post(
        f"{ORACLE_BASE_URL}/fscmRestApi/resources/11.13.18.05/journalBatches",
        headers=headers,
        json=payload
    )
    
    if resp.status_code in (200, 201):
        oracle_id = resp.json().get("JournalBatchId")
        return {"success": True, "oracle_id": oracle_id}
    else:
        return {"success": False, "error": resp.text, "status": resp.status_code}
```

**Expected output:** Oracle journal batch ID for each invoice posted.

---

### Step 4 — Update Yve.01 Status After Posting
**File to build:** Addition to `asignador_cuentas.py` or new `oracle_actualizar_estado.py`

After a successful Oracle POST, write back the Oracle journal ID to
`facturas_contabilizadas_*.xlsx` so the dashboard shows "Contabilizada":

```python
def marcar_contabilizada(df, numero_factura, oracle_id):
    """Write oracle_id and status back to the Excel file."""
    mask = df["numero_factura"] == numero_factura
    df.loc[mask, "oracle_status"]   = "CONTABILIZADA"
    df.loc[mask, "oracle_id"]       = oracle_id
    df.loc[mask, "fecha_oracle"]    = datetime.now().strftime("%Y-%m-%d %H:%M")
    return df
```

The dashboard (`dashboard.py`) should show a new badge:
- 🟢 CONTABILIZADA — posted to Oracle with ID
- 🟡 PENDIENTE — approved but not yet posted
- 🔴 ERROR_ORACLE — post attempted but failed

---

### Step 5 — Automated Daily Run + Oracle Posting Confirmation
**File to build:** `oracle_pipeline.py`

Full end-to-end automation with error handling and confirmation emails:

```python
def run_oracle_pipeline():
    print("=== Yve.01 — Oracle Pipeline ===")
    
    # 1. Auth
    token = get_token()
    if not test_connection(token):
        raise ConnectionError("No se pudo conectar a Oracle")
    print("✓ Autenticado en Oracle")
    
    # 2. Load approved invoices
    df = cargar_facturas_para_oracle()
    cta_cble = pd.read_excel("datos-referencia/cta_cble.xlsx")
    print(f"  Facturas aprobadas pendientes: {len(df)}")
    
    resultados = []
    for _, row in df.iterrows():
        lineas = construir_lineas_oracle(row, cta_cble)
        resultado = crear_asiento_oracle(token, row, lineas)
        
        if resultado["success"]:
            print(f"  ✓ {row['numero_factura']} → Oracle ID {resultado['oracle_id']}")
            df = marcar_contabilizada(df, row["numero_factura"], resultado["oracle_id"])
        else:
            print(f"  ✗ {row['numero_factura']} → ERROR: {resultado['error'][:80]}")
        
        resultados.append({**row.to_dict(), **resultado})
    
    # 3. Save updated statuses
    guardar_excel_actualizado(df)
    
    # 4. Summary
    ok  = sum(1 for r in resultados if r.get("success"))
    err = len(resultados) - ok
    print(f"\n✅ Contabilizadas: {ok} | ❌ Errores: {err}")
    return resultados


if __name__ == "__main__":
    run_oracle_pipeline()
```

**Expected output:** All approved invoices posted to Oracle, Excel updated with
Oracle journal IDs, dashboard shows "CONTABILIZADA" badges.

---

## Environment Variables Required

Add to `.env`:

```
ORACLE_BASE_URL=https://your-hotel.oraclecloud.com
ORACLE_CLIENT_ID=your_client_id
ORACLE_CLIENT_SECRET=your_client_secret
ORACLE_LEDGER_NAME=Hilton Barcelona
```

---

## Oracle API Endpoints Reference

| Operation | Method | Endpoint |
|---|---|---|
| Get token | POST | `/oauth/token` |
| List ledgers | GET | `/fscmRestApi/resources/.../ledgers` |
| Create journal batch | POST | `/fscmRestApi/resources/.../journalBatches` |
| Get journal status | GET | `/fscmRestApi/resources/.../journalBatches/{id}` |
| Post journal to GL | POST | `/fscmRestApi/resources/.../journalBatches/{id}/action/post` |
| List suppliers | GET | `/fscmRestApi/resources/.../suppliers` |
| Create supplier invoice | POST | `/fscmRestApi/resources/.../supplierInvoices` |

---

## Integration Notes

1. **Approval gate:** Never post to Oracle without `accion == "APROBADA"`. The department
   head signature is a legal requirement in Spain.

2. **Date format:** Oracle REST API expects dates as `YYYY-MM-DD`. Convert from `DD/MM/YYYY`.

3. **Currency:** All amounts in EUR. Oracle field is `EnteredCurrencyCode: "EUR"`.

4. **Idempotency:** Check `oracle_status` before posting. Never post the same invoice twice.
   Use `JournalBatchName` as unique key — include `numero_factura`.

5. **CtaCble dependency:** The integration requires the hotel's CtaCble sheet loaded
   into `datos-referencia/cta_cble.xlsx`. This is the first thing to configure for each
   new hotel client during onboarding.

6. **Test environment:** Oracle provides a sandbox environment. Always test there first.
   Add `ORACLE_ENV=sandbox` to `.env` and route to the sandbox base URL.

7. **Posting vs. saving:** Oracle journals must be explicitly "posted" after creation
   (Step 5 action/post call). Created journals that are not posted exist as drafts only.

---

## Files to Build (Phase 3)

In priority order:

| File | Description |
|---|---|
| `oracle_auth.py` | OAuth token + connection test |
| `oracle_lector_facturas.py` | Read Yve.01 output → Oracle line dicts |
| `oracle_crear_asientos.py` | POST journal batches to Oracle GL |
| `oracle_actualizar_estado.py` | Write Oracle ID back to Excel + dashboard |
| `oracle_pipeline.py` | Full automated daily run |
| `datos-referencia/cta_cble.xlsx` | Hotel-specific CtaCble mapping (one per client) |
