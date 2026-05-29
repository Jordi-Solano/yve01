"""
oracle_auth.py — Yve.01 Módulo Oracle
Gestión de autenticación OAuth 2.0 con Oracle Fusion Cloud Finance.
Si ORACLE_CLIENT_ID / ORACLE_CLIENT_SECRET no están configurados,
activa modo SIMULACIÓN automáticamente.
"""

import os
import time
import json
from pathlib import Path
from datetime import datetime

# ─── Carga de .env ────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
env_path = BASE_DIR / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

# ─── Config Oracle ────────────────────────────────────────────────────────────
ORACLE_BASE_URL     = os.getenv("ORACLE_BASE_URL",     "https://your-hotel.oraclecloud.com")
ORACLE_CLIENT_ID    = os.getenv("ORACLE_CLIENT_ID",    "")
ORACLE_CLIENT_SECRET= os.getenv("ORACLE_CLIENT_SECRET","")
ORACLE_LEDGER_NAME  = os.getenv("ORACLE_LEDGER_NAME",  "Hilton Barcelona")
ORACLE_SCOPE        = os.getenv("ORACLE_SCOPE",        "")

# ─── Modo simulación ──────────────────────────────────────────────────────────
SIMULATION_MODE = not (ORACLE_CLIENT_ID and ORACLE_CLIENT_SECRET
                       and "your-hotel" not in ORACLE_BASE_URL)

# Cache de token en memoria
_token_cache = {
    "access_token": None,
    "expires_at":   0,
}

SIM_TOKEN = "SIM-TOKEN-YVE01-NOPRODUCTION"


def is_simulation() -> bool:
    """True si el módulo opera en modo simulación."""
    return SIMULATION_MODE


def get_token() -> str:
    """
    Obtiene (o renueva) el token OAuth 2.0.
    En modo simulación devuelve un token ficticio.
    """
    if SIMULATION_MODE:
        return SIM_TOKEN

    # ── Usar token cacheado si sigue vigente (margen 60s) ──
    now = time.time()
    if (_token_cache["access_token"]
            and _token_cache["expires_at"] - now > 60):
        return _token_cache["access_token"]

    try:
        import requests
        auth_url = f"{ORACLE_BASE_URL}/oauth/token"
        data = {
            "grant_type":    "client_credentials",
            "client_id":     ORACLE_CLIENT_ID,
            "client_secret": ORACLE_CLIENT_SECRET,
        }
        if ORACLE_SCOPE:
            data["scope"] = ORACLE_SCOPE

        resp = requests.post(
            auth_url,
            data=data,
            auth=(ORACLE_CLIENT_ID, ORACLE_CLIENT_SECRET),
            timeout=30,
        )
        resp.raise_for_status()
        payload = resp.json()
        token   = payload["access_token"]
        expires_in = int(payload.get("expires_in", 3600))

        _token_cache["access_token"] = token
        _token_cache["expires_at"]   = now + expires_in
        return token

    except Exception as e:
        raise ConnectionError(f"Oracle auth failed: {e}") from e


def get_headers() -> dict:
    """Cabeceras HTTP estándar para la REST API de Oracle."""
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type":  "application/json",
        "Accept":        "application/json",
    }


def test_connection() -> dict:
    """
    Prueba la conexión con Oracle.
    En modo simulación siempre devuelve éxito.
    Retorna dict con: ok (bool), mode (str), message (str), ledgers (list)
    """
    if SIMULATION_MODE:
        msg_parts = []
        if not ORACLE_CLIENT_ID:
            msg_parts.append("ORACLE_CLIENT_ID no configurado")
        if not ORACLE_CLIENT_SECRET:
            msg_parts.append("ORACLE_CLIENT_SECRET no configurado")
        if "your-hotel" in ORACLE_BASE_URL:
            msg_parts.append("ORACLE_BASE_URL es el valor por defecto")
        return {
            "ok":      True,
            "mode":    "SIMULACION",
            "message": f"Modo simulación activo — {'; '.join(msg_parts) or 'sin credenciales'}",
            "ledgers": [{"name": ORACLE_LEDGER_NAME, "id": "SIM-001", "status": "simulado"}],
        }

    try:
        import requests
        url = (f"{ORACLE_BASE_URL}/fscmRestApi/resources/"
               "11.13.18.05/ledgers?limit=5")
        resp = requests.get(url, headers=get_headers(), timeout=15)
        resp.raise_for_status()
        data    = resp.json()
        ledgers = [
            {"name": item.get("LedgerName"), "id": item.get("LedgerId"),
             "currency": item.get("CurrencyCode")}
            for item in data.get("items", [])
        ]
        return {
            "ok":      True,
            "mode":    "PRODUCCION",
            "message": f"Conexión Oracle OK — {len(ledgers)} ledger(s) accesibles",
            "ledgers": ledgers,
        }
    except Exception as e:
        return {
            "ok":      False,
            "mode":    "PRODUCCION",
            "message": f"Error de conexión Oracle: {e}",
            "ledgers": [],
        }


def print_status():
    """Imprime el estado del módulo de autenticación."""
    modo = "⚠  SIMULACIÓN" if SIMULATION_MODE else "✅ PRODUCCIÓN"
    print(f"  Oracle Auth  — Modo: {modo}")
    print(f"  Base URL:    {ORACLE_BASE_URL}")
    print(f"  Ledger:      {ORACLE_LEDGER_NAME}")
    if SIMULATION_MODE:
        print("  ┌─────────────────────────────────────────────────────┐")
        print("  │  MODO SIMULACIÓN ACTIVO                             │")
        print("  │  Para activar producción, añade en .env:            │")
        print("  │    ORACLE_BASE_URL=https://tu-hotel.oraclecloud.com │")
        print("  │    ORACLE_CLIENT_ID=tu_client_id                    │")
        print("  │    ORACLE_CLIENT_SECRET=tu_client_secret            │")
        print("  └─────────────────────────────────────────────────────┘")


if __name__ == "__main__":
    print("=" * 60)
    print("  Yve.01 — Oracle Auth: Test de Conexión")
    print("=" * 60)
    print_status()
    print()
    result = test_connection()
    estado = "✅ OK" if result["ok"] else "❌ ERROR"
    print(f"  Estado:  {estado}")
    print(f"  Modo:    {result['mode']}")
    print(f"  Mensaje: {result['message']}")
    if result.get("ledgers"):
        print("  Ledgers disponibles:")
        for l in result["ledgers"]:
            print(f"    - {l.get('name')} (ID: {l.get('id')})")
    print("=" * 60)
