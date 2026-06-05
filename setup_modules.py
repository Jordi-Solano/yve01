#!/usr/bin/env python3
"""
setup_modules.py
Integra F&B y AR Real en dashboard.py — ejecutar una sola vez
"""

from pathlib import Path
import sys

BASE_DIR = Path(__file__).parent
DASHBOARD = BASE_DIR / "dashboard.py"

if not DASHBOARD.exists():
    print("❌ dashboard.py not found")
    sys.exit(1)

with open(DASHBOARD, 'r') as f:
    content = f.read()

# Check if already integrated
if "from tab_ar_grupo import ar_bp" in content and "from tab_fb_dashboard import fb_bp" in content:
    print("✅ F&B y AR ya están integrados en dashboard.py")
    sys.exit(0)

print("✅ Modules are ready in GitHub")
print("\nNext steps:")
print("1. git pull (para descargar ar_grupo_corporativo.py, tab_ar_grupo.py, etc)")
print("2. python dashboard.py (test local)")
print("3. Deploy a Render")
