# -*- coding: utf-8 -*-
"""Los emails de reclamacion AP a proveedores firmaban como "Hilton Barcelona"
con finanzas@hiltonbarcelona.com, a fuego. Ahora firman con el hotel del tenant
(hotel_config.json) o el hotel activo, y sin email si no hay.

  python3.12 tests/test_emails_sin_hilton.py
  python3.12 tests/test_emails_sin_hilton.py --sabotaje
"""
import importlib, json, os, shutil, subprocess, sys, tempfile
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv
CFG = os.path.join(BASE, 'datos-referencia', 'hotel_config.json')


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    prev = open(CFG, encoding='utf-8').read() if os.path.exists(CFG) else None
    try:
        cfg = json.loads(prev) if prev else {}
        cfg.update({"hotel_nombre": "Hotel Els Pins", "hotel_email": "admin@elspins.cat"})
        json.dump(cfg, open(CFG, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        import generador_emails_ap as G
        importlib.reload(G)
        if SABOTAJE:
            G.HOTEL_NOMBRE, G.HOTEL_EMAIL = "Hilton Barcelona", "finanzas@hiltonbarcelona.com"
        fac = {"numero_factura": "DG-2026-0812", "nombre_proveedor": "Distribucions Garraf", "total_factura": 704.0, "importe_po": 640.0, "fecha": "12/08/2026", "descripcion_concepto": "Mercancia"}
        textos = [G.template_discrepancia_po(fac), G.template_sin_po(fac), G.template_alerta_consumo(fac)]
        ok(all("Hilton" not in t and "hiltonbarcelona" not in t for t in textos), "ninguna plantilla menciona Hilton")
        ok(all("Hotel Els Pins" in t for t in textos), "firman con el hotel de hotel_config.json")
        ok(all("admin@elspins.cat" in t for t in textos), "y con su email")
        # sin email en la config: no se inventa uno
        cfg["hotel_email"] = ""; json.dump(cfg, open(CFG, 'w', encoding='utf-8'), ensure_ascii=False)
        importlib.reload(G)
        if SABOTAJE:
            G.HOTEL_NOMBRE, G.HOTEL_EMAIL = "Hilton Barcelona", "finanzas@hiltonbarcelona.com"
        t = G.template_sin_po(fac)
        ok("@" not in t.split("Atentamente")[-1] and "a su disposición" in t, "sin email configurado: no se inventa ninguno")
        src = open(os.path.join(BASE, 'generador_emails_ap.py'), encoding='utf-8').read()
        ok('"Hilton Barcelona"' not in src and 'hiltonbarcelona' not in src, "el fuente ya no lleva el hotel a fuego")
    finally:
        if prev is None:
            if os.path.exists(CFG): os.remove(CFG)
        else:
            open(CFG, 'w', encoding='utf-8').write(prev)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
