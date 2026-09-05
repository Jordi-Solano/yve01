# -*- coding: utf-8 -*-
"""DRR en formato español (Jordi, sep 2026): las tarjetas salen '16.360 €' /
'79,20 €' y `num_drr` (lo que relee el agregador del grupo y el cierre)
entiende ese formato Y los dos antiguos ('€16,360', '40,130.50 EUR').

  python3.12 tests/test_drr_formato_es.py
  python3.12 tests/test_drr_formato_es.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    num = D.num_drr
    if SABOTAJE:
        def num(s):     # el parser viejo: quita '€', '%' y comas
            t = str(s).replace("€", "").replace("%", "").replace(",", "").replace("EUR", "").strip()
            try:
                return float(t)
            except ValueError:
                return None
    casos = {"€16,360": 16360.0, "16,360 EUR": 16360.0, "40,130.50 EUR": 40130.5, "83.70": 83.7,
             "16.360 €": 16360.0, "83,70 €": 83.7, "40.130,50 €": 40130.5, "1.234.567,89 €": 1234567.89,
             "€402,515 ~": 402515.0, "402.515 € ~": 402515.0, "7,200": 7200.0, "N/D": None, "": None}
    mal = {k: num(k) for k, v in casos.items() if num(k) != v}
    ok(not mal, f"num_drr entiende los tres formatos: {mal or 'todos'}")
    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
    ok('return _eur_es(f, dec)' in src and 'f"{_eur_es(rev_val*p, 0)} ~"' in src and 'return f"€{f:,.{dec}f}"' not in src, "las tarjetas del DRR y el GOP derivado salen por _eur_es")
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
