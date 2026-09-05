# -*- coding: utf-8 -*-
"""Emitir factura corporativa daba 403 "CSRF invalido": el fetch no mandaba el
header. Ahora todos los POST JSON del panel pasan por _postJson.

  python3.12 tests/test_emitir_csrf.py
  python3.12 tests/test_emitir_csrf.py --sabotaje
"""
import os, re, subprocess, sys
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    import dashboard as D
    cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
    html = cl.get('/').get_data(as_text=True)
    if SABOTAJE:
        html = html.replace("_postJson('/api/ar_real/emitir_factura'", "fetch('/api/ar_real/emitir_factura', {method:'POST', headers:{'Content-Type':'application/json'}, body:''")
    ok("_postJson('/api/ar_real/emitir_factura'" in html, "Emitir factura va por _postJson (con CSRF)")
    # ningun fetch POST con JSON sin el header CSRF en el HTML servido
    sospechosos = [m.start() for m in re.finditer(r"method:\s*'POST'", html)]
    sin_csrf = [html[max(0, i-200):i+160] for i in sospechosos if 'X-CSRF-Token' not in html[max(0, i-200):i+260] and 'FormData' not in html[max(0, i-300):i+300]]
    ok(not sin_csrf, f"ningun POST JSON suelto sin CSRF ({len(sin_csrf)})" + (": " + sin_csrf[0][-120:].replace(chr(10), ' ') if sin_csrf else ""))
    tok = (cl.get('/api/csrf_token').get_json() or {}).get('token')
    r = cl.post('/api/ar_real/emitir_factura', json={"cliente": "", "fecha_entrada": "2026-08-01", "fecha_salida": "2026-08-02", "habitaciones": 1, "precio_noche": 100, "fb_extras": 0, "total": 110}, headers={'X-CSRF-Token': tok})
    ok(r.status_code != 403, f"con CSRF el endpoint ya no responde 403 ({r.status_code})")
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
