# -*- coding: utf-8 -*-
"""El cliente AR nace solo del bono, con ficha "pendiente de completar",
editable desde AR Real (Jordi, sep 2026).

  python3.12 tests/test_cliente_desde_bono.py
  python3.12 tests/test_cliente_desde_bono.py --sabotaje
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DATOS = os.path.join(BASE, 'datos-referencia')
CLI = os.path.join(DATOS, 'clientes_credito.xlsx')
BON = os.path.join(DATOS, 'bonos_agencia.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import tab_ar_real as AR
    if SABOTAJE:
        AR.alta_cliente_desde_bono = lambda *a, **k: "YA_EXISTE"     # el bono deja de crear clientes

    tmp = tempfile.mkdtemp(prefix='cb_'); copias = {}
    for f in (CLI, BON):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        for f in (CLI, BON):
            if os.path.exists(f):
                os.remove(f)
        import dashboard as D
        # 1. el bono entra por el punto unico de guardado y crea la ficha
        r = D._guardar_bono({"clave": "VM-7781", "numero_bono": "VM-7781", "agencia": "Viatges Mediterrani", "NIF_agencia": "B-62233445",
                             "huesped": "Familia Puig", "fecha_entrada": "14/08/2026", "fecha_salida": "17/08/2026", "importe_total": 1320.0})
        ok(r == "CREADO", f"bono nuevo → cliente CREADO ({r})")
        g = pd.read_excel(CLI) if os.path.exists(CLI) else pd.DataFrame()
        fila = g[g["nombre_cliente"] == "Viatges Mediterrani"].iloc[0] if len(g) else None
        ok(fila is not None and str(fila.get("nif")) == "B-62233445" and str(fila.get("estado_ficha")) == "PENDIENTE" and float(fila.get("credito_limite")) == 0,
           f"ficha: nif={fila.get('nif') if fila is not None else '?'} estado={fila.get('estado_ficha') if fila is not None else '?'} limite=0")
        ok(fila is not None and str(fila.get("origen")) == "bono VM-7781", f"origen anotado: {fila.get('origen') if fila is not None else '?'}")
        # 2. el mismo bono otra vez: no duplica
        r2 = D._guardar_bono({"clave": "VM-7781", "numero_bono": "VM-7781", "agencia": "viatges mediterrani", "NIF_agencia": "B-62233445", "importe_total": 1320.0})
        g = pd.read_excel(CLI) if os.path.exists(CLI) else pd.DataFrame(columns=["nombre_cliente"])
        ok(r2 == "YA_EXISTE" and int((g["nombre_cliente"].str.lower() == "viatges mediterrani").sum()) == 1, f"reprocesar el bono no duplica ({r2}, {len(g)} fila(s))")
        # 3. el panel lo marca como pendiente y una persona lo completa desde AR Real
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        d = cl.get('/api/ar_real/clientes').get_json() or {}
        c = next((x for x in d.get('clientes', []) if x['nombre'] == 'Viatges Mediterrani'), None)
        ok(c is not None and c.get('pendiente_completar') is True and c.get('NIF') == 'B-62233445' and 'bono' in (c.get('origen') or ''), f"/api/ar_real/clientes: pendiente_completar={c and c.get('pendiente_completar')}")
        r = cl.post('/api/ar_real/cliente', json={"nombre": "Viatges Mediterrani", "nif": "B-62233445", "limite": 15000, "dias_pago": 45, "email": "admin@viatges.cat"}, headers=H)
        ok(r.status_code == 200 and (r.get_json() or {}).get('ok'), f"completar la ficha desde AR Real → {r.status_code}")
        g = pd.read_excel(CLI); f2 = g[g["nombre_cliente"] == "Viatges Mediterrani"].iloc[0]
        ok(float(f2["credito_limite"]) == 15000 and str(f2["estado_ficha"]) == "COMPLETA" and str(f2.get("origen")) == "bono VM-7781" and len(g) == 1,
           f"ficha completada: limite {f2['credito_limite']}, estado {f2['estado_ficha']}, origen conservado, sin duplicar")
        d = cl.get('/api/ar_real/clientes').get_json() or {}
        c = next((x for x in d.get('clientes', []) if x['nombre'] == 'Viatges Mediterrani'), None)
        ok(c is not None and c.get('pendiente_completar') is False, "ya no esta pendiente")
        # 4. bono de una agencia que YA tiene ficha sin NIF: se rellena el NIF y nada mas
        pd.DataFrame([{"nombre_cliente": "Empresa Ficha SA", "nif": "", "credito_limite": 9000, "dias_pago": 30}]).to_excel(CLI, index=False)
        r3 = AR.alta_cliente_desde_bono("Empresa Ficha SA", "B-11111111", "X-1", DATOS, "H1")
        g = pd.read_excel(CLI)
        ok(r3 in ("NIF_RELLENADO", "YA_EXISTE") and str(g.iloc[0]["nif"]) == "B-11111111" and float(g.iloc[0]["credito_limite"]) == 9000, f"ficha existente sin NIF: {r3}, limite intacto")
        # 5. la pantalla
        html = cl.get('/').get_data(as_text=True)
        ok("editarClienteAR(" in html and "ar.pendienteCompletar" in html and "function abrirNuevoCliente(cli)" in html, "AR Real: badge 'pendiente de completar' y boton ✎ que abre el modal prellenado")
        ok("completa su ficha" in open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read(), "el log del lote avisa del cliente nuevo")
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_cb.js', 'w', encoding='utf-8').write(b)
            rc = subprocess.run(['node', '--check', '/tmp/_cb.js'], capture_output=True, text=True)
            if rc.returncode:
                ok(False, f"JS roto: {rc.stderr[:100]}"); break
    finally:
        for f in (CLI, BON):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True)
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
