# -*- coding: utf-8 -*-
"""El Excel de la pestaña Banco: sin datos devolvia un JSON 404 (el navegador lo
"descargaba"), y con un hotel elegido mandaba los movimientos de TODOS los
hoteles (cajon 28). Ahora: Excel siempre, y el mismo filtro que la pestaña.

  python3.12 tests/test_export_banco.py
  python3.12 tests/test_export_banco.py --sabotaje
"""
import io, json, os, shutil, subprocess, sys, tempfile
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
import pandas as pd            # noqa: E402
SABOTAJE = '--sabotaje' in sys.argv
EXT = os.path.join(BASE, 'datos-referencia', 'extracto_banco.xlsx')
CFG = os.path.join(BASE, 'datos-referencia', 'config_banco.json')


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    import config_banco as CB
    if SABOTAJE:
        CB.modo = lambda: "grupo"     # el export deja de mirar la config: manda todo
    tmp = tempfile.mkdtemp(prefix='xb_'); copias = {}
    for f in (EXT, CFG):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        for f in (EXT, CFG):
            if os.path.exists(f): os.remove(f)
        import dashboard as D
        cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
        r = cl.get('/api/exportar/banco')
        ok(r.status_code == 200 and r.data[:2] == b'PK', f"sin extracto: un Excel igualmente ({r.status_code}), no un JSON de error")
        pd.DataFrame([{'fecha': '05/08/2026', 'concepto': 'LUZ H1', 'importe': -100.0, 'saldo': 900.0, 'hotel_id': 'H1'},
                      {'fecha': '06/08/2026', 'concepto': 'COBRO H2', 'importe': 50.0, 'saldo': 950.0, 'hotel_id': 'H2'}]).to_excel(EXT, index=False)
        r = cl.get('/api/exportar/banco'); df = pd.read_excel(io.BytesIO(r.data))
        ok(r.status_code == 200 and len(df) == 2, f"modo grupo (por defecto): los 2 movimientos ({len(df)})")
        json.dump({"modo": "por_hotel", "elegido": True}, open(CFG, 'w'))
        import censo_hoteles as CH
        _act = CH.activo
        CH.activo = lambda: 'H1'
        try:
            r = cl.get('/api/exportar/banco'); df = pd.read_excel(io.BytesIO(r.data))
            ok(len(df) == 1 and str(df.iloc[0]['concepto']) == 'LUZ H1', f"modo por hotel con H1 activo: solo lo de H1 ({len(df)} fila(s))")
        finally:
            CH.activo = _act
    finally:
        for f in (EXT, CFG):
            if os.path.exists(f): os.remove(f)
            if f in copias: shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
