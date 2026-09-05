# -*- coding: utf-8 -*-
"""Re-conciliar pisaba las asignaciones manuales del banco (cajon 8): el
informe se regeneraba desde cero. Ahora `conservar_manuales` las vuelve a
aplicar por la clave del movimiento.

  python3.12 tests/test_reconciliar_manuales.py
  python3.12 tests/test_reconciliar_manuales.py --sabotaje
"""
import glob, os, shutil, subprocess, sys, tempfile
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
import pandas as pd            # noqa: E402
SABOTAJE = '--sabotaje' in sys.argv
REP = os.path.join(BASE, 'reportes')


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    import conciliacion_bancaria as C
    import almacen_datos as ALM
    if SABOTAJE:
        C.conservar_manuales = lambda df, reportes_dir=None: 0
    tmp = tempfile.mkdtemp(prefix='rc_'); copias = {}
    prev = glob.glob(os.path.join(REP, 'conciliacion_*.xlsx'))
    for f in prev:
        copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        for f in prev: os.remove(f)
        ext = pd.DataFrame([{'fecha': '05/08/2026', 'concepto': 'TRANSF SIN NUMERO', 'importe': -704.0, 'saldo': 900.0, 'tipo': '', 'referencia': ''},
                            {'fecha': '06/08/2026', 'concepto': 'RECIBO LUZ FRA EL-88213', 'importe': -1175.8, 'saldo': -275.8, 'tipo': '', 'referencia': ''}])
        # informe anterior: el primero asignado A MANO a DG-2026-0812
        viejo = ext.copy(); viejo['estado'] = 'PENDIENTE'; viejo['factura_ref'] = ''; viejo['origen'] = ''; viejo['match_proveedor'] = ''; viejo['diferencia'] = 0.0
        viejo.loc[0, 'estado'] = 'CONCILIADO'; viejo.loc[0, 'factura_ref'] = 'DG-2026-0812'; viejo.loc[0, 'origen'] = 'MANUAL'
        os.makedirs(REP, exist_ok=True)
        viejo.to_excel(os.path.join(REP, 'conciliacion_20260901.xlsx'), index=False)
        # re-conciliar sin facturas: el cruce automatico deja todo PENDIENTE...
        res = C.conciliar(ext.copy(), [])
        ruta = C.generar_reporte(res)
        nuevo = pd.read_excel(ruta)
        f0 = nuevo[nuevo['concepto'] == 'TRANSF SIN NUMERO'].iloc[0]
        ok(str(f0['estado']) == 'CONCILIADO' and str(f0['factura_ref']) == 'DG-2026-0812' and str(f0['origen']) == 'MANUAL',
           f"...pero la asignacion manual sobrevive: {f0['estado']} / {f0['factura_ref']} / {f0['origen']}")
        f1 = nuevo[nuevo['concepto'].astype(str).str.contains('LUZ')].iloc[0]
        ok(str(f1['estado']) != 'CONCILIADO' or str(f1['origen']) != 'MANUAL', "y el otro movimiento no hereda nada")
    finally:
        for f in glob.glob(os.path.join(REP, 'conciliacion_*.xlsx')): os.remove(f)
        for f, c in copias.items(): shutil.copy(c, f)
        shutil.rmtree(tmp, ignore_errors=True)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
