# -*- coding: utf-8 -*-
"""Bug 11 — que una factura ya contabilizada NO se contabilice dos veces.

El caso, reproducido antes de arreglarlo:

  1. día 1 · Oracle contabiliza F-001 y F-002. El marcador `oracle_status`
     queda escrito DENTRO de `facturas_contabilizadas_20260801.xlsx`.
  2. ese fichero se corrompe (un `BadZipFile` con tres guardados a la vez no es
     teórico: pasó en producción antes del candado).
  3. el asignador regenera el informe de hoy leyendo por `almacen_datos`, que
     salta el fichero ilegible CALLANDO. Encuentra las mismas facturas en su
     `facturas_ap_*` de origen — sin marcador — y las da por nuevas.
  4. Oracle abre solo ese informe, no ve marcador, y vuelve a montar el asiento
     de F-001 y F-002 en el libro mayor.

Se comprueban las tres cosas que Jordi pidió:
  · que NO se recontabilizan;
  · que el `oracle_status` sigue intacto donde sí se puede leer;
  · que el candado de la Fase 0 (nada sin APROBADA) sigue cerrado.

  python3.12 tests/test_bug11_no_recontabiliza.py
  python3.12 tests/test_bug11_no_recontabiliza.py --sabotaje
"""
import os
import shutil
import sys
import tempfile
from pathlib import Path

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_TENANT'] = 'default'

import pandas as pd                                    # noqa: E402
import almacen_datos as ALM                            # noqa: E402
import app_aprobacion_ap as PANEL                      # noqa: E402
import oracle_actualizar_estado as OAE                 # noqa: E402
import oracle_lector_facturas as ORA                   # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def fila(n, prov, tot, orc='', cuenta='600'):
    return {'numero_factura': n, 'nombre_proveedor': prov, 'total_factura': tot,
            'base_imponible': round(tot / 1.21, 2), 'iva': round(tot - tot / 1.21, 2),
            'hotel_id': 'H6F22C9', 'fecha_factura': '01/08/2026', 'fecha': '01/08/2026',
            'oracle_status': orc, 'cuenta_contable': cuenta,
            'estado_asignacion': 'ASIGNADA', 'cuenta_iva': '472'}


def montar(raiz, aprobadas=('F-001', 'F-002', 'F-003')):
    """El árbol del día 2, con dos facturas ya contabilizadas el día 1."""
    pdir = os.path.join(raiz, 'facturas-procesadas')
    apdir = os.path.join(raiz, 'aprobaciones')
    rdir = os.path.join(raiz, 'reportes')
    ddir = os.path.join(raiz, 'datos-referencia')
    for d in (pdir, apdir, rdir, ddir):
        os.makedirs(d, exist_ok=True)

    pd.DataFrame([fila('F-001', 'Makro', 121, 'CONTABILIZADA'),
                  fila('F-002', 'Lavanderia', 242, 'CONTABILIZADA')]
                 ).to_excel(os.path.join(pdir, 'facturas_contabilizadas_20260801.xlsx'),
                            index=False)
    # las mismas facturas siguen en su fichero de origen, SIN marcador
    pd.DataFrame([fila('F-001', 'Makro', 121), fila('F-002', 'Lavanderia', 242)]
                 ).to_excel(os.path.join(pdir, 'facturas_ap_20260801.xlsx'), index=False)
    pd.DataFrame([fila('F-003', 'Bebidas', 363)]
                 ).to_excel(os.path.join(pdir, 'facturas_ap_20260802.xlsx'), index=False)
    # Las columnas REALES que lee `cargar_aprobaciones_ap`: `accion` (no
    # `estado`) y `fecha_hora`, que es por donde ordena para quedarse con la
    # ultima decision de cada factura.
    pd.DataFrame([{'numero_factura': n, 'accion': 'APROBADA', 'aprobador': 'jefe',
                   'fecha_hora': '2026-08-02 09:00'} for n in aprobadas]
                 ).to_excel(os.path.join(apdir, 'aprobaciones_ap.xlsx'), index=False)

    PANEL.PROCESADAS_DIR = pdir
    PANEL.REPORTES_DIR = rdir
    PANEL.APROBACIONES_DIR = apdir
    PANEL.APRO_FILE = os.path.join(apdir, 'aprobaciones_ap.xlsx')
    ORA.PROCESADAS_DIR = Path(pdir)
    ORA.REPORTES_DIR = Path(rdir)
    ORA.APROBACIONES_DIR = Path(apdir)
    OAE.REGISTRO_FILE = Path(ddir) / 'oracle_contabilizadas.json'
    return pdir, rdir


def contabilizar(bypass=False):
    batches, bloq, df = ORA.preparar_facturas_para_oracle(bypass_aprobacion=bypass)
    return (sorted(str(b['numero_factura']) for b in batches),
            sorted((str(x['numero_factura']), str(x['motivo'])[:45]) for x in bloq),
            df)


def romper_el_marcador(pdir, rdir):
    """El paso 2 y 3: se corrompe el fichero del marcador y el asignador
    regenera el informe de hoy exactamente como lo hace de verdad."""
    with open(os.path.join(pdir, 'facturas_contabilizadas_20260801.xlsx'), 'wb') as f:
        f.write(b'no soy un zip')
    regen = ALM.facturas_ap(pdir, rdir)          # lo que hace cargar_todas_facturas_ap
    regen.to_excel(os.path.join(pdir, 'facturas_contabilizadas_20260802.xlsx'), index=False)
    return regen


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: se ignora el registro de contabilizadas ***')
        OAE.ya_contabilizadas_registro = lambda: set()
    fallos = 0
    raiz = tempfile.mkdtemp(prefix='b11t_')
    try:
        pdir, rdir = montar(raiz)

        # ── 1 · con todo sano, solo se contabiliza la nueva ──────────
        hacer, bloq, df = contabilizar()
        ok1 = hacer == ['F-003']
        print(f"  {'OK ' if ok1 else 'FALLA'}  con todo sano se contabiliza solo la nueva: "
              f"{hacer}")
        if not ok1:
            fallos += 1

        # ── 2 · el oracle_status llega intacto ──────────────────────
        marc = {str(r['numero_factura']): str(r.get('oracle_status', '')).upper()
                for _, r in df.iterrows()}
        ok2 = marc.get('F-001') == 'CONTABILIZADA' and marc.get('F-002') == 'CONTABILIZADA'
        print(f"  {'OK ' if ok2 else 'FALLA'}  el oracle_status llega intacto: "
              f"F-001={marc.get('F-001')} F-002={marc.get('F-002')}")
        if not ok2:
            fallos += 1

        # ── 3 · EL CASO · se pierde el marcador ─────────────────────
        regen = romper_el_marcador(pdir, rdir)
        sin_marcador = sum(1 for v in regen.get('oracle_status', pd.Series(dtype=object))
                           if str(v).upper() != 'CONTABILIZADA')
        ileg = ALM.ficheros_ilegibles()
        print(f"  ··  reproducido: el informe de hoy sale con {sin_marcador}/3 filas sin "
              f"marcador; ficheros ilegibles detectados: {ileg}")

        hacer2, bloq2, df2 = contabilizar()
        ok3 = hacer2 == ['F-003']
        print(f"  {'OK ' if ok3 else 'FALLA'}  perdido el marcador, NO se recontabiliza: "
              f"{hacer2}  (esperaba ['F-003'])")
        if not ok3:
            fallos += 1

        ok3b = 'F-001' not in hacer2 and 'F-002' not in hacer2
        print(f"  {'OK ' if ok3b else 'FALLA'}  F-001 y F-002 no vuelven al libro mayor")
        if not ok3b:
            fallos += 1

        # ── 4 · el aviso existe (no era un print perdido) ───────────
        ok4 = len(ileg) == 1 and 'contabilizadas_20260801' in ileg[0]
        print(f"  {'OK ' if ok4 else 'FALLA'}  el fichero ilegible se puede nombrar: {ileg}")
        if not ok4:
            fallos += 1

        # ── 5 · el candado de la Fase 0: nada sin APROBADA ──────────
        # Desde el 3 sep 2026 la simulacion respeta la puerta igual que el
        # real; se sigue apagando is_simulation para probar el camino de
        # VERDAD, el que llega al libro mayor, sin depender de esa igualdad.
        _sim = ORA.is_simulation
        ORA.is_simulation = lambda: False
        shutil.rmtree(raiz, ignore_errors=True)
        raiz2 = tempfile.mkdtemp(prefix='b11t2_')
        try:
            pdir2, rdir2 = montar(raiz2, aprobadas=('F-001',))   # F-003 SIN aprobar
            hacer3, bloq3, _ = contabilizar()
            ok5 = hacer3 == [] and any(x[0] == 'F-003' and 'no aprobada' in x[1].lower()
                                       for x in bloq3)
            print(f"  {'OK ' if ok5 else 'FALLA'}  sin APROBADA no pasa nada: "
                  f"contabiliza {hacer3}, bloquea {[x[0] for x in bloq3]}")
            if not ok5:
                fallos += 1

            # y ni siquiera perdiendo el marcador se relaja la puerta
            romper_el_marcador(pdir2, rdir2)
            hacer4, bloq4, _ = contabilizar()
            ok6 = hacer4 == []
            print(f"  {'OK ' if ok6 else 'FALLA'}  la puerta APROBADA aguanta también con "
                  f"el marcador perdido: {hacer4}")
            if not ok6:
                fallos += 1
        finally:
            ORA.is_simulation = _sim
            shutil.rmtree(raiz2, ignore_errors=True)

        # ── 6 · el panel ve lo MISMO que Oracle ─────────────────────
        raiz3 = tempfile.mkdtemp(prefix='b11t3_')
        try:
            pdir3, rdir3 = montar(raiz3)
            romper_el_marcador(pdir3, rdir3)
            _, _, df_ora = contabilizar()
            df_pan = PANEL._facturas_crudas()
            v_ora = sorted(df_ora['numero_factura'].astype(str))
            v_pan = sorted(df_pan['numero_factura'].astype(str))
            ok7 = v_ora == v_pan == ['F-001', 'F-002', 'F-003']
            print(f"  {'OK ' if ok7 else 'FALLA'}  el panel y Oracle ven lo mismo: "
                  f"panel {v_pan} · oracle {v_ora}")
            if not ok7:
                fallos += 1
        finally:
            shutil.rmtree(raiz3, ignore_errors=True)
    finally:
        shutil.rmtree(raiz, ignore_errors=True)

    print()
    if SABOTAJE:
        if fallos:
            print(f'SABOTAJE OK: {fallos} en rojo.')
            return 0
        print('SABOTAJE MAL.')
        return 1
    if fallos:
        print(f'{fallos} en rojo')
        return 1
    print('Todo OK. Una factura ya contabilizada no vuelve al libro mayor aunque '
          'se pierda su marcador, y la puerta APROBADA sigue cerrada.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
