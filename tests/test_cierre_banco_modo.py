# -*- coding: utf-8 -*-
"""b80 — el Cierre ve el banco en modo 'grupo' aunque haya un hotel elegido.

Visto al repasar el plan de pruebas de Els Pins: con el banco en modo grupo
(el extracto no lleva hotel_id) y un hotel activo, el cuadre de banco, los
asientos 572 y la reconciliacion del Cierre salian VACIOS, mientras la
pestaña Banco si enseñaba los 12 movimientos. Regla: el Cierre acota el
extracto por hotel SOLO cuando el banco va por hotel (config_banco), igual
que hace /api/stats_banco.

  python3.12 tests/test_cierre_banco_modo.py
  python3.12 tests/test_cierre_banco_modo.py --sabotaje
"""
import json
import os
import shutil
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes', 'aprobaciones']


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import almacen_datos as ALM
    if SABOTAJE:
        ALM.banco_del_hotel = lambda bk, hotel: ALM._filtrar_hotel(bk, hotel)   # vuelve a filtrar siempre
        import tab_cierre, cierre_mes  # noqa: F401  (usan ALM.banco_del_hotel por atributo)
    copia = tempfile.mkdtemp(prefix='cbm_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token')
        H = {'X-CSRF-Token': tok}
        r = cl.post('/admin/api/hoteles/crear', json={'nombre': 'Hotel Modo Banco', 'ciudad': 'X', 'grupo': 'G', 'habitaciones': 10}).get_json()
        hid = r['id']
        assert cl.post('/api/hotel_activo', json={'hotel': hid}, headers=H).status_code == 200
        # extracto del grupo: SIN hotel_id, como lo deja Procesar archivos en modo grupo
        bk = pd.DataFrame([
            {'fecha': '05/08/2026', 'concepto': 'TRANSFERENCIA BOOKING.COM B.V. LIQUIDACION', 'importe': 21450.0, 'saldo': 60000.0},
            {'fecha': '14/08/2026', 'concepto': 'TRANSF. DISTRIBUCIONS GARRAF SL FRA DG-1', 'importe': -704.0, 'saldo': 59296.0},
            {'fecha': '20/08/2026', 'concepto': 'NOMINAS AGOSTO', 'importe': -23410.0, 'saldo': 35886.0},
        ])
        bk.to_excel(os.path.join('datos-referencia', 'extracto_banco.xlsx'), index=False)
        for f in os.listdir('reportes'):
            if f.startswith('conciliacion_'):
                os.remove(os.path.join('reportes', f))

        # ── modo grupo: el cierre ve los 3 movimientos con el hotel elegido ──
        assert cl.post('/api/config_banco', json={'modo': 'grupo'}, headers=H).get_json().get('modo') == 'grupo'
        cu = cl.get('/api/cuadre_banco?mes=2026-08').get_json()
        ok(cu.get('ok_api') and cu.get('n') == 3 and cu['pestanas']['AR']['n'] == 1 and cu['pestanas']['AP']['n'] == 1 and cu['pestanas']['VARIOS']['n'] == 1,
           f"modo grupo + hotel activo: el cuadre de banco ve los 3 movimientos (AR 1 / AP 1 / VARIOS 1) → n={cu.get('n')}")
        sb = cl.get('/api/stats_banco').get_json() or {}
        ok(sb.get('total') == 3 or len(sb.get('movimientos') or []) == 3, f"la pestaña Banco ve lo mismo ({sb.get('total')})")
        import cierre_mes as CM
        fu = CM.recoger_fuentes('2026-08', hotel=hid, procesadas_dir='facturas-procesadas', reportes_dir='reportes', datos_dir='datos-referencia')
        ok(len(fu['banco']) == 3, f"recoger_fuentes (asientos, reconciliacion, paquete): banco con 3 filas ({len(fu['banco'])})")

        # ── modo por hotel: solo lo etiquetado con ese hotel ──
        assert cl.post('/api/config_banco', json={'modo': 'por_hotel'}, headers=H).get_json().get('modo') == 'por_hotel'
        bk2 = bk.copy(); bk2['hotel_id'] = [hid, hid, 'HOTRO']
        bk2.to_excel(os.path.join('datos-referencia', 'extracto_banco.xlsx'), index=False)
        cu = cl.get('/api/cuadre_banco?mes=2026-08').get_json()
        ok(cu.get('n') == 2, f"modo por hotel: solo los 2 movimientos del hotel ({cu.get('n')})")
        fu = CM.recoger_fuentes('2026-08', hotel=hid, procesadas_dir='facturas-procesadas', reportes_dir='reportes', datos_dir='datos-referencia')
        ok(len(fu['banco']) == 2, f"recoger_fuentes por hotel: 2 filas ({len(fu['banco'])})")
        # sin hotel elegido: todo, en cualquier modo
        assert cl.post('/api/hotel_activo', json={'hotel': ''}, headers=H).status_code == 200
        cu = cl.get('/api/cuadre_banco?mes=2026-08').get_json()
        ok(cu.get('n') == 3, f"vista de grupo (sin hotel): los 3 ({cu.get('n')})")
    finally:
        for d in DIRS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)):
                shutil.copytree(os.path.join(copia, d), d)
        shutil.rmtree(copia, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
