# -*- coding: utf-8 -*-
"""La mitad AR del Libro Diario + el alta de clientes de credito.

LIBRO DIARIO — `_get_facturas_ar()` leia `facturas_ota_demo.xlsx`, un fichero
de demo que ya no existe en el repo: la mitad AR del Libro Diario salia
SIEMPRE a cero, en silencio (sin error, sin aviso). Ahora lee de
`almacen_datos.facturas_ar()`, el punto unico de lectura.

CLIENTES DE CREDITO — no habia forma de dar de alta ninguno: el unico que
escribia `clientes_credito.xlsx` era el generador de demo. Sin clientes no hay
limite de credito ni aviso de riesgo.

  python3.12 tests/test_ar_libro_y_clientes.py
  python3.12 tests/test_ar_libro_y_clientes.py --sabotaje
"""
import json
import os
import shutil
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import pandas as pd                                   # noqa: E402
import almacen_datos                                  # noqa: E402
import exportador_asientos as EX                      # noqa: E402
import tab_ar_real as T                               # noqa: E402
from tenant_dirs import procesadas_dir, datos_dir     # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
HOY = datetime.now().strftime('%Y%m%d')

# Dos facturas de OTA, como las que escribe el procesado real: la columna se
# llama `nombre_ota` y `importe_comision`, NO `ota` y `comision`.
FILAS_AR = [
    {'numero_factura': 'BK-2026-001', 'nombre_ota': 'Booking.com',
     'periodo_inicio': '2026-07-01', 'importe_bruto': 12000.0,
     'importe_comision': 1800.0, 'hotel_id': 'HTEST01'},
    {'numero_factura': 'EX-2026-002', 'nombre_ota': 'Expedia',
     'periodo_inicio': '2026-07-01', 'importe_bruto': 5000.0,
     'importe_comision': 750.0, 'hotel_id': 'HTEST01'},
]


def _guardado(ruta):
    """Guarda el fichero de al lado para devolverlo tal cual al terminar."""
    copia = ruta + '.bak-ar'
    habia = os.path.exists(ruta)
    if habia:
        shutil.copy(ruta, copia)
    return habia, copia


def _restaurar(ruta, habia, copia):
    if os.path.exists(ruta):
        os.remove(ruta)
    if habia:
        shutil.move(copia, ruta)


def main():
    if SABOTAJE:
        print('*** MODO SABOTAJE: el Libro Diario vuelve al fichero de demo y'
              ' el alta de clientes no escribe ***')
    fallos = 0

    ruta_ar = os.path.join(procesadas_dir(), f'facturas_procesadas_{HOY}.xlsx')
    ruta_cli = os.path.join(datos_dir(), 'clientes_credito.xlsx')
    hab_ar, cop_ar = _guardado(ruta_ar)
    hab_cli, cop_cli = _guardado(ruta_cli)

    try:
        # ══ 1 · LIBRO DIARIO ═══════════════════════════════════════════
        pd.DataFrame(FILAS_AR).to_excel(ruta_ar, index=False)
        almacen_datos.facturas_ar.__wrapped__ if False else None

        lista = [] if SABOTAJE else EX._get_facturas_ar()

        # a · llegan las dos
        ok_n = len(lista) == 2
        print(f"  {'OK ' if ok_n else 'FALLA'}  el Libro Diario ve las facturas AR: "
              f"{len(lista)} (esperaba 2)")
        if not ok_n:
            fallos += 1

        # b · los nombres de columna traducidos
        d = lista[0] if lista else {}
        ok_map = (str(d.get('ota')) == 'Booking.com'
                  and round(float(d.get('comision') or 0), 2) == 1800.0
                  and str(d.get('numero_factura')) == 'BK-2026-001')
        print(f"  {'OK ' if ok_map else 'FALLA'}  nombre_ota->ota e importe_comision->comision "
              f"({d.get('ota')!r} · {d.get('comision')!r})")
        if not ok_map:
            fallos += 1

        # c · el asiento sale, con su cuenta y cuadrado
        asientos = EX.generar_libro_diario()
        if SABOTAJE:
            asientos = [a for a in asientos if 'Comisión' not in str(a.get('concepto'))]
        ar = [a for a in asientos if 'Comisión' in str(a.get('concepto'))]
        cuentas = sorted({str(a.get('cuenta')) for a in ar})
        ok_asi = len(ar) >= 6 and '628' in cuentas          # 3 lineas x 2 facturas
        print(f"  {'OK ' if ok_asi else 'FALLA'}  se generan los asientos de comision: "
              f"{len(ar)} lineas, cuentas {cuentas}")
        if not ok_asi:
            fallos += 1

        debe = round(sum(float(a.get('debe') or 0) for a in ar), 2)
        haber = round(sum(float(a.get('haber') or 0) for a in ar), 2)
        ok_cuadra = bool(ar) and abs(debe - haber) < 0.01 and abs(haber - 2550.0) < 0.01
        print(f"  {'OK ' if ok_cuadra else 'FALLA'}  cuadran y suman la comision real: "
              f"debe {debe} / haber {haber} (esperaba 2550.0 = 1800 + 750)")
        if not ok_cuadra:
            fallos += 1

        # d · sin facturas AR, cero asientos AR — y NINGUN error
        os.remove(ruta_ar)
        vacio = EX.generar_libro_diario()
        ar0 = [a for a in vacio if 'Comisión' in str(a.get('concepto'))]
        ok_vacio = len(ar0) == 0
        print(f"  {'OK ' if ok_vacio else 'FALLA'}  sin facturas AR no revienta: "
              f"{len(ar0)} asientos de comision")
        if not ok_vacio:
            fallos += 1

        # ══ 2 · ALTA DE CLIENTES DE CREDITO ════════════════════════════
        if os.path.exists(ruta_cli):
            os.remove(ruta_cli)
        # Contra la app DE VERDAD, no un Flask pelado: asi pasa por el login y
        # por el guardia de CSRF, que es lo que se va a encontrar en Render.
        import dashboard
        app = dashboard.app
        app.config['TESTING'] = True
        c = app.test_client()
        assert c.post('/api/login', json={'username': 'admin',
                                          'password': 'admin123'}).status_code == 200
        TOK = (c.get('/api/csrf_token').get_json() or {}).get('token')
        assert TOK, 'sin token de CSRF el POST daria 403'

        def alta(payload, tok=None):
            if SABOTAJE:
                class _R:
                    status_code = 200

                    def get_json(self):
                        return {'ok': True, 'cliente': payload.get('nombre'), 'total': 0}
                return _R()
            return c.post('/api/ar_real/cliente', json=payload,
                          headers={'X-CSRF-Token': TOK if tok is None else tok})

        r = alta({'nombre': 'Viajes Meridiano S.A.', 'nif': 'A28004556',
                  'limite': 25000, 'dias_pago': 45, 'email': 'cuentas@meridiano.es'})
        j = r.get_json()
        df = pd.read_excel(ruta_cli) if os.path.exists(ruta_cli) else pd.DataFrame()
        ok_alta = (j.get('ok') and len(df) == 1
                   and str(df.iloc[0]['nombre_cliente']) == 'Viajes Meridiano S.A.'
                   and float(df.iloc[0]['credito_limite']) == 25000.0
                   and int(df.iloc[0]['dias_pago']) == 45)
        print(f"  {'OK ' if ok_alta else 'FALLA'}  se da de alta el cliente y se guarda "
              f"({len(df)} fila/s en el fichero)")
        if not ok_alta:
            fallos += 1

        # el nombre es la identidad: el mismo cliente se ACTUALIZA
        alta({'nombre': 'Viajes Meridiano S.A.', 'limite': 40000})
        df2 = pd.read_excel(ruta_cli) if os.path.exists(ruta_cli) else pd.DataFrame()
        if SABOTAJE:
            df2 = df
        ok_upd = (len(df2) == 1
                  and float(df2.iloc[0]['credito_limite']) == (25000.0 if SABOTAJE else 40000.0))
        print(f"  {'OK ' if ok_upd else 'FALLA'}  dar de alta el MISMO nombre actualiza, "
              f"no duplica ({len(df2)} fila/s, límite "
              f"{df2.iloc[0]['credito_limite'] if len(df2) else '—'})")
        if not ok_upd:
            fallos += 1

        # otro cliente distinto SI se añade
        alta({'nombre': 'Corporate Travel S.L.', 'limite': 10000})
        df3 = pd.read_excel(ruta_cli) if os.path.exists(ruta_cli) else pd.DataFrame()
        ok_dos = len(df3) == (1 if SABOTAJE else 2)
        print(f"  {'OK ' if ok_dos else 'FALLA'}  un cliente distinto se añade "
              f"({len(df3)} fila/s)")
        if not ok_dos:
            fallos += 1

        # y la pantalla lo ve
        if not SABOTAJE:
            rr = c.get('/api/ar_real/clientes')
            jj = rr.get_json() or {}
            nombres = sorted(x.get('nombre') for x in (jj.get('clientes') or []))
        else:
            nombres = []
        ok_ve = nombres == ['Corporate Travel S.L.', 'Viajes Meridiano S.A.']
        print(f"  {'OK ' if ok_ve else 'FALLA'}  la pantalla de AR Real los enseña: {nombres}")
        if not ok_ve:
            fallos += 1

        # sin nombre o sin limite, no entra basura
        if SABOTAJE:
            malos = [400, 400]
        else:
            malos = [alta({'limite': 100}).status_code,
                     alta({'nombre': 'X', 'limite': 0}).status_code]
        ok_mal = malos == [400, 400]
        print(f"  {'OK ' if ok_mal else 'FALLA'}  sin nombre o con límite 0 se rechaza: {malos}")
        if not ok_mal:
            fallos += 1

        # y sin token de CSRF no entra nadie (el guardia de siempre sigue ahí)
        sin_tok = 403 if SABOTAJE else alta({'nombre': 'Coladero S.L.', 'limite': 1},
                                            tok='invalido').status_code
        ok_csrf = sin_tok == 403
        print(f"  {'OK ' if ok_csrf else 'FALLA'}  sin token válido de CSRF se rechaza: {sin_tok}")
        if not ok_csrf:
            fallos += 1

    finally:
        _restaurar(ruta_ar, hab_ar, cop_ar)
        _restaurar(ruta_cli, hab_cli, cop_cli)

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
    print('Todo OK. El Libro Diario ve las facturas AR reales y se pueden dar '
          'de alta clientes de crédito.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
