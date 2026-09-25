# -*- coding: utf-8 -*-
"""b93 — la factura del grupo lleva numero CORRELATIVO de la serie al emitirla.

Lo que se comprueba:
  - una sola serie FAC-<año>-CORP-<nnnn> para el grupo: el siguiente = el MAYOR del año
    + 1 (antes len(fichero)+1: las filas GRP- pendientes dejaban huecos y podia repetir);
  - emitir la factura de un contrato (GRP-<contrato>) le da el siguiente numero
    (`numero_factura`); el GRP- sigue siendo su referencia (`numero_reserva`); la factura a
    mano va a la MISMA serie y no repite;
  - el numero legal se ve (lista AR, PDF), se asienta (documento del cierre) y el contrato
    lo apunta (la factura de comision que lo cite se une por referencia);
  - pantalla y JS.

  python3.12 tests/test_numero_factura.py
  python3.12 tests/test_numero_factura.py --sabotaje
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DIRS = ['datos-referencia', 'facturas-procesadas', 'reportes']
DD = os.path.join(BASE, 'datos-referencia')
PROC = os.path.join(BASE, 'facturas-procesadas')
ANIO = date.today().year


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import tab_ar_real as T
    import contratos_grupo as CG
    import lector_contratos_grupo as L
    import cierre_mes as CM
    import almacen_datos as ALM
    if SABOTAJE:
        # vuelve el contador de antes: numero de filas + 1
        T.siguiente_numero_factura = lambda anio=None, df=None: f"FAC-{int(anio or ANIO)}-CORP-{len(T._get_reservas_todas() if df is None else df) + 1:04d}"

    # 1. la serie (pura)
    df = pd.DataFrame([{"numero_reserva": f"FAC-{ANIO}-CORP-0001"}, {"numero_reserva": f"FAC-{ANIO}-CORP-0003"},
                       {"numero_reserva": "GRP-X"}, {"numero_reserva": f"FAC-{ANIO - 1}-CORP-0009"},
                       {"numero_reserva": "GRP-Y", "numero_factura": f"FAC-{ANIO}-CORP-0004"}])
    ok(T.siguiente_numero_factura(ANIO, df) == f"FAC-{ANIO}-CORP-0005", f"el siguiente es el mayor del año + 1 ({T.siguiente_numero_factura(ANIO, df)})")
    ok(T.siguiente_numero_factura(ANIO - 1, df) == f"FAC-{ANIO - 1}-CORP-0010", 'cada año su serie')
    ok(T.siguiente_numero_factura(ANIO + 1, df) == f"FAC-{ANIO + 1}-CORP-0001", 'año nuevo empieza en 0001')

    copia = tempfile.mkdtemp(prefix='nfac_')
    for d in DIRS:
        if os.path.isdir(d):
            shutil.copytree(d, os.path.join(copia, d))
    try:
        for f in os.listdir(PROC):
            if f.startswith(('facturas_ap_', 'facturas_contabilizadas_')):
                os.remove(os.path.join(PROC, f))
        for f in ('contratos_grupo.json', 'reservas_credito.xlsx', 'clientes_credito.xlsx', 'compensaciones_ar.json', 'ajustes_ap.json'):
            if os.path.exists(os.path.join(DD, f)):
                os.remove(os.path.join(DD, f))
        # dos contratos de grupo (pendientes de emitir) y una factura a mano ya emitida
        for num, ag in (("CG-2026-0951", "Viajes Meridiano S.L."), ("CG-2026-0952", "Eventos Costa DMC S.L.")):
            dt = {"es_contrato_grupo": True, "evento": {"id": "", "nombre": "Evento " + num}, "contrato_numero": num,
                  "cliente": {"nombre": "Laboratorios Norte S.A."}, "agencia": {"nombre": ag, "cif": "B28004554"},
                  "alojamiento": {"fecha_entrada": "2026-08-07", "fecha_salida": "2026-08-10", "noches": 3, "habitaciones": 10,
                                  "total_habitaciones": 5500, "iva_pct": 10},
                  "fb": {"total": 0}, "salas": {"total": 0}, "comisiones": {"modo": "porcentaje", "alojamiento_pct": 10},
                  "facturacion": {"pagador": "agencia"}}
            t = L.transformar(dt, hotel_id="")
            L.guardar(t, DD)
            CG.registrar(dt, t, hotel_id="", datos_dir=DD)
        rv = pd.read_excel(os.path.join(DD, 'reservas_credito.xlsx'))
        rv = pd.concat([rv, pd.DataFrame([{"numero_reserva": f"FAC-{ANIO}-CORP-0001", "cliente": "Otra SA", "total": 100.0,
                                            "estado": "FACTURADO", "fecha_emision": f"{ANIO}-01-10", "habitaciones": 1, "hotel_id": ""}])], ignore_index=True)
        rv.to_excel(os.path.join(DD, 'reservas_credito.xlsx'), index=False)

        import dashboard as D
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client(); assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token'); H = {'X-CSRF-Token': tok}
        # 2. emitir: numero de la serie
        r1 = cl.post('/api/ar_real/emitir_pendiente', json={'numero': 'GRP-CG-2026-0951'}, headers=H).get_json()
        ok(r1.get('ok') and r1.get('numero_factura') == f"FAC-{ANIO}-CORP-0002" and r1.get('numero') == 'GRP-CG-2026-0951',
           f"la factura del contrato toma el siguiente numero de la serie ({r1.get('numero_factura')}) y conserva su referencia")
        r2 = cl.post('/api/ar_real/emitir_factura', json={'cliente': 'Otra SA', 'fecha_entrada': '2026-10-01', 'fecha_salida': '2026-10-02', 'total': 50}, headers=H).get_json()
        ok(r2.get('ok') and r2.get('numero') == f"FAC-{ANIO}-CORP-0003",
           f"la factura a mano va a la MISMA serie y no repite ({r2.get('numero')}; antes: filas + 1)")
        r3 = cl.post('/api/ar_real/emitir_pendiente', json={'numero': 'GRP-CG-2026-0952'}, headers=H).get_json()
        ok(r3.get('ok') and r3.get('numero_factura') == f"FAC-{ANIO}-CORP-0004", f"y la siguiente del otro contrato ({r3.get('numero_factura')})")
        todos = pd.read_excel(os.path.join(DD, 'reservas_credito.xlsx'))
        nums = [n for n in todos.get('numero_factura', pd.Series(dtype=object)).tolist() if isinstance(n, str) and n.startswith('FAC-')] + \
               [n for n in todos['numero_reserva'].tolist() if isinstance(n, str) and n.startswith('FAC-') and n not in todos.get('numero_factura', pd.Series(dtype=object)).tolist()]
        ok(len(nums) == len(set(nums)) == 4, f"ningun numero repetido ({sorted(set(nums))})")
        g1 = todos[todos['numero_reserva'] == 'GRP-CG-2026-0951'].iloc[0]
        ok(g1['numero_factura'] == f"FAC-{ANIO}-CORP-0002" and g1['estado'] == 'FACTURADO', 'la fila del contrato guarda su numero legal')
        # 3. se ve, se imprime, se asienta
        fr = cl.get('/api/ar_real/facturas').get_json()
        ok(fr.get('ok'), f"lista de AR ({fr.get('error')})")
        f1 = next(x for x in fr.get('facturas', []) if x['numero'] == 'GRP-CG-2026-0951')
        ok(f1['numero_factura'] == f"FAC-{ANIO}-CORP-0002", 'la lista de AR ensena el numero legal')
        pdf = cl.get('/api/ar_real/pdf/GRP-CG-2026-0951')
        ruta = tempfile.mktemp(suffix='.pdf'); open(ruta, 'wb').write(pdf.data)
        txt = subprocess.run(['pdftotext', ruta, '-'], capture_output=True, text=True).stdout
        os.remove(ruta)
        ok(pdf.status_code == 200 and f"FAC-{ANIO}-CORP-0002" in txt, 'el PDF de la factura lleva el numero legal')
        ok(cl.get(f'/api/ar_real/pdf/FAC-{ANIO}-CORP-0002').status_code == 200, 'y se encuentra tambien por el numero legal')
        mes = date.today().strftime('%Y-%m')
        res = CM.generar_asientos(mes, CM.recoger_fuentes(mes, None, PROC, 'reportes', DD), CM.plan_cuentas(DD), CM.config_cierre(DD))
        docs = {x['documento'] for x in res['asientos'] if x['origen'] == 'AR'}
        ok(f"FAC-{ANIO}-CORP-0002" in docs and 'GRP-CG-2026-0951' not in docs, f"el cierre asienta con el numero legal ({sorted(docs)})")
        # 4. el contrato lo apunta y la comision que lo cite se une
        c1 = next(c for c in CG.leer(DD) if c['contrato'] == 'CG-2026-0951')
        ok(c1.get('factura_numero') == f"FAC-{ANIO}-CORP-0002", 'el contrato apunta el numero legal de su factura')
        pd.DataFrame([{"archivo": "VM-9.pdf", "numero_factura": "VM-9", "nombre_proveedor": "Viajes Meridiano SL", "NIF_proveedor": "B28004554",
                       "fecha_factura": date.today().strftime('%d/%m/%Y'), "descripcion_concepto": f"Comisión 10 % factura FAC-{ANIO}-CORP-0002",
                       "base_imponible": 500.0, "porcentaje_iva": 21, "cuota_iva": 105.0, "total_factura": 605.0, "tipo_proveedor": "OTRAS",
                       "cuenta_contable": "629", "hotel_id": ""}]).to_excel(os.path.join(PROC, f"facturas_ap_{date.today().strftime('%Y%m%d')}.xlsx"), index=False)
        fap = next(x for x in ALM.facturas_ap().to_dict('records') if x['numero_factura'] == 'VM-9')
        ok(bool(fap.get('es_comision_agencia')) and fap.get('comision_contrato') == c1['id'] and fap.get('comision_vinculo') == 'referencia',
           f"la factura de comision que cita el numero legal se une a su contrato ({fap.get('comision_vinculo')})")
        # 5. pantalla
        html = cl.get('/').get_data(as_text=True)
        ok('f.numero_factura || f.numero' in html and 'd.numero_factura || numero' in html, 'Aging AR ensena el numero legal (y la referencia GRP- debajo)')
        malos = 0
        for bl in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_nfac.js', 'w', encoding='utf-8').write(bl)
            if subprocess.run(['node', '--check', '/tmp/_nfac.js'], capture_output=True, text=True).returncode:
                malos += 1
        ok(malos == 0, 'el JS servido parsea')
        diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
        ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    except Exception as e:
        import traceback; traceback.print_exc()
        ok(False, f'excepcion en la prueba: {type(e).__name__}: {e}')
    finally:
        for d in DIRS:
            shutil.rmtree(d, ignore_errors=True)
            if os.path.isdir(os.path.join(copia, d)):
                shutil.copytree(os.path.join(copia, d), d)
            else:
                os.makedirs(d, exist_ok=True)
        shutil.rmtree(copia, ignore_errors=True)

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
