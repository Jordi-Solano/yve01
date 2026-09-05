# -*- coding: utf-8 -*-
"""Lote 2 de Jordi (fases 1-4): euros en español en el log, incidencias de
albaran CON nombre, y servicio → NO_REQUIERE_ALBARAN aunque sea anterior al
primer albaran del hotel.

  python3.12 tests/test_lote2_log.py
  python3.12 tests/test_lote2_log.py --sabotaje
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    import matching_ap_albaran as MA
    if SABOTAJE:
        D._eur_es = lambda v, dec=2: f"€{float(v):,.{dec}f}"
        MA._exige_albaran = lambda fila: True

    ok(D._eur_es(1175.8) == '1.175,80 €' and D._eur_es(704) == '704,00 €' and D._eur_es(23410, 0) == '23.410 €', f"formato: {D._eur_es(1175.8)} · {D._eur_es(23410, 0)}")
    ok(D._resumen_factura_ap([{'nombre_proveedor': 'Energia Llevant SA', 'total_factura': 1175.8}]) == 'Energia Llevant SA — 1.175,80 €', "log de factura en español")
    ok('640,00 €' in D._resumen_albaran({'numero_albaran': '5531', 'nombre_proveedor': 'X', 'total_albaran': 640.0}, [1, 2]), "log de albaran en español")
    ok('1.320,00 €' in D._resumen_bono({'numero_bono': 'VM-7781', 'agencia': 'V', 'importe_total': 1320.0}), "log de bono en español")
    src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
    import re
    restos = [l for l in src.splitlines() if re.search(r"€\{", l) and 'is_eur' not in l and 'GOP' not in l and 'dec}f' not in l]   # DRR: su parser espera '€402,515'
    ok(not restos, f"no queda ningun '€{{...}}' americano en los mensajes ({len(restos)})")

    # servicio anterior al primer albaran del hotel → NO_REQUIERE_ALBARAN, no ANTERIOR_AL_REGISTRO
    df_alb = pd.DataFrame(columns=["numero_albaran", "nombre_proveedor", "total_albaran", "fecha_entrega", "hotel_id"])
    import datetime as _dt
    serv = pd.Series({"numero_factura": "AM-1", "nombre_proveedor": "Assegurances Mar SA", "fecha": "01/08/2026", "base_imponible": 890.0,
                      "tipo_proveedor": "OTRAS", "cuenta_contable": "625", "hotel_id": "H1"})
    merc = pd.Series({"numero_factura": "DG-0", "nombre_proveedor": "Garraf", "fecha": "01/08/2026", "base_imponible": 100.0,
                      "tipo_proveedor": "FB", "cuenta_contable": "600", "hotel_id": "H1"})
    corte = {"H1": _dt.date(2026, 8, 11)}
    rs = MA.analizar_factura(serv, [], df_alb, {}, 0, cortes=corte, con_albaran={"H1"})
    rm = MA.analizar_factura(merc, [], df_alb, {}, 1, cortes=corte, con_albaran={"H1"})
    ok(rs["estado_matching"] == "NO_REQUIERE_ALBARAN", f"seguro anterior al primer albaran → {rs['estado_matching']}")
    ok(rm["estado_matching"] == "ANTERIOR_AL_REGISTRO", f"mercancia anterior al primer albaran → {rm['estado_matching']} (sin cambios)")
    ok('INCIDENCIAS_DETALLE' in open(os.path.join(BASE, 'matching_ap_albaran.py'), encoding='utf-8').read() and '_NOMBRE_INC' in src,
       "el cruce dice QUE facturas tienen incidencia y el log las nombra")
    # el formateador unico del panel, ejecutado de verdad con node sobre el HTML servido
    cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
    html = cl.get('/').get_data(as_text=True)
    m1 = re.search(r"function _fmtEurES\(v, dec\) \{.*?\n\}", html, re.S); m2 = re.search(r"function _numES\(v\) \{.*?\n\}", html, re.S)
    ok(m1 and m2, "el panel sirve _fmtEurES y _numES")
    if m1 and m2:
        js = m1.group(0) + "\n" + m2.group(0) + "\nconsole.log(JSON.stringify([_fmtEurES(1175.8), _fmtEurES(12400), _fmtEurES(248.0), _fmtEurES(23410, 0), _numES('12.400,00'), _numES('1,175.80'), _numES('248.0'), _numES('')]));"
        open('/tmp/_fmt.js', 'w').write(js)
        out = subprocess.run(['node', '/tmp/_fmt.js'], capture_output=True, text=True).stdout.strip()
        ok(out == '["1.175,80 €","12.400,00 €","248,00 €","23.410 €",12400,1175.8,248,null]', f"node: {out}")
    ok(html.count("'€' + ") <= 3, f"sin '€' delante en el panel (quedan {html.count(chr(39)+'€'+chr(39)+' + ')}: solo los compactos '€12K' de graficas y multi-hotel)")
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
