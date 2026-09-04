# -*- coding: utf-8 -*-
"""Ronda de pruebas de Jordi · lote 1 (fases 0-2).

  1 foto solo en movil · 2 grupo al crear hotel · 3 servicios no exigen albaran
  4 una cuenta por proveedor, servicios nunca a 600 · 5 pantalla de albaranes
  6 euros con decimales en AP · 7 duplicados independientes del orden

  python3.12 tests/test_ronda1_ap.py
  python3.12 tests/test_ronda1_ap.py --sabotaje
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd                     # noqa: E402
import cuentas_proveedor as CP          # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    if SABOTAJE:
        CP.evidencia_fb = lambda *a, **k: True      # todo es comida: la lavanderia vuelve a 600

    tmp = tempfile.mkdtemp(prefix='r1_')
    M = ("OTRAS", CP.NF)     # proveedor desconocido para proveedores.xlsx
    # ── 4 · una cuenta por proveedor ─────────────────────────────────
    f1 = {"nombre_proveedor": "Distribucions Garraf SL", "descripcion_concepto": "productos alimentación", "_lineas": [{"descripcion": "Pollo entero fresco"}]}
    f2 = {"nombre_proveedor": "Distribucions Garraf, S.L.", "descripcion_concepto": "", "_lineas": [{"descripcion": "Merluza fresca"}, {"descripcion": "Limones"}]}
    f3 = {"nombre_proveedor": "Bugaderia Sitges SL", "descripcion_concepto": "Estovalles restaurant", "cuenta_contable": "600", "_lineas": [{"descripcion": "Roba de llit (kg)"}]}
    f4 = {"nombre_proveedor": "Energia Llevant SA", "descripcion_concepto": "Consumo eléctrico julio 2026", "_lineas": []}
    f5 = {"nombre_proveedor": "Instal·lacions Vila SL", "descripcion_concepto": "Reparación climatización", "_lineas": []}
    f6 = {"nombre_proveedor": "Empresa Rara SL", "descripcion_concepto": "Parte de trabajo", "_lineas": []}
    ok(CP.decidir(f1, None, {}, M)[:2] == ("FB", "600") and CP.decidir(f2, None, {}, M)[:2] == ("FB", "600"), "distribuidor de alimentacion: FB / 600 en las dos facturas")
    ok(CP.decidir(f3, None, {}, M)[:2] == ("OTRAS", "629"), f"lavanderia con 'restaurant' en el concepto: {CP.decidir(f3, None, {}, M)} (nunca 600)")
    ok(CP.decidir(f4, None, {}, M)[1] == "629" and CP.decidir(f5, None, {}, M)[1] == "622", "luz → 629, climatizacion → 622")
    ok(CP.decidir(f6, None, {}, M)[:2] == ("OTRAS", "629"), "sin pista → 629, no 600")
    ok(CP.clave_proveedor("Distribucions Garraf, S.L.") == CP.clave_proveedor("DISTRIBUCIONS GARRAF SL"), "misma clave de proveedor con o sin S.L.")
    # normalizar aprende y la segunda factura hereda aunque no traiga pistas
    CP._maestro = lambda nombre: M
    filas = [dict(f1), {"nombre_proveedor": "Distribucions Garraf SL", "descripcion_concepto": "Factura", "_lineas": []}, dict(f3)]
    CP.normalizar(filas, tmp)
    ok(filas[0]["cuenta_contable"] == "600" and filas[1]["cuenta_contable"] == "600" and filas[1]["tipo_proveedor"] == "FB", "la segunda factura del proveedor hereda la cuenta aunque no traiga lineas")
    ok(filas[2]["cuenta_contable"] == "629" and filas[2]["tipo_proveedor"] == "OTRAS", "la lavanderia se guarda como OTRAS / 629")
    apr = CP.aprendidos(tmp)
    ok("distribucions garraf" in apr and apr["distribucions garraf"]["cuenta"] == "600", f"aprendido en proveedores_aprendidos.json: {list(apr)}")
    apr["distribucions garraf"]["cuenta"] = "601"; CP.guardar_aprendidos(apr, tmp)
    f7 = [{"nombre_proveedor": "Distribucions Garraf SL", "descripcion_concepto": "", "_lineas": [{"descripcion": "Pollo"}]}]
    CP.normalizar(f7, tmp)
    ok(f7[0]["cuenta_contable"] == "601", "lo editado a mano en el json manda sobre las palabras clave")

    # ── 3 · exige albaran solo la mercancia ──────────────────────────
    ok(CP.exige_albaran({"tipo_proveedor": "FB", "cuenta_contable": "600"}) and CP.exige_albaran({"tipo_proveedor": "OTRAS", "cuenta_contable": 600.0})
       and not CP.exige_albaran({"tipo_proveedor": "OTRAS", "cuenta_contable": "629"}), "exige_albaran: FB/60x si, servicios no")
    import matching_ap_albaran as MA
    df_alb = pd.DataFrame(columns=["numero_albaran", "nombre_proveedor", "total_albaran", "fecha_entrega", "hotel_id"])
    serv = pd.Series({"numero_factura": "EL-1", "nombre_proveedor": "Energia Llevant SA", "fecha": "14/08/2026", "base_imponible": 980.0,
                      "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "hotel_id": ""})
    merc = pd.Series({"numero_factura": "DG-1", "nombre_proveedor": "Distribucions Garraf SL", "fecha": "12/08/2026", "base_imponible": 640.0,
                      "tipo_proveedor": "FB", "cuenta_contable": "600", "hotel_id": ""})
    rs = MA.analizar_factura(serv, [], df_alb, {}, 0, cortes={}, con_albaran={""})
    rm = MA.analizar_factura(merc, [], df_alb, {}, 1, cortes={}, con_albaran={""})
    ok(rs["estado_matching"] == "NO_REQUIERE_ALBARAN", f"servicio sin albaran → {rs['estado_matching']} (no es incidencia)")
    ok(rm["estado_matching"] == "FACTURA_SIN_ALBARAN", f"mercancia sin albaran → {rm['estado_matching']} (sigue siendo incidencia)")
    import app_aprobacion_ap as P
    ok("NO_REQUIERE_ALBARAN" not in P._ESTADOS_INCIDENCIA and "FACTURA_SIN_ALBARAN" in P._ESTADOS_INCIDENCIA, "reclamaciones: solo la mercancia sin albaran se reclama")

    # ── 7 · duplicados independientes del orden ──────────────────────
    import almacen_datos as A
    rows = [{"numero_factura": "NC-118", "nombre_proveedor": "Neteges", "archivo": "nc_copia.pdf", "fecha": "05/08/2026", "total_factura": 314.6, "hotel_id": "", "_etapa": 0},
            {"numero_factura": "NC-118", "nombre_proveedor": "Neteges", "archivo": "nc.pdf", "fecha": "05/08/2026", "total_factura": 290.4, "hotel_id": "", "_etapa": 0},
            {"numero_factura": "IV-733", "nombre_proveedor": "Vila", "archivo": "iv.pdf", "fecha": "20/08/2026", "total_factura": 1512.5, "hotel_id": "", "_etapa": 0}]
    r1 = A._consolidar(pd.DataFrame(rows), A._ID_AP); r2 = A._consolidar(pd.DataFrame(rows[::-1]), A._ID_AP)
    g1 = r1.set_index("numero_factura"); g2 = r2.set_index("numero_factura")
    ok(g1.loc["NC-118", "archivo"] == g2.loc["NC-118", "archivo"] and g1.loc["NC-118", "total_factura"] == g2.loc["NC-118", "total_factura"],
       f"mismo resultado suban en el orden que suban: {g1.loc['NC-118', 'archivo']} / {g1.loc['NC-118', 'total_factura']}")
    ok(int(g1.loc["NC-118", "duplicados"]) == 2 and g1.loc["NC-118", "duplicado_de"] and int(g1.loc["IV-733", "duplicados"]) == 0,
       f"la fila avisa: duplicados={g1.loc['NC-118', 'duplicados']}, de {g1.loc['NC-118', 'duplicado_de']}")
    # etapa 0 = la mas avanzada (contabilizadas); las crudas van con _etapa 1
    rows2 = [dict(r, _etapa=1) for r in rows[:2]] + [{"numero_factura": "NC-118", "nombre_proveedor": "Neteges", "archivo": "nc.pdf", "fecha": "05/08/2026", "total_factura": 290.4, "hotel_id": "", "_etapa": 0, "estado_matching": "X"}]
    r3 = A._consolidar(pd.DataFrame(rows2), A._ID_AP).set_index("numero_factura")
    ok(r3.loc["NC-118", "estado_matching"] == "X" and int(r3.loc["NC-118", "duplicados"]) == 2, "la etapa mas avanzada sigue ganando y conserva el aviso")

    # ── endpoints y HTML (1, 2, 5, 6) ────────────────────────────────
    import dashboard as D
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    html = cl.get('/').get_data(as_text=True)
    ok('class="show-mobile" onclick="event.stopPropagation();document.getElementById(\'upload-photo-input\').click()"' in html, "1 · el boton de fotos solo sale en movil")
    ok('accept=".pdf,.xlsm,.xlsx,.xls,.csv,.jpg,.jpeg,.png,.webp,.heic"' in html, "1 · en escritorio las fotos entran por 'Seleccionar archivos'")
    ok("maximumFractionDigits:2}).format(v) + ' €'" in html and 'maximumFractionDigits:0}).format(v)' not in html, "6 · fmtEurAP con dos decimales")
    ok('id="card-albaranes"' in html and 'function cargarAlbaranes' in html and 'cargarAlbaranes();' in html, "5 · tarjeta de albaranes en AP")
    a = cl.get('/api/albaranes').get_json() or {}
    ok(a.get('ok') and 'albaranes' in a and 'resumen' in a, f"5 · /api/albaranes responde ({a.get('resumen', {}).get('n')} albaranes)")
    x = cl.get('/api/exportar/albaranes')
    ok(x.status_code == 200 and x.data[:2] == b'PK', "5 · /api/exportar/albaranes da un xlsx")
    adm = cl.get('/admin/').get_data(as_text=True)
    ok('id="hn-grupo"' in adm and 'hn-cat' not in adm and '<th>Grupo</th>' in adm, "2 · el alta de hotel pide grupo y no estrellas")
    lst = cl.get('/api/facturas_ap').get_json()
    ok(isinstance(lst, list) and all('duplicados' in f for f in lst), "7 · /api/facturas_ap trae el aviso de duplicados")
    # alta real con grupo (fichero de hoteles del tenant, con copia)
    import tenant_dirs
    hj = os.path.join(str(tenant_dirs.datos_dir()), 'hoteles.json')
    bak = open(hj, encoding='utf-8').read() if os.path.exists(hj) else None
    try:
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token')
        r = cl.post('/admin/api/hoteles/crear', json={'nombre': 'Hotel Test Grupo', 'ciudad': 'Sitges', 'habitaciones': 10, 'grupo': 'Cadena Llevant'}, headers={'X-CSRF-Token': tok})
        d = r.get_json() or {}
        hs = json.load(open(hj, encoding='utf-8')) if os.path.exists(hj) else []
        mio = next((h for h in hs if h.get('id') == d.get('id')), {})
        ok(r.status_code == 200 and d.get('ok') and mio.get('grupo') == 'Cadena Llevant', f"2 · el hotel se guarda con su grupo: {mio.get('grupo')}")
    finally:
        if bak is None:
            if os.path.exists(hj):
                os.remove(hj)
        else:
            open(hj, 'w', encoding='utf-8').write(bak)
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
