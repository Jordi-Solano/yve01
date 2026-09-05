# -*- coding: utf-8 -*-
"""Lote 3 de Jordi (fases 5-7): el cierre respeta el mes, el paquete no dice
"listo" sin datos, euros en español en el paquete y en F&B, y el cruce con
albaranes dice que mercancia NO reclama y por que.

  python3.12 tests/test_lote3_cierre.py
  python3.12 tests/test_lote3_cierre.py --sabotaje
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
PROC = os.path.join(BASE, 'facturas-procesadas')
RAW = os.path.join(PROC, 'facturas_ap_20260102.xlsx')
DATOS = os.path.join(BASE, 'datos-referencia')
INV = os.path.join(DATOS, 'inventario.xlsx')


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import inventarios as I
    import paquete_cierre as PQ
    if SABOTAJE:
        I.filtrar_mes = lambda df, mes: (df, [])                   # el mes deja de contar
        PQ._eur = lambda v, dec=2: f"{float(v):,.2f} EUR"           # formato americano otra vez

    # 1. mes a partir de nombres de fichero / celdas
    casos = {'inventario_agosto_2026.xlsx': '2026-08', 'recuento_inventario_2026-08.xlsx': '2026-08', '08/2026': '2026-08',
             '31/08/2026': '2026-08', 'Setembre 2026': '2026-09', 'stock.xlsx': '', None: '', float('nan'): ''}
    mal = {k: I.mes_de_texto(k) for k, v in casos.items() if I.mes_de_texto(k) != v}
    ok(not mal, f"mes_de_texto: {mal or 'todos bien'}")
    ok(I.mes_de_texto(pd.Timestamp('2026-08-31')) == '2026-08', "mes_de_texto con Timestamp")

    # 2. valorar: un inventario de agosto NO vale para septiembre
    df = pd.DataFrame([
        {"ingrediente": "Pollo", "categoria": "Carnes", "coste_unitario": 4.5, "stock_inicial_kg_l": 60, "stock_actual_kg_l": 42, "unidad": "kg", "mes": "2026-08"},
        {"ingrediente": "Merluza", "categoria": "Pescados", "coste_unitario": 9.0, "stock_inicial_kg_l": 40, "stock_actual_kg_l": 31, "unidad": "kg", "mes": "2026-08"},
    ])
    ago = I.valorar("2026-08", df); sep = I.valorar("2026-09", df)
    ok(ago["resumen"]["n_articulos"] == 2 and ago["resumen"]["valor_final"] == 468.0, f"agosto: 2 articulos, existencias {ago['resumen']['valor_final']}")
    ok(sep["resumen"]["n_articulos"] == 0 and not sep["asientos"], f"septiembre sin recuento: {sep['resumen']['n_articulos']} articulos, {len(sep['asientos'])} asientos")
    ok(sep["resumen"].get("otros_meses") == ["2026-08"] and "2026-08" in sep["resumen"]["nota"], f"y dice que el guardado es de agosto: {sep['resumen'].get('otros_meses')}")
    sin_mes = I.valorar("2026-09", df.drop(columns=["mes"]))
    ok(sin_mes["resumen"]["n_articulos"] == 2, "un inventario SIN mes estampado sigue valiendo para cualquier mes (datos viejos)")

    # 3. paquete: sin asientos no esta "listo", la reconciliacion es SIN_DATO y los euros en español
    p_vacio = PQ.montar("2026-09", {"n_asientos": 0, "debe": 0, "cuadra": True}, {"ok": True, "resumen": {"CUADRA": 7, "SIN_DATO": 1}, "checks": []}, None, None, sep, None, None, None)
    est = {c["clave"]: c for c in p_vacio["checklist"]}
    ok(not p_vacio["listo"] and p_vacio.get("sin_datos"), f"mes vacio: listo={p_vacio['listo']} sin_datos={p_vacio.get('sin_datos')}")
    ok(est["reconciliacion"]["estado"] == "SIN_DATO", f"reconciliacion sin asientos → {est['reconciliacion']['estado']}")
    ok(est["inventarios"]["estado"] == "SIN_DATO" and "2026-08" in est["inventarios"]["cifra"], f"inventarios → {est['inventarios']['estado']} · {est['inventarios']['cifra']}")
    p_ok = PQ.montar("2026-08", {"n_asientos": 39, "debe": 62261.3, "cuadra": True}, {"ok": True, "resumen": {"CUADRA": 7}, "checks": []}, None, None, ago, None, None, None)
    ok(p_ok["listo"], "con asientos y sin pendientes: listo")
    cifras = " | ".join(c["cifra"] for c in p_ok["checklist"])
    ok("62.261,30 €" in cifras and "468,00 €" in cifras and "EUR" not in cifras and not re.search(r"\d,\d{3}\.\d\d", cifras), f"cifras del paquete en español: {cifras[:120]}")

    # 4. subir un recuento de septiembre sobre el inventario de agosto: el final pasa a inicial y el mes cambia
    tmp = tempfile.mkdtemp(prefix='l3_'); copias = {}
    for f in (INV, RAW):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        import dashboard as D
        D._guardar_fb_del_hotel(df.drop(columns=["mes"]), 'inventario.xlsx', mes=D._mes_de_nombre('inventario_agosto_2026.xlsx'))
        g = pd.read_excel(INV)
        ok('mes' in g.columns and set(g['mes'].astype(str)) == {'2026-08'}, f"subida por Procesar Archivos estampa el mes del nombre: {sorted(set(g['mes'].astype(str))) if 'mes' in g.columns else 'sin columna'}")
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        r9 = cl.get('/api/inventarios?mes=2026-09').get_json()
        ok(r9["resumen"]["n_articulos"] == 0, f"/api/inventarios?mes=2026-09 no pinta agosto ({r9['resumen']['n_articulos']} articulos)")
        pq = cl.get('/api/cierre/paquete?mes=2026-10').get_json()
        ok(not pq["listo"] and pq.get("sin_datos"), f"/api/cierre/paquete?mes=2026-10 no esta 'listo': listo={pq['listo']}")
        from io import BytesIO
        buf = BytesIO()
        pd.DataFrame([{"ingrediente": "Pollo", "recuento": 30}, {"ingrediente": "Merluza", "recuento": 20}]).to_excel(buf, index=False); buf.seek(0)
        tok = (cl.get('/api/csrf_token').get_json() or {}).get('token')
        r = cl.post('/api/inventarios/recuento?mes=2026-09', data={'archivo': (buf, 'recuento_2026-09.xlsx')}, content_type='multipart/form-data', headers={'X-CSRF-Token': tok})
        ok(r.status_code == 200, f"subir recuento de septiembre → {r.status_code} {str(r.get_json())[:80]}")
        g2 = pd.read_excel(INV)
        pollo = g2[g2['ingrediente'] == 'Pollo'].iloc[0]
        ok(str(pollo.get('mes')) == '2026-09' and float(pollo['stock_inicial_kg_l']) == 42 and float(pollo['stock_actual_kg_l']) == 30,
           f"Pollo: mes={pollo.get('mes')} inicial={pollo['stock_inicial_kg_l']} (era el final de agosto) final={pollo['stock_actual_kg_l']}")
        r9 = cl.get('/api/inventarios?mes=2026-09').get_json()
        ok(r9["resumen"]["n_articulos"] == 2 and r9["resumen"]["valor_final"] == 315.0, f"ahora septiembre si tiene recuento: existencias {r9['resumen']['valor_final']}")
        # F&B: ningun '€' delante en recetas / inventario / mermas
        html = cl.get('/').get_data(as_text=True)
        restos = [l.strip()[:80] for l in html.splitlines() if ("'€' +" in l or "€' + " in l) and "+ 'K'" not in l]   # '€12K' de graficas: compacto a proposito
        ok(not restos, f"sin '€' delante en F&B ni en el detalle ({len(restos)}): {restos[:2]}")
        ok("_fmtEurES(item.coste_unitario)" in html and "_fmtEurES(m.coste)" in html and "_fmtEurES(r.precio_venta)" in html, "recetas, inventario y mermas usan el formateador unico")
        ok("paq.sinDatos" in html, "el paquete distingue 'sin datos de este mes' de 'listo'")
        for bloque in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html, re.S):
            open('/tmp/_l3.js', 'w', encoding='utf-8').write(bloque)
            rc = subprocess.run(['node', '--check', '/tmp/_l3.js'], capture_output=True, text=True)
            if rc.returncode != 0:
                ok(False, f"JS servido con error: {rc.stderr[:120]}"); break
        else:
            ok(True, "todos los <script> del panel pasan node --check")

        # 5. cruce con albaranes sin ningun albaran: lo dice, no lo calla
        os.makedirs(PROC, exist_ok=True)
        pd.DataFrame([
            {"archivo": "fl.pdf", "numero_factura": "FL-TEST-1180", "fecha": "20/08/2026", "nombre_proveedor": "Fruites Test SL", "base_imponible": 200.0,
             "porcentaje_iva": 4, "cuota_iva": 8.0, "total_factura": 208.0, "tipo_proveedor": "FB", "cuenta_contable": "600", "descripcion_concepto": "Fruta", "hotel_id": ""},
            {"archivo": "el.pdf", "numero_factura": "EL-TEST-1", "fecha": "20/08/2026", "nombre_proveedor": "Energia Test SA", "base_imponible": 100.0,
             "porcentaje_iva": 21, "cuota_iva": 21.0, "total_factura": 121.0, "tipo_proveedor": "OTRAS", "cuenta_contable": "629", "descripcion_concepto": "Luz", "hotel_id": ""},
        ]).to_excel(RAW, index=False)
        import almacen_datos as ALM
        alb = ALM.albaranes(PROC, os.path.join(BASE, 'reportes'))
        if alb is None or alb.empty:
            out = subprocess.run([sys.executable, 'matching_ap_albaran.py'], cwd=BASE, capture_output=True, text=True, timeout=120).stdout
            ok('INCIDENCIAS: 0|0' in out and re.search(r"SIN_REGISTRO:.*FL-TEST-1180", out) and 'EL-TEST-1' not in (re.search(r"SIN_REGISTRO:(.*)", out) or [None, ''])[1],
               "sin albaranes: INCIDENCIAS 0 y SIN_REGISTRO con la mercancia (no con la luz)")
        else:
            ok(True, "(hay albaranes en el checkout: el caso 'sin ninguno' no se puede reproducir aqui)")
        src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
        ok("SIN_REGISTRO:" in src and "NO se reclaman" in src, "el log del cierre explica la mercancia que no se reclama")
    finally:
        for f in (INV, RAW):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
        for f in ('matching_albaran_report.xlsx',):
            pass
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
