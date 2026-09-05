# -*- coding: utf-8 -*-
"""Pieza 9 (Jordi, sep 2026): UNA sola entrada de documentos.
Todo entra por Procesar archivos: la hoja de recuento y el escandallo
(recetas) tambien, con el mes pedido en la lista de subida cuando el nombre
no lo lleva. Fuera los botones de subida de DRR, F&B, Cierre y /conciliacion.

  python3.12 tests/test_entrada_unica.py
  python3.12 tests/test_entrada_unica.py --sabotaje
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DATOS = os.path.join(BASE, 'datos-referencia')
INV = os.path.join(DATOS, 'inventario.xlsx')
REC = os.path.join(DATOS, 'recetas.xlsx')
REG = os.path.join(DATOS, 'archivos_procesados.json')
NOMBRES = ['recuento_inventario_2026-08.xlsx', 'hoja_contada.xlsx', 'escandallo_restaurante.xlsx']


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import dashboard as D
    if SABOTAJE:
        # la capa 1 vuelve a no conocer el recuento ni las recetas
        D._CAB_KEYWORDS = [k for k in D._CAB_KEYWORDS if k[0] not in ('RECUENTO', 'RECETAS')]
    tmp = tempfile.mkdtemp(prefix='eu_'); copias = {}
    for f in (INV, REC, REG):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        for f in (INV, REC):
            if os.path.exists(f):
                os.remove(f)
        # inventario previo del hotel activo, para que el recuento tenga que pisar el stock
        app = D.app; app.config['TESTING'] = True
        cl = app.test_client()
        assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
        D._guardar_fb_del_hotel(pd.DataFrame([
            {'ingrediente': 'Tomate', 'categoria': 'Verdura', 'stock_actual_kg_l': 10.0, 'coste_unitario': 2.0, 'unidad': 'kg'},
            {'ingrediente': 'Aceite', 'categoria': 'Despensa', 'stock_actual_kg_l': 5.0, 'coste_unitario': 6.0, 'unidad': 'l'},
        ]), 'inventario.xlsx', mes='2026-07')

        # 1. capa 1: el nombre propone y las cabeceras confirman
        hoja = pd.DataFrame([
            {'ingrediente': 'Tomate', 'categoria': 'Verdura', 'familia': 'F&B', 'unidad': 'kg', 'stock_sistema': 10.0, 'recuento': 7.5, 'coste_unitario': 2.0, 'observaciones': ''},
            {'ingrediente': 'Aceite', 'categoria': 'Despensa', 'familia': 'F&B', 'unidad': 'l', 'stock_sistema': 5.0, 'recuento': None, 'coste_unitario': 6.0, 'observaciones': 'no contado'},
            {'ingrediente': 'Limones', 'categoria': 'Fruta', 'familia': 'F&B', 'unidad': 'kg', 'stock_sistema': None, 'recuento': 3.0, 'coste_unitario': 1.5, 'observaciones': ''},
        ])
        esc = pd.DataFrame([
            {'receta': 'Gazpacho', 'ingrediente': 'Tomate', 'cantidad': 0.4, 'unidad': 'kg', 'coste': 2.0, 'pvp': 8.5, 'categoria': 'Entrantes'},
            {'receta': 'Gazpacho', 'ingrediente': 'Aceite', 'cantidad': 0.05, 'unidad': 'l', 'coste': 6.0, 'pvp': 8.5, 'categoria': 'Entrantes'},
            {'receta': 'Limonada', 'ingrediente': 'Limones', 'cantidad': 0.2, 'unidad': 'kg', 'coste': 1.5, 'pvp': 4.0, 'categoria': 'Bebidas'},
        ])
        rutas = {}
        for nombre, df in ((NOMBRES[0], hoja), (NOMBRES[1], hoja), (NOMBRES[2], esc)):
            rutas[nombre] = os.path.join(tmp, nombre); df.to_excel(rutas[nombre], index=False)
        ok(D._destino_capa1(NOMBRES[0], rutas[NOMBRES[0]]) == 'RECUENTO', "recuento_inventario_2026-08.xlsx → RECUENTO (no INVENTARIO)")
        ok(D._destino_capa1(NOMBRES[2], rutas[NOMBRES[2]]) == 'RECETAS', "escandallo_restaurante.xlsx → RECETAS (no F&B)")
        ok(D._destino_capa1(NOMBRES[1], rutas[NOMBRES[1]]) == 'IA', "hoja_contada.xlsx (sin keyword) → IA, como siempre")

        def lote(nombres, meses=None):
            for n in nombres:
                with open(rutas[n], 'rb') as fh:
                    cl.post('/api/upload_facturas', data={'files': [(fh, n)]}, content_type='multipart/form-data')
            q = '/api/procesar_batch_stream?archivos=' + urllib.parse.quote(json.dumps(nombres))
            if meses:
                q += '&meses=' + urllib.parse.quote(json.dumps(meses))
            r = cl.get(q); txt = r.get_data(as_text=True); r.close()
            return txt

        # 2. el recuento con mes en el nombre entra y pisa el stock; lo no contado no se toca
        txt = lote([NOMBRES[0]])
        ok('✓ Recuento' in txt and '2 articulos contados' in txt and '1 nuevos' in txt, f"lote: {[l for l in txt.splitlines() if 'Recuento' in l][:1]}")
        inv = pd.read_excel(INV) if os.path.exists(INV) else pd.DataFrame()
        from inventarios import filtrar_mes
        ago, _ = filtrar_mes(inv, '2026-08')
        st = {str(r['ingrediente']): float(r['stock_actual_kg_l']) for _, r in ago.iterrows()}
        ok(st.get('Tomate') == 7.5 and st.get('Limones') == 3.0 and 'Aceite' not in st, f"agosto: Tomate 7.5, Limones 3.0 nuevo, Aceite sin contar → {st}")
        # el inventario guarda UNA fila por articulo (clave ingrediente+hotel): el
        # final de julio pasa a ser el inicial de agosto, como hace la ruta del cierre
        _t = ago[ago['ingrediente'] == 'Tomate'] if 'ingrediente' in ago.columns else ago.iloc[0:0]
        tom = _t.iloc[0].to_dict() if len(_t) else {}
        ok(tom and float(tom.get('stock_inicial_kg_l') or 0) == 10.0 and str(tom.get('mes')) == '2026-08', f"Tomate: inicial de agosto = final de julio (10) → {tom.get('stock_inicial_kg_l')}, mes {tom.get('mes')}")
        # 3. el escandallo entra por el mismo camino
        txt = lote([NOMBRES[2]])
        ok('Recetas escandallo_restaurante.xlsx: 2 recetas (3 ingredientes)' in txt, f"lote recetas: {[l for l in txt.splitlines() if 'Recetas' in l][:1]}")
        rec = pd.read_excel(REC) if os.path.exists(REC) else pd.DataFrame()
        ok(len(rec) == 2 and set(rec.get('nombre', [])) == {'Gazpacho', 'Limonada'}, f"recetas.xlsx: {list(rec.get('nombre', []))}")
        # 4. el mes del recuento llega por ?meses= cuando el nombre no lo lleva (lo pide la lista de subida);
        #    y una hoja sin keyword pasa por la IA — aqui se simula con un nombre que si propone recuento pero sin mes
        rutas['recuento_hotel.xlsx'] = rutas[NOMBRES[1]]
        txt = lote(['recuento_hotel.xlsx'])
        ok('✗ Recuento recuento_hotel.xlsx: falta el mes' in txt, "sin mes en el nombre ni en la lista: se pide, no se inventa")
        txt = lote(['recuento_hotel.xlsx'], {'recuento_hotel.xlsx': '2026-09'})
        inv = pd.read_excel(INV) if os.path.exists(INV) else pd.DataFrame(); sep_, _ = filtrar_mes(inv, '2026-09')
        ok('✓ Recuento' in txt and 'para 2026-09' in txt and len(sep_) >= 2, f"con ?meses= entra en 2026-09 ({len(sep_)} filas)")
        # 5. la pantalla: un solo sitio para subir
        html = cl.get('/').get_data(as_text=True)
        ok(all(x not in html for x in ('drr-file-input', 'inv-file"', 'fb-upload-input', 'fb-rec-input', 'uploadDRR(', 'fbUploadPOS(', 'fbUploadRecetas(', '_invSubir(')),
           "fuera los botones de subida de DRR, recuento, POS y recetario")
        ok('_recibirEnProcesar(event.dataTransfer.files)' in html and 'onclick="openUploadModal()"' in html, "la zona del DRR y los botones de F&B/Cierre abren Procesar archivos")
        ok("upload.mesRecuento" in html and "_pideMes(f) === 'Recuento' && !f._mes" in html and "'&meses=' + encodeURIComponent(JSON.stringify(_mesesSubida" in html,
           "la lista pide el mes del recuento y lo manda al lote")
        ok("'Recuento'" in html.split('function _detectType')[1][:1500] and "'Recetas'" in html.split('function _detectType')[1][:1500], "_detectType conoce Recuento y Recetas")
        conc = cl.get('/conciliacion/').get_data(as_text=True)
        ok('upload-input' not in conc and 'uploadFile(' not in conc and 'Procesar archivos' in conc, "/conciliacion ya no sube extractos: manda al dashboard")
        for b in re.findall(r"<script(?![^>]*src)[^>]*>(.*?)</script>", html + conc, re.S):
            open('/tmp/_eu.js', 'w', encoding='utf-8').write(b)
            rc = subprocess.run(['node', '--check', '/tmp/_eu.js'], capture_output=True, text=True)
            if rc.returncode:
                ok(False, f"JS roto: {rc.stderr[:100]}"); break
        for lang in ('en', 'ca', 'fr', 'de', 'it', 'pt'):
            d = json.load(open(os.path.join(BASE, 'static', 'i18n', f'{lang}.json'), encoding='utf-8'))
            if not all(k in d for k in ('upload.mesRecuento', 'upload.faltaMes', 'btn.importarFB', 'drr.arrastra')):
                ok(False, f"i18n {lang} sin las claves nuevas"); break
        else:
            ok(True, "i18n: 6 idiomas con las claves nuevas")
    finally:
        for f in (INV, REC, REG):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
        up = os.path.join(BASE, 'facturas-entrada')
        for n in NOMBRES + ['recuento_hotel.xlsx']:
            p = os.path.join(up, n)
            if os.path.exists(p):
                os.remove(p)
        import gc as _gc; _gc.collect()
        if REG in copias:
            shutil.copy(copias[REG], REG)
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
