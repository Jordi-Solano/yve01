import os, sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sim_common import *
FASES = sys.argv[1:] or ['2', '3', '4', '5', '6', '7']
D, cl = empezar()


def cliente(user, pwd):
    c = D.app.test_client()
    assert c.post('/api/login', json={'username': user, 'password': pwd}).status_code == 200
    tok = (c.get('/api/csrf_token').get_json() or {}).get('token')
    r = gj(cl, '/api/hotel_activo')
    c.post('/api/hotel_activo', json={'hotel': r['hotel']}, headers={'X-CSRF-Token': tok})
    return c, tok


def accion(c, tok, clave, tipo='APROBADA'):
    return c.post('/aprobaciones-ap/api/accion', json={'clave': clave, 'numero_factura': clave, 'accion': tipo, 'comentario': 'ok', 'departamento': 'Administracion'}, headers={'X-CSRF-Token': tok}).get_json()


def sse(cl, url):
    r = cl.get(url); txt = r.get_data(as_text=True); r.close()
    ls = [l[6:] for l in txt.splitlines() if l.startswith('data: ')]
    for l in ls:
        print('   o|', l[:200])
    return ls


def ap_resumen():
    ap = gj(cl, '/api/facturas_ap')
    for a in ap:
        print(f"   AP {a['numero_factura']:14} {a['proveedor']:26} {a['total']:>9} {a['cuenta_contable']:>4} {a['tipo']:5} {a['estado']:22} dup={a['duplicados']} acc={a['accion']} imp={a['importes_cuadran']} {a['aviso_importes'][:50]}")


try:
    if '2' in FASES:
        print('\n======== FASE 2')
        F = FIX + '/fase2_aprobaciones/'
        lote(cl, [F + 'factura_neteges_costa_NC-2026-118.pdf', F + 'factura_installacions_vila_IV-0733.pdf', F + 'factura_neteges_costa_NC-2026-118_copia.pdf'])
        ap_resumen()
        print('-- duplicados'); pj(gj(cl, '/aprobaciones-ap/api/duplicados'), 2500)
        print('-- aging total'); ag = gj(cl, '/api/aging_ap'); pj(ag.get('resumen') or ag.get('totales') or {k: v for k, v in ag.items() if k != 'filas'}, 800)
        print('-- aprobar lote'); pj(post(cl, '/api/ap/aprobar_lote', {}).get_json(), 600)
        d = gj(cl, '/aprobaciones-ap/api/duplicados')
        dup = (d.get('duplicados') or d.get('grupos') or [])
        if dup:
            g = dup[0]; docs = g.get('documentos') or g.get('docs') or []
            bueno = [x for x in docs if abs(float(x.get('total') or x.get('total_factura') or 0) - 290.40) < 0.01]
            print('-- elegir buena', bueno[0] if bueno else docs)
            if bueno:
                pj(post(cl, '/aprobaciones-ap/api/duplicados/elegir', {'clave': g.get('clave'), 'archivo': bueno[0].get('archivo')}).get_json(), 600)
        ap_resumen()
        print('-- aging'); ag = gj(cl, '/api/aging_ap'); pj({k: v for k, v in ag.items() if k != 'filas'}, 800)
        admin, ta = cliente('admin', 'admin123'); fc, tf = cliente('fc_user', 'hotel2024')
        print('-- admin aprueba NC:', accion(admin, ta, 'NC-2026-118'))
        print('-- admin aprueba IV:', accion(admin, ta, 'IV-0733'))
        print('-- admin otra vez IV:', accion(admin, ta, 'IV-0733'))
        print('-- fc_user aprueba IV:', accion(fc, tf, 'IV-0733'))
        print('-- fc_user rechaza NC (ya aprobada):', accion(fc, tf, 'NC-2026-118', 'RECHAZADA'))
        print('-- historial'); pj([{k: h.get(k) for k in ('numero_factura', 'clave', 'accion', 'aprobador', 'fecha', 'vigente')} for h in (gj(cl, '/aprobaciones-ap/api/historial') or [])][:8], 1500)
        print('-- stats'); pj(gj(cl, '/aprobaciones-ap/api/stats'), 800)
        ap_resumen()
        print('-- panel facturas'); pj([{k: f.get(k) for k in ('numero_factura', 'total', 'estado', 'accion', 'firmas', 'doble_firma', 'duplicados', 'bloqueada', 'aviso')} for f in (gj(cl, '/aprobaciones-ap/api/facturas') or [])], 1500)

    if '3' in FASES:
        print('\n======== FASE 3')
        F = FIX + '/fase3_AR_bonos/'
        lote(cl, [F + 'contrato_booking_tarifas_pactadas.xlsx'])
        lote(cl, [F + 'booking_comision_agosto_2026.pdf'])
        print('-- facturas AR'); pj(gj(cl, '/api/facturas'), 3000)
        print('-- stats'); pj(gj(cl, '/api/stats'), 1200)
        print('-- reclamaciones ota'); pj(gj(cl, '/api/reclamaciones_ota/list'), 1500)
        print('-- crear cliente'); pj(post(cl, '/api/ar_real/cliente', {'nombre': 'Viatges Mediterrani SL', 'nif': 'B-67890123', 'email': 'a@b.c', 'credito_limite': 5000}).get_json(), 400)
        print('-- emitir 1'); pj(post(cl, '/api/ar_real/emitir_factura', {'cliente': 'Viatges Mediterrani SL', 'fecha_entrada': '2026-08-10', 'fecha_salida': '2026-08-14', 'habitaciones': 2, 'precio_noche': 150, 'fb_extras': 0, 'total': 1320.0}).get_json(), 400)
        print('-- emitir 2'); pj(post(cl, '/api/ar_real/emitir_factura', {'cliente': 'Viatges Mediterrani SL', 'fecha_entrada': '2026-08-21', 'fecha_salida': '2026-08-24', 'habitaciones': 1, 'precio_noche': 140, 'fb_extras': 0, 'total': 462.0}).get_json(), 400)
        lote(cl, [F + 'bono_viatges_mediterrani_VM-7781.pdf', F + 'bono_viatges_mediterrani_VM-7790.pdf'])
        print('-- bonos'); pj(gj(cl, '/api/ar_real/bonos'), 2500)
        print('-- ar_real facturas'); pj(gj(cl, '/api/ar_real/facturas'), 1500)
        print('-- cobrar'); pj(post(cl, '/api/ar_real/cobrar', {'numero': 'FAC-2026-CORP-0001'}).get_json(), 300)
        print('-- bonos tras cobrar'); b = gj(cl, '/api/ar_real/bonos'); pj([{k: x.get(k) for k in ('numero_bono', 'estado', 'detalle', 'factura', 'numero_factura')} for x in (b.get('bonos') if isinstance(b, dict) else b)], 800)
        print('-- clientes'); pj(gj(cl, '/api/ar_real/clientes'), 800)

    if '4' in FASES:
        print('\n======== FASE 4')
        F = FIX + '/fase4_oracle/'
        lote(cl, [F + 'factura_seguretat_garraf_SG-2026-0455.pdf', F + 'factura_assegurances_mar_AM-2026-9107.pdf'])
        ap_resumen()
        print('-- oracle status'); pj(gj(cl, '/api/oracle/status'), 500)
        print('-- oracle 1 (sin aprobar)'); sse(cl, '/api/procesar_oracle')
        r = cl.get('/api/oracle/export_excel'); print('   GL:', r.status_code, r.headers.get('Content-Type'), len(r.data))
        if r.status_code == 200:
            x = pd.read_excel(io.BytesIO(r.data), sheet_name=None); [print('   hoja', k, len(v), list(v.columns)[:8]) for k, v in x.items()]
        admin, ta = cliente('admin', 'admin123'); fc, tf = cliente('fc_user', 'hotel2024')
        print('-- aprobar SG:', accion(admin, ta, 'SG-2026-0455'))
        print('-- aprobar AM admin:', accion(admin, ta, 'AM-2026-9107'))
        print('-- aprobar AM fc:', accion(fc, tf, 'AM-2026-9107'))
        print('-- oracle 2'); sse(cl, '/api/procesar_oracle')
        r = cl.get('/api/oracle/export_excel'); print('   GL:', r.status_code, len(r.data))
        if r.status_code == 200:
            x = pd.read_excel(io.BytesIO(r.data), sheet_name=None)
            for k, v in x.items():
                print('   hoja', k, len(v)); print(v.to_string()[:2500])
        print('-- oracle 3'); sse(cl, '/api/procesar_oracle')
        r = cl.get('/api/oracle/export_excel'); x = pd.read_excel(io.BytesIO(r.data), sheet_name=None); print('   GL tras 3a:', {k: len(v) for k, v in x.items()})
        admin, ta = cliente('admin', 'admin123'); print('-- aprobar SG otra vez:', accion(admin, ta, 'SG-2026-0455'))

    if '5' in FASES:
        print('\n======== FASE 5')
        F = FIX + '/fase5_FB_inventarios/'
        lote(cl, [F + 'escandallo_recetas_els_pins.xlsx'])
        print('-- recetas'); rc = gj(cl, '/fb/api/recetas'); pj(rc if not isinstance(rc, dict) else {k: (v if k != 'recetas' else [{kk: x.get(kk) for kk in ('nombre', 'pvp', 'coste', 'coste_total', 'food_cost_pct', 'fc_pct', 'categoria')} for x in v]) for k, v in rc.items()}, 2500)
        lote(cl, [F + 'ventas_pos_agosto_2026.xlsx', F + 'inventario_agosto_2026.xlsx', F + 'mermas_agosto_2026.xlsx'])
        print('-- resultados'); res = gj(cl, '/fb/api/resultados?mes=2026-08'); pj({k: v for k, v in res.items() if k not in ('platos', 'ventas_por_dia', 'top_platos', 'por_categoria', 'ingredientes', 'series')}, 3000)
        print('   claves:', list(res.keys()))
        for k in ('por_categoria', 'top_platos', 'sin_escandallo', 'platos_sin_receta'):
            if k in res: print('  ', k, str(res[k])[:400])
        print('-- resultados sin mes'); res2 = gj(cl, '/fb/api/resultados'); print('   ', {k: res2.get(k) for k in ('resumen',)} if 'resumen' in res2 else str(res2)[:300])
        print('-- mermas'); pj(gj(cl, '/fb/api/mermas?mes=2026-08'), 1500); print('-- mermas sin mes'); pj(gj(cl, '/fb/api/mermas'), 600)
        print('-- inventarios cierre'); inv = gj(cl, '/api/inventarios?mes=2026-08'); pj({k: v for k, v in inv.items() if k not in ('articulos', 'items', 'filas')}, 2500)
        r = cl.get('/api/inventarios/hoja_recuento?mes=2026-08'); print('   hoja recuento:', r.status_code, len(r.data))
        if r.status_code == 200:
            h = pd.read_excel(io.BytesIO(r.data)); print('   ', len(h), list(h.columns), h['recuento'].isna().sum() if 'recuento' in h.columns else '')
        lote(cl, [F + 'recuento_inventario_2026-08.xlsx'])
        inv = gj(cl, '/api/inventarios?mes=2026-08'); pj({k: v for k, v in inv.items() if k not in ('articulos', 'items', 'filas')}, 2500)
        print('-- fb inventario'); fi = gj(cl, '/fb/api/inventario'); pj([{k: x.get(k) for k in ('ingrediente', 'stock_actual_kg_l', 'stock_inicial_kg_l')} for x in (fi.get('items') or fi.get('inventario') or fi if isinstance(fi, list) else [])][:12], 1200)

    if '6' in FASES:
        print('\n======== FASE 6')
        F = FIX + '/fase6_banco_cierre/'
        lote(cl, [FIX + '/fase1_AP/factura_distribucions_garraf_DG-2026-0812.pdf', FIX + '/fase1_AP/factura_energia_llevant_EL-88213.pdf', F + 'extracto_banco_agosto_2026.xlsx', F + 'DRR_els_pins_agosto_2026.xlsm'])
        print('-- drr'); pj(gj(cl, '/api/stats_drr'), 2500)
        print('-- banco antes'); sb = gj(cl, '/api/stats_banco'); pj({k: v for k, v in sb.items() if k not in ('movimientos', 'filas')}, 1200); print('   movs:', len(sb.get('movimientos') or []))
        print('-- config banco'); pj(post(cl, '/api/config_banco', {'modo': 'grupo'}).get_json(), 300)
        print('-- conciliar'); r = post(cl, '/api/conciliar', {}); pj(r.get_json() if r.is_json else r.get_data(as_text=True)[:500], 1500)
        sb = gj(cl, '/api/stats_banco'); pj({k: v for k, v in sb.items() if k not in ('movimientos', 'filas')}, 1200)
        for m in (sb.get('movimientos') or []):
            print('   mov', {k: m.get(k) for k in ('fecha', 'concepto', 'importe', 'estado', 'factura', 'conciliado_con', 'numero_factura')})
        print('-- cuadre'); cu = gj(cl, '/api/cuadre_banco?mes=2026-08'); pj(cu, 3500)
        print('-- asientos'); asi = gj(cl, '/api/cierre/asientos?mes=2026-08'); pj({k: v for k, v in asi.items() if k not in ('asientos', 'lineas')}, 2500)
        for a in (asi.get('asientos') or [])[:30]:
            print('   as', str(a)[:200])
        print('-- paquete 2026-08'); pj(gj(cl, '/api/cierre/paquete?mes=2026-08'), 3000)
        print('-- paquete 2026-09'); pj(gj(cl, '/api/cierre/paquete?mes=2026-09'), 800)
        print('-- comentario'); pj(post(cl, '/api/cierre/comentario', {'mes': '2026-08', 'bloque': 'banco', 'texto': 'pago duplicado a Garraf, pedir devolucion'}).get_json(), 300)
        r = cl.get('/api/exportar/cierre_paquete?mes=2026-08'); print('   paquete xlsx:', r.status_code, len(r.data))
        if r.status_code == 200:
            x = pd.read_excel(io.BytesIO(r.data), sheet_name=None); print('   hojas', {k: len(v) for k, v in x.items()})
        r = cl.get('/api/exportar/cierre?mes=2026-08'); print('   excel cierre:', r.status_code, len(r.data))

    if '7' in FASES:
        print('\n======== FASE 7')
        F = FIX + '/fase7_fiscal/'
        lote(cl, [F + 'factura_vins_penedes_VP-2291.pdf', F + 'factura_fruites_llobregat_FL-1180.pdf', F + 'booking_comision_agosto_2026_fiscal.pdf', F + 'expedia_commission_august_2026.pdf', F + 'ventas_pos_fiscal_agosto_2026.xlsx'])
        print('-- facturas AR'); pj([{k: f.get(k) for k in ('numero_factura', 'nombre_ota', 'importe_bruto', 'porcentaje_comision', 'importe_comision', 'estado', 'estado_di', 'motivo_di', 'tipo_mercado', 'mercado')} for f in (gj(cl, '/api/facturas') or [])], 1500)
        print('-- fiscal'); fi = gj(cl, '/api/fiscal?mes=2026-08'); pj(fi, 6000)
        r = cl.get('/api/exportar/fiscal?mes=2026-08'); print('   excel fiscal:', r.status_code, len(r.data))
        if r.status_code == 200:
            x = pd.read_excel(io.BytesIO(r.data), sheet_name=None); print('   hojas', {k: len(v) for k, v in x.items()})
            if 'SII recibidas' in x: print(x['SII recibidas'].to_string()[:1500])
        for lib in ('emitidas', 'recibidas'):
            r = cl.get(f'/api/exportar/sii?libro={lib}&mes=2026-08'); print(f'   sii {lib}:', r.status_code, len(r.data), r.headers.get('Content-Disposition'))
        print('-- paquete fiscal'); p = gj(cl, '/api/cierre/paquete?mes=2026-08'); pj([b for b in (p.get('bloques') or p.get('checklist') or []) if 'iscal' in str(b)], 800)
finally:
    terminar()
