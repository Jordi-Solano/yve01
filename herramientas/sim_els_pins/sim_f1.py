import os, sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sim_common import *
D, cl = empezar()
try:
    F1 = FIX+'/fase1_AP/'
    print('== fase 1')
    lote(cl, [F1+f for f in ['factura_distribucions_garraf_DG-2026-0812.pdf','albaran_garraf_5531.pdf','factura_energia_llevant_EL-88213.pdf','factura_distribucions_garraf_DG-2026-0819.pdf','albaran_garraf_5540.pdf']])
    foto(cl, F1+'foto_factura_bugaderia_sitges.jpg')
    ap = gj(cl, '/api/facturas_ap'); pj(ap, 6000)
    print('-- stats_ap'); pj(gj(cl,'/api/stats_ap'), 1500)
    print('-- aging'); pj(gj(cl,'/api/aging_ap'), 1500)
    print('-- albaranes'); pj(gj(cl,'/api/albaranes'), 2500)
    print('-- reclamaciones'); r=cl.get('/api/reclamaciones_ap/list'); print(r.status_code, r.get_data(as_text=True)[:1200])
    print('== resubir DG-0812')
    lote(cl, [F1+'factura_distribucions_garraf_DG-2026-0812.pdf'])
    ap = gj(cl, '/api/facturas_ap'); print('n facturas', len(ap), [a['numero_factura'] for a in ap])
    df = pd.read_excel('facturas-procesadas/facturas_ap.xlsx') if os.path.exists('facturas-procesadas/facturas_ap.xlsx') else None
    print(df.columns.tolist() if df is not None else 'no facturas_ap.xlsx'); 
    if df is not None: print(df[['numero_factura','NIF_proveedor','total_factura','cuenta_contable','tipo_proveedor']].to_string() if 'NIF_proveedor' in df.columns else df.iloc[:, :8].to_string())
    print(open('datos-referencia/proveedores_aprendidos.json').read()[:600] if os.path.exists('datos-referencia/proveedores_aprendidos.json') else 'sin proveedores_aprendidos.json')
finally:
    terminar()
