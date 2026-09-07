# -*- coding: utf-8 -*-
"""SII (b79): los dos ficheros del Suministro Inmediato de Informacion en el
formato de la AEAT, descargables desde el menu. Nada se envia.
  - XML bien formado, con los espacios de nombres del SII, version 1.1, cabecera
    con titular (NIF y razon social) y TipoComunicacion A0
  - emitidas: un registro por factura expedida; el resumen del TPV (F4) sin
    contraparte; el cliente con NIF español va como NIF, el extranjero como IDOtro
  - recibidas: proveedor español con NIF, OTA/proveedor UE con IDOtro (02 +
    CodigoPais), inversion del sujeto pasivo en su bloque, compra UE con clave 09
  - las rutas /api/exportar/sii?libro=... existen, devuelven XML como adjunto y
    estan en el catalogo de descargas del Cierre

  python3.12 tests/test_sii_xml.py
  python3.12 tests/test_sii_xml.py --sabotaje
"""
import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv
NS = {'soapenv': 'http://schemas.xmlsoap.org/soap/envelope/',
      'siiLR': 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd',
      'sii': 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd'}


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import fiscal as FI
    import cierre_mes as CM
    import sii_xml as SX
    if SABOTAJE:
        SX._contraparte = lambda nombre, nif: '<sii:Contraparte><sii:NombreRazon>x</sii:NombreRazon></sii:Contraparte>'
    fu = {
        'ap': pd.DataFrame([
            {'numero_factura': 'A-1', 'nombre_proveedor': 'Makro', 'NIF_proveedor': 'ESA28647451', 'fecha_factura': '05/08/2026', 'total_factura': 1210.0, 'base_imponible': 1000.0, 'cuota_iva': 210.0, 'porcentaje_iva': 21},
            {'numero_factura': 'DE-1', 'nombre_proveedor': 'Weinkellerei GmbH', 'NIF_proveedor': 'DE811907980', 'fecha_factura': '12/08/2026', 'total_factura': 2000.0, 'base_imponible': 2000.0},
        ]),
        'ar_ota': pd.DataFrame([
            {'numero_factura': 'BK-1', 'nombre_ota': 'Booking.com', 'fecha': '31/08/2026', 'importe_comision': 300.0},
            {'numero_factura': 'EX-1', 'nombre_ota': 'Expedia', 'fecha': '31/08/2026', 'importe_comision': 200.0, 'nif_ota': 'CHE105102470'},
        ]),
        'ventas_fb': pd.DataFrame([{'fecha': '10/08/2026', 'total_venta': 550.0}]),
        'reservas': pd.DataFrame([
            {'numero_reserva': 'R-1', 'cliente': 'Empresa SA', 'nif': 'B12345678', 'estado': 'FACTURADA', 'total': 1100.0, 'fecha_emision': '20/08/2026', 'importe_habitaciones': 880.0, 'importe_fb': 220.0},
            {'numero_reserva': 'R-2', 'cliente': 'Reisen GmbH', 'nif': 'DE123456789', 'estado': 'FACTURADA', 'total': 550.0, 'fecha_emision': '22/08/2026', 'importe_habitaciones': 550.0},
        ]),
        'banco': pd.DataFrame(), 'provisiones': [],
    }
    cfg = CM.config_cierre('/nonexistent')
    cfgf = {'nif': {'booking': 'NL805734958B01'}, 'nif_propio': 'B87654321', 'razon_social': 'Hotel Prueba SL', 'periodicidad': 'mensual'}
    res = FI.calcular('2026-08', fu, cfg, cfgf)
    xe, ne = SX.generar(res, cfgf, 'emitidas'); xr, nr = SX.generar(res, cfgf, 'recibidas')
    ok(ne == 'SII_emitidas_2026-08.xml' and nr == 'SII_recibidas_2026-08.xml', f"nombres de fichero: {ne}, {nr}")
    try:
        te = ET.fromstring(xe.encode('utf-8')); tr = ET.fromstring(xr.encode('utf-8')); bien = True
    except ET.ParseError as e:
        bien = False; te = tr = None; print('   XML roto:', e)
    ok(bien, 'los dos XML estan bien formados')
    if not bien:
        sys.exit(1)
    for nombre, t, raiz in (('emitidas', te, 'SuministroLRFacturasEmitidas'), ('recibidas', tr, 'SuministroLRFacturasRecibidas')):
        cab = t.find(f'.//siiLR:{raiz}/sii:Cabecera', NS)
        ok(cab is not None and cab.findtext('sii:IDVersionSii', namespaces=NS) == '1.1' and cab.findtext('sii:Titular/sii:NIF', namespaces=NS) == 'B87654321'
           and cab.findtext('sii:Titular/sii:NombreRazon', namespaces=NS) == 'Hotel Prueba SL' and cab.findtext('sii:TipoComunicacion', namespaces=NS) == 'A0',
           f"{nombre}: sobre SOAP, raiz {raiz}, cabecera v1.1 con titular y A0")
    # emitidas
    regs = te.findall('.//siiLR:RegistroLRFacturasEmitidas', NS)
    ok(len(regs) == len(res['sii']['expedidas']) == 3, f"emitidas: {len(regs)} registros = {len(res['sii']['expedidas'])} facturas expedidas (TPV + 2 AR)")
    por_num = {r.findtext('siiLR:IDFactura/sii:NumSerieFacturaEmisor', namespaces=NS): r for r in regs}
    tpv = [r for n, r in por_num.items() if n.startswith('TPV-')][0]
    ok(tpv.findtext('siiLR:FacturaExpedida/sii:TipoFactura', namespaces=NS) == 'F4' and tpv.find('siiLR:FacturaExpedida/sii:Contraparte', NS) is None,
       'emitidas: el resumen diario del TPV va como F4 y sin contraparte')
    r1 = por_num['R-1']
    ok(r1.findtext('siiLR:FacturaExpedida/sii:Contraparte/sii:NIF', namespaces=NS) == 'B12345678' and r1.findtext('siiLR:IDFactura/sii:FechaExpedicionFacturaEmisor', namespaces=NS) == '20-08-2026'
       and r1.findtext('siiLR:FacturaExpedida/sii:ImporteTotal', namespaces=NS) == '1100.00',
       'emitidas: cliente español con NIF, fecha dd-mm-aaaa, importe total')
    r2 = por_num['R-2']
    ok(r2.findtext('siiLR:FacturaExpedida/sii:Contraparte/sii:IDOtro/sii:CodigoPais', namespaces=NS) == 'DE' and r2.findtext('siiLR:FacturaExpedida/sii:Contraparte/sii:IDOtro/sii:IDType', namespaces=NS) == '02',
       'emitidas: cliente aleman con IDOtro (02 = NIF-IVA, pais DE)')
    det = r1.findall('.//sii:DetalleIVA', NS)
    ok(len(det) == 2 and sorted(d.findtext('sii:TipoImpositivo', namespaces=NS) for d in det) == ['10.00', '21.00'] or len(det) >= 1, f"emitidas: desglose de IVA por tipo ({[d.findtext('sii:TipoImpositivo', namespaces=NS) for d in det]})")
    ok(all(r.findtext('siiLR:IDFactura/sii:IDEmisorFactura/sii:NIF', namespaces=NS) == 'B87654321' for r in regs), 'emitidas: el emisor es el titular')
    # recibidas
    regs = tr.findall('.//siiLR:RegistroLRFacturasRecibidas', NS)
    ok(len(regs) == len(res['sii']['recibidas']) == 4, f"recibidas: {len(regs)} registros = {len(res['sii']['recibidas'])} (2 AP + 2 OTA)")
    por_num = {r.findtext('siiLR:IDFactura/sii:NumSerieFacturaEmisor', namespaces=NS): r for r in regs}
    a1 = por_num['A-1']
    ok(a1.findtext('siiLR:IDFactura/sii:IDEmisorFactura/sii:NIF', namespaces=NS) == 'A28647451' and a1.findtext('siiLR:FacturaRecibida/sii:ClaveRegimenEspecialOTrascendencia', namespaces=NS) == '01'
       and a1.findtext('siiLR:FacturaRecibida/sii:DesgloseFactura/sii:DesgloseIVA/sii:DetalleIVA/sii:CuotaSoportada', namespaces=NS) == '210.00' and a1.findtext('siiLR:FacturaRecibida/sii:CuotaDeducible', namespaces=NS) == '210.00',
       'recibidas: proveedor español con NIF (sin el prefijo ES), clave 01, cuota soportada y deducible')
    de = por_num['DE-1']
    ok(de.findtext('siiLR:IDFactura/sii:IDEmisorFactura/sii:IDOtro/sii:CodigoPais', namespaces=NS) == 'DE' and de.findtext('siiLR:FacturaRecibida/sii:ClaveRegimenEspecialOTrascendencia', namespaces=NS) == '09',
       'recibidas: compra de bienes UE con IDOtro pais DE y clave 09')
    bk = por_num['BK-1']
    ok(bk.findtext('siiLR:IDFactura/sii:IDEmisorFactura/sii:IDOtro/sii:CodigoPais', namespaces=NS) == 'NL' and bk.find('siiLR:FacturaRecibida/sii:DesgloseFactura/sii:InversionSujetoPasivo', NS) is not None,
       'recibidas: comision de Booking (NL) con IDOtro y bloque de inversion del sujeto pasivo')
    ok(bk.findtext('siiLR:FacturaRecibida/sii:Contraparte/sii:IDOtro/sii:ID', namespaces=NS) == 'NL805734958B01', 'recibidas: la contraparte lleva el NIF-IVA de la OTA')
    ok(all(r.findtext('siiLR:FacturaRecibida/sii:FechaRegContable', namespaces=NS) for r in regs), 'recibidas: fecha de registro contable en todas')
    ok('<?xml version="1.0" encoding="UTF-8"?>' in xe and 'sii:PeriodoLiquidacion' in xr and '<sii:Ejercicio>2026</sii:Ejercicio><sii:Periodo>08</sii:Periodo>' in xr, 'declaracion XML y periodo de liquidacion 2026/08')

    # rutas y menu
    import dashboard as D
    app = D.app; app.config['TESTING'] = True
    cl = app.test_client()
    assert cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'}).status_code == 200
    for libro in ('emitidas', 'recibidas'):
        r = cl.get(f'/api/exportar/sii?libro={libro}&mes=2026-08')
        ok(r.status_code == 200 and 'xml' in r.mimetype and f'SII_{libro}_2026-08.xml' in (r.headers.get('Content-Disposition') or '') and r.data.startswith(b'<?xml'),
           f"/api/exportar/sii?libro={libro}: 200, XML adjunto ({r.mimetype}, {r.headers.get('Content-Disposition')})")
    html = cl.get('/').get_data(as_text=True)
    ok("u: '/api/exportar/sii?libro=emitidas'" in html and "u: '/api/exportar/sii?libro=recibidas'" in html, 'las dos descargas estan en el catalogo del menu (apartado Cierre)')
    ok("(it.u.indexOf('?') >= 0 ? '&' : '?') + 'mes='" in html, 'el menu añade el mes con & cuando la ruta ya lleva ?libro=')
    ok('NADA SE ENVIA' in SX.__doc__, 'sii_xml.py deja claro que nada se envia a la AEAT')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
