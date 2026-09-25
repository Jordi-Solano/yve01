# -*- coding: utf-8 -*-
"""b96 — el IVA de la factura del grupo por concepto: las SALAS al 21 %.

Visto en produccion el 25 sep 2026 (contrato CG-2026-0917: alojamiento 11.000 y
F&B 2.200 con el 10 % incluido, salas 1.210 con el 21 % incluido): la factura
FAC-2026-CORP-0001 decia "Base 13.100 / IVA (10%) 1.310", y el asiento del cierre
y el 303 lo mismo. Lo correcto: base 13.000 (10.000 + 2.000 + 1.000) e IVA 1.410
(1.000 + 200 al 10 % y 210 al 21 %). 100 EUR de IVA repercutido de menos por
cada 1.210 de salas, y el SII repartia la base a partes iguales entre los tipos.

Se comprueba, con UN solo calculo (cierre_mes.desglose_factura_ar) para todos:
  - el lector de contratos deja el tipo de cada concepto en la fila de la factura;
  - el asiento: 430 D 14.410 / 705 H 11.000 / 700 H 2.000 / 477 H 1.410;
  - el 303: 10 % base 12.000 cuota 1.200; 21 % base 1.000 cuota 210;
  - el SII: un DetalleIVA por tipo con SUS importes (no la base a partes iguales);
  - el PDF de la factura: base 13.000, "IVA (10%)" 1.200 y "IVA (21%)" 210;
  - una factura de grupo de antes (sin los tipos en la fila) tambien va al 21 %;
  - una factura que NO es de grupo sale EXACTAMENTE como antes (todo al 10 %).

  python3.12 tests/test_iva_salas.py
  python3.12 tests/test_iva_salas.py --sabotaje   (las salas vuelven al 10 %: tiene que fallar)
"""
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
DD = os.path.join(BASE, 'datos-referencia')
RES = os.path.join(DD, 'reservas_credito.xlsx')
CLI = os.path.join(DD, 'clientes_credito.xlsx')
NS = {'sii': 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd',
      'siiLR': 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd'}


def contrato():
    return {"es_contrato_grupo": True, "evento": {"id": "EV-77", "nombre": "Congreso Cardio 2026"},
            "contrato_numero": "CG-2026-0917",
            "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000001"},
            "agencia": {"nombre": "Viajes Meridiano S.L.", "cif": "B28004554"},
            "alojamiento": {"fecha_entrada": "2026-10-05", "fecha_salida": "2026-10-08", "noches": 3,
                            "habitaciones": 20, "total_habitaciones": 11000, "iva_pct": 10},
            "fb": {"total": 2200, "iva_pct": 10}, "salas": {"total": 1210},
            "comisiones": {"modo": "porcentaje", "alojamiento_pct": 10}, "facturacion": {"pagador": "agencia"}}


def _viejo(r, pct_h=10.0, pct_f=10.0):
    """El calculo de ANTES (cierre_mes hasta b95), para comprobar que lo que no es de grupo no cambia."""
    def n(v):
        try:
            x = float(v); return 0.0 if x != x else x
        except (TypeError, ValueError):
            return 0.0
    rr = lambda x: round(float(x or 0), 2)
    total = n(r.get("total")) or n(r.get("importe"))
    hab, fb, ext = n(r.get("importe_habitaciones")), n(r.get("importe_fb")), n(r.get("importe_extras"))
    if not (hab or fb or ext):
        hab = total
    s = rr(hab + fb + ext)
    if s and abs(s - total) > 0.011:
        k = total / s; hab, fb, ext = rr(hab * k), rr(fb * k), rr(ext * k)
    b_h = rr((hab + ext) / (1 + pct_h / 100)); i_h = rr(hab + ext - b_h)
    b_f = rr(fb / (1 + pct_f / 100)); i_f = rr(fb - b_f)
    iva = rr(i_h + i_f); dif = rr(total - (b_h + b_f + iva)); iva = rr(iva + dif)
    return b_h, b_f, iva


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import cierre_mes as CM
    import fiscal as FI
    import sii_xml as SX
    import lector_contratos_grupo as L
    if SABOTAJE:
        CM.IVA_SALAS = 10.0                                   # las salas al tipo del alojamiento, como antes
        CM._pct_de = lambda r, col, defecto: float(defecto)   # y sin mirar lo que trae la fila

    cfg = CM.config_cierre('/nonexistent')
    fila = L.transformar(contrato(), hotel_id="")["reserva"]
    ok(fila.get("iva_habitaciones_pct") == 10 and fila.get("iva_fb_pct") == 10 and fila.get("iva_extras_pct") == 21,
       f"el lector deja el IVA de cada concepto en la factura (10/10/21: {fila.get('iva_habitaciones_pct')}/{fila.get('iva_fb_pct')}/{fila.get('iva_extras_pct')})")
    fila.update({"estado": "FACTURADO", "fecha_emision": "2026-09-25", "numero_factura": "FAC-2026-CORP-0001"})

    # 1 · el asiento de la factura
    fu = {'ap': pd.DataFrame(), 'ota': pd.DataFrame(), 'ventas_fb': pd.DataFrame(), 'banco': pd.DataFrame(),
          'provisiones': [], 'reservas': pd.DataFrame([fila])}
    res = CM.generar_asientos('2026-09', fu, CM.plan_cuentas('/nonexistent'), cfg)
    lin = {(a['cuenta'], a['debe'], a['haber']) for a in res['asientos'] if a['documento'] == 'FAC-2026-CORP-0001'}
    ok(lin == {('430', 14410.0, 0.0), ('705', 0.0, 11000.0), ('700', 0.0, 2000.0), ('477', 0.0, 1410.0)},
       f"asiento: 430 D 14.410 / 705 H 11.000 / 700 H 2.000 / 477 H 1.410 ({sorted(lin)})")

    # 2 · el 303
    fi = FI.calcular('2026-09', fu, cfg, {'periodicidad': 'mensual', 'nif_propio': 'B87654321', 'razon_social': 'Hotel Prueba SL'})
    c = {x['clave']: x for x in fi['m303']['casillas']}
    ok(c['dev_10']['base'] == 12000.0 and c['dev_10']['cuota'] == 1200.0 and c['dev_21']['base'] == 1000.0 and c['dev_21']['cuota'] == 210.0,
       f"303: 10 % base 12.000 / 1.200 y 21 % base 1.000 / 210 ({c['dev_10']['base']}/{c['dev_10']['cuota']} · {c['dev_21']['base']}/{c['dev_21']['cuota']})")

    # 3 · el SII: un DetalleIVA por tipo, con SUS importes
    xe, _n = SX.generar(fi, {'nif_propio': 'B87654321', 'razon_social': 'Hotel Prueba SL'}, 'emitidas')
    det = []
    try:
        for d in ET.fromstring(xe.encode('utf-8')).findall('.//sii:DetalleIVA', NS):
            det.append((d.findtext('sii:TipoImpositivo', namespaces=NS), d.findtext('sii:BaseImponible', namespaces=NS),
                        d.findtext('sii:CuotaRepercutida', namespaces=NS)))
    except ET.ParseError as e:
        det = [('XML roto', str(e), '')]
    ok(sorted(det) == [('10.00', '12000.00', '1200.00'), ('21.00', '1000.00', '210.00')],
       f"SII: 10 % 12.000/1.200 y 21 % 1.000/210, no la base a partes iguales ({sorted(det)})")

    # 4 · una factura de grupo de ANTES (sin los tipos en la fila) tambien lleva las salas al 21 %
    vieja = {k: v for k, v in fila.items() if not k.startswith("iva_")}
    dv = CM.desglose_factura_ar(vieja, cfg)
    ok(dv["cuota"] == 1410.0 and dv["b705"] == 11000.0, f"factura de grupo sin los tipos en la fila: salas al 21 % igual ({dv['cuota']}, {dv['b705']})")

    # 5 · lo que NO es de grupo sale EXACTAMENTE como antes
    rnd = random.Random(96)
    iguales = True
    for _ in range(300):
        h = round(rnd.uniform(0, 5000), 2) * rnd.choice([0, 1, 1]); f_ = round(rnd.uniform(0, 900), 2) * rnd.choice([0, 1])
        e = round(rnd.uniform(0, 400), 2) * rnd.choice([0, 1]); tot = round(h + f_ + e + rnd.choice([0, 0, 0.01, -0.02, 3.5]), 2)
        r = {"total": tot or 100.0, "importe_habitaciones": h, "importe_fb": f_, "importe_extras": e}
        d = CM.desglose_factura_ar(r, cfg)
        if (d["b705"], d["b700"], d["cuota"]) != _viejo(r):
            iguales = False
            print("     distinta:", r, (d["b705"], d["b700"], d["cuota"]), _viejo(r))
            break
    ok(iguales, "300 facturas que no son de grupo (con extras, sin F&B, totales descuadrados): mismo 705/700/477 que antes")

    # 6 · el PDF de la factura
    import exportador_pdf as EP
    tmp = tempfile.mkdtemp(prefix='iva_'); copias = {}
    for f in (RES, CLI):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        pd.DataFrame([fila]).to_excel(RES, index=False)
        pd.DataFrame([{'nombre_cliente': 'Viajes Meridiano S.L.', 'nif': 'B28004554', 'credito_limite': 20000, 'dias_pago': 30}]).to_excel(CLI, index=False)
        buf, err = EP.export_invoice_pdf('FAC-2026-CORP-0001')
        txt = ''
        if buf is not None:
            pdf = os.path.join(tmp, 'f.pdf'); open(pdf, 'wb').write(buf.getvalue())
            txt = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True).stdout
        t = re.sub(r'\s+', ' ', txt)
        ok(err is None and 'Base imponible: 13.000,00' in t and 'IVA (10%): 1.200,00' in t and 'IVA (21%): 210,00' in t,
           f"PDF: base 13.000, IVA (10%) 1.200 e IVA (21%) 210 ({err or re.findall(r'(Base imponible:[^€]*€|IVA \\([^)]*\\):[^€]*€)', t)})")
        ok('Salas · IVA 21%' in t and 'Habitaciones (20 hab.) · IVA 10%' in t, "PDF: cada linea dice su tipo cuando hay mas de uno")
    finally:
        for f in (RES, CLI):
            if os.path.exists(f):
                os.remove(f)
            if f in copias:
                shutil.copy(copias[f], f)
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
