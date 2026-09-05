# -*- coding: utf-8 -*-
"""El PDF de la factura corporativa (/api/ar_real/pdf/<n>) daba 500 siempre:
TableStyle mal formado, rutas a la raiz del repo, KeyError con el fichero vacio
y "Hotel / Yve.01 Demo" como emisor. Ahora sale un PDF con el hotel activo.

  python3.12 tests/test_pdf_factura_ar.py
  python3.12 tests/test_pdf_factura_ar.py --sabotaje
"""
import os, shutil, subprocess, sys, tempfile
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE); os.chdir(BASE)
import pandas as pd            # noqa: E402
SABOTAJE = '--sabotaje' in sys.argv
RES = os.path.join(BASE, 'datos-referencia', 'reservas_credito.xlsx')
CLI = os.path.join(BASE, 'datos-referencia', 'clientes_credito.xlsx')


def main():
    fallos = 0
    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond: fallos += 1
    import exportador_pdf as EP
    if SABOTAJE:
        _orig = EP.export_invoice_pdf
        EP.export_invoice_pdf = lambda n: (None, "TableStyle roto")
    tmp = tempfile.mkdtemp(prefix='pdf_'); copias = {}
    for f in (RES, CLI):
        if os.path.exists(f):
            copias[f] = os.path.join(tmp, os.path.basename(f)); shutil.copy(f, copias[f])
    try:
        import dashboard as D
        cl = D.app.test_client(); cl.post('/api/login', json={'username': 'admin', 'password': 'admin123'})
        for f in (RES, CLI):
            if os.path.exists(f): os.remove(f)
        r = cl.get('/api/ar_real/pdf/NADA')
        ok(r.status_code == 404, f"sin fichero de reservas: 404 'no encontrada', no 500 ({r.status_code})")
        pd.DataFrame([{'numero_reserva': 'FAC-2026-001', 'cliente': 'Viatges Mediterrani', 'estado': 'FACTURADO', 'total': 1320.0, 'fecha_emision': '2026-08-20',
                       'fecha_entrada': '2026-08-14', 'fecha_salida': '2026-08-17', 'habitaciones': 2, 'importe_habitaciones': 1100.0, 'importe_fb': 100.0, 'importe_extras': 0}]).to_excel(RES, index=False)
        pd.DataFrame([{'nombre_cliente': 'Viatges Mediterrani', 'nif': 'B-62233445', 'credito_limite': 15000, 'dias_pago': 30}]).to_excel(CLI, index=False)
        buf, err = EP.export_invoice_pdf('FAC-2026-001')
        ok(err is None and buf is not None and buf.getvalue()[:4] == b'%PDF', f"export_invoice_pdf genera un PDF ({err or 'ok'})")
        r = cl.get('/api/ar_real/pdf/FAC-2026-001')
        ok(r.status_code == 200 and r.data[:4] == b'%PDF' and 'pdf' in r.headers.get('Content-Type', ''), f"/api/ar_real/pdf → {r.status_code} {r.headers.get('Content-Type','')[:20]}")
        try:
            from pypdf import PdfReader
            import io
            txt = "".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(r.data)).pages)
            ok('Viatges Mediterrani' in txt and 'B-62233445' in txt and 'Yve.01 Demo' not in txt and '1.320,00' in txt, "el PDF lleva cliente, NIF, euros en español y NO 'Yve.01 Demo'")
        except ImportError:
            print("  (sin pypdf: no se lee el texto del PDF)")
        r = cl.get('/api/ar_real/pdf/NO-EXISTE')
        ok(r.status_code == 404, f"factura inexistente → 404 ({r.status_code})")
    finally:
        for f in (RES, CLI):
            if os.path.exists(f): os.remove(f)
            if f in copias: shutil.copy(copias[f], f)
        shutil.rmtree(tmp, ignore_errors=True)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***'); sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS'); sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
