"""
oracle_export_dryrun.py — Yve.01
Genera asientos contables en formato Oracle GL sin necesidad de credenciales.
Exporta a Excel para importación manual.
"""
import os, json
import pandas as pd
from datetime import datetime, date
from flask import Blueprint, jsonify, send_file, request
from io import BytesIO

oracle_export_bp = Blueprint('oracle_export', __name__)
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATOS      = os.path.join(BASE_DIR, 'datos-referencia')
REPORTES   = os.path.join(BASE_DIR, 'reportes')

# Oracle GL column mapping (Fusion format)
GL_COLUMNS = [
    'STATUS', 'LEDGER_NAME', 'ACCOUNTING_DATE', 'CURRENCY_CODE',
    'USER_JE_CATEGORY_NAME', 'USER_JE_SOURCE_NAME', 'REFERENCE_DATE',
    'ACTUAL_FLAG', 'SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4',
    'ENTERED_DR', 'ENTERED_CR', 'ACCOUNTED_DR', 'ACCOUNTED_CR',
    'DESCRIPTION', 'ATTRIBUTE1', 'ATTRIBUTE2', 'CONVERSION_TYPE',
    'CONVERSION_RATE', 'PERIOD_NAME'
]

def _get_config():
    cfg_path = os.path.join(DATOS, 'hotel_config.json')
    if os.path.exists(cfg_path):
        with open(cfg_path) as f: return json.load(f)
    return {'hotel_nombre': 'Hotel Demo', 'hotel_habitaciones': 142}

def _period_name(dt):
    """Format date as Oracle period: MAY-26"""
    months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
    d = pd.Timestamp(dt)
    return f"{months[d.month-1]}-{str(d.year)[2:]}"

def _make_gl_row(ledger, date_str, cat, seg1, seg2, seg3, seg4, dr, cr, desc, ref1='', ref2=''):
    period = _period_name(date_str)
    return {
        'STATUS': 'NEW',
        'LEDGER_NAME': ledger,
        'ACCOUNTING_DATE': date_str,
        'CURRENCY_CODE': 'EUR',
        'USER_JE_CATEGORY_NAME': cat,
        'USER_JE_SOURCE_NAME': 'YVE01',
        'REFERENCE_DATE': date_str,
        'ACTUAL_FLAG': 'A',
        'SEGMENT1': seg1, 'SEGMENT2': seg2, 'SEGMENT3': seg3, 'SEGMENT4': seg4,
        'ENTERED_DR': round(dr, 2) if dr else '',
        'ENTERED_CR': round(cr, 2) if cr else '',
        'ACCOUNTED_DR': round(dr, 2) if dr else '',
        'ACCOUNTED_CR': round(cr, 2) if cr else '',
        'DESCRIPTION': desc[:240],
        'ATTRIBUTE1': ref1, 'ATTRIBUTE2': ref2,
        'CONVERSION_TYPE': 'User', 'CONVERSION_RATE': 1.0,
        'PERIOD_NAME': period
    }

def _generate_journal_from_ap():
    """Generate Oracle GL journal entries from AP matching reports."""
    import glob as _g
    cfg = _get_config()
    ledger = cfg.get('ledger_name', f"{cfg.get('hotel_nombre','Hotel')} Ledger")
    
    rows = []
    hits = _g.glob(os.path.join(REPORTES, 'matching_*.xlsx'))
    
    if not hits:
        # Generate demo entries from proveedores.xlsx
        ruta = os.path.join(DATOS, 'proveedores.xlsx')
        if not os.path.exists(ruta): return []
        df = pd.read_excel(ruta)
        today = date.today().isoformat()
        for _, prov in df.iterrows():
            importe = float(prov.get('importe_mensual_estimado', 5000) or 5000)
            base    = round(importe / 1.10, 2)
            iva     = round(importe - base, 2)
            cuenta  = str(prov.get('cuenta_contable', '6000'))
            nombre  = str(prov.get('nombre', 'Proveedor'))
            
            # DEBE: Gasto
            rows.append(_make_gl_row(ledger, today, 'Accounts Payable',
                '01', '1000', cuenta, '0000',
                dr=base, cr=0, desc=f'Gasto {nombre}', ref1=nombre))
            # DEBE: IVA soportado
            rows.append(_make_gl_row(ledger, today, 'Accounts Payable',
                '01', '1000', '4720', '0000',
                dr=iva, cr=0, desc=f'IVA soportado {nombre}', ref1=nombre))
            # HABER: Proveedor
            rows.append(_make_gl_row(ledger, today, 'Accounts Payable',
                '01', '1000', '4000', '0000',
                dr=0, cr=importe, desc=f'Factura {nombre}', ref1=nombre))
        return rows
    
    for ruta in hits[:3]:  # Max 3 report files
        df = pd.read_excel(ruta)
        for _, row in df.iterrows():
            if str(row.get('aprobacion','')) != 'APROBADA': continue
            importe = float(row.get('importe_con_iva', 0) or 0)
            base    = round(importe / 1.10, 2)
            iva     = round(importe - base, 2)
            cuenta  = str(row.get('cuenta_contable', '6000'))
            prov    = str(row.get('proveedor', 'Proveedor'))[:30]
            num_fac = str(row.get('numero_factura', ''))[:20]
            fecha   = str(row.get('fecha_factura', date.today()))[:10]
            
            rows.append(_make_gl_row(ledger, fecha, 'Accounts Payable',
                '01', '1000', cuenta, '0000', dr=base, cr=0,
                desc=f'GASTO {prov}', ref1=num_fac, ref2=prov))
            rows.append(_make_gl_row(ledger, fecha, 'Accounts Payable',
                '01', '1000', '4720', '0000', dr=iva, cr=0,
                desc=f'IVA {prov}', ref1=num_fac))
            rows.append(_make_gl_row(ledger, fecha, 'Accounts Payable',
                '01', '1000', '4000', '0000', dr=0, cr=importe,
                desc=f'PROV {prov}', ref1=num_fac, ref2=prov))
    return rows

@oracle_export_bp.route('/api/oracle/dryrun')
def api_oracle_dryrun():
    """Preview Oracle journal entries as JSON."""
    rows = _generate_journal_from_ap()
    total_dr = sum(float(r.get('ENTERED_DR') or 0) for r in rows)
    total_cr = sum(float(r.get('ENTERED_CR') or 0) for r in rows)
    return jsonify({
        'ok': True,
        'mode': 'dry_run',
        'entries': len(rows),
        'total_debe': round(total_dr, 2),
        'total_haber': round(total_cr, 2),
        'balanced': abs(total_dr - total_cr) < 0.01,
        'sample': rows[:5],
    })

@oracle_export_bp.route('/api/oracle/export_excel')
def api_oracle_export_excel():
    """Export Oracle GL journal entries as Excel for manual import."""
    rows = _generate_journal_from_ap()
    if not rows:
        return jsonify({'error': 'Sin asientos que exportar'}), 404
    
    df = pd.DataFrame(rows, columns=GL_COLUMNS)
    
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='GL_INTERFACE', index=False)
        ws = writer.sheets['GL_INTERFACE']
        # Style header
        from openpyxl.styles import PatternFill, Font, Alignment
        header_fill = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(color='FFFFFF', bold=True, size=9)
            cell.alignment = Alignment(horizontal='center')
        ws.freeze_panes = 'A2'
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 16
    
    buf.seek(0)
    fname = f"oracle_gl_journal_{date.today().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=fname)

@oracle_export_bp.route('/api/oracle/status')
def api_oracle_status():
    """Check Oracle connectivity (sim vs real)."""
    oracle_url = os.environ.get('ORACLE_BASE_URL', '')
    return jsonify({
        'mode': 'real' if oracle_url else 'simulation',
        'connected': bool(oracle_url),
        'message': 'Oracle Fusion conectado' if oracle_url else 'Modo simulación activo — exporta a Excel para importar manualmente',
        'export_url': '/api/oracle/export_excel',
    })
