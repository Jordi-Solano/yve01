"""
exportador_pdf.py — Yve PDF export
Genera informes PDF de AR, AP, DRR y F&B.
"""
import os, io
from pathlib import Path
from datetime import datetime
from flask import Blueprint, send_file, Response

pdf_bp = Blueprint('pdf', __name__)
BASE_DIR = Path(__file__).parent

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, HRFlowable)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

YVE_BLUE  = colors.HexColor('#3b82f6')
YVE_DARK  = colors.HexColor('#0f172a')
YVE_SLATE = colors.HexColor('#1e293b')
YVE_GRAY  = colors.HexColor('#334155')
YVE_GREEN = colors.HexColor('#22c55e')
YVE_RED   = colors.HexColor('#ef4444')
YVE_ORA   = colors.HexColor('#f97316')
WHITE     = colors.white
LIGHT     = colors.HexColor('#f1f5f9')
MUT       = colors.HexColor('#94a3b8')


def _build_pdf(elements, filename):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    doc.build(elements)
    buf.seek(0)
    return send_file(buf, mimetype='application/pdf',
                     as_attachment=True, download_name=filename)


def _header(styles, title, subtitle=''):
    from reportlab.platypus import Table as T, TableStyle as TS
    date_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    header_data = [[
        Paragraph('<font color="#3b82f6"><b>Yve.01</b></font> — ' + title, styles['h1']),
        Paragraph(f'<font color="#94a3b8">{subtitle}<br/>{date_str}</font>', styles['right']),
    ]]
    t = T(header_data, colWidths=['70%','30%'])
    t.setStyle(TS([
        ('BACKGROUND', (0,0), (-1,-1), YVE_DARK),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [YVE_DARK]),
        ('TOPPADDING', (0,0), (-1,-1), 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 12),
        ('LEFTPADDING', (0,0), (0,-1), 14),
        ('RIGHTPADDING', (-1,0), (-1,-1), 14),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    return t


def _styles():
    s = getSampleStyleSheet()
    # Remove existing 'h1' alias if present to avoid duplicate error
    if 'h1' in s.byName: del s.byName['h1']
    if 'h1' in s.byAlias: del s.byAlias['h1']
    s.add(ParagraphStyle('h1', parent=s['Normal'],
          fontSize=15, textColor=WHITE, fontName='Helvetica-Bold', leading=18))
    if 'right' in s.byName: del s.byName['right']
    if 'right' in s.byAlias: del s.byAlias['right']
    s.add(ParagraphStyle('right', parent=s['Normal'],
          fontSize=9, textColor=MUT, alignment=TA_RIGHT, leading=12))
    if 'section' in s.byName: del s.byName['section']
    if 'section' in s.byAlias: del s.byAlias['section']
    s.add(ParagraphStyle('section', parent=s['Normal'],
          fontSize=11, textColor=YVE_BLUE, fontName='Helvetica-Bold',
          spaceBefore=16, spaceAfter=8))
    if 'body' in s.byName: del s.byName['body']
    if 'body' in s.byAlias: del s.byAlias['body']
    s.add(ParagraphStyle('body', parent=s['Normal'],
          fontSize=9, textColor=LIGHT, leading=12))
    if 'small' in s.byName: del s.byName['small']
    if 'small' in s.byAlias: del s.byAlias['small']
    s.add(ParagraphStyle('small', parent=s['Normal'],
          fontSize=8, textColor=MUT, leading=10))
    return s


def _kpi_table(kpis):
    """kpis = [(label, value, color), ...]"""
    from reportlab.platypus import Table as T, TableStyle as TS
    cols = min(len(kpis), 4)
    rows = [kpis[i:i+cols] for i in range(0, len(kpis), cols)]
    table_data = []
    for row in rows:
        table_data.append([
            Paragraph(
                f'<font color="#94a3b8" size="8">{lbl.upper()}</font><br/>'
                f'<font color="{col}" size="16"><b>{val}</b></font>',
                ParagraphStyle('kpi', fontSize=9, leading=20)
            )
            for lbl, val, col in row
        ])
    t = T(table_data, colWidths=[4.25*cm]*cols)
    t.setStyle(TS([
        ('BACKGROUND', (0,0), (-1,-1), YVE_SLATE),
        ('BOX', (0,0), (-1,-1), 0.5, YVE_GRAY),
        ('INNERGRID', (0,0), (-1,-1), 0.5, YVE_GRAY),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
    ]))
    return t


# ── AR PDF ─────────────────────────────────────────────────────────────────
@pdf_bp.route('/api/exportar/ar/pdf')
def export_ar_pdf():
    if not REPORTLAB_OK:
        return Response('reportlab no instalado', status=500)
    try:
        import pandas as pd, json
        styles = _styles()
        elements = []
        elements.append(_header(styles, 'Informe AR — OTAs', 'Comisiones y Facturas'))
        elements.append(Spacer(1, 0.4*cm))

        # Try to load AR data
        try:
            from notificaciones import escanear_alertas
            alertas = escanear_alertas()
        except Exception:
            alertas = []

        elements.append(Paragraph('Resumen del ciclo AR', styles['section']))
        kpis = [
            ('Estado', 'Activo', '#22c55e'),
            ('Fecha informe', datetime.now().strftime('%d/%m/%Y'), '#60a5fa'),
        ]
        elements.append(_kpi_table(kpis))
        elements.append(Spacer(1, 0.3*cm))

        # Alertas table
        if alertas:
            elements.append(Paragraph('Alertas activas', styles['section']))
            headers = [['Tipo', 'Descripción', 'Nivel']]
            rows = [[a.get('tipo',''), a.get('descripcion','')[:60], a.get('nivel','')] for a in alertas[:20]]
            t_data = headers + rows
            t = Table(t_data, colWidths=[3*cm, 10*cm, 3*cm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), YVE_BLUE),
                ('TEXTCOLOR', (0,0), (-1,0), WHITE),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [YVE_DARK, YVE_SLATE]),
                ('TEXTCOLOR', (0,1), (-1,-1), LIGHT),
                ('GRID', (0,0), (-1,-1), 0.3, YVE_GRAY),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ]))
            elements.append(t)
        else:
            elements.append(Paragraph('Sin alertas activas en este ciclo.', styles['body']))

        return _build_pdf(elements, f'yve_ar_{datetime.now().strftime("%Y%m%d")}.pdf')
    except Exception as e:
        return Response(f'Error: {e}', status=500)


# ── F&B PDF ────────────────────────────────────────────────────────────────
@pdf_bp.route('/api/exportar/fb/pdf')
def export_fb_pdf():
    if not REPORTLAB_OK:
        return Response('reportlab no instalado', status=500)
    try:
        import pandas as pd, json
        styles = _styles()
        elements = []
        elements.append(_header(styles, 'Informe F&B Cost Control', 'Food Cost · Inventario · Mermas'))
        elements.append(Spacer(1, 0.4*cm))

        # Load F&B data
        datos = BASE_DIR / 'datos-referencia'
        df_ven = pd.read_excel(datos / 'ventas_fb_diarias.xlsx')
        df_mer = pd.read_excel(datos / 'mermas.xlsx')
        df_rec = pd.read_excel(datos / 'recetas.xlsx')
        df_inv = pd.read_excel(datos / 'inventario.xlsx')

        total_ventas   = float(df_ven['total_venta'].sum())
        total_mermas   = float(df_mer['coste_merma'].sum()) if 'coste_merma' in df_mer else 0

        elements.append(Paragraph('KPIs del período', styles['section']))
        kpis = [
            ('Ventas F&B', f'€{total_ventas:,.0f}', '#60a5fa'),
            ('Mermas', f'€{total_mermas:.2f}', '#ef4444'),
            ('Recetas activas', str(len(df_rec)), '#22c55e'),
            ('Items inventario', str(len(df_inv)), '#f97316'),
        ]
        elements.append(_kpi_table(kpis))
        elements.append(Spacer(1, 0.3*cm))

        # Top ventas por plato
        elements.append(Paragraph('Top Platos — Ventas del Período', styles['section']))
        rank = df_ven.groupby('nombre_plato')['total_venta'].sum().sort_values(ascending=False).head(10)
        headers = [['Plato', 'Categoría', 'Ventas €']]
        rows_data = []
        cats = df_ven.groupby('nombre_plato')['categoria'].first()
        for plato, venta in rank.items():
            rows_data.append([plato, cats.get(plato, ''), f'€{venta:,.0f}'])
        t = Table(headers + rows_data, colWidths=[7*cm, 5*cm, 4*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), YVE_BLUE),
            ('TEXTCOLOR', (0,0), (-1,0), WHITE),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [YVE_DARK, YVE_SLATE]),
            ('TEXTCOLOR', (0,1), (-1,-1), LIGHT),
            ('GRID', (0,0), (-1,-1), 0.3, YVE_GRAY),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('ALIGN', (2,0), (2,-1), 'RIGHT'),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 0.3*cm))

        # Mermas
        elements.append(Paragraph('Registro de Mermas', styles['section']))
        m_headers = [['Fecha', 'Ingrediente', 'Causa', 'Coste €']]
        m_rows = [[str(r['fecha'])[:10], str(r['ingrediente']),
                   str(r['causa']), f'€{float(r["coste_merma"]):.2f}']
                  for _, r in df_mer.iterrows()]
        t2 = Table(m_headers + m_rows, colWidths=[3*cm, 5*cm, 5*cm, 3*cm])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), YVE_BLUE),
            ('TEXTCOLOR', (0,0), (-1,0), WHITE),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [YVE_DARK, YVE_SLATE]),
            ('TEXTCOLOR', (0,1), (-1,-1), LIGHT),
            ('GRID', (0,0), (-1,-1), 0.3, YVE_GRAY),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('ALIGN', (3,0), (3,-1), 'RIGHT'),
        ]))
        elements.append(t2)

        return _build_pdf(elements, f'yve_fb_{datetime.now().strftime("%Y%m%d")}.pdf')
    except Exception as e:
        return Response(f'Error: {e}', status=500)


# ── DRR PDF ────────────────────────────────────────────────────────────────
@pdf_bp.route('/api/exportar/drr/pdf')
def export_drr_pdf():
    if not REPORTLAB_OK:
        return Response('reportlab no instalado', status=500)
    try:
        import pandas as pd, glob
        styles = _styles()
        elements = []
        elements.append(_header(styles, 'Daily Revenue Report', 'Análisis de Revenue Diario'))
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph('Informe generado automáticamente por Yve.01', styles['body']))
        elements.append(Spacer(1, 0.3*cm))

        # Try to read DRR processed file
        reportes = BASE_DIR / 'reportes'
        drr_files = sorted(glob.glob(str(reportes / 'drr_procesado_*.xlsx')))
        if drr_files:
            df = pd.read_excel(drr_files[-1], sheet_name='Resumen', header=None)
            elements.append(Paragraph('Métricas del período', styles['section']))
            data_table = []
            for _, row in df.iterrows():
                if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
                    data_table.append([str(row[0])[:40], str(row[1])[:40]])
            if data_table:
                t = Table(data_table[:20], colWidths=[8*cm, 8*cm])
                t.setStyle(TableStyle([
                    ('ROWBACKGROUNDS', (0,0), (-1,-1), [YVE_DARK, YVE_SLATE]),
                    ('TEXTCOLOR', (0,0), (-1,-1), LIGHT),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('GRID', (0,0), (-1,-1), 0.3, YVE_GRAY),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                elements.append(t)
        else:
            elements.append(Paragraph('Procesa el DRR primero desde la pestaña DRR.', styles['body']))

        return _build_pdf(elements, f'yve_drr_{datetime.now().strftime("%Y%m%d")}.pdf')
    except Exception as e:
        return Response(f'Error: {e}', status=500)


def export_invoice_pdf(numero_factura):
    """Genera PDF de una factura corporativa."""
    import os as _os
    import pandas as pd
    from io import BytesIO
    from datetime import datetime

    ruta = _os.path.join(_os.path.dirname(__file__), 'datos-referencia', 'reservas_credito.xlsx')
    ruta_c = _os.path.join(_os.path.dirname(__file__), 'datos-referencia', 'clientes_credito.xlsx')
    
    df_r = pd.read_excel(ruta)
    df_c = pd.read_excel(ruta_c)
    
    row = df_r[df_r['numero_reserva'].astype(str) == str(numero_factura)]
    if row.empty:
        return None, f"Factura {numero_factura} no encontrada"
    
    row = row.iloc[0]
    cliente = str(row.get('cliente', ''))
    c_data = df_c[df_c['nombre_cliente'] == cliente]
    nif = str(c_data.iloc[0].get('NIF','')) if len(c_data) else '—'
    
    total = float(row.get('total', 0))
    base = round(total / 1.10, 2)
    iva = round(total - base, 2)
    
    buf = BytesIO()
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
        from reportlab.lib.units import cm
        
        doc = SimpleDocTemplate(buf, pagesize=A4,
            topMargin=2*cm, bottomMargin=2*cm,
            leftMargin=2.5*cm, rightMargin=2.5*cm)
        
        s = getSampleStyleSheet()
        story = []
        
        # Header
        header_data = [
            [Paragraph('<font size="22" color="#3b82f6"><b>Yve.01</b></font>', s['Normal']),
             Paragraph(f'<font size="9" color="#64748b">FACTURA<br/><font size="18" color="#0f172a"><b>{numero_factura}</b></font></font>', s['Normal'])]
        ]
        ht = Table(header_data, colWidths=[10*cm, 5.5*cm])
        ht.setStyle(TableStyle([('ALIGN',(1,0),'RIGHT'),('VALIGN',(0,0),'TOP'),('VALIGN',(1,0),'TOP')]))
        story.append(ht)
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#3b82f6")))
        story.append(Spacer(1, 0.5*cm))
        
        # Dates and parties
        fecha_em = str(row.get('fecha_emision',''))[:10]
        fecha_venc = fecha_em  # simplify
        
        parties = [
            [Paragraph('<b><font color="#64748b" size="8">DE</font></b>', s['Normal']),
             Paragraph('<b><font color="#64748b" size="8">PARA</font></b>', s['Normal'])],
            [Paragraph('<b>Hotel / Yve.01 Demo</b><br/><font size="8" color="#64748b">Barcelona, España</font>', s['Normal']),
             Paragraph(f'<b>{cliente}</b><br/><font size="8" color="#64748b">NIF: {nif}</font>', s['Normal'])],
        ]
        pt = Table(parties, colWidths=[7.75*cm, 7.75*cm])
        pt.setStyle(TableStyle([('VALIGN',(0,0),'TOP'),('VALIGN',(1,0),'TOP'),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
        story.append(pt)
        story.append(Spacer(1, 0.5*cm))
        
        # Dates info
        info_data = [
            ['Fecha emisión:', fecha_em, 'Entrada:', str(row.get('fecha_entrada',''))[:10]],
            ['Habitaciones:', str(row.get('habitaciones',1)), 'Salida:', str(row.get('fecha_salida',''))[:10]],
        ]
        it = Table(info_data, colWidths=[3.5*cm, 3.5*cm, 3*cm, 5.5*cm])
        it.setStyle(TableStyle([
            ('FONTSIZE',(0,0),(-1,-1),9),
            ('TEXTCOLOR',(0,0),(0,-1),colors.HexColor("#64748b")),
            ('TEXTCOLOR',(2,0),(2,-1),colors.HexColor("#64748b")),
            ('FONTNAME',(1,0),(1,-1),'Helvetica-Bold'),
            ('FONTNAME',(3,0),(3,-1),'Helvetica-Bold'),
        ]))
        story.append(it)
        story.append(Spacer(1, 0.5*cm))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e2e8f0")))
        story.append(Spacer(1, 0.3*cm))
        
        # Line items
        items_header = [['Descripción', 'Importe']]
        items = []
        hab = float(row.get('importe_habitaciones', 0))
        fb  = float(row.get('importe_fb', 0))
        ext = float(row.get('importe_extras', 0))
        if hab > 0: items.append([f"Habitaciones ({row.get('habitaciones',1)} hab.)", f"€{hab:,.2f}"])
        if fb > 0:  items.append(["F&B y restauración", f"€{fb:,.2f}"])
        if ext > 0: items.append(["Extras y servicios", f"€{ext:,.2f}"])
        
        items_data = items_header + items + [
            ['', ''],
            ['Base imponible:', f"€{base:,.2f}"],
            ['IVA (10%):', f"€{iva:,.2f}"],
            ['', ''],
            [Paragraph('<b>TOTAL</b>', s['Normal']), Paragraph(f'<b>€{total:,.2f}</b>', s['Normal'])],
        ]
        lw = Table(items_data, colWidths=[12*cm, 3.5*cm])
        lw.setStyle(TableStyle([
            ('FONTSIZE',(0,0),(-1,-1),10),
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#0f172a")),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('ALIGN',(1,0),(1,-1),'RIGHT'),
            ('LINEBELOW',(0,-1),(-1,-1),1,colors.HexColor("#3b82f6")),
            ('LINEBELOW',(0,0),(-1,0),0.5,colors.HexColor("#1e293b")),
            ('BACKGROUND',(0,-1),(-1,-1),colors.HexColor("#eff6ff")),
            ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,-1),(-1,-1),12),
            ('ROWBACKGROUNDS',(0,1),(-1,-5),[colors.white, colors.HexColor("#f8fafc")]),
            ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6),
        ]))
        story.append(lw)
        story.append(Spacer(1, 1*cm))
        
        # Footer
        story.append(Paragraph(
            '<font size="8" color="#64748b">Factura emitida por Yve.01 — Sistema de gestión financiera hotelera · yve01.onrender.com<br/>'
            'Condiciones: Pago según términos contractuales · En caso de discrepancia contacte a jordi@yve01.com</font>',
            s['Normal']))
        
        doc.build(story)
        buf.seek(0)
        return buf, None
    except ImportError:
        return None, "reportlab no instalado"
    except Exception as e:
        return None, str(e)
