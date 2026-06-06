"""
Generador de Reportes PDF automáticos
Diarios, semanales, mensuales
"""
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from datetime import datetime
import os

class GeneradorReportePDF:
    def __init__(self, titulo, tipo="diario"):
        self.titulo = titulo
        self.tipo = tipo
        self.fecha = datetime.now()
        self.styles = getSampleStyleSheet()
        self.story = []
    
    def agregar_titulo(self):
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1F4E78'),
            spaceAfter=30,
            alignment=1
        )
        self.story.append(Paragraph(self.titulo, title_style))
        fecha_str = self.fecha.strftime("%d/%m/%Y")
        self.story.append(Paragraph(f"<font size=10 color=#666666>Generado: {fecha_str}</font>", self.styles['Normal']))
        self.story.append(Spacer(1, 0.3*inch))
    
    def agregar_kpis(self, kpis_dict):
        self.story.append(Paragraph("KPIs Principales", self.styles['Heading2']))
        self.story.append(Spacer(1, 0.2*inch))
        
        data = [["Métrica", "Valor"]]
        for k, v in kpis_dict.items():
            data.append([k.replace("_", " ").title(), str(v)])
        
        table = Table(data, colWidths=[3*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f0f0')])
        ]))
        self.story.append(table)
        self.story.append(Spacer(1, 0.3*inch))
    
    def agregar_resumen(self, texto):
        self.story.append(Paragraph("Resumen", self.styles['Heading2']))
        self.story.append(Spacer(1, 0.1*inch))
        self.story.append(Paragraph(texto, self.styles['Normal']))
        self.story.append(Spacer(1, 0.3*inch))
    
    def generar(self, output_path=None):
        if output_path is None:
            output_path = f"/tmp/yve01/reportes/Reporte_{self.tipo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        doc.build(self.story)
        return output_path

def generar_reporte_diario():
    gen = GeneradorReportePDF("Reporte Diario YVE", "diario")
    gen.agregar_titulo()
    kpis = {"Revenue": "€45,230", "Ocupacion": "88.5%", "Facturas": "12"}
    gen.agregar_kpis(kpis)
    gen.agregar_resumen("Día operativo normal. Todas las facturas procesadas.")
    return gen.generar()

def generar_reporte_semanal():
    gen = GeneradorReportePDF("Reporte Semanal YVE", "semanal")
    gen.agregar_titulo()
    kpis = {"Revenue MTD": "€320,610", "Ocupacion Avg": "87.3%", "Facturas": "85"}
    gen.agregar_kpis(kpis)
    gen.agregar_resumen("Semana productiva con buena gestión de facturas.")
    return gen.generar()

def generar_reporte_mensual():
    gen = GeneradorReportePDF("Reporte Mensual YVE", "mensual")
    gen.agregar_titulo()
    kpis = {"Revenue Total": "€1,918,400", "Ocupacion": "87.5%", "GOP%": "21.6%", "Facturas": "342"}
    gen.agregar_kpis(kpis)
    gen.agregar_resumen("Mes cierra exitosamente. Revenue 98% del presupuesto.")
    return gen.generar()

if __name__ == "__main__":
    print("Generando reportes PDF...")
    print(f"✓ Diario: {generar_reporte_diario()}")
    print(f"✓ Semanal: {generar_reporte_semanal()}")
    print(f"✓ Mensual: {generar_reporte_mensual()}")
