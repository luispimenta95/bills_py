import pdfplumber
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.colors import black

def reduce_font_size(input_pdf, output_pdf, scale_factor):
    # Abrir o PDF existente
    with pdfplumber.open(input_pdf) as pdf:
        # Criar um novo PDF
        doc = SimpleDocTemplate(output_pdf, pagesize=letter)
        story = []

        # Estilo padrão para o texto das tabelas
        styles = getSampleStyleSheet()
        normal_style = styles['Normal']

        # Iterar por cada página no PDF original
        for page in pdf.pages:
            tables = page.extract_tables()

            # Ajustar o tamanho da fonte para o texto das tabelas
            adjusted_style = ParagraphStyle(
                name='AdjustedStyle',
                parent=normal_style,
                fontSize=normal_style.fontSize * scale_factor
            )

            # Adicionar as tabelas ajustadas ao documento
            for table in tables:
                data = [[Paragraph(cell, adjusted_style) for cell in row] for row in table]

                # Estilizar a tabela
                table_style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), black),
                    ('TEXTCOLOR', (0, 0), (-1, 0), black),
                    ('GRID', (0, 0), (-1, -1), 0.5, black),
                ])

                table_obj = Table(data)
                table_obj.setStyle(table_style)

                story.append(table_obj)

        # Construir o novo PDF
        doc.build(story)

# Exemplo de uso
input_pdf = "doc.pdf"
output_pdf = "output.pdf"
scale_factor = 0.8  # Reduz o tamanho da fonte em 20%

reduce_font_size(input_pdf, output_pdf, scale_factor)
