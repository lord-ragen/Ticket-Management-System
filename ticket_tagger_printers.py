import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Printers ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(jam|paper)\b', subject_lower):
        return 'Printers - Paper Jam'
    elif re.search(r'\b(faulty|not working|error)\b', subject_lower):
        return 'Printers - Faulty'
    elif re.search(r'\b(toner|cartridge|refill)\b', subject_lower):
        return 'Printers - Toner Request'
    elif re.search(r'\b(mapping|map|setup)\b', subject_lower):
        return 'Printers - Mapping/Setup'
    elif re.search(r'\b(offline|down)\b', subject_lower):
        return 'Printers - Offline'
    elif re.search(r'\b(scan|scan|scanning)\b', subject_lower):
        return 'Printers - Scan Issues'
    elif re.search(r'\b(color|colour|print)\b', subject_lower):
        return 'Printers - Color Printing'
    else:
        return 'Printers - General Issue'

def generate_printers_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Printers Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/printers_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/printers_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Printers Ticket Tag Report", styles['Title']),
        Spacer(1, 12),
        Image(pie_chart_path, width=300, height=300),
        Spacer(1, 12),
        Table(table_data, style=[
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
    ]
    doc.build(elements)
