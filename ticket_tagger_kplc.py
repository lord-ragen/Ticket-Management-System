import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a KPLC ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(new|connection|connect)\b', subject_lower):
        return 'KPLC - New Connection'
    elif re.search(r'\b(reconnect|reconnection)\b', subject_lower):
        return 'KPLC - Reconnection'
    elif re.search(r'\b(bill|settlement|payment)\b', subject_lower):
        return 'KPLC - Bill Settlement'
    else:
        return 'KPLC - General Issue'

def generate_kplc_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('KPLC Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/kplc_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/kplc_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("KPLC Ticket Tag Report", styles['Title']),
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
