import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Data Center ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(dc ac|air|cooling|temperature)\b', subject_lower):
        return 'Data Center - AC Issue'
    elif re.search(r'\b(dr|disaster recovery|failover)\b', subject_lower):
        if re.search(r'\b(server)\b', subject_lower):
            return 'Data Center - DR Server Issue'
        else:
            return 'Data Center - DR Network Issue'
    elif re.search(r'\b(ho|head office)\b', subject_lower):
        if re.search(r'\b(server)\b', subject_lower):
            return 'Data Center - HO Server Issue'
        else:
            return 'Data Center - HO Network Issue'
    else:
        return 'Data Center - General Issue'

def generate_datacenter_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Data Center Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/datacenter_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/datacenter_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Data Center Ticket Tag Report", styles['Title']),
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
