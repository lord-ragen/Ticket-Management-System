import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a General Repairs ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(plumbing)\b', subject_lower):
        return 'Repairs - Plumbing'
    elif re.search(r'\b(furniture)\b', subject_lower):
        return 'Repairs - Furniture'
    elif re.search(r'\b(security|door)\b', subject_lower):
        return 'Repairs - Security Doors'
    elif re.search(r'\b(window|glass)\b', subject_lower):
        return 'Repairs - Windows/Glass'
    elif re.search(r'\b(steel)\b', subject_lower):
        return 'Repairs - Steel Works'
    elif re.search(r'\b(paint|painting)\b', subject_lower):
        return 'Repairs - Painting'
    elif re.search(r'\b(electrical|electric)\b', subject_lower):
        return 'Repairs - Electrical'
    elif re.search(r'\b(refurbish|renovation)\b', subject_lower):
        return 'Repairs - Refurbishment'
    else:
        return 'Repairs - General'

def generate_generalrepairs_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('General Repairs Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/generalrepairs_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/generalrepairs_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("General Repairs Ticket Tag Report", styles['Title']),
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
