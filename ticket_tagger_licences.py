import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Licences ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(business|permit)\b', subject_lower):
        return 'Licences - Business Permit'
    elif re.search(r'\b(osha)\b', subject_lower):
        return 'Licences - OSHA Permit'
    elif re.search(r'\b(fire)\b', subject_lower):
        return 'Licences - Fire Permit'
    elif re.search(r'\b(cbk)\b', subject_lower):
        return 'Licences - CBK Permit'
    elif re.search(r'\b(trade)\b', subject_lower):
        return 'Licences - Trade Permit'
    elif re.search(r'\b(hire|purchase)\b', subject_lower):
        return 'Licences - Hire Purchase'
    elif re.search(r'\b(signage)\b', subject_lower):
        return 'Licences - Signage'
    elif re.search(r'\b(parking)\b', subject_lower):
        return 'Licences - Parking'
    elif re.search(r'\b(land|rate|rent)\b', subject_lower):
        return 'Licences - Land Related'
    else:
        return 'Licences - General'

def generate_licences_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Licences Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/licences_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/licences_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Licences Ticket Tag Report", styles['Title']),
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
