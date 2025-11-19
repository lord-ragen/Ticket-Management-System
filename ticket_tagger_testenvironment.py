import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Test Environment ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(patch|install|upgrade)\b', subject_lower):
        return 'Test Env - Patch Installation'
    elif re.search(r'\b(restore|backup)\b', subject_lower):
        return 'Test Env - Restore'
    elif re.search(r'\b(eod|end.*day)\b', subject_lower):
        return 'Test Env - End of Day'
    elif re.search(r'\b(create|creation|new)\b', subject_lower):
        return 'Test Env - Creation'
    elif re.search(r'\b(performance|slow)\b', subject_lower):
        return 'Test Env - Performance'
    elif re.search(r'\b(login|access)\b', subject_lower):
        return 'Test Env - Login Issue'
    else:
        return 'Test Env - General Issue'

def generate_testenvironment_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Test Environment Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/testenvironment_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/testenvironment_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Test Environment Ticket Tag Report", styles['Title']),
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
