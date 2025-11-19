import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Desktop Hardware/Software ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(hardware|monitor|keyboard|mouse|laptop|computer)\b', subject_lower):
        return 'Desktop - Hardware Issue'
    elif re.search(r'\b(software|application|program|install)\b', subject_lower):
        return 'Desktop - Software Issue'
    elif re.search(r'\b(network|connection|internet|lan)\b', subject_lower):
        return 'Desktop - Network Connection'
    elif re.search(r'\b(system|computer|pc|performance)\b', subject_lower):
        return 'Desktop - System Performance'
    elif re.search(r'\b(hrms)\b', subject_lower):
        return 'Desktop - HRMS Access'
    elif re.search(r'\b(projector|laptop|setup)\b', subject_lower):
        return 'Desktop - Setup/Mapping'
    else:
        return 'Desktop - General Issue'

def generate_deskhardwaresoftware_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Desktop Hardware/Software Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/deskhardwaresoftware_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/deskhardwaresoftware_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Desktop Hardware/Software Ticket Tag Report", styles['Title']),
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
