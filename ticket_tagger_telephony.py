import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags a Telephony ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(cisco|phone)\b', subject_lower):
        return 'Telephony - Cisco Phones'
    elif re.search(r'\b(teleconference|conference)\b', subject_lower):
        return 'Telephony - Teleconference'
    elif re.search(r'\b(external|calling|code)\b', subject_lower):
        return 'Telephony - External Calling Code'
    elif re.search(r'\b(call|center|screen|cti)\b', subject_lower):
        return 'Telephony - Call Center'
    elif re.search(r'\b(finesse)\b', subject_lower):
        return 'Telephony - Finesse'
    else:
        return 'Telephony - General Issue'

def generate_telephony_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Telephony Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/telephony_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/telephony_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Telephony Ticket Tag Report", styles['Title']),
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
