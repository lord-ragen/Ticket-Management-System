import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """Tags an Email ticket"""
    subject_lower = subject.lower()
    
    if re.search(r'\b(outlook|connection|connect)\b', subject_lower):
        return 'Email - Outlook Connection'
    elif re.search(r'\b(office 365|365)\b', subject_lower):
        return 'Email - Office 365 Access'
    elif re.search(r'\b(blocked|release|hold|quarantine)\b', subject_lower):
        return 'Email - Blocked Mail'
    elif re.search(r'\b(not responding|slow|hang)\b', subject_lower):
        return 'Email - Performance Issue'
    elif re.search(r'\b(archive|storage)\b', subject_lower):
        return 'Email - Archive Issue'
    elif re.search(r'\b(shared|mailbox)\b', subject_lower):
        return 'Email - Shared Mailbox'
    elif re.search(r'\b(attachment|size)\b', subject_lower):
        return 'Email - Attachment Issue'
    elif re.search(r'\b(calendar)\b', subject_lower):
        return 'Email - Calendar Issue'
    elif re.search(r'\b(recipient|restricted)\b', subject_lower):
        return 'Email - Recipient Issue'
    else:
        return 'Email - General Issue'

def generate_email_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)
    tag_counts = df['issue_tag'].value_counts()

    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Email Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/email_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    pdf_path = f"{save_dir}/email_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Email Ticket Tag Report", styles['Title']),
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
