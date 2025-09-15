import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- Define the Rule-Based Tagging Logic ---
def tag_ticket(subcategory, subject):
    """
    Tags a CRM ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()  # Convert to lowercase for case-insensitive matching

    # ========== CRM ERROR ==========
    if subcategory == 'CRM ERROR':
        if re.search(r'\b(login|log in|log-in|access|credentials|password)\b', subject_lower):
            return 'CRM - Login/Access Issue'
        elif re.search(r'\b(reassign|re-assign|reassigning|assignment|assign)\b', subject_lower):
            return 'CRM - Request Reassignment Issue'
        elif re.search(r'\b(new|create|setup|configuration)\b', subject_lower):
            return 'CRM - New Setup/Configuration'
        elif re.search(r'\b(error|issue|problem|fail|not working|bug)\b', subject_lower):
            return 'CRM - General System Error'
        else:
            return 'CRM - Other Error'

    # ========== For any other CRM subcategories ==========
    else:
        return f'CRM - {subcategory}'

def generate_crm_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('CRM Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/crm_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/crm_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("CRM Ticket Tag Report", styles['Title']),
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