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
    Tags an ACs Service/Repair ticket based on its subcategory and subject text.
    """
    subject_lower = subject.lower()

    if subcategory == 'Installation':
        if re.search(r'\b(new|install|set up|setup)\b', subject_lower):
            return 'ACs - New Installation'
        else:
            return 'ACs - Installation Request'
    
    elif subcategory == 'Maintenance':
        if re.search(r'\b(service|inspect|check|maintenance|clean)\b', subject_lower):
            return 'ACs - Maintenance Service'
        elif re.search(r'\b(repair|fix|broken|not working|faulty)\b', subject_lower):
            return 'ACs - Repair Service'
        else:
            return 'ACs - Maintenance Request'
    
    else:
        return f'ACs - {subcategory}'

def generate_acsservicerepair_pdf(save_dir, csv_path):
    """Generate PDF report for ACs Service/Repair tickets"""
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('ACs Service/Repair Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/acsservicerepair_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/acsservicerepair_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("ACs Service/Repair Ticket Tag Report", styles['Title']),
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
