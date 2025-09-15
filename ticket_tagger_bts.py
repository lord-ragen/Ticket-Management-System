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
    Tags a BTS ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()  # Convert to lowercase for case-insensitive matching

    # 1. ========== PROFIT CENTERS ==========
    if subcategory == 'Profit Centers':
        # Assuming most requests are for adding new centers. Add more rules if needed.
        if re.search(r'\b(add|new|additional|create)\b', subject_lower):
            return 'BTS - Profit Center Creation'
        elif re.search(r'\b(modify|update|change|edit)\b', subject_lower):
            return 'BTS - Profit Center Modification'
        else:
            return 'BTS - Profit Center Management'

    # 2. ========== INCORRECT SERVER NAME ==========
    elif subcategory == 'Incorrect Server Name':
        # This subcategory name is a misnomer. We re-tag based on the actual subject.
        if re.search(r'\b(password|vision.*password|request.*password)\b', subject_lower):
            return 'BTS - Vision Password Reset'
        elif re.search(r'\b(listing|report|premium|bond|discount|statement)\b', subject_lower):
            return 'BTS - Report Generation'
        elif re.search(r'\b(error|issue|problem|fail|not.*working)\b', subject_lower):
            return 'BTS - System Error / Issue'
        else:
            # Fallback for any other request in this subcategory
            return 'BTS - General/Other Request'

    # 3. ========== For any other BTS subcategories ==========
    else:
        return f'BTS - {subcategory}'

def generate_bts_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('BTS Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/bts_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/bts_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("BTS Ticket Tag Report", styles['Title']),
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