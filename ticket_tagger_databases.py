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
    Tags a Databases ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()  # Convert to lowercase for case-insensitive matching

    # ========== UPGRADES ==========
    if subcategory == 'Upgrades':
        if re.search(r'\b(password|login|access|authentication|credentials)\b', subject_lower):
            return 'Databases - Access/Authentication Issue'
        elif re.search(r'\b(upgrade|update|migration|version|patch)\b', subject_lower):
            return 'Databases - Upgrade/Migration Request'
        elif re.search(r'\b(performance|slow|optimization|tuning|index)\b', subject_lower):
            return 'Databases - Performance Optimization'
        elif re.search(r'\b(backup|recovery|restore|dr|disaster)\b', subject_lower):
            return 'Databases - Backup/Recovery'
        elif re.search(r'\b(connection|connectivity|timeout|network|host)\b', subject_lower):
            return 'Databases - Connection Issues'
        elif re.search(r'\b(error|fail|crash|corrupt|corruption)\b', subject_lower):
            return 'Databases - General Error'
        elif re.search(r'\b(space|storage|capacity|disk|memory)\b', subject_lower):
            return 'Databases - Storage/Space Issues'
        elif re.search(r'\b(schema|table|structure|design)\b', subject_lower):
            return 'Databases - Schema/Structure Changes'
        elif re.search(r'\b(whizz|mips|profits|crm|iapply)\b', subject_lower):
            return 'Databases - Application Database Issue'
        else:
            return 'Databases - Other Database Request'

    # ========== For any other Databases subcategories ==========
    else:
        return f'Databases - {subcategory}'

def generate_databases_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Databases Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/databases_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/databases_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Databases Ticket Tag Report", styles['Title']),
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