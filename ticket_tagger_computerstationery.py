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
    Tags a Computer Stationery ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()  # Convert to lowercase for case-insensitive matching

    # ========== PRINTER TONERS ==========
    if subcategory == 'Printer Toners':
        if re.search(r'\b(toner|cartridge|ink|printer.*ink)\b', subject_lower):
            return 'Computer Stationery - Toner/Cartridge Request'
        elif re.search(r'\b(headset|headphone|earphone|headphone)\b', subject_lower):
            return 'Computer Stationery - Headset/Audio Equipment'
        elif re.search(r'\b(mouse|keyboard|key pad)\b', subject_lower):
            return 'Computer Stationery - Input Devices'
        elif re.search(r'\b(cable|adapter|charger|power.*cord)\b', subject_lower):
            return 'Computer Stationery - Cables & Adapters'
        elif re.search(r'\b(request|need|require|requisition)\b', subject_lower):
            return 'Computer Stationery - General Request'
        else:
            return 'Computer Stationery - Other Accessory'

    # ========== For any other Computer Stationery subcategories ==========
    else:
        return f'Computer Stationery - {subcategory}'

def generate_computer_stationery_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Computer Stationery Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/computer_stationery_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/computer_stationery_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Computer Stationery Ticket Tag Report", styles['Title']),
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