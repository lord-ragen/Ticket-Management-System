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
    Tags a ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()  # Convert to lowercase for case-insensitive matching

    # 1. ========== WHIZZ ==========
    if subcategory == 'Whizz':
        if re.search(r'\b(delink|delete|de-linking|remove|deletion)\b', subject_lower):
            return 'Whizz - Delinking/Deletion Request'
        elif re.search(r'\b(loan|limit|repay)\b', subject_lower):
            return 'Whizz - Loan-Related Issue'
        elif re.search(r'\b(otp|not receiving|sms|message|alert)\b', subject_lower):
            return 'Whizz - OTP/SMS Issue'
        elif re.search(r'\b(not visible|not viewing|invisible|view account|view.*account)\b', subject_lower):
            return 'Whizz - Account Visibility Issue'
        elif re.search(r'\b(failing|declining|transaction|paybill|failed)\b', subject_lower):
            return 'Whizz - Transaction Failure'
        elif re.search(r'\b(statement|estatement)\b', subject_lower):
            return 'Whizz - Statement Issue'
        elif re.search(r'\b(ussd)\b', subject_lower):
            return 'Whizz - USSD Issue'
        elif re.search(r'\b(setup|onboarding|download|install|reset|register)\b', subject_lower):
            return 'Whizz - Setup/Onboarding Issue'
        else:
            # Catch-all for any other Whizz error
            return 'Whizz - General App Error'

    # 2. ========== MIPS ==========
    elif subcategory == 'MIPS':
        if re.search(r'\b(password|reset|unlock)\b', subject_lower):
            return 'MIPS - Password Reset/Unlock'
        elif re.search(r'\b(install|setup|map|mapping|rights)\b', subject_lower):
            return 'MIPS - Installation/Setup/Mapping'
        elif re.search(r'\b(slow|not working|error|abort|lock)\b', subject_lower):
            return 'MIPS - General Performance Error'
        else:
            return 'MIPS - Other Request'

    # 3. ========== PESALINK ==========
    elif subcategory == 'Pesalink':
        if re.search(r'\b(error|fail|not working|unable)\b', subject_lower):
            return 'Pesalink - Transaction Error'
        elif re.search(r'\b(block|login|credential|portal)\b', subject_lower):
            return 'Pesalink - Portal/Access Issue'
        else:
            return 'Pesalink - Other Request'

    # 4. ========== ATM ==========
    elif subcategory == 'ATM':
        if re.search(r'\b(down|out of service|not working)\b', subject_lower):
            return 'ATM - Down/Out of Service'
        elif re.search(r'\b(error|card|transaction|reverse|powercard)\b', subject_lower):
            return 'ATM - Transaction/Card Error'
        else:
            return 'ATM - Other Request'

    # 5. ========== Western Union / Money Gram ==========
    elif subcategory in ['Western Union', 'Money Gram']:
        if re.search(r'\b(password|reset|login)\b', subject_lower):
            return f'{subcategory} - Password/Login Issue'
        elif re.search(r'\b(install|setup|creation|new)\b', subject_lower):
            return f'{subcategory} - Installation/Setup'
        elif re.search(r'\b(error|fail)\b', subject_lower):
            return f'{subcategory} - Transaction Error'
        else:
            return f'{subcategory} - Other Request'

    # 6. ========== Internet Banking ==========
    elif subcategory == 'Internet Banking':
        if re.search(r'\b(otp|credential|login|reset)\b', subject_lower):
            return 'Internet Banking - Login/OTP Issue'
        elif re.search(r'\b(statement)\b', subject_lower):
            return 'Internet Banking - Statement Issue'
        else:
            return 'Internet Banking - Other Request'

    # 7. ========== RTGS ==========
    elif subcategory == 'RTGS':
        if re.search(r'\b(error|fail|not working|unable)\b', subject_lower):
            return 'RTGS - Transaction Error'
        else:
            return 'RTGS - Other Request'

    # 8. ========== For all other subcategories (KPrinter, USSD, TransUnion, etc.) ==========
    else:
        return f'{subcategory} - General Request'

def generate_channels_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Channels Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/channels_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # Generate PDF
    pdf_path = f"{save_dir}/channels_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Channels Ticket Tag Report", styles['Title']),
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