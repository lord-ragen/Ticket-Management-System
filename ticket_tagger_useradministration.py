import pandas as pd
import re
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def tag_ticket(subcategory, subject):
    """
    Tags a User Administration ticket based on its subcategory and subject text using predefined rules.
    """
    subject_lower = subject.lower()
   
    if subcategory == 'Password Reset':
        if re.search(r'\b(expired|forgot|reset|change)\b', subject_lower):
            return 'Password Reset Request'
        else:
            return 'Password Reset - Other'
    elif subcategory == 'User Creation':
        if re.search(r'\b(new|create|add)\b', subject_lower):
            return 'User Creation Request'
        else:
            return 'User Creation - Other'
    elif subcategory == 'Rights Change':
        if re.search(r'\b(rights|permission|access)\b', subject_lower):
            return 'Rights Change Request'
        else:
            return 'Rights Change - Other'
    elif subcategory == 'Account Activation':
        if re.search(r'\b(activate|activation)\b', subject_lower):
            return 'Account Activation Request'
        else:
            return 'Account Activation - Other'
    elif subcategory == 'AD':
        if re.search(r'\b(active directory|ad)\b', subject_lower):
            return 'AD Request'
        else:
            return 'AD - Other'
    elif subcategory == 'User Relocation':
        if re.search(r'\b(relocate|move|transfer)\b', subject_lower):
            return 'User Relocation Request'
        else:
            return 'User Relocation - Other'
    elif subcategory == 'Access to FTP':
        if re.search(r'\b(ftp|access)\b', subject_lower):
            return 'FTP Access Request'
        else:
            return 'FTP Access - Other'
    else:
        return f'{subcategory} - General Request'

def extract_system(subject):
    subject_lower = subject.lower()
    if re.search(r'\b(iprofits|profits)\b', subject_lower):
        return 'IPROFITS/PROFITS'
    elif re.search(r'\b(mips|mipps)\b', subject_lower):
        return 'MIPS/MIPPS'
    elif re.search(r'\biapply\b', subject_lower):
        return 'IAPPLY'
    elif re.search(r'\bcrm\b', subject_lower):
        return 'CRM'
    elif re.search(r'\b(windows|ad|active directory)\b', subject_lower):
        return 'WINDOWS/AD'
    else:
        return 'Other'

def generate_useradmin_pdf(save_dir, csv_path):
    df = pd.read_csv(csv_path)

    # Tag tickets
    df['issue_tag'] = df.apply(lambda row: tag_ticket(row['subcategory'], row['subject']), axis=1)

    # Count tags
    tag_counts = df['issue_tag'].value_counts()

    # Create pie chart for all tags
    plt.figure(figsize=(6, 6))
    tag_counts.plot.pie(autopct='%1.1f%%')
    plt.title('User Administration Ticket Tag Distribution')
    pie_chart_path = f"{save_dir}/useradmin_pie_chart.png"
    plt.savefig(pie_chart_path)
    plt.close()

    # Prepare table data
    table_data = [['Tag', 'Count']]
    for tag, count in tag_counts.items():
        table_data.append([tag, count])

    # --- Password Reset System Pie Chart ---
    password_reset_df = df[df['issue_tag'].str.startswith('Password Reset')]
    password_reset_df['system'] = password_reset_df['subject'].apply(extract_system)
    system_counts = password_reset_df['system'].value_counts()

    # Map system names for the table
    system_name_map = {
        'IPROFITS/PROFITS': 'PROFITS',
        'IAPPLY': 'IAPPLY',
        'WINDOWS/AD': 'AD',
        'MIPS/MIPPS': 'MIPS',
        'CRM': 'CRM',
        'Other': 'Other'
    }
    # Remap index for table display
    system_table_data = [['System', 'Count']]
    for system, count in system_counts.items():
        display_name = system_name_map.get(system, system)
        system_table_data.append([display_name, count])

    plt.figure(figsize=(6, 6))
    system_counts.plot.pie(autopct='%1.1f%%')
    plt.title('Password Reset Tickets by System')
    system_pie_chart_path = f"{save_dir}/password_reset_system_pie_chart.png"
    plt.savefig(system_pie_chart_path)
    plt.close()
    # --- End Password Reset System Pie Chart ---

    # Generate PDF
    pdf_path = f"{save_dir}/useradmin_report.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("User Administration Ticket Tag Report", styles['Title']),
        Spacer(1, 12),
        Image(pie_chart_path, width=300, height=300),
        Spacer(1, 12),
        Table(table_data, style=[
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]),
        Spacer(1, 24),
        Paragraph("Password Reset Tickets by System", styles['Heading2']),
        Spacer(1, 12),
        Image(system_pie_chart_path, width=300, height=300),
        Spacer(1, 12),
        Table(system_table_data, style=[
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]),
    ]
    doc.build(elements)