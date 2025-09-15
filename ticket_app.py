import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.widgets import Frame, Button, Label, Combobox
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from tkinter import Toplevel, BooleanVar, Checkbutton

# PDF reporting imports
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt
import os
import tempfile

# Import your PDF generation functions (make sure these exist)
from ticket_tagger_channels import generate_channels_pdf
from ticket_tagger_useradministration import generate_useradmin_pdf

# ----------------------
# Category configuration
# ----------------------
CATEGORIES = {
    "ACs Service/Repair": ["installation", "maintenance"],
    "Antivirus": ["antivirus"],
    "BTS": ["profit centers", "incorrect server name"],
    "Channels": ["mips", "atm", "waptx", "internet banking", "western union", "money gram", 
                "cms", "citrix", "kprinter", "fimi", "hfdi crm", "internet banking", "iprs", 
                "kprinter", "mbanking admin portal", "mips", "moneygram", "pesalink", 
                "rates upload for emp", "rtgs", "swift", "transunion", "ussd", "visacard", 
                "waptx", "western union", "whatsapp", "whizz"],
    "Computer Stationery": ["printer toners"],
    "CRM": ["crm error"],
    "Data Center": ["dc ac", "dr server issues", "dr network equipment", "ho servers", 
                   "ho network equipment"],
    "Databases": ["upgrades"],
    "Desktop Hardware/Software": ["hardware", "basic software", "network connection", 
                                 "computer functionality", "system mapping", "usb bypass code", 
                                 "projector and laptop setup", "hrms"],
    "Email": ["outlook connection", "office 365 access", "release blocked mails", 
             "outlook not responding", "archive creation", "shared", "attachment size", 
             "outlook calendar", "archive attachment", "restricted recipient"],
    "Enterprise Service Bus (ESB)": ["pesalink", "mpesa c2b", "mpesa b2c", 
                                    "internal funds transfer (ift)", "host to host", "hf leads"],
    "EOD": ["eod errors"],
    "Equinox": ["mapping"],
    "General Repairs": ["plumbing works", "furniture works", "security doors", 
                       "window/glass works", "steel works", "painting works", 
                       "electrical works", "refurbishment"],
    "Generator": ["maintenance", "fueling", "installation"],
    "Helpdesk": ["helpdesk"],
    "Iapply": ["application re-route", "profits error message", "connection marked dead", 
              "virtual user error", "db error:login failed"],
    "Information Provision": ["information to user"],
    "Internet": ["internet connection"],
    "Intranet/Corporate Website": ["not accessible", "upload files-intranet"],
    "iProfits": ["unable to log in"],
    "IPRS": ["iprs"],
    "JIRA": ["link access", "username and password"],
    "K-Printer": ["k-printer"],
    "KPLC": ["new conection", "reconnection", "bill settlement"],
    "Licences": ["business permit", "osha permit", "fire permit", "cbk permit", 
                "trade permit", "hire purchase", "signage", "parking", "land rates", 
                "land rents"],
    "Lift": ["maintenance installation"],
    "Market Place and Lead Management": ["market place", "lead management"],
    "MIPS": ["mips"],
    "Money Gram": ["money gram"],
    "Network": ["hq lan", "wireless", "safaricom link", "cbk connection", "comtec connection", 
               "mpls branch connection", "kenswitch connection", "internet", 
               "access kenya link", "remote connection"],
    "Notification and Alerts": ["sms alerts", "email alerts", "disable wrong alerts"],
    "Online Bid Bond": ["online bid bond"],
    "Printed Stationery": ["validation stationery", "white envelopes b4 hfc", 
                          "brown env b5 hfc", "white robin envelop(dl) hfc", 
                          "brown env a3 hfc", "window envelopes hfc", "letterheads hfc", 
                          "conti. letterheads hfc", "complimentary slips hfc", 
                          "white envelopes b4 hf group"],
    "Printers": ["paper jam", "faulty printer", "toner request", "printer mapping", 
                "printer id", "scan issues", "printer offline", "colour printing"],
    "Production Scheduler Run": ["profits scheduler"],
    "Profits": ["profits loan", "profits deposit", "transaction alerts", "till number", 
               "seller code creation/migration", "profits performance"],
    "Records & Case Management": ["records", "case management"],
    "Reports": ["bi reports", "profits", "crm", "iapply", "mobile banking", "emp", 
               "equinox", "reports and other related issues", "toad", "eft", "manageengine"],
    "SAP BI": ["sap bi"],
    "SAP Fi": ["gl", "tb"],
    "SWIFT": ["swift"],
    "Taxis/Cabs": ["bill settlement"],
    "Team Mate": ["team mate"],
    "Telephony": ["cisco phones", "teleconference", "external calling code", 
                 "call center screens", "cti phones", "finesse"],
    "Test Environment": ["patch installation", "profits", "crm", "iapply", "restore", 
                        "end of day", "creation", "environment performance", 
                        "iapply", "profits", "login"],
    "UPS": ["ups"],
    "User Administration": ["ad", "password resets", "rights change", "access to shared folder", 
                           "system user creation", "account activation", "user relocation", 
                           "access to ftp"],
    "Vehicle": ["fueling", "purchase", "maintenance", "policy enhancement"],
    "Vendor Support": ["hfc partner company"],
    "Water": ["bill settlement", "drinking water", "dispenser installation"],
    "Western Union": ["western union"],
    "Zendesk": ["zendesk"],
    "Other": []
}

CATEGORY_COLORS = {
    "ACs Service/Repair": "#F8D7DA",
    "Antivirus": "#FFF3CD",
    "BTS": "#D1ECF1",
    "Channels": "#D4EDDA",
    "Computer Stationery": "#E2E3E5",
    "CRM": "#FDFD96",
    "Data Center": "#D1E7FF",
    "Databases": "#FFE4E1",
    "Desktop Hardware/Software": "#F0FFF0",
    "Email": "#F5F5DC",
    "Enterprise Service Bus (ESB)": "#E6E6FA",
    "EOD": "#FFF0F5",
    "Equinox": "#F0F8FF",
    "General Repairs": "#FAEBD7",
    "Generator": "#F0FFFF",
    "Helpdesk": "#F5F5F5",
    "Iapply": "#FFFACD",
    "Information Provision": "#E0FFFF",
    "Internet": "#FAFAD2",
    "Intranet/Corporate Website": "#F8F8FF",
    "iProfits": "#F5DEB3",
    "IPRS": "#FFF5EE",
    "JIRA": "#F0FFF0",
    "K-Printer": "#F5F5DC",
    "KPLC": "#FFEFD5",
    "Licences": "#FFE4B5",
    "Lift": "#FFDAB9",
    "Market Place and Lead Management": "#FFE4E1",
    "MIPS": "#FFEFD5",
    "Money Gram": "#FFF0F5",
    "Network": "#E6E6FA",
    "Notification and Alerts": "#F0F8FF",
    "Online Bid Bond": "#FAEBD7",
    "Printed Stationery": "#F0FFFF",
    "Printers": "#F5F5F5",
    "Production Scheduler Run": "#FFFACD",
    "Profits": "#E0FFFF",
    "Records & Case Management": "#FAFAD2",
    "Reports": "#F8F8FF",
    "SAP BI": "#F5DEB3",
    "SAP Fi": "#FFF5EE",
    "SWIFT": "#F0FFF0",
    "Taxis/Cabs": "#F5F5DC",
    "Team Mate": "#FFEFD5",
    "Telephony": "#FFE4B5",
    "Test Environment": "#FFDAB9",
    "UPS": "#FFE4E1",
    "User Administration": "#FFEFD5",
    "Vehicle": "#FFF0F5",
    "Vendor Support": "#E6E6FA",
    "Water": "#F0F8FF",
    "Western Union": "#FAEBD7",
    "Zendesk": "#F0FFFF",
    "Other": "#CCCCCC"
}

class TicketApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ticket Categorizer")
        self.style = Style("flatly")
        self.df = None
        self.filtered_df = None
        self.actual_subcategory_col = None

        # Frame Setup
        frame = Frame(root, padding=20)
        frame.pack(fill="both", expand=True)

        # Buttons
        Button(frame, text="Load Excel", command=self.load_excel,
               bootstyle="primary").grid(row=0, column=0, padx=5, pady=5)
        Button(frame, text="Export PDF Report", command=self.export_pdf,
               bootstyle="success").grid(row=0, column=1, padx=5, pady=5)

        # Filter by Category
        Label(frame, text="Filter by Category:", bootstyle="inverse").grid(row=0, column=2, padx=5, pady=5)
        self.category_var = tk.StringVar()
        self.category_dropdown = Combobox(frame, textvariable=self.category_var, state="readonly", width=20)
        self.category_dropdown.grid(row=0, column=3, padx=5, pady=5)
        self.category_dropdown.bind("<<ComboboxSelected>>", self.apply_filters)

        # Filter by Subcategory
        Label(frame, text="Filter by Subcategory:", bootstyle="inverse").grid(row=0, column=4, padx=5, pady=5)
        self.subcategory_var = tk.StringVar()
        self.subcategory_dropdown = Combobox(frame, textvariable=self.subcategory_var, state="readonly", width=20)
        self.subcategory_dropdown.grid(row=0, column=5, padx=5, pady=5)
        self.subcategory_dropdown.bind("<<ComboboxSelected>>", self.apply_filters)

        # Clear Filters Button
        Button(frame, text="Clear Filters", command=self.clear_filters,
               bootstyle="warning").grid(row=0, column=6, padx=5, pady=5)
        Button(frame, text="Requester Stats", command=self.show_requester_stats,
               bootstyle="info").grid(row=0, column=7, padx=5, pady=5)

        # Table (Treeview)
        self.tree = ttk.Treeview(frame, columns=("Subcategory", "Category", "Count"), show="headings", height=15)
        self.tree.heading("Subcategory", text="Subcategory")
        self.tree.heading("Category", text="Category")
        self.tree.heading("Count", text="Count")
        self.tree.grid(row=1, column=0, columnspan=7, sticky="nsew", pady=10)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=1, column=7, sticky="ns")

        # Graph Area
        self.fig, self.ax = plt.subplots(figsize=(8, 4))
        self.ax.clear()
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.ax.set_frame_on(False)
        self.ax.set_title("")
        self.ax.text(0.5, 0.5, '', ha='center', va='center')  # Blank message
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().grid(row=2, column=0, columnspan=8, pady=10)

        # Grid expand
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        # Initialize dropdowns
        self.category_dropdown["values"] = ["All Categories"] + list(CATEGORIES.keys())
        self.category_dropdown.set("All Categories")
        self.subcategory_dropdown["values"] = ["All Subcategories"]
        self.subcategory_dropdown.set("All Subcategories")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path, header=9)
            self.df.columns = (
                self.df.columns.astype(str)
                .str.strip()
                .str.lower()
                .str.replace(' ', '_')
                .str.replace('[^a-zA-Z0-9_]', '', regex=True)
            )

            # Find the subcategory column
            self.actual_subcategory_col = None
            for col in self.df.columns:
                if 'subcategory' in col.lower():
                    self.actual_subcategory_col = col
                    break

            if self.actual_subcategory_col is None:
                messagebox.showerror("Error", f"No 'Subcategory' column found! Available: {', '.join(self.df.columns)}")
                return

            # Categorize
            self.df["Category"] = self.df[self.actual_subcategory_col].apply(self.categorize_ticket)

            # Summarize
            self.filtered_df = (
                self.df.groupby([self.actual_subcategory_col, "Category"])
                .size()
                .reset_index(name="Count")
            )

            self.update_filters()
            self.update_table(self.filtered_df)
            self.plot_category_pie(self.filtered_df)

            messagebox.showinfo("Success", "Excel file loaded successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel: {e}")

    def categorize_ticket(self, subcategory):
        subcategory = str(subcategory).lower()
        for category, keywords in CATEGORIES.items():
            if any(keyword in subcategory for keyword in keywords):
                return category
        return "Other"

    def update_filters(self):
        categories = ["All Categories"] + sorted(self.filtered_df["Category"].unique().tolist())
        self.category_dropdown["values"] = categories

        selected_category = self.category_var.get()
        if selected_category != "All Categories" and selected_category in self.filtered_df["Category"].values:
            subcategories = self.filtered_df[self.filtered_df["Category"] == selected_category][self.actual_subcategory_col].unique()
            subcategories = ["All Subcategories"] + sorted([str(s) for s in subcategories])
        else:
            subcategories = ["All Subcategories"] + sorted(self.filtered_df[self.actual_subcategory_col].unique().tolist())

        self.subcategory_dropdown["values"] = subcategories

    def apply_filters(self, event=None):
        if self.filtered_df is None:
            return

        data = self.filtered_df.copy()

        selected_category = self.category_var.get()
        if selected_category != "All Categories":
            data = data[data["Category"] == selected_category]

        selected_subcategory = self.subcategory_var.get()
        if selected_subcategory != "All Subcategories":
            data = data[data[self.actual_subcategory_col] == selected_subcategory]

        self.update_table(data)
        self.plot_category_pie(data)

    def clear_filters(self):
        self.category_var.set("All Categories")
        self.subcategory_var.set("All Subcategories")
        if self.filtered_df is not None:
            self.update_table(self.filtered_df)
            self.plot_category_pie(self.filtered_df)

    def update_table(self, data):
        for row in self.tree.get_children():
            self.tree.delete(row)

        if data is None or len(data) == 0:
            return

        for _, row in data.iterrows():
            tag = row["Category"].replace(" ", "_")
            self.tree.insert("", "end", values=(row[self.actual_subcategory_col], row["Category"], row["Count"]), tags=(tag,))

        for category, color in CATEGORY_COLORS.items():
            self.tree.tag_configure(category.replace(" ", "_"), background=color)

    def plot_category_pie(self, data):
        self.ax.clear()

        if data is None or len(data) == 0:
            self.ax.text(0.5, 0.5, 'No data to display', ha='center', va='center')
            self.canvas.draw()
            return

        category_counts = data.groupby("Category")["Count"].sum()
        percentages = (category_counts / category_counts.sum()) * 100

        wedges, texts, autotexts = self.ax.pie(
            percentages,
            labels=category_counts.index,
            autopct="%1.1f%%",
            startangle=140,
            colors=[CATEGORY_COLORS.get(cat, "lightgrey") for cat in category_counts.index],
            wedgeprops={'edgecolor': 'white'}
        )

        for text in texts:
            text.set_fontsize(8)
        for autotext in autotexts:
            autotext.set_color("white")
            autotext.set_fontsize(7)

        self.ax.set_title(f"Ticket Categories (%) - {int(category_counts.sum())} records", fontsize=10, fontweight="bold")
        self.fig.tight_layout()
        self.canvas.draw()

    def show_requester_stats(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            return

        # Find requester column
        requester_col = None
        for col in self.df.columns:
            if "requester" in col.lower() or "user" in col.lower():
                requester_col = col
                break

        if requester_col is None:
            messagebox.showerror("Error", "No 'Requester' column found in the data!")
            return

        # Find subject column
        subject_col = None
        for col in self.df.columns:
            if "subject" in col.lower():
                subject_col = col
                break

        # Mapping of systems (for User Administration category only)
        system_name_map = {
            'IPROFITS/PROFITS': 'PROFITS',
            'IAPPLY': 'IAPPLY',
            'WINDOWS/AD': 'AD',
            'MIPS/MIPPS': 'MIPS',
            'CRM': 'CRM',
            'Other': 'Other'
        }

        def detect_system(subject):
            if pd.isna(subject):
                return "Other"
            text = str(subject).upper()
            for key, sys in system_name_map.items():
                # allow multiple keywords in key e.g. "IPROFITS/PROFITS"
                for option in key.split("/"):
                    if option in text:
                        return sys
            return "Other"

        # Ask for PDF save location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if not save_path:
            return

        try:
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            elements.append(Paragraph("Requester Ticket Stats by Category", styles['Title']))
            elements.append(Spacer(1, 12))

            # For each category, show requester stats grouped by subcategory
            for category in sorted(self.df["Category"].unique()):
                cat_df = self.df[self.df["Category"] == category]
                elements.append(Paragraph(f"Category: {category}", styles['Heading1']))
                elements.append(Spacer(1, 8))

                # Special handling for User Administration (Password reset / Rights change)
                if category == "User Administration" and subject_col is not None:
                    ua_df = cat_df.copy()
                    ua_df[self.actual_subcategory_col] = ua_df[self.actual_subcategory_col].str.lower()

                    ua_df["System"] = ua_df.apply(
                        lambda r: detect_system(r[subject_col]) 
                        if r[self.actual_subcategory_col] in ["password resets", "rights change"]
                        else "N/A", axis=1
                    )

                    # If system = "N/A", keep normal grouping, else add system
                    stats_df = (
                        ua_df.groupby([requester_col, self.actual_subcategory_col, "System"])
                        .size()
                        .reset_index(name="Ticket Count")
                        .sort_values("Ticket Count", ascending=False)
                    )

                    # Table header
                    table_data = [[requester_col, self.actual_subcategory_col, "System", "Ticket Count"]]
                    table_data += stats_df.values.tolist()

                else:
                    # Default grouping for other categories
                    stats_df = (
                        cat_df.groupby([requester_col, self.actual_subcategory_col])
                        .size()
                        .reset_index(name="Ticket Count")
                        .sort_values("Ticket Count", ascending=False)
                    )

                    # Table header
                    table_data = [[requester_col, self.actual_subcategory_col, "Ticket Count"]]
                    table_data += stats_df.values.tolist()

                if stats_df.empty:
                    elements.append(Paragraph("No tickets for this category.", styles['Normal']))
                    elements.append(Spacer(1, 8))
                    continue

                # Build table
                table = Table(table_data)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 16))

            doc.build(elements)
            messagebox.showinfo("Success", f"Requester stats PDF exported to {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export requester stats PDF: {e}")


    def export_pdf(self):
        if self.filtered_df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if not save_path:
            return

        try:
            temp_files = []  # keep track of all temp chart files

            # 1️⃣ Overall summary
            summary = self.filtered_df.groupby("Category")["Count"].sum().reset_index()
            summary["Percentage"] = (summary["Count"] / summary["Count"].sum() * 100).round(2)

            # 2️⃣ Create main pie chart in temp dir
            plt.figure(figsize=(5, 5))
            plt.pie(summary["Count"], labels=summary["Category"], autopct="%1.1f%%", startangle=140)
            plt.axis("equal")

            overall_chart = os.path.join(tempfile.gettempdir(), "overall_chart.png")
            plt.savefig(overall_chart, bbox_inches="tight")
            plt.close()
            temp_files.append(overall_chart)

            # 3️⃣ Setup PDF
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()

            # Report Title
            elements.append(Paragraph("Ticket Category Report", styles['Title']))
            elements.append(Spacer(1, 12))

            # Overall Summary Table
            table_data = [["Category", "Count", "Percentage (%)"]] + summary.values.tolist()
            table = Table(table_data)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 20))

            # Overall Pie Chart
            elements.append(Paragraph("Overall Category Distribution", styles['Heading2']))
            elements.append(Spacer(1, 12))
            elements.append(Image(overall_chart, width=300, height=300))
            elements.append(PageBreak())

            # 4️⃣ Per-Category Sections
            for _, row in summary.iterrows():
                category = row["Category"]
                cat_df = self.filtered_df[self.filtered_df["Category"] == category]

                # Subcategory summary
                sub_summary = cat_df.groupby(self.actual_subcategory_col)["Count"].sum().reset_index()
                sub_summary["Percentage"] = (sub_summary["Count"] / sub_summary["Count"].sum() * 100).round(2)

                # Subcategory Pie Chart in temp dir
                plt.figure(figsize=(5, 5))
                plt.pie(sub_summary["Count"], labels=sub_summary[self.actual_subcategory_col],
                        autopct="%1.1f%%", startangle=140)
                plt.axis("equal")

                safe_category = category.replace(" ", "_").replace("/", "_")
                sub_chart = os.path.join(tempfile.gettempdir(), f"{safe_category}_chart.png")
                plt.savefig(sub_chart, bbox_inches="tight")
                plt.close()
                temp_files.append(sub_chart)

                # Section Title
                elements.append(Paragraph(f"{category} Breakdown", styles['Heading1']))
                elements.append(Spacer(1, 12))

                # Subcategory Table
                sub_table_data = [[self.actual_subcategory_col, "Count", "Percentage (%)"]] + sub_summary.values.tolist()
                sub_table = Table(sub_table_data)
                sub_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#9BBB59")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                elements.append(sub_table)
                elements.append(Spacer(1, 20))

                # Add Subcategory Pie Chart
                elements.append(Paragraph("Subcategory Distribution", styles['Heading2']))
                elements.append(Spacer(1, 12))
                elements.append(Image(sub_chart, width=300, height=300))
                elements.append(PageBreak())

            # Build PDF
            doc.build(elements)

            # Cleanup temp chart files
            for f in temp_files:
                if os.path.exists(f):
                    os.remove(f)

            messagebox.showinfo("Success", f"PDF report exported to {save_path}")

            # Prompt for additional PDF reports
            dialog = AdditionalPDFDialog(self.root)
            self.root.wait_window(dialog)
            selected_reports = dialog.selected

            # Generate additional PDFs if the functions exist
            if selected_reports:
                self.generate_additional_reports(selected_reports, save_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")

    def generate_additional_reports(self, selected_reports, save_path):
        """Generate additional specialized reports"""
        try:
            temp_csvs = {}
            
            # Prepare temp CSVs
            if "channels" in selected_reports:
                channels_df = self.df[self.df["Category"] == "Channels"]
                if not channels_df.empty:
                    temp_channels_csv = os.path.join(tempfile.gettempdir(), "channels_data.csv")
                    channels_df.to_csv(temp_channels_csv, index=False)
                    temp_csvs["channels"] = temp_channels_csv
            
            if "useradmin" in selected_reports:
                useradmin_df = self.df[self.df["Category"] == "User Administration"]
                if not useradmin_df.empty:
                    temp_useradmin_csv = os.path.join(tempfile.gettempdir(), "useradmin_data.csv")
                    useradmin_df.to_csv(temp_useradmin_csv, index=False)
                    temp_csvs["useradmin"] = temp_useradmin_csv

            # Generate reports (you'll need to implement these functions)
            for report in selected_reports:
                if report == "channels" and "channels" in temp_csvs:
                    generate_channels_pdf(os.path.dirname(save_path), temp_csvs["channels"])
                    messagebox.showinfo("Info", "Channels PDF generated")
                    
                elif report == "useradmin" and "useradmin" in temp_csvs:
                    generate_useradmin_pdf(os.path.dirname(save_path), temp_csvs["useradmin"])
                    messagebox.showinfo("Info", "User Administration PDF generated")

            # Cleanup temp CSV files
            for f in temp_csvs.values():
                if os.path.exists(f):
                    os.remove(f)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate additional reports: {e}")


class AdditionalPDFDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Generate Additional PDF Reports")
        self.selected = []
        self.vars = []
        options = [
            ("Channels Report", "channels"),
            ("User Administration Report", "useradmin")
        ]
        tk.Label(self, text="Select additional PDF reports to generate:").pack(pady=10)
        for label, key in options:
            var = BooleanVar()
            chk = Checkbutton(self, text=label, variable=var)
            chk.pack(anchor="w")
            self.vars.append((key, var))
        btn = tk.Button(self, text="Generate", command=self.on_generate)
        btn.pack(pady=10)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_generate(self):
        self.selected = [key for key, var in self.vars if var.get()]
        self.destroy()

    def on_close(self):
        self.selected = []
        self.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TicketApp(root)
    root.mainloop()