import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from tkinter import ttk
from tkinter.ttk import Frame, Button, Label, Combobox
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
from ticket_tagger_acsservicerepair import generate_acsservicerepair_pdf
from ticket_tagger_antivirus import generate_antivirus_pdf
from ticket_tagger_bts import generate_bts_pdf
from ticket_tagger_channels import generate_channels_pdf
from ticket_tagger_computerstationery import generate_computerstationery_pdf
from ticket_tagger_crm import generate_crm_pdf
from ticket_tagger_datacenter import generate_datacenter_pdf
from ticket_tagger_databases import generate_databases_pdf
from ticket_tagger_deskhardwaresoftware import generate_deskhardwaresoftware_pdf
from ticket_tagger_email import generate_email_pdf
from ticket_tagger_esb import generate_esb_pdf
from ticket_tagger_eod import generate_eod_pdf
from ticket_tagger_equinox import generate_equinox_pdf
from ticket_tagger_generalrepairs import generate_generalrepairs_pdf
from ticket_tagger_generator import generate_generator_pdf
from ticket_tagger_helpdesk import generate_helpdesk_pdf
from ticket_tagger_iapply import generate_iapply_pdf
from ticket_tagger_informationprovision import generate_informationprovision_pdf
from ticket_tagger_internet import generate_internet_pdf
from ticket_tagger_intranet import generate_intranet_pdf
from ticket_tagger_iprofits import generate_iprofits_pdf
from ticket_tagger_iprs import generate_iprs_pdf
from ticket_tagger_jira import generate_jira_pdf
from ticket_tagger_kprinter import generate_kprinter_pdf
from ticket_tagger_kplc import generate_kplc_pdf
from ticket_tagger_licences import generate_licences_pdf
from ticket_tagger_lift import generate_lift_pdf
from ticket_tagger_marketplace import generate_marketplace_pdf
from ticket_tagger_mips import generate_mips_pdf
from ticket_tagger_moneygram import generate_moneygram_pdf
from ticket_tagger_network import generate_network_pdf
from ticket_tagger_notifications import generate_notifications_pdf
from ticket_tagger_onlinebiddbond import generate_onlinebiddbond_pdf
from ticket_tagger_printedstationery import generate_printedstationery_pdf
from ticket_tagger_printers import generate_printers_pdf
from ticket_tagger_productionscheduler import generate_productionscheduler_pdf
from ticket_tagger_profits import generate_profits_pdf
from ticket_tagger_recordsmanagement import generate_recordsmanagement_pdf
from ticket_tagger_reports import generate_reports_pdf
from ticket_tagger_sapbi import generate_sapbi_pdf
from ticket_tagger_sapfi import generate_sapfi_pdf
from ticket_tagger_swift import generate_swift_pdf
from ticket_tagger_taxiscabs import generate_taxiscabs_pdf
from ticket_tagger_teammate import generate_teammate_pdf
from ticket_tagger_telephony import generate_telephony_pdf
from ticket_tagger_testenvironment import generate_testenvironment_pdf
from ticket_tagger_ups import generate_ups_pdf
from ticket_tagger_useradministration import generate_useradmin_pdf
from ticket_tagger_vehicle import generate_vehicle_pdf
from ticket_tagger_vendorsupport import generate_vendorsupport_pdf
from ticket_tagger_water import generate_water_pdf
from ticket_tagger_westernunion import generate_westernunion_pdf
from ticket_tagger_zendesk import generate_zendesk_pdf

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
        self.root.title("Ticket Management System - v1.2")
        self.root.geometry("1400x800")
        
        # Modern minimalistic style
        self.style = Style("litera")  # Light, modern theme
        self.df = None
        self.filtered_df = None
        self.actual_subcategory_col = None
        self.requester_col = None
        self.technician_col = None

        # Main Container
        main_container = Frame(root)
        main_container.pack(fill="both", expand=True)

        # Header with Title, Creator Info, and Version
        header = Frame(main_container, height=80)
        header.pack(fill="x", padx=20, pady=15)
        header.configure(relief="flat")

        # Title
        title_label = Label(header, text="Ticket Management System", font=("Segoe UI", 22, "bold"))
        title_label.pack(side="left", anchor="nw")

        # Creator and Version Info (right side)
        info_frame = Frame(header)
        info_frame.pack(side="right", anchor="ne")
        
        creator_label = Label(info_frame, text="Created by Thomas Ragen", font=("Segoe UI", 10))
        creator_label.pack(anchor="e")
        
        version_label = Label(info_frame, text="Version 1.2", font=("Segoe UI", 9), foreground="gray")
        version_label.pack(anchor="e")

        # Separator line
        separator = ttk.Separator(main_container, orient="horizontal")
        separator.pack(fill="x", padx=20)

        # Main content frame
        content_frame = Frame(main_container)
        content_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # Top toolbar - Control Panel
        toolbar = Frame(content_frame)
        toolbar.pack(fill="x", pady=(0, 20))

        # Action buttons (minimalistic spacing)
        Button(toolbar, text="📁 Load Excel", command=self.load_excel,
               bootstyle="primary", width=12).pack(side="left", padx=5)
        Button(toolbar, text="📊 Export PDF", command=self.export_pdf,
               bootstyle="success", width=12).pack(side="left", padx=5)
        Button(toolbar, text="📈 Requester Stats", command=self.show_requester_stats,
               bootstyle="info", width=14).pack(side="left", padx=5)
        Button(toolbar, text="👤 Technician Stats", command=self.show_technician_stats,
               bootstyle="info", width=14).pack(side="left", padx=5)
        Button(toolbar, text="🔄 Clear Filters", command=self.clear_filters,
               bootstyle="warning", width=12).pack(side="left", padx=5)
        Button(toolbar, text="📄 Category Reports", command=self.show_additional_reports_dialog,
               bootstyle="secondary", width=15).pack(side="left", padx=5)

        # Filters section - Minimalistic with two rows
        filters_frame = Frame(content_frame)
        filters_frame.pack(fill="x", pady=(0, 15))

        filter_label = Label(filters_frame, text="Filters:", font=("Segoe UI", 10, "bold"))
        filter_label.pack(side="left", padx=(0, 15))

        # First row - Category and Subcategory
        filters_row1 = Frame(filters_frame)
        filters_row1.pack(side="left", fill="x", padx=(0, 20))

        # Category filter
        Label(filters_row1, text="Category:", font=("Segoe UI", 9)).pack(side="left", padx=(0, 5))
        self.category_var = tk.StringVar()
        self.category_dropdown = Combobox(filters_row1, textvariable=self.category_var, state="readonly", width=20)
        self.category_dropdown.pack(side="left", padx=5)
        self.category_dropdown.bind("<<ComboboxSelected>>", self.on_category_changed)

        # Subcategory filter (disabled until category selected)
        Label(filters_row1, text="Subcategory:", font=("Segoe UI", 9)).pack(side="left", padx=(0, 5))
        self.subcategory_var = tk.StringVar()
        self.subcategory_dropdown = Combobox(filters_row1, textvariable=self.subcategory_var, state="disabled", width=20)
        self.subcategory_dropdown.pack(side="left", padx=5)
        self.subcategory_dropdown.bind("<<ComboboxSelected>>", self.apply_filters)

        # Second row - Requester and Technician
        filters_row2 = Frame(filters_frame)
        filters_row2.pack(side="left", fill="x")

        # Requester filter
        Label(filters_row2, text="Requester:", font=("Segoe UI", 9)).pack(side="left", padx=(0, 5))
        self.requester_var = tk.StringVar()
        self.requester_dropdown = Combobox(filters_row2, textvariable=self.requester_var, state="readonly", width=20)
        self.requester_dropdown.pack(side="left", padx=5)
        self.requester_dropdown.bind("<<ComboboxSelected>>", self.apply_filters)

        # Technician filter
        Label(filters_row2, text="Technician:", font=("Segoe UI", 9)).pack(side="left", padx=(0, 5))
        self.technician_var = tk.StringVar()
        self.technician_dropdown = Combobox(filters_row2, textvariable=self.technician_var, state="readonly", width=20)
        self.technician_dropdown.pack(side="left", padx=5)
        self.technician_dropdown.bind("<<ComboboxSelected>>", self.apply_filters)

        # Main content area (Table + Chart)
        content_split = Frame(content_frame)
        content_split.pack(fill="both", expand=True)

        # Left side - Table
        table_frame = Frame(content_split)
        table_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        table_title = Label(table_frame, text="Ticket Breakdown", font=("Segoe UI", 11, "bold"))
        table_title.pack(anchor="w", pady=(0, 8))

        # Table (Treeview) with modern styling
        self.tree = ttk.Treeview(table_frame, columns=("Subcategory", "Category", "Count"), show="headings", height=20)
        self.tree.heading("Subcategory", text="Subcategory")
        self.tree.heading("Category", text="Category")
        self.tree.heading("Count", text="Count")
        self.tree.column("Subcategory", width=200)
        self.tree.column("Category", width=180)
        self.tree.column("Count", width=80)
        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar for table
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Right side - Chart
        chart_frame = Frame(content_split)
        chart_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        chart_title = Label(chart_frame, text="Category Distribution", font=("Segoe UI", 11, "bold"))
        chart_title.pack(anchor="w", pady=(0, 8))

        # Graph Area
        self.fig, self.ax = plt.subplots(figsize=(6, 4.5), facecolor='white')
        self.fig.tight_layout(pad=1)
        self.ax.clear()
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.ax.set_frame_on(False)
        self.ax.text(0.5, 0.5, 'Load data to see chart', ha='center', va='center', fontsize=11, color='gray')
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Status bar at bottom
        status_frame = Frame(main_container, relief="sunken", height=25)
        status_frame.pack(fill="x", side="bottom", padx=20, pady=(10, 5))
        
        self.status_label = Label(status_frame, text="Ready", font=("Segoe UI", 9), foreground="gray")
        self.status_label.pack(anchor="w", padx=10)

        # Initialize dropdowns
        self.category_dropdown["values"] = ["All Categories"] + list(CATEGORIES.keys())
        self.category_dropdown.set("All Categories")
        self.subcategory_dropdown["values"] = ["All Subcategories"]
        self.subcategory_dropdown.set("All Subcategories")
        self.requester_dropdown["values"] = ["All Requesters"]
        self.requester_dropdown.set("All Requesters")
        self.technician_dropdown["values"] = ["All Technicians"]
        self.technician_dropdown.set("All Technicians")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.status_label.config(text="Loading Excel file...")
            self.root.update()
            
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
                self.status_label.config(text="Error: Subcategory column not found")
                return

            # Find the requester column
            self.requester_col = None
            for col in self.df.columns:
                if 'requester' in col.lower():
                    self.requester_col = col
                    break

            # Find the technician column
            self.technician_col = None
            for col in self.df.columns:
                if 'technician' in col.lower() or 'assigned_to' in col.lower():
                    self.technician_col = col
                    break

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

            self.status_label.config(text=f"✓ Loaded {len(self.df)} tickets | {len(self.filtered_df)} unique categories")
            messagebox.showinfo("Success", f"Excel file loaded successfully!\n\n{len(self.df)} tickets processed")

        except Exception as e:
            self.status_label.config(text=f"✗ Error loading file: {str(e)[:40]}")
            messagebox.showerror("Error", f"Failed to read Excel: {e}")

    def categorize_ticket(self, subcategory):
        subcategory = str(subcategory).lower()
        for category, keywords in CATEGORIES.items():
            if any(keyword in subcategory for keyword in keywords):
                return category
        return "Other"

    def update_filters(self):
        """Update the filter dropdowns with available options"""
        categories = ["All Categories"] + sorted(self.filtered_df["Category"].unique().tolist())
        self.category_dropdown["values"] = categories

        # Populate requester dropdown
        requesters = ["All Requesters"]
        if self.requester_col and self.requester_col in self.df.columns:
            requesters += sorted([str(r) for r in self.df[self.requester_col].unique() if pd.notna(r)])
        self.requester_dropdown["values"] = requesters

        # Populate technician dropdown
        technicians = ["All Technicians"]
        if self.technician_col and self.technician_col in self.df.columns:
            technicians += sorted([str(t) for t in self.df[self.technician_col].unique() if pd.notna(t)])
        self.technician_dropdown["values"] = technicians

    def on_category_changed(self, event=None):
        """Handle category selection change - update subcategories"""
        selected_category = self.category_var.get()
        
        # Enable subcategory dropdown only if category is selected
        if selected_category != "All Categories":
            self.subcategory_dropdown.config(state="readonly")
            # Filter subcategories for this category
            subcategories = self.filtered_df[self.filtered_df["Category"] == selected_category][self.actual_subcategory_col].unique()
            subcategories = ["All Subcategories"] + sorted([str(s) for s in subcategories])
        else:
            # Reset subcategory dropdown
            self.subcategory_dropdown.config(state="readonly")
            subcategories = ["All Subcategories"] + sorted(self.filtered_df[self.actual_subcategory_col].unique().tolist())
        
        self.subcategory_dropdown["values"] = subcategories
        self.subcategory_var.set("All Subcategories")
        
        # Apply filters after updating
        self.apply_filters()

    def apply_filters(self, event=None):
        """Apply all active filters to the data"""
        if self.df is None:
            return

        data = self.df.copy()

        # Apply category filter
        selected_category = self.category_var.get()
        if selected_category != "All Categories":
            data = data[data["Category"] == selected_category]

        # Apply subcategory filter (only if category is selected)
        selected_subcategory = self.subcategory_var.get()
        if selected_subcategory != "All Subcategories" and selected_category != "All Categories":
            data = data[data[self.actual_subcategory_col] == selected_subcategory]

        # Apply requester filter
        selected_requester = self.requester_var.get()
        if selected_requester != "All Requesters" and self.requester_col:
            data = data[data[self.requester_col] == selected_requester]

        # Apply technician filter
        selected_technician = self.technician_var.get()
        if selected_technician != "All Technicians" and self.technician_col:
            data = data[data[self.technician_col] == selected_technician]

        # Summarize filtered data for display
        if len(data) > 0:
            filtered_summary = (
                data.groupby([self.actual_subcategory_col, "Category"])
                .size()
                .reset_index(name="Count")
            )
        else:
            filtered_summary = pd.DataFrame()

        self.update_table(filtered_summary)
        self.plot_category_pie(filtered_summary)

    def clear_filters(self):
        self.category_var.set("All Categories")
        self.subcategory_var.set("All Subcategories")
        self.subcategory_dropdown.config(state="readonly")
        self.requester_var.set("All Requesters")
        self.technician_var.set("All Technicians")
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
        """Plot pie chart showing category or subcategory breakdown based on filters"""
        self.ax.clear()

        if data is None or len(data) == 0:
            self.ax.text(0.5, 0.5, 'No data to display', ha='center', va='center', fontsize=11, color='gray')
            self.canvas.draw()
            return

        # Determine what to display based on current filters
        selected_category = self.category_var.get()
        
        # If a specific category is selected, show subcategory breakdown
        if selected_category != "All Categories":
            # Filter data for selected category
            category_data = data[data["Category"] == selected_category]
            
            if len(category_data) == 0:
                self.ax.text(0.5, 0.5, 'No data for selected category', ha='center', va='center', fontsize=11, color='gray')
                self.canvas.draw()
                return
            
            # Group by subcategory
            subcategory_counts = category_data.groupby(self.actual_subcategory_col)["Count"].sum()
            percentages = (subcategory_counts / subcategory_counts.sum()) * 100
            
            wedges, texts, autotexts = self.ax.pie(
                percentages,
                labels=subcategory_counts.index,
                autopct="%1.1f%%",
                startangle=140,
                colors=plt.cm.Set3(range(len(subcategory_counts))),
                wedgeprops={'edgecolor': 'white', 'linewidth': 2}
            )
            
            for text in texts:
                text.set_fontsize(8)
            for autotext in autotexts:
                autotext.set_color("black")
                autotext.set_fontsize(7)
                autotext.set_fontweight("bold")
            
            title_parts = [f"Subcategory Breakdown - {selected_category}"]
            title_parts.append(f"({int(subcategory_counts.sum())} tickets)")
            self.ax.set_title(" ".join(title_parts), fontsize=10, fontweight="bold")
        else:
            # Show category breakdown
            category_counts = data.groupby("Category")["Count"].sum()
            percentages = (category_counts / category_counts.sum()) * 100

            wedges, texts, autotexts = self.ax.pie(
                percentages,
                labels=category_counts.index,
                autopct="%1.1f%%",
                startangle=140,
                colors=[CATEGORY_COLORS.get(cat, "lightgrey") for cat in category_counts.index],
                wedgeprops={'edgecolor': 'white', 'linewidth': 2}
            )

            for text in texts:
                text.set_fontsize(8)
            for autotext in autotexts:
                autotext.set_color("white")
                autotext.set_fontsize(7)
                autotext.set_fontweight("bold")

            self.ax.set_title(f"Category Distribution ({int(category_counts.sum())} tickets)", fontsize=10, fontweight="bold")
        
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

    def show_technician_stats(self):
        """Show and export technician statistics by category"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            self.status_label.config(text="⚠ No data loaded")
            return

        # Find technician column
        technician_col = None
        for col in self.df.columns:
            if "technician" in col.lower() or "assigned_to" in col.lower():
                technician_col = col
                break

        if technician_col is None:
            messagebox.showerror("Error", "No 'Technician' or 'Assigned To' column found in the data!")
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
            elements.append(Paragraph("Technician Ticket Stats by Category", styles['Title']))
            elements.append(Spacer(1, 12))

            # For each category, show technician stats grouped by subcategory
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
                        ua_df.groupby([technician_col, self.actual_subcategory_col, "System"])
                        .size()
                        .reset_index(name="Ticket Count")
                        .sort_values("Ticket Count", ascending=False)
                    )

                    # Table header
                    table_data = [[technician_col, self.actual_subcategory_col, "System", "Ticket Count"]]
                    table_data += stats_df.values.tolist()

                else:
                    # Default grouping for other categories
                    stats_df = (
                        cat_df.groupby([technician_col, self.actual_subcategory_col])
                        .size()
                        .reset_index(name="Ticket Count")
                        .sort_values("Ticket Count", ascending=False)
                    )

                    # Table header
                    table_data = [[technician_col, self.actual_subcategory_col, "Ticket Count"]]
                    table_data += stats_df.values.tolist()

                if stats_df.empty:
                    elements.append(Paragraph("No tickets for this category.", styles['Normal']))
                    elements.append(Spacer(1, 8))
                    continue

                # Build table
                table = Table(table_data)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#70AD47")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 16))

            doc.build(elements)
            self.status_label.config(text="✓ Technician stats PDF generated")
            messagebox.showinfo("Success", f"Technician stats PDF exported to {save_path}")

        except Exception as e:
            self.status_label.config(text="✗ Error generating technician stats")
            messagebox.showerror("Error", f"Failed to export technician stats PDF: {e}")

    def show_additional_reports_dialog(self):
        """Open dialog to generate additional category-specific reports"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            self.status_label.config(text="⚠ No data loaded")
            return
        
        # First ask for save location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if not save_path:
            return
        
        # Then show dialog to select which reports to generate
        dialog = AdditionalPDFDialog(self.root)
        self.root.wait_window(dialog)
        selected_reports = dialog.selected
        
        # Generate the selected reports
        if selected_reports:
            self.generate_additional_reports(selected_reports, save_path)

    def export_pdf(self):
        if self.filtered_df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            self.status_label.config(text="⚠ No data loaded")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if not save_path:
            return

        try:
            self.status_label.config(text="📊 Generating PDF report...")
            self.root.update()
            
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

            self.status_label.config(text=f"✓ Main report saved: {os.path.basename(save_path)}")
            messagebox.showinfo("Success", f"PDF report exported successfully!")

            # Prompt for additional PDF reports
            dialog = AdditionalPDFDialog(self.root)
            self.root.wait_window(dialog)
            selected_reports = dialog.selected

            # Generate additional PDFs if the functions exist
            if selected_reports:
                self.generate_additional_reports(selected_reports, save_path)

        except Exception as e:
            self.status_label.config(text=f"✗ Error exporting PDF")
            messagebox.showerror("Error", f"Failed to export PDF: {e}")

    def generate_additional_reports(self, selected_reports, save_path):
        """Generate additional specialized reports"""
        try:
            self.status_label.config(text="📄 Generating category reports...")
            self.root.update()
            
            category_map = {
                "acsservicerepair": ("ACs Service/Repair", generate_acsservicerepair_pdf),
                "antivirus": ("Antivirus", generate_antivirus_pdf),
                "bts": ("BTS", generate_bts_pdf),
                "channels": ("Channels", generate_channels_pdf),
                "computerstationery": ("Computer Stationery", generate_computerstationery_pdf),
                "crm": ("CRM", generate_crm_pdf),
                "datacenter": ("Data Center", generate_datacenter_pdf),
                "databases": ("Databases", generate_databases_pdf),
                "deskhardwaresoftware": ("Desktop Hardware/Software", generate_deskhardwaresoftware_pdf),
                "email": ("Email", generate_email_pdf),
                "esb": ("Enterprise Service Bus (ESB)", generate_esb_pdf),
                "eod": ("EOD", generate_eod_pdf),
                "equinox": ("Equinox", generate_equinox_pdf),
                "generalrepairs": ("General Repairs", generate_generalrepairs_pdf),
                "generator": ("Generator", generate_generator_pdf),
                "helpdesk": ("Helpdesk", generate_helpdesk_pdf),
                "iapply": ("Iapply", generate_iapply_pdf),
                "informationprovision": ("Information Provision", generate_informationprovision_pdf),
                "internet": ("Internet", generate_internet_pdf),
                "intranet": ("Intranet/Corporate Website", generate_intranet_pdf),
                "iprofits": ("iProfits", generate_iprofits_pdf),
                "iprs": ("IPRS", generate_iprs_pdf),
                "jira": ("JIRA", generate_jira_pdf),
                "kprinter": ("K-Printer", generate_kprinter_pdf),
                "kplc": ("KPLC", generate_kplc_pdf),
                "licences": ("Licences", generate_licences_pdf),
                "lift": ("Lift", generate_lift_pdf),
                "marketplace": ("Market Place and Lead Management", generate_marketplace_pdf),
                "mips": ("MIPS", generate_mips_pdf),
                "moneygram": ("Money Gram", generate_moneygram_pdf),
                "network": ("Network", generate_network_pdf),
                "notifications": ("Notification and Alerts", generate_notifications_pdf),
                "onlinebiddbond": ("Online Bid Bond", generate_onlinebiddbond_pdf),
                "printedstationery": ("Printed Stationery", generate_printedstationery_pdf),
                "printers": ("Printers", generate_printers_pdf),
                "productionscheduler": ("Production Scheduler Run", generate_productionscheduler_pdf),
                "profits": ("Profits", generate_profits_pdf),
                "recordsmanagement": ("Records & Case Management", generate_recordsmanagement_pdf),
                "reports": ("Reports", generate_reports_pdf),
                "sapbi": ("SAP BI", generate_sapbi_pdf),
                "sapfi": ("SAP Fi", generate_sapfi_pdf),
                "swift": ("SWIFT", generate_swift_pdf),
                "taxiscabs": ("Taxis/Cabs", generate_taxiscabs_pdf),
                "teammate": ("Team Mate", generate_teammate_pdf),
                "telephony": ("Telephony", generate_telephony_pdf),
                "testenvironment": ("Test Environment", generate_testenvironment_pdf),
                "ups": ("UPS", generate_ups_pdf),
                "useradmin": ("User Administration", generate_useradmin_pdf),
                "vehicle": ("Vehicle", generate_vehicle_pdf),
                "vendorsupport": ("Vendor Support", generate_vendorsupport_pdf),
                "water": ("Water", generate_water_pdf),
                "westernunion": ("Western Union", generate_westernunion_pdf),
                "zendesk": ("Zendesk", generate_zendesk_pdf),
            }
            
            generated_count = 0
            skipped_count = 0
            failed_count = 0
            skipped_reports = []
            
            # Generate reports for each selected category
            for report_key in selected_reports:
                if report_key not in category_map:
                    failed_count += 1
                    continue
                    
                category_name, generate_func = category_map[report_key]
                
                try:
                    # Filter data for this category
                    df = self.df[self.df["Category"] == category_name]
                    
                    # Check if category has any tickets
                    if df.empty:
                        skipped_count += 1
                        skipped_reports.append(category_name)
                        continue
                    
                    # Create temp CSV
                    temp_csv = os.path.join(tempfile.gettempdir(), f"{report_key}_data.csv")
                    df.to_csv(temp_csv, index=False)
                    
                    # Generate the PDF report
                    generate_func(os.path.dirname(save_path), temp_csv)
                    generated_count += 1
                    messagebox.showinfo("Success", f"✓ {category_name} PDF generated successfully")
                    
                    # Cleanup temp CSV
                    if os.path.exists(temp_csv):
                        os.remove(temp_csv)
                        
                except Exception as e:
                    failed_count += 1
                    messagebox.showerror("Error", f"Failed to generate {category_name} report:\n{str(e)}")
            
            # Final summary
            summary_msg = f"Report Generation Summary:\n\n✓ Generated: {generated_count}\n✗ Failed: {failed_count}\n⊘ Skipped (No tickets): {skipped_count}"
            if skipped_reports:
                summary_msg += f"\n\nSkipped Reports:\n" + "\n".join([f"  • {r}" for r in skipped_reports])
            summary_msg += "\n\nAll selected reports have been processed."
            
            self.status_label.config(text=f"✓ Generated {generated_count} reports | {skipped_count} skipped")
            messagebox.showinfo("Summary", summary_msg)
                    
        except Exception as e:
            self.status_label.config(text="✗ Error generating reports")
            messagebox.showerror("Error", f"Failed to generate additional reports: {e}")


class AdditionalPDFDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Generate Additional PDF Reports")
        self.selected = []
        self.vars = []
        
        # All available report options
        options = [
            ("ACs Service/Repair Report", "acsservicerepair"),
            ("Antivirus Report", "antivirus"),
            ("BTS Report", "bts"),
            ("Channels Report", "channels"),
            ("Computer Stationery Report", "computerstationery"),
            ("CRM Report", "crm"),
            ("Data Center Report", "datacenter"),
            ("Databases Report", "databases"),
            ("Desktop Hardware/Software Report", "deskhardwaresoftware"),
            ("Email Report", "email"),
            ("Enterprise Service Bus Report", "esb"),
            ("EOD Report", "eod"),
            ("Equinox Report", "equinox"),
            ("General Repairs Report", "generalrepairs"),
            ("Generator Report", "generator"),
            ("Helpdesk Report", "helpdesk"),
            ("iApply Report", "iapply"),
            ("Information Provision Report", "informationprovision"),
            ("Internet Report", "internet"),
            ("Intranet/Corporate Website Report", "intranet"),
            ("iProfits Report", "iprofits"),
            ("IPRS Report", "iprs"),
            ("JIRA Report", "jira"),
            ("K-Printer Report", "kprinter"),
            ("KPLC Report", "kplc"),
            ("Licences Report", "licences"),
            ("Lift Report", "lift"),
            ("Market Place/Lead Management Report", "marketplace"),
            ("MIPS Report", "mips"),
            ("Money Gram Report", "moneygram"),
            ("Network Report", "network"),
            ("Notifications/Alerts Report", "notifications"),
            ("Online Bid Bond Report", "onlinebiddbond"),
            ("Printed Stationery Report", "printedstationery"),
            ("Printers Report", "printers"),
            ("Production Scheduler Report", "productionscheduler"),
            ("Profits Report", "profits"),
            ("Records & Case Management Report", "recordsmanagement"),
            ("Reports Report", "reports"),
            ("SAP BI Report", "sapbi"),
            ("SAP Fi Report", "sapfi"),
            ("SWIFT Report", "swift"),
            ("Taxis/Cabs Report", "taxiscabs"),
            ("Team Mate Report", "teammate"),
            ("Telephony Report", "telephony"),
            ("Test Environment Report", "testenvironment"),
            ("UPS Report", "ups"),
            ("User Administration Report", "useradmin"),
            ("Vehicle Report", "vehicle"),
            ("Vendor Support Report", "vendorsupport"),
            ("Water Report", "water"),
            ("Western Union Report", "westernunion"),
            ("Zendesk Report", "zendesk"),
        ]
        
        tk.Label(self, text="Select additional PDF reports to generate:").pack(pady=10)
        
        # Create a scrollable frame for checkboxes
        from tkinter import Canvas
        canvas = Canvas(self, height=300, width=400)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        for label, key in options:
            var = BooleanVar()
            chk = Checkbutton(scrollable_frame, text=label, variable=var)
            chk.pack(anchor="w", padx=10)
            self.vars.append((key, var))
        
        canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y")
        
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