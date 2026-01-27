import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import smtplib
import schedule
import threading
import time
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import json
import re
import webbrowser
from tkinter import scrolledtext
import sys
import keyring  # For secure credential storage
import uuid

class JobApplicationSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Application Assistant")
        self.root.geometry("1200x800")
        self.root.configure(bg='white')
        
        # Generate a unique app ID for this installation
        self.app_id = self.get_or_create_app_id()
        
        # Set application icon
        self.set_application_icon()
        
        # Clean color scheme
        self.colors = {
            "bg": "#ffffff",
            "bg_light": "#f5f7fa",
            "bg_dark": "#e8ecf1",
            "primary": "#2c3e50",
            "secondary": "#3498db",
            "accent": "#2980b9",
            "success": "#27ae60",
            "warning": "#f39c12",
            "error": "#e74c3c",
            "text": "#2c3e50",
            "text_light": "#7f8c8d",
            "border": "#bdc3c7"
        }
        
        # Fonts
        self.fonts = {
            "title": ("Segoe UI", 20, "bold"),
            "heading": ("Segoe UI", 11, "bold"),
            "normal": ("Segoe UI", 10),
            "small": ("Segoe UI", 9)
        }
        
        # Data storage
        self.recipients = []
        self.attachments = []  # Store all attachments without duplicates
        self.sent_emails = []  # Track sent emails
        self.selected_recipients = []  # For multi-select
        
        # FIX: Use user's home directory for settings files to avoid permission issues
        home_dir = os.path.expanduser("~")
        app_data_dir = os.path.join(home_dir, "JobApplicationAssistant")
        
        # Create directory if it doesn't exist
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
        
        # Note: We're NOT storing passwords in these files anymore
        self.settings_file = os.path.join(app_data_dir, "job_app_settings.json")
        self.history_file = os.path.join(app_data_dir, "sent_history.csv")
        self.drafts_file = os.path.join(app_data_dir, "email_drafts.json")
        self.recipients_file = os.path.join(app_data_dir, "recipients_data.json")  # NEW: Recipients storage file
        
        # Connection status
        self.connection_status_var = None
        self.connection_status_label = None
        
        # FIX: Initialize tkinter variables BEFORE loading settings
        self.initialize_variables()
        
        # Create GUI
        self.create_widgets()
        
        # Now load settings AFTER widgets are created
        self.load_settings()
        self.load_drafts()
        self.load_recipients_data()  # NEW: Load saved recipients
        self.load_history()  # NEW: Load history immediately
        
        # Start scheduler thread
        self.scheduler_running = True
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()
        
        # Bind window close event to save data
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_or_create_app_id(self):
        """Get or create a unique app ID for this installation"""
        try:
            # Try to get existing app ID from Windows Credential Manager
            existing_id = keyring.get_password("JobAppAssistant", "app_id")
            if existing_id:
                return existing_id
            
            # Create new unique ID
            new_id = str(uuid.uuid4())
            keyring.set_password("JobAppAssistant", "app_id", new_id)
            return new_id
        except Exception as e:
            # Fallback to generating a simple ID
            import hashlib
            import socket
            import getpass
            # Create ID based on username and hostname
            user_info = f"{getpass.getuser()}@{socket.gethostname()}"
            return hashlib.sha256(user_info.encode()).hexdigest()[:32]

    def save_credentials_to_vault(self, email, password):
        """Securely save email credentials to Windows Credential Manager"""
        try:
            # Use app_id + email as service name to make it unique to this installation
            service_name = f"JobAppAssistant_{self.app_id}"
            
            # Save to Windows Credential Manager
            keyring.set_password(service_name, email, password)
            
            # Also store email in settings (without password)
            self.save_email_to_settings(email)
            
            print(f"✅ Credentials saved securely for: {email}")
            return True
        except Exception as e:
            print(f"❌ Error saving credentials: {e}")
            messagebox.showerror("Error", f"Could not save credentials securely: {str(e)}")
            return False

    def get_credentials_from_vault(self, email):
        """Get email credentials from Windows Credential Manager"""
        try:
            service_name = f"JobAppAssistant_{self.app_id}"
            password = keyring.get_password(service_name, email)
            
            if password:
                print(f"✅ Retrieved credentials for: {email}")
                return password
            else:
                print(f"⚠️ No saved credentials found for: {email}")
                return None
        except Exception as e:
            print(f"❌ Error retrieving credentials: {e}")
            return None

    def delete_credentials_from_vault(self, email):
        """Delete email credentials from Windows Credential Manager"""
        try:
            service_name = f"JobAppAssistant_{self.app_id}"
            keyring.delete_password(service_name, email)
            print(f"✅ Deleted credentials for: {email}")
            return True
        except Exception as e:
            print(f"❌ Error deleting credentials: {e}")
            return False

    def save_email_to_settings(self, email):
        """Save email (without password) to settings file"""
        try:
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
            
            settings['sender_email'] = email
            settings['smtp_server'] = self.smtp_server_var.get()
            settings['smtp_port'] = self.smtp_port_var.get()
            settings['sender_name'] = self.your_name_var.get()
            settings['sender_title'] = self.your_title_var.get()
            settings['send_time'] = self.send_time_var.get()
            settings['interval_days'] = self.interval_days_var.get()
            settings['max_resends'] = self.max_resends_var.get()
            settings['show_password'] = self.show_password_var.get()
            
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
            
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            
            print(f"✅ Settings saved (without password) for: {email}")
            return True
            
        except Exception as e:
            print(f"❌ Error saving settings: {e}")
            return False

    def initialize_variables(self):
        """Initialize all tkinter variables BEFORE creating widgets"""
        # Email settings variables
        self.sender_email_var = tk.StringVar()
        self.sender_pass_var = tk.StringVar()
        self.your_name_var = tk.StringVar()
        self.your_title_var = tk.StringVar()
        self.smtp_server_var = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port_var = tk.StringVar(value="587")
        self.send_time_var = tk.StringVar(value="09:00")
        self.interval_days_var = tk.StringVar(value="1")  # Changed from interval_var
        self.max_resends_var = tk.StringVar(value="1")  # NEW: Maximum number of resends
        self.show_password_var = tk.BooleanVar(value=False)
        
        # Recipient tab variables
        self.company_var = tk.StringVar()
        self.hr_name_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.position_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        
        # Email template variables
        self.subject_var = tk.StringVar(value="Application for {position} position at {company}")
        
        # Send tab variables
        self.send_mode_var = tk.StringVar(value="send")
        self.send_status_var = tk.StringVar(value="No recipients selected")
        
        # Connection status
        self.connection_status_var = tk.StringVar(value="❌ Not configured")

    def set_application_icon(self):
        """Set the application icon for the window and taskbar"""
        try:
            # Try multiple paths for the icon
            icon_paths = [
                r"C:\Users\RiDDL3TeCH\Desktop\JobApplicationAssistant\JobApplicationAssistant.png",
                "JobApplicationAssistant.png",
                os.path.join(os.path.dirname(__file__), "JobApplicationAssistant.png")
            ]
            
            icon_found = False
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    try:
                        # Load the image and set as icon
                        icon_image = tk.PhotoImage(file=icon_path)
                        self.root.iconphoto(True, icon_image)
                        self.root.iconbitmap(icon_path)  # Alternative method for .ico files
                        print(f"Icon loaded from: {icon_path}")
                        icon_found = True
                        break
                    except Exception as e:
                        print(f"Could not load icon from {icon_path}: {e}")
                        continue
            
            if not icon_found:
                print("Icon not found. Using default icon.")
                
        except Exception as e:
            print(f"Error setting application icon: {e}")

    def on_closing(self):
        """Save all data before closing the application"""
        try:
            self.save_recipients_data()
            self.log_message("Application closing - data saved")
            print("✅ Data saved before closing")
        except Exception as e:
            print(f"❌ Error saving data before closing: {e}")
        
        # Stop scheduler thread
        self.scheduler_running = False
        
        # Destroy the window
        self.root.destroy()

    def save_recipients_data(self):
        """Save recipients data to JSON file"""
        try:
            # Prepare recipients data for saving
            recipients_data = []
            for recipient in self.recipients:
                recipients_data.append({
                    'company': recipient.get('company', ''),
                    'hr_name': recipient.get('hr_name', ''),
                    'email': recipient.get('email', ''),
                    'position': recipient.get('position', ''),
                    'notes': recipient.get('notes', ''),
                    'status': recipient.get('status', 'PENDING'),
                    'last_sent': recipient.get('last_sent', ''),
                    'added_date': recipient.get('added_date', ''),
                    'send_count': recipient.get('send_count', 0),
                    'max_sends': recipient.get('max_sends', 1),
                    'stop_resend': recipient.get('stop_resend', False),  # NEW: Save stop resend flag
                    'email_template': recipient.get('email_template', ''),  # NEW: Save email template
                    'email_subject': recipient.get('email_subject', ''),  # NEW: Save email subject
                    'attachments': recipient.get('attachments', [])  # NEW: Save attachments for this recipient
                })
            
            # Save to JSON file
            with open(self.recipients_file, 'w') as f:
                json.dump(recipients_data, f, indent=2)
            
            print(f"✅ Recipients data saved: {len(recipients_data)} recipients")
            return True
        except Exception as e:
            print(f"❌ Error saving recipients data: {e}")
            return False

    def load_recipients_data(self):
        """Load recipients data from JSON file"""
        try:
            if os.path.exists(self.recipients_file):
                with open(self.recipients_file, 'r') as f:
                    recipients_data = json.load(f)
                
                # Clear current recipients
                self.recipients.clear()
                
                # Load recipients from file
                for recipient_data in recipients_data:
                    recipient = {
                        "company": recipient_data.get('company', ''),
                        "hr_name": recipient_data.get('hr_name', ''),
                        "email": recipient_data.get('email', ''),
                        "position": recipient_data.get('position', ''),
                        "notes": recipient_data.get('notes', ''),
                        "status": recipient_data.get('status', 'PENDING'),
                        "last_sent": recipient_data.get('last_sent', ''),
                        "added_date": recipient_data.get('added_date', datetime.now().strftime("%Y-%m-%d")),
                        "send_count": recipient_data.get('send_count', 0),
                        "max_sends": recipient_data.get('max_sends', 1),
                        "stop_resend": recipient_data.get('stop_resend', False),  # NEW: Load stop resend flag
                        "email_template": recipient_data.get('email_template', ''),  # NEW: Load email template
                        "email_subject": recipient_data.get('email_subject', ''),  # NEW: Load email subject
                        "attachments": recipient_data.get('attachments', [])  # NEW: Load attachments for this recipient
                    }
                    self.recipients.append(recipient)
                    
                    # Add to treeview with status showing send count and stop status
                    if recipient['stop_resend']:
                        status_display = f"⏹️ STOPPED ({recipient['send_count']}/{recipient['max_sends']})"
                    else:
                        status_display = f"{recipient['status']} ({recipient['send_count']}/{recipient['max_sends']})"
                    
                    self.recipient_tree.insert("", tk.END, values=(
                        "",  # Empty checkbox (not selected by default)
                        recipient['company'],
                        recipient['hr_name'],
                        recipient['email'],
                        recipient['position'],
                        status_display,
                        recipient['added_date'],
                        "⏹️" if recipient['stop_resend'] else ""
                    ))
                
                print(f"✅ Recipients data loaded: {len(self.recipients)} recipients")
                self.log_message(f"Loaded {len(self.recipients)} recipients from saved data")
                
                # Update the selected listbox
                self.update_selected_listbox()
                
                return True
            else:
                print("ℹ️ No recipients data file found. Starting fresh.")
                return False
        except Exception as e:
            print(f"❌ Error loading recipients data: {e}")
            self.log_message(f"Error loading recipients: {str(e)}")
            return False

    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self.root, bg=self.colors["primary"], height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Title with icon
        title_frame = tk.Frame(header_frame, bg=self.colors["primary"])
        title_frame.pack(side=tk.LEFT, padx=30, pady=20)
        
        title_label = tk.Label(title_frame,
                             text="📧 Job Application Assistant",
                             font=self.fonts["title"],
                             bg=self.colors["primary"],
                             fg='white')
        title_label.pack(side=tk.LEFT)
        
        # Version label
        version_label = tk.Label(title_frame,
                               text="v1.0",
                               font=("Segoe UI", 8),
                               bg=self.colors["primary"],
                               fg="#95a5a6")
        version_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Status indicators
        status_frame = tk.Frame(header_frame, bg=self.colors["primary"])
        status_frame.pack(side=tk.RIGHT, padx=30, pady=20)
        
        self.status_indicator = tk.Label(status_frame,
                                       text="●",
                                       font=("Segoe UI", 20),
                                       bg=self.colors["primary"],
                                       fg=self.colors["success"])
        self.status_indicator.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.status_label = tk.Label(status_frame,
                                   text="READY",
                                   font=self.fonts["heading"],
                                   bg=self.colors["primary"],
                                   fg='white')
        self.status_label.pack(side=tk.RIGHT)
        
        # Main container
        main_container = tk.Frame(self.root, bg='white')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Notebook
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.create_recipient_tab()
        self.create_email_tab()
        self.create_attachment_tab()
        self.create_schedule_tab()  # Settings tab
        self.create_send_tab()
        self.create_history_tab()
        self.create_developer_tab()

    def create_card_frame(self, parent, title):
        frame = tk.Frame(parent, bg='white', highlightbackground=self.colors["border"],
                        highlightthickness=1, relief="solid")
        
        # Title
        title_frame = tk.Frame(frame, bg='white')
        title_frame.pack(fill=tk.X, padx=15, pady=10)
        
        title_label = tk.Label(title_frame,
                             text=title,
                             font=self.fonts["heading"],
                             bg='white',
                             fg=self.colors["primary"])
        title_label.pack(side=tk.LEFT)
        
        return frame

    def create_scrollable_frame(self, parent):
        """Create a scrollable frame with canvas and scrollbar"""
        # Create a canvas and scrollbar
        canvas = tk.Canvas(parent, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        
        # Create a frame inside the canvas
        scrollable_frame = tk.Frame(canvas, bg='white')
        
        # Configure the canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create a window in the canvas for the scrollable frame
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Configure the scrollable frame
        def configure_scrollable_frame(event):
            # Update the scrollregion to encompass the inner frame
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Make the canvas window the same width as the canvas
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())
        
        # Bind events
        scrollable_frame.bind("<Configure>", configure_scrollable_frame)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_frame, width=e.width))
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        return scrollable_frame, canvas, scrollbar

    def create_recipient_tab(self):
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📋 Recipients")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        # Main container for content
        content_frame = tk.Frame(scrollable_tab, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left panel - Add recipient
        left_panel = tk.Frame(content_frame, bg='white')
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        add_frame = self.create_card_frame(left_panel, "➕ Add New Recipient")
        add_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Form
        form_frame = tk.Frame(add_frame, bg='white')
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        fields = [
            ("Company Name*:", "company_var"),
            ("HR/Manager Name:", "hr_name_var"),
            ("Email Address*:", "email_var"),
            ("Job Position:", "position_var"),
            ("Notes:", "notes_var")
        ]
        
        for i, (label, var_name) in enumerate(fields):
            tk.Label(form_frame,
                    text=label,
                    bg='white',
                    fg=self.colors["text"],
                    font=self.fonts["normal"]).grid(row=i, column=0, sticky=tk.W, pady=8)
            
            # Get the already initialized variable
            var = getattr(self, var_name)
            
            entry = tk.Entry(form_frame,
                           textvariable=var,
                           bg='white',
                           fg=self.colors["text"],
                           font=self.fonts["normal"],
                           relief="solid",
                           borderwidth=1)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=8, padx=(10, 0), ipady=5)
        
        # Configure grid weight
        form_frame.columnconfigure(1, weight=1)
        
        # Add button
        add_btn = tk.Button(form_frame,
                          text="➕ Add Recipient",
                          command=self.add_recipient,
                          bg=self.colors["secondary"],
                          fg='white',
                          font=self.fonts["normal"],
                          relief="flat",
                          cursor="hand2",
                          padx=20,
                          pady=8)
        add_btn.grid(row=5, column=1, sticky=tk.E, pady=20)
        
        # Right panel - Recipients list
        right_panel = tk.Frame(content_frame, bg='white')
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create a main frame for the right panel with proper layout
        right_main_frame = tk.Frame(right_panel, bg='white')
        right_main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Recipients list card
        list_frame = self.create_card_frame(right_main_frame, "📋 Recipients List")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Treeview with checkboxes - ADDED "Stop Resend" column
        columns = ("✓", "Company", "Contact", "Email", "Position", "Status", "Added Date", "⏹️ Stop")
        self.recipient_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        # Style the treeview
        style = ttk.Style()
        style.configure("Treeview",
                       background='white',
                       foreground=self.colors["text"],
                       fieldbackground='white',
                       rowheight=25)
        style.configure("Treeview.Heading",
                       background=self.colors["bg_light"],
                       foreground=self.colors["primary"],
                       relief="flat",
                       font=self.fonts["heading"])
        
        # Define columns
        for col in columns:
            self.recipient_tree.heading(col, text=col)
            if col == "✓":
                self.recipient_tree.column(col, width=50, anchor='center')
            elif col == "Company":
                self.recipient_tree.column(col, width=150)
            elif col == "Contact":
                self.recipient_tree.column(col, width=120)
            elif col == "Email":
                self.recipient_tree.column(col, width=200)
            elif col == "Position":
                self.recipient_tree.column(col, width=150)
            elif col == "Status":
                self.recipient_tree.column(col, width=100)
            elif col == "Added Date":
                self.recipient_tree.column(col, width=100)
            elif col == "⏹️ Stop":
                self.recipient_tree.column(col, width=70, anchor='center')
        
        self.recipient_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add checkbox functionality - FIXED: Only toggle when clicking the checkbox column
        self.recipient_tree.bind('<Button-1>', self.on_recipient_click)
        
        # Scrollbar for treeview
        tree_scrollbar = ttk.Scrollbar(list_frame, command=self.recipient_tree.yview)
        self.recipient_tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Action buttons at the bottom - OUTSIDE the recipients list card
        # Create a separate container for buttons
        button_container = tk.Frame(right_main_frame, bg='white')
        button_container.pack(fill=tk.X, pady=(0, 10))
        
        # Add visual separator
        separator = tk.Frame(button_container, bg=self.colors["border"], height=1)
        separator.pack(fill=tk.X, pady=(0, 15))
        
        # Label to indicate horizontal scrolling
        scroll_label = tk.Label(button_container,
                              text="➖➡️ Use horizontal scrollbar below to access all buttons",
                              bg='white',
                              fg=self.colors["text_light"],
                              font=("Segoe UI", 9, "bold"))
        scroll_label.pack(pady=(0, 10))
        
        # Create a frame for the horizontal scrollbar and buttons
        button_outer_frame = tk.Frame(button_container, bg='white')
        button_outer_frame.pack(fill=tk.X, expand=True)
        
        # Create a canvas for horizontal scrolling of buttons
        btn_canvas = tk.Canvas(button_outer_frame, bg='white', height=50, highlightthickness=0)
        btn_horizontal_scrollbar = tk.Scrollbar(button_outer_frame, orient="horizontal", command=btn_canvas.xview)
        
        btn_canvas.configure(xscrollcommand=btn_horizontal_scrollbar.set)
        btn_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)
        btn_horizontal_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Create a frame inside the canvas for buttons
        btn_frame = tk.Frame(btn_canvas, bg='white')
        btn_canvas.create_window((0, 0), window=btn_frame, anchor="nw")
        
        # Configure the button frame to update canvas scrollregion
        def configure_btn_frame(event):
            btn_canvas.configure(scrollregion=btn_canvas.bbox("all"))
        
        btn_frame.bind("<Configure>", configure_btn_frame)
        
        buttons = [
            ("📥 Import CSV", self.import_csv, self.colors["secondary"]),
            ("📤 Export CSV", self.export_csv, self.colors["accent"]),
            ("✓ Select All", self.select_all_recipients, self.colors["primary"]),
            ("✗ Deselect All", self.deselect_all_recipients, self.colors["warning"]),
            ("⏹️ Stop Resend", self.stop_resend_selected, self.colors["error"]),
            ("▶️ Resume Resend", self.resume_resend_selected, self.colors["success"]),
            ("🗑️ Remove", self.remove_recipient, self.colors["error"]),
            ("🔥 Clear All", self.clear_recipients, self.colors["error"]),
            ("💾 Save Now", self.save_recipients_data, self.colors["success"])
        ]
        
        for text, command, color in buttons:
            btn = tk.Button(btn_frame,
                          text=text,
                          command=command,
                          bg=color,
                          fg='white',
                          font=self.fonts["normal"],
                          relief="flat",
                          padx=20,
                          pady=8,
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Center the buttons in the canvas
        btn_frame.update_idletasks()
        btn_canvas.configure(scrollregion=btn_canvas.bbox("all"))
        
        # Add help text
        help_text = tk.Label(button_container,
                           text="Tip: Click on the checkbox (✓) column to select/deselect recipients | Click on ⏹️ to stop/resume resend",
                           bg='white',
                           fg=self.colors["text_light"],
                           font=("Segoe UI", 8))
        help_text.pack(pady=(10, 0))

    def on_recipient_click(self, event):
        """Handle clicks on recipient tree to toggle selection - FIXED: Only toggle checkbox column"""
        region = self.recipient_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.recipient_tree.identify_column(event.x)
            item = self.recipient_tree.identify_row(event.y)
            
            # Get column index
            col_index = int(column.replace("#", "")) - 1
            
            # Only toggle if clicking on the first column (checkbox column) or stop column
            if col_index == 0:  # First column (Select/Checkbox column)
                values = self.recipient_tree.item(item, "values")
                current_value = values[0] if values else ""
                
                # Toggle selection
                if current_value == "✓":
                    new_value = ""  # Unchecked
                else:
                    new_value = "✓"  # Checked
                
                self.recipient_tree.set(item, column="✓", value=new_value)
                self.update_selected_listbox()
            
            elif col_index == 7:  # Stop Resend column (8th column)
                values = self.recipient_tree.item(item, "values")
                email = values[3] if len(values) > 3 else ""
                
                # Find recipient and toggle stop_resend
                for recipient in self.recipients:
                    if recipient['email'] == email:
                        recipient['stop_resend'] = not recipient['stop_resend']
                        
                        # Update treeview
                        if recipient['stop_resend']:
                            new_status = f"⏹️ STOPPED ({recipient['send_count']}/{recipient['max_sends']})"
                            self.recipient_tree.set(item, column="⏹️ Stop", value="⏹️")
                        else:
                            new_status = f"{recipient['status']} ({recipient['send_count']}/{recipient['max_sends']})"
                            self.recipient_tree.set(item, column="⏹️ Stop", value="")
                        
                        self.recipient_tree.set(item, column="Status", value=new_status)
                        self.log_message(f"{'Stopped' if recipient['stop_resend'] else 'Resumed'} resend for {recipient['company']}")
                        break

    def create_email_tab(self):
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📝 Email Template")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        main_frame = self.create_card_frame(scrollable_tab, "📝 Email Template Editor")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Tags display
        tags_frame = tk.Frame(main_frame, bg=self.colors["bg_light"])
        tags_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tags_info = [
            ("{company}", "Company Name"),
            ("{hr_name}", "HR/Manager Name"),
            ("{position}", "Job Position"),
            ("{date}", "Current Date"),
            ("{your_name}", "Your Name"),
            ("{your_title}", "Your Title")
        ]
        
        tk.Label(tags_frame,
                text="🎯 Available tags: ",
                bg=self.colors["bg_light"],
                fg=self.colors["text_light"],
                font=self.fonts["normal"]).pack(side=tk.LEFT)
        
        for tag, desc in tags_info:
            tag_frame = tk.Frame(tags_frame, bg=self.colors["bg_light"])
            tag_frame.pack(side=tk.LEFT, padx=3)
            
            tag_label = tk.Label(tag_frame,
                               text=tag,
                               bg='white',
                               fg=self.colors["primary"],
                               font=self.fonts["small"],
                               padx=8,
                               pady=3,
                               relief="solid",
                               borderwidth=1)
            tag_label.pack()
            
            # Tooltip
            self.create_tooltip(tag_label, desc)
        
        # Subject field
        subject_frame = tk.Frame(main_frame, bg='white')
        subject_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(subject_frame,
                text="📌 Subject:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(side=tk.LEFT)
        
        subject_entry = tk.Entry(subject_frame,
                               textvariable=self.subject_var,
                               bg='white',
                               fg=self.colors["text"],
                               font=self.fonts["normal"],
                               width=60)
        subject_entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # Email body text area
        body_frame = tk.Frame(main_frame, bg='white')
        body_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(body_frame,
                text="📝 Email Body:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 5))
        
        # Use scrolledtext for better text editing
        self.body_text = scrolledtext.ScrolledText(body_frame,
                                                 bg='white',
                                                 fg=self.colors["text"],
                                                 font=self.fonts["normal"],
                                                 relief="solid",
                                                 borderwidth=1,
                                                 height=15,
                                                 wrap=tk.WORD)
        self.body_text.pack(fill=tk.BOTH, expand=True)
        
        # Default template
        default_template = """Dear {hr_name},

I am writing to express my interest in the {position} position at {company}. 

With my background and experience in [Your Field], I believe I would be a valuable addition to your team.

Attached are my resume and cover letter for your review. I would welcome the opportunity to discuss how my skills and experience align with the needs of your team.

Thank you for considering my application.

Best regards,
{your_name}
{your_title}"""
        
        self.body_text.insert("1.0", default_template)
        
        # Control buttons
        control_frame = tk.Frame(main_frame, bg='white')
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        
        control_buttons = [
            ("💾 Save Template", self.save_template, self.colors["secondary"]),
            ("📂 Load Template", self.load_template, self.colors["accent"]),
            ("👁️ Preview Email", self.preview_email, self.colors["primary"]),
            ("💾 Save as Draft", self.save_as_draft, self.colors["warning"]),
            ("📋 Load Draft", self.load_draft, self.colors["success"])
        ]
        
        for text, command, color in control_buttons:
            btn = tk.Button(control_frame,
                          text=text,
                          command=command,
                          bg=color,
                          fg='white',
                          font=self.fonts["normal"],
                          padx=15,
                          pady=8,
                          relief="flat",
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)

    def create_attachment_tab(self):
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📎 Attachments")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        # Main container
        main_container = tk.Frame(scrollable_tab, bg='white')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a single card frame for attachments
        attach_frame = self.create_card_frame(main_container, "📎 Current Attachments")
        attach_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top: List of attachments
        list_frame = tk.Frame(attach_frame, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(list_frame,
                text="📁 Selected Files:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 5))
        
        # Listbox with scrollbar
        listbox_container = tk.Frame(list_frame, bg='white')
        listbox_container.pack(fill=tk.BOTH, expand=True)
        
        self.attach_listbox = tk.Listbox(listbox_container,
                                       bg='white',
                                       fg=self.colors["text"],
                                       font=self.fonts["normal"],
                                       selectbackground=self.colors["bg_light"],
                                       relief="solid",
                                       borderwidth=1,
                                       height=8)
        self.attach_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        list_scrollbar = tk.Scrollbar(listbox_container, command=self.attach_listbox.yview)
        self.attach_listbox.configure(yscrollcommand=list_scrollbar.set)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preview area
        preview_frame = tk.Frame(attach_frame, bg='white')
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        tk.Label(preview_frame,
                text="📄 File Preview:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 5))
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame,
                                                    bg=self.colors["bg_light"],
                                                    fg=self.colors["text"],
                                                    font=self.fonts["normal"],
                                                    height=10,
                                                    wrap=tk.WORD,
                                                    state='disabled')
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        # Set initial preview message
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(1.0, "Select a file from the list above to preview its contents.")
        self.preview_text.config(state='disabled')
        
        # Bind listbox selection to preview
        self.attach_listbox.bind('<<ListboxSelect>>', self.show_file_preview)
        
        # Attachment buttons
        btn_frame = tk.Frame(attach_frame, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        attach_buttons = [
            ("📄 Add Resume", lambda: self.add_attachment("resume")),
            ("📝 Cover Letter", lambda: self.add_attachment("cover")),
            ("➕ Add File", self.add_attachment),
            ("👁️ Preview", self.preview_selected_file),
            ("➖ Remove", self.remove_attachment),
            ("🗑️ Clear All", self.clear_attachments)
        ]
        
        for text, command in attach_buttons:
            btn = tk.Button(btn_frame,
                          text=text,
                          command=command,
                          bg=self.colors["secondary"],
                          fg='white',
                          font=self.fonts["small"],
                          relief="flat",
                          padx=10,
                          pady=5,
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)

    def create_schedule_tab(self):
        """Settings tab with secure credential management"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="⚙️ Settings")
        
        # Create a canvas and scrollbar for the entire settings tab
        canvas = tk.Canvas(tab, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')
        
        # Configure the canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create a window in the canvas for the scrollable frame
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Configure the scrollable frame
        def configure_scrollable_frame(event):
            # Update the scrollregion to encompass the inner frame
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Make the canvas window the same width as the canvas
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())
        
        # Bind events
        scrollable_frame.bind("<Configure>", configure_scrollable_frame)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_frame, width=e.width))
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Main content frame inside the scrollable frame
        main_frame = self.create_card_frame(scrollable_frame, "⚙️ Email Account Settings (Secure)")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Security info
        security_frame = tk.Frame(main_frame, bg=self.colors["bg_light"])
        security_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(security_frame,
                text="🔒 Secure Credential Storage:",
                bg=self.colors["bg_light"],
                fg=self.colors["success"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        security_text = """✅ Your credentials are stored securely using Windows Credential Manager.

• Passwords are NEVER saved to files
• Each user has their own isolated credentials
• Credentials stay on YOUR PC only
• If you share the app, others won't see your credentials

📌 For Gmail: Use "App Password" (16 characters), NOT your regular password."""
        
        security_scroll = scrolledtext.ScrolledText(security_frame,
                                                   bg='white',
                                                   fg=self.colors["text"],
                                                   font=self.fonts["small"],
                                                   height=6,
                                                   wrap=tk.WORD,
                                                   state='disabled')
        security_scroll.pack(fill=tk.BOTH, expand=True, pady=5)
        
        security_scroll.config(state='normal')
        security_scroll.delete(1.0, tk.END)
        security_scroll.insert(1.0, security_text)
        security_scroll.config(state='disabled')
        
        # Gmail Configuration Guide
        guide_frame = tk.Frame(main_frame, bg=self.colors["bg_dark"])
        guide_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(guide_frame,
                text="🔧 Gmail SMTP Configuration Guide:",
                bg=self.colors["bg_dark"],
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        # Instructions in a scrollable text widget
        instructions_text = """📌 IMPORTANT: For Gmail accounts, you need to use an "App Password" instead of your regular password.

Follow these steps:
1. Enable 2-Step Verification on your Google Account:
   • Go to: https://myaccount.google.com/security
   • Turn on "2-Step Verification"

2. Generate an App Password:
   • Visit: https://myaccount.google.com/apppasswords
   • Select "Mail" as the app
   • Select "Other" as the device and name it "Job Application Assistant"
   • Click "Generate"
   • Copy the 16-character password

3. Use the App Password below:
   • Email: Your Gmail address
   • Password: The 16-character app password (NOT your regular password)

4. Test the connection using the "Test Connection" button below.

Note: Regular password authentication will NOT work due to Google's security policies."""
        
        instructions_scroll = scrolledtext.ScrolledText(guide_frame,
                                                       bg='white',
                                                       fg=self.colors["text"],
                                                       font=self.fonts["small"],
                                                       height=12,
                                                       wrap=tk.WORD,
                                                       state='disabled')
        instructions_scroll.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Enable text widget to insert content
        instructions_scroll.config(state='normal')
        instructions_scroll.delete(1.0, tk.END)
        instructions_scroll.insert(1.0, instructions_text)
        
        # Make URLs clickable
        instructions_scroll.tag_config("url", foreground="blue", underline=True)
        
        # Function to find and make URLs clickable
        def make_urls_clickable():
            urls = [
                ("https://myaccount.google.com/security", "Google Security Settings"),
                ("https://myaccount.google.com/apppasswords", "Google App Passwords")
            ]
            
            for url, text in urls:
                start_pos = "1.0"
                while True:
                    start_pos = instructions_scroll.search(url, start_pos, stopindex=tk.END)
                    if not start_pos:
                        break
                    end_pos = f"{start_pos}+{len(url)}c"
                    
                    instructions_scroll.tag_add("url", start_pos, end_pos)
                    
                    # Bind click event
                    def open_url(url_to_open=url):
                        webbrowser.open(url_to_open)
                    
                    instructions_scroll.tag_bind("url", "<Button-1>", lambda e, u=url: webbrowser.open(u))
                    instructions_scroll.tag_bind("url", "<Enter>", 
                                               lambda e: instructions_scroll.config(cursor="hand2"))
                    instructions_scroll.tag_bind("url", "<Leave>", 
                                               lambda e: instructions_scroll.config(cursor=""))
                    
                    start_pos = end_pos
        
        make_urls_clickable()
        instructions_scroll.config(state='disabled')
        
        # Quick action buttons for Google setup
        setup_frame = tk.Frame(guide_frame, bg=self.colors["bg_dark"])
        setup_frame.pack(fill=tk.X, pady=10)
        
        setup_buttons = [
            ("🔓 Enable 2-Step Verification", "https://myaccount.google.com/security"),
            ("🔑 Generate App Password", "https://myaccount.google.com/apppasswords"),
            ("📖 Google Help Guide", "https://support.google.com/accounts/answer/185833")
        ]
        
        for text, url in setup_buttons:
            btn = tk.Button(setup_frame,
                          text=text,
                          command=lambda u=url: webbrowser.open(u),
                          bg=self.colors["accent"],
                          fg='white',
                          font=self.fonts["small"],
                          relief="flat",
                          padx=10,
                          pady=5,
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)
        
        # Email configuration
        config_frame = tk.Frame(main_frame, bg='white')
        config_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Title for email settings
        tk.Label(config_frame,
                text="📧 Your Email Configuration:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        configs = [
            ("📧 Your Email*:", "sender_email_var"),
            ("👤 Your Name*:", "your_name_var"),
            ("💼 Your Title:", "your_title_var")
        ]
        
        for i, (label, var_name) in enumerate(configs):
            row_frame = tk.Frame(config_frame, bg='white')
            row_frame.pack(fill=tk.X, pady=8)
            
            tk.Label(row_frame,
                    text=label,
                    bg='white',
                    fg=self.colors["text"],
                    font=self.fonts["normal"],
                    width=15,
                    anchor=tk.W).pack(side=tk.LEFT)
            
            # Get the already initialized variable
            var = getattr(self, var_name)
            
            entry = tk.Entry(row_frame,
                           textvariable=var,
                           bg='white',
                           fg=self.colors["text"],
                           font=self.fonts["normal"],
                           width=40)
            entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # Password with show/hide option
        row_frame = tk.Frame(config_frame, bg='white')
        row_frame.pack(fill=tk.X, pady=8)
        
        tk.Label(row_frame,
                text="🔐 App Password*:",
                bg='white',
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=15,
                anchor=tk.W).pack(side=tk.LEFT)
        
        self.password_entry = tk.Entry(row_frame,
                                     textvariable=self.sender_pass_var,
                                     show="*",
                                     bg='white',
                                     fg=self.colors["text"],
                                     font=self.fonts["normal"],
                                     width=35)
        self.password_entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # Show/Hide password checkbox
        show_password_cb = tk.Checkbutton(row_frame,
                                        text="👁️ Show",
                                        variable=self.show_password_var,
                                        command=self.toggle_password_visibility,
                                        bg='white',
                                        fg=self.colors["text"],
                                        font=self.fonts["small"])
        show_password_cb.pack(side=tk.LEFT)
        
        # Password help text
        pass_help = tk.Label(row_frame,
                            text="(16-character App Password, NOT your regular password)",
                            bg='white',
                            fg=self.colors["text_light"],
                            font=("Segoe UI", 8))
        pass_help.pack(side=tk.LEFT, padx=10)
        
        # SMTP Settings
        smtp_frame = tk.Frame(main_frame, bg=self.colors["bg_light"])
        smtp_frame.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Label(smtp_frame,
                text="🔧 SMTP Server Settings (Configured for Gmail):",
                bg=self.colors["bg_light"],
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        smtp_row = tk.Frame(smtp_frame, bg=self.colors["bg_light"])
        smtp_row.pack(fill=tk.X, pady=5)
        
        tk.Label(smtp_row,
                text="SMTP Server:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=15,
                anchor=tk.W).pack(side=tk.LEFT)
        
        smtp_entry = tk.Entry(smtp_row,
                            textvariable=self.smtp_server_var,
                            bg='white',
                            fg=self.colors["text"],
                            font=self.fonts["normal"],
                            width=40)
        smtp_entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        smtp_row2 = tk.Frame(smtp_frame, bg=self.colors["bg_light"])
        smtp_row2.pack(fill=tk.X, pady=5)
        
        tk.Label(smtp_row2,
                text="SMTP Port:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=15,
                anchor=tk.W).pack(side=tk.LEFT)
        
        port_entry = tk.Entry(smtp_row2,
                            textvariable=self.smtp_port_var,
                            bg='white',
                            fg=self.colors["text"],
                            font=self.fonts["normal"],
                            width=10)
        port_entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # Test connection immediately after entering credentials
        test_now_frame = tk.Frame(smtp_frame, bg=self.colors["bg_light"])
        test_now_frame.pack(fill=tk.X, pady=10)
        
        # Auto-detect Gmail button
        tk.Button(test_now_frame,
                 text="🎯 Auto-configure for Gmail",
                 command=self.auto_configure_gmail,
                 bg=self.colors["success"],
                 fg='white',
                 font=self.fonts["small"],
                 relief="flat",
                 padx=10,
                 pady=5,
                 cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Scheduling Options
        schedule_frame = tk.Frame(main_frame, bg='white')
        schedule_frame.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Label(schedule_frame,
                text="⏰ Scheduling & Resend Options:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        # Interval days selection
        interval_frame = tk.Frame(schedule_frame, bg='white')
        interval_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(interval_frame,
                text="Resend every:",
                bg='white',
                fg=self.colors["text"],
                font=self.fonts["normal"]).pack(side=tk.LEFT)
        
        intervals = [("1 day", "1"), ("2 days", "2"), ("3 days", "3"), ("Week", "7")]
        
        for text, value in intervals:
            tk.Radiobutton(interval_frame,
                         text=text,
                         variable=self.interval_days_var,
                         value=value,
                         bg='white',
                         fg=self.colors["text"],
                         selectcolor=self.colors["secondary"],
                         font=self.fonts["small"]).pack(side=tk.LEFT, padx=10)
        
        # Maximum resends selection
        resend_frame = tk.Frame(schedule_frame, bg='white')
        resend_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(resend_frame,
                text="Maximum resends:",
                bg='white',
                fg=self.colors["text"],
                font=self.fonts["normal"]).pack(side=tk.LEFT)
        
        max_resends_options = [("Send once", "1"), ("2 times", "2"), ("3 times", "3"), 
                              ("4 times", "4"), ("5 times", "5"), ("6 times", "6"), 
                              ("7 times (1 week)", "7")]
        
        resend_options_frame = tk.Frame(resend_frame, bg='white')
        resend_options_frame.pack(side=tk.LEFT, padx=10)
        
        row_count = 0
        col_count = 0
        for text, value in max_resends_options:
            rb = tk.Radiobutton(resend_options_frame,
                              text=text,
                              variable=self.max_resends_var,
                              value=value,
                              bg='white',
                              fg=self.colors["text"],
                              selectcolor=self.colors["secondary"],
                              font=self.fonts["small"])
            rb.grid(row=row_count, column=col_count, sticky=tk.W, padx=5, pady=2)
            col_count += 1
            if col_count > 3:  # 4 columns per row
                col_count = 0
                row_count += 1
        
        # Time selection
        time_frame = tk.Frame(schedule_frame, bg='white')
        time_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(time_frame,
                text="⏱️ Send time (24h format):",
                bg='white',
                fg=self.colors["text"],
                font=self.fonts["normal"]).pack(side=tk.LEFT)
        
        time_entry = tk.Entry(time_frame,
                            textvariable=self.send_time_var,
                            bg='white',
                            fg=self.colors["text"],
                            font=self.fonts["normal"],
                            width=10,
                            justify="center")
        time_entry.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # Help text for scheduling
        schedule_help = tk.Label(schedule_frame,
                                text="📌 Example: 'Resend every 2 days, Maximum 3 times' = Email sent on Day 1, Day 3, Day 5",
                                bg='white',
                                fg=self.colors["text_light"],
                                font=("Segoe UI", 9))
        schedule_help.pack(anchor=tk.W, pady=(10, 0))
        
        # Credential management buttons
        cred_frame = tk.Frame(main_frame, bg='white')
        cred_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(cred_frame,
                text="🔐 Credential Management:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W, pady=(0, 10))
        
        cred_buttons_frame = tk.Frame(cred_frame, bg='white')
        cred_buttons_frame.pack()
        
        # Load saved credentials button
        load_cred_btn = tk.Button(cred_buttons_frame,
                                text="🔄 Load Saved Credentials",
                                command=self.load_saved_credentials,
                                bg=self.colors["secondary"],
                                fg='white',
                                font=self.fonts["small"],
                                relief="flat",
                                padx=15,
                                pady=5,
                                cursor="hand2")
        load_cred_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear credentials button
        clear_cred_btn = tk.Button(cred_buttons_frame,
                                 text="🗑️ Clear Saved Credentials",
                                 command=self.clear_saved_credentials,
                                 bg=self.colors["warning"],
                                 fg='white',
                                 font=self.fonts["small"],
                                 relief="flat",
                                 padx=15,
                                 pady=5,
                                 cursor="hand2")
        clear_cred_btn.pack(side=tk.LEFT, padx=5)
        
        # Control buttons
        control_frame = tk.Frame(main_frame, bg='white')
        control_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Left side: Save Settings button
        left_control_frame = tk.Frame(control_frame, bg='white')
        left_control_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        save_btn = tk.Button(left_control_frame,
                          text="💾 Save Settings",
                          command=self.save_settings,
                          bg=self.colors["secondary"],
                          fg='white',
                          font=self.fonts["normal"],
                          relief="flat",
                          padx=20,
                          pady=10,
                          cursor="hand2")
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Center: Test Connection and Send Test Email buttons
        center_control_frame = tk.Frame(control_frame, bg='white')
        center_control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20)
        
        test_btn = tk.Button(center_control_frame,
                          text="🔗 Test Connection",
                          command=self.test_connection,
                          bg=self.colors["accent"],
                          fg='white',
                          font=self.fonts["normal"],
                          relief="flat",
                          padx=20,
                          pady=10,
                          cursor="hand2")
        test_btn.pack(side=tk.LEFT, padx=5)
        
        test_email_btn = tk.Button(center_control_frame,
                                text="📧 Send Test Email",
                                command=self.send_test_email,
                                bg=self.colors["success"],
                                fg='white',
                                font=self.fonts["normal"],
                                relief="flat",
                                padx=20,
                                pady=10,
                                cursor="hand2")
        test_email_btn.pack(side=tk.LEFT, padx=5)
        
        # Right side: Connection status label
        right_control_frame = tk.Frame(control_frame, bg='white')
        right_control_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.connection_status_label = tk.Label(right_control_frame,
                                              textvariable=self.connection_status_var,
                                              bg='white',
                                              fg=self.colors["error"],
                                              font=self.fonts["small"])
        self.connection_status_label.pack(side=tk.RIGHT, padx=10)

    def load_saved_credentials(self):
        """Load saved credentials from Windows Credential Manager"""
        email = self.sender_email_var.get().strip()
        
        if not email:
            messagebox.showinfo("Load Credentials", 
                              "Please enter your email address first, then click 'Load Saved Credentials'.")
            return
        
        password = self.get_credentials_from_vault(email)
        
        if password:
            self.sender_pass_var.set(password)
            messagebox.showinfo("Success", 
                              f"✅ Credentials loaded for:\n{email}\n\nPassword retrieved securely from Windows Credential Manager.")
            self.log_message(f"Loaded saved credentials for {email}")
            
            # Update connection status
            self.connection_status_var.set("⚙️ Credentials loaded")
            self.connection_status_label.config(fg=self.colors["warning"])
        else:
            messagebox.showinfo("No Saved Credentials", 
                              f"No saved credentials found for:\n{email}\n\nPlease enter your password and click 'Save Settings' to store them.")
            self.log_message(f"No saved credentials found for {email}")

    def clear_saved_credentials(self):
        """Clear saved credentials from Windows Credential Manager"""
        email = self.sender_email_var.get().strip()
        
        if not email:
            messagebox.showwarning("Clear Credentials", 
                                 "Please enter the email address whose credentials you want to clear.")
            return
        
        if messagebox.askyesno("Confirm Clear", 
                              f"Are you sure you want to delete saved credentials for:\n{email}\n\nThis cannot be undone!"):
            success = self.delete_credentials_from_vault(email)
            
            if success:
                # Clear the password field
                self.sender_pass_var.set("")
                messagebox.showinfo("Success", 
                                  f"✅ Credentials cleared for:\n{email}\n\nYou'll need to enter your password again.")
                self.log_message(f"Cleared saved credentials for {email}")
                
                # Update connection status
                self.connection_status_var.set("❌ Credentials cleared")
                self.connection_status_label.config(fg=self.colors["error"])
            else:
                messagebox.showerror("Error", 
                                   f"❌ Could not clear credentials for:\n{email}\n\nThey may not exist or there was an error.")

    def create_send_tab(self):
        """Separate tab for sending emails"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="🚀 Send Emails")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        main_frame = self.create_card_frame(scrollable_tab, "🚀 Send Emails to Selected Recipients")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Top section: Selected recipients
        selected_frame = tk.Frame(main_frame, bg='white')
        selected_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(selected_frame,
                text="👥 Selected Recipients:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W)
        
        # Refresh button
        refresh_btn = tk.Button(selected_frame,
                              text="🔄 Refresh List",
                              command=self.update_selected_listbox,
                              bg=self.colors["accent"],
                              fg='white',
                              font=self.fonts["small"],
                              relief="flat",
                              padx=10,
                              cursor="hand2")
        refresh_btn.pack(side=tk.RIGHT)
        
        # Selected recipients listbox
        list_frame = tk.Frame(selected_frame, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.selected_listbox = tk.Listbox(list_frame,
                                         bg='white',
                                         fg=self.colors["text"],
                                         font=self.fonts["normal"],
                                         height=8,
                                         selectmode=tk.MULTIPLE)
        self.selected_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        selected_scrollbar = tk.Scrollbar(list_frame, command=self.selected_listbox.yview)
        self.selected_listbox.configure(yscrollcommand=selected_scrollbar.set)
        selected_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Middle section: Email preview
        preview_frame = tk.Frame(main_frame, bg='white')
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(preview_frame,
                text="📧 Email Preview:",
                bg='white',
                fg=self.colors["primary"],
                font=self.fonts["heading"]).pack(anchor=tk.W)
        
        preview_text_frame = tk.Frame(preview_frame, bg='white')
        preview_text_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.send_preview_text = scrolledtext.ScrolledText(preview_text_frame,
                                                         bg=self.colors["bg_light"],
                                                         fg=self.colors["text"],
                                                         font=self.fonts["normal"],
                                                         height=10,
                                                         wrap=tk.WORD,
                                                         state='disabled')
        self.send_preview_text.pack(fill=tk.BOTH, expand=True)
        
        # Bottom section: Send controls
        control_frame = tk.Frame(main_frame, bg='white')
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Update preview button
        tk.Button(control_frame,
                 text="🔄 Update Preview",
                 command=self.update_email_preview,
                 bg=self.colors["accent"],
                 fg='white',
                 font=self.fonts["normal"],
                 relief="flat",
                 cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Send options frame
        send_options_frame = tk.Frame(control_frame, bg='white')
        send_options_frame.pack(side=tk.RIGHT)
        
        tk.Radiobutton(send_options_frame,
                     text="📤 Send Now",
                     variable=self.send_mode_var,
                     value="send",
                     bg='white',
                     fg=self.colors["text"],
                     font=self.fonts["normal"]).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(send_options_frame,
                     text="📝 Save as Draft",
                     variable=self.send_mode_var,
                     value="draft",
                     bg='white',
                     fg=self.colors["text"],
                     font=self.fonts["normal"]).pack(side=tk.LEFT, padx=10)
        
        # Send button
        self.send_button = tk.Button(control_frame,
                                   text="🚀 SEND EMAILS",
                                   command=self.send_selected_emails,
                                   bg=self.colors["success"],
                                   fg='white',
                                   font=self.fonts["heading"],
                                   relief="flat",
                                   padx=30,
                                   pady=12,
                                   cursor="hand2")
        self.send_button.pack(side=tk.RIGHT, padx=10)
        
        # Status label for sending
        self.send_status_label = tk.Label(control_frame,
                                        textvariable=self.send_status_var,
                                        bg='white',
                                        fg=self.colors["text_light"],
                                        font=self.fonts["small"])
        self.send_status_label.pack(side=tk.RIGHT, padx=10)

    def create_history_tab(self):
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📊 History & Logs")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        # Two columns for history
        left_frame = tk.Frame(scrollable_tab, bg='white')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        right_frame = tk.Frame(scrollable_tab, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Sent emails history
        sent_frame = self.create_card_frame(left_frame, "📨 Sent Emails History")
        sent_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview for sent emails
        columns = ("Date", "Time", "To", "Company", "Subject", "Status")
        self.history_tree = ttk.Treeview(sent_frame, columns=columns, show="headings", height=15)
        
        style = ttk.Style()
        style.configure("History.Treeview",
                       background='white',
                       foreground=self.colors["text"],
                       fieldbackground='white',
                       rowheight=25)
        
        for col in columns:
            self.history_tree.heading(col, text=col)
            if col == "Date":
                self.history_tree.column(col, width=100)
            elif col == "Time":
                self.history_tree.column(col, width=80)
            elif col == "To":
                self.history_tree.column(col, width=150)
            elif col == "Company":
                self.history_tree.column(col, width=120)
            elif col == "Subject":
                self.history_tree.column(col, width=200)
            elif col == "Status":
                self.history_tree.column(col, width=100)
        
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar
        tree_scrollbar = ttk.Scrollbar(sent_frame, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # History controls
        history_controls = tk.Frame(sent_frame, bg='white')
        history_controls.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        history_buttons = [
            ("🔄 Refresh", self.load_history, self.colors["secondary"]),
            ("📊 Export", self.export_history, self.colors["accent"]),
            ("🗑️ Clear", self.clear_history, self.colors["warning"])
        ]
        
        for text, command, color in history_buttons:
            btn = tk.Button(history_controls,
                          text=text,
                          command=command,
                          bg=color,
                          fg='white',
                          font=self.fonts["small"],
                          relief="flat",
                          padx=15,
                          pady=5,
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)
        
        # Activity Log
        log_frame = self.create_card_frame(right_frame, "📝 Activity Log")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_content = tk.Frame(log_frame, bg='white')
        log_content.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_content,
                                                bg=self.colors["bg_light"],
                                                fg=self.colors["text"],
                                                font=self.fonts["normal"],
                                                height=20,
                                                wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state='disabled')
        
        # Initial log message
        self.log_message("Application started with secure credential storage")
        
        # Stats panel
        stats_frame = self.create_card_frame(left_frame, "📈 Statistics")
        stats_frame.pack(fill=tk.X)
        
        stats_content = tk.Frame(stats_frame, bg='white')
        stats_content.pack(fill=tk.X, padx=20, pady=15)
        
        # Initialize stats variables
        self.stats_vars = {
            "total_sent": tk.StringVar(value="0"),
            "success_rate": tk.StringVar(value="0%"),
            "pending": tk.StringVar(value="0"),
            "failed": tk.StringVar(value="0"),
            "resends": tk.StringVar(value="0")  # NEW: Resend count
        }
        
        stats_labels = [
            ("Total Sent:", "total_sent", self.colors["primary"]),
            ("Success Rate:", "success_rate", self.colors["success"]),
            ("Pending:", "pending", self.colors["warning"]),
            ("Failed:", "failed", self.colors["error"]),
            ("Resends:", "resends", self.colors["accent"])  # NEW: Resends stat
        ]
        
        for i, (label, var_name, color) in enumerate(stats_labels):
            stat_frame = tk.Frame(stats_content, bg='white')
            stat_frame.grid(row=0, column=i, padx=20)
            
            tk.Label(stat_frame,
                    text=label,
                    bg='white',
                    fg=self.colors["text_light"],
                    font=self.fonts["small"]).pack()
            
            tk.Label(stat_frame,
                    textvariable=self.stats_vars[var_name],
                    bg='white',
                    fg=color,
                    font=("Segoe UI", 14, "bold")).pack()

    def create_developer_tab(self):
        """Developer information tab with clickable social links"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="👨‍💻 Developer Info")
        
        # Create scrollable frame for the entire tab
        scrollable_tab, canvas, scrollbar = self.create_scrollable_frame(tab)
        
        main_frame = self.create_card_frame(scrollable_tab, "👨‍💻 About Job Application Assistant")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Developer info
        info_frame = tk.Frame(main_frame, bg='white')
        info_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Logo/Title
        title_label = tk.Label(info_frame,
                             text="📧 Job Application Assistant",
                             font=("Segoe UI", 24, "bold"),
                             bg='white',
                             fg=self.colors["primary"])
        title_label.pack(pady=(0, 20))
        
        # Version
        version_label = tk.Label(info_frame,
                               text="Version 1.0.0",
                               font=self.fonts["heading"],
                               bg='white',
                               fg=self.colors["text_light"])
        version_label.pack(pady=(0, 30))
        
        # Description
        description = """This application helps job seekers automate their application process.
It allows you to:
• Store and manage recipient information
• Create personalized email templates
• Schedule automatic email sending (with resend options)
• Track sent applications
• Manage attachments (resume, cover letter, etc.)

🔒 SECURITY FEATURE:
• Credentials stored securely in Windows Credential Manager
• Each user's credentials stay on their own PC
• Passwords are NEVER saved to files

🔄 NEW: Resend Feature with Stop Control
• Set how many times to resend emails (1-7 times)
• Choose resend interval (1-7 days)
• Stop resend for any recipient at any time
• Automatic scheduling based on last sent date"""
        
        desc_label = tk.Label(info_frame,
                            text=description,
                            font=self.fonts["normal"],
                            bg='white',
                            fg=self.colors["text"],
                            justify=tk.LEFT)
        desc_label.pack(pady=(0, 30))
        
        # Developer info
        dev_frame = tk.Frame(info_frame, bg=self.colors["bg_light"])
        dev_frame.pack(fill=tk.X, pady=10, ipadx=10, ipady=10)
        
        tk.Label(dev_frame,
                text="👨‍💻 Developer Information:",
                font=self.fonts["heading"],
                bg=self.colors["bg_light"],
                fg=self.colors["primary"]).pack(anchor=tk.W, pady=(0, 10))
        
        # Create a frame for developer info with clickable links
        dev_info_frame = tk.Frame(dev_frame, bg=self.colors["bg_light"])
        dev_info_frame.pack(anchor=tk.W, pady=5)
        
        # Name
        tk.Label(dev_info_frame,
                text="Name:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=10,
                anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Label(dev_info_frame,
                text="RiDDL3TeCH [EPHANTUS MWAURA - KENYA ]",
                bg=self.colors["bg_light"],
                fg=self.colors["primary"],
                font=self.fonts["normal"]).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        # Email (clickable)
        tk.Label(dev_info_frame,
                text="Email:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=10,
                anchor=tk.W).grid(row=1, column=0, sticky=tk.W, pady=2)
        email_btn = tk.Button(dev_info_frame,
                            text="snaidermwaura104@gmail.com",
                            command=lambda: webbrowser.open("mailto:snaidermwaura104@gmail.com"),
                            bg=self.colors["bg_light"],
                            fg="#3498db",
                            font=self.fonts["normal"],
                            bd=0,
                            cursor="hand2",
                            relief="flat",
                            anchor=tk.W)
        email_btn.grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # GitHub (clickable)
        tk.Label(dev_info_frame,
                text="GitHub:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=10,
                anchor=tk.W).grid(row=2, column=0, sticky=tk.W, pady=2)
        github_btn = tk.Button(dev_info_frame,
                             text="https://github.com/Riddlemwas",
                             command=lambda: webbrowser.open("https://github.com/Riddlemwas"),
                             bg=self.colors["bg_light"],
                             fg="#3498db",
                             font=self.fonts["normal"],
                             bd=0,
                             cursor="hand2",
                             relief="flat",
                             anchor=tk.W)
        github_btn.grid(row=2, column=1, sticky=tk.W, pady=2)
        
        # Created
        tk.Label(dev_info_frame,
                text="Created:",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"],
                width=10,
                anchor=tk.W).grid(row=3, column=0, sticky=tk.W, pady=2)
        tk.Label(dev_info_frame,
                text="January 2026",
                bg=self.colors["bg_light"],
                fg=self.colors["text"],
                font=self.fonts["normal"]).grid(row=3, column=1, sticky=tk.W, pady=2)
        
        # Social Media Links - UPDATED WITH CLICKABLE LINKS
        social_frame = tk.Frame(info_frame, bg='white')
        social_frame.pack(fill=tk.X, pady=20)
        
        tk.Label(social_frame,
                text="🌐 Connect with Developer:",
                font=self.fonts["heading"],
                bg='white',
                fg=self.colors["primary"]).pack(pady=(0, 10))
        
        # Social buttons frame
        social_buttons_frame = tk.Frame(social_frame, bg='white')
        social_buttons_frame.pack()
        
        # WhatsApp button with clickable link
        whatsapp_btn = tk.Button(social_buttons_frame,
                               text="💬 WhatsApp",
                               command=lambda: webbrowser.open("https://wa.me/+254743156789?text=Hello%20from%20Job%20Application%20Assistant"),
                               bg="#25D366",  # WhatsApp green
                               fg='white',
                               font=self.fonts["heading"],
                               relief="flat",
                               padx=20,
                               pady=10,
                               cursor="hand2")
        whatsapp_btn.pack(side=tk.LEFT, padx=5)
        
        # Facebook button with clickable link
        facebook_btn = tk.Button(social_buttons_frame,
                               text="👥 Facebook",
                               command=lambda: webbrowser.open("https://www.facebook.com/riddletech104"),
                               bg="#1877F2",  # Facebook blue
                               fg='white',
                               font=self.fonts["heading"],
                               relief="flat",
                               padx=20,
                               pady=10,
                               cursor="hand2")
        facebook_btn.pack(side=tk.LEFT, padx=5)
        
        # GitHub button with clickable link
        github_btn = tk.Button(social_buttons_frame,
                             text="🐙 GitHub",
                             command=lambda: webbrowser.open("https://github.com/Riddlemwas"),
                             bg="#333333",  # GitHub black
                             fg='white',
                             font=self.fonts["heading"],
                               relief="flat",
                               padx=20,
                               pady=10,
                               cursor="hand2")
        github_btn.pack(side=tk.LEFT, padx=5)
        
        # Instagram button with clickable link
        instagram_btn = tk.Button(social_buttons_frame,
                                text="📷 Instagram",
                                command=lambda: webbrowser.open("https://instagram.com/riddl3tech"),
                                bg="#E4405F",  # Instagram pink
                                fg='white',
                                font=self.fonts["heading"],
                                relief="flat",
                                padx=20,
                                pady=10,
                                cursor="hand2")
        instagram_btn.pack(side=tk.LEFT, padx=5)
        
        # Support info
        support_frame = tk.Frame(info_frame, bg='white')
        support_frame.pack(fill=tk.X, pady=20)
        
        tk.Label(support_frame,
                text="❓ Need help or want to contribute?",
                font=self.fonts["heading"],
                bg='white',
                fg=self.colors["primary"]).pack(pady=(0, 10))
        
        # Support buttons
        button_frame = tk.Frame(support_frame, bg='white')
        button_frame.pack()
        
        support_buttons = [
            ("📧 Contact Developer", lambda: webbrowser.open("mailto:snaidermwaura104@gmail.com")),
            ("🐛 Report Bug", lambda: webbrowser.open("https://www.facebook.com/me/")),
            ("📚 Documentation", lambda: webbrowser.open("https://www.facebook.com/me/")),
            ("⭐ Rate App", lambda: webbrowser.open("https://example.com/rate"))
        ]
        
        for text, command in support_buttons:
            btn = tk.Button(button_frame,
                          text=text,
                          command=command,
                          bg=self.colors["secondary"],
                          fg='white',
                          font=self.fonts["normal"],
                          relief="flat",
                          padx=15,
                          pady=8,
                          cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)
        
        # Copyright
        copyright_frame = tk.Frame(info_frame, bg='white')
        copyright_frame.pack(fill=tk.X, pady=20)
        
        tk.Label(copyright_frame,
                text="© 2026 Job Application Assistant. All rights reserved.",
                font=self.fonts["small"],
                bg='white',
                fg=self.colors["text_light"]).pack()

    def update_selected_listbox(self):
        """Update the selected recipients listbox - FIXED: Only show checked recipients"""
        self.selected_listbox.delete(0, tk.END)
        self.selected_recipients.clear()
        
        for item in self.recipient_tree.get_children():
            values = self.recipient_tree.item(item, "values")
            if values:  # Make sure values exist
                # Check if the first column (checkbox) is checked
                is_checked = values[0] == "✓" if len(values) > 0 else False
                
                if is_checked:  # Only selected recipients (with checkmark)
                    company = values[1] if len(values) > 1 else ""
                    email = values[3] if len(values) > 3 else ""
                    
                    if company and email:  # Only add if we have valid data
                        display_text = f"{company} - {email}"
                        self.selected_listbox.insert(tk.END, display_text)
                        self.selected_recipients.append({
                            "company": company,
                            "hr_name": values[2] if len(values) > 2 else "",
                            "email": email,
                            "position": values[4] if len(values) > 4 else "",
                            "status": values[5] if len(values) > 5 else "PENDING"
                        })
        
        # Update status
        selected_count = len(self.selected_recipients)
        if selected_count == 0:
            self.send_status_var.set("No recipients selected")
            self.send_status_label.config(fg=self.colors["text_light"])
        else:
            self.send_status_var.set(f"Selected: {selected_count} recipient(s)")
            self.send_status_label.config(fg=self.colors["success"])
        
        # Auto-update the email preview when selection changes
        self.update_email_preview()

    def select_all_recipients(self):
        """Select all recipients"""
        for item in self.recipient_tree.get_children():
            self.recipient_tree.set(item, column="✓", value="✓")
        self.update_selected_listbox()
        self.log_message("Selected all recipients")

    def deselect_all_recipients(self):
        """Deselect all recipients"""
        for item in self.recipient_tree.get_children():
            self.recipient_tree.set(item, column="✓", value="")
        self.update_selected_listbox()
        self.log_message("Deselected all recipients")

    def stop_resend_selected(self):
        """Stop resend for selected recipients"""
        selected_items = self.recipient_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select recipients to stop resend.")
            return
        
        for item in selected_items:
            values = self.recipient_tree.item(item, "values")
            email = values[3] if len(values) > 3 else ""
            
            for recipient in self.recipients:
                if recipient['email'] == email:
                    recipient['stop_resend'] = True
                    
                    # Update treeview
                    status_display = f"⏹️ STOPPED ({recipient['send_count']}/{recipient['max_sends']})"
                    self.recipient_tree.set(item, column="Status", value=status_display)
                    self.recipient_tree.set(item, column="⏹️ Stop", value="⏹️")
                    break
        
        self.save_recipients_data()
        messagebox.showinfo("Success", "✅ Stopped resend for selected recipients")
        self.log_message("Stopped resend for selected recipients")

    def resume_resend_selected(self):
        """Resume resend for selected recipients"""
        selected_items = self.recipient_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select recipients to resume resend.")
            return
        
        for item in selected_items:
            values = self.recipient_tree.item(item, "values")
            email = values[3] if len(values) > 3 else ""
            
            for recipient in self.recipients:
                if recipient['email'] == email:
                    recipient['stop_resend'] = False
                    
                    # Update treeview
                    status_display = f"{recipient['status']} ({recipient['send_count']}/{recipient['max_sends']})"
                    self.recipient_tree.set(item, column="Status", value=status_display)
                    self.recipient_tree.set(item, column="⏹️ Stop", value="")
                    break
        
        self.save_recipients_data()
        messagebox.showinfo("Success", "✅ Resumed resend for selected recipients")
        self.log_message("Resumed resend for selected recipients")

    def add_recipient(self):
        """Add new recipient WITHOUT auto-selecting it"""
        company = self.company_var.get().strip()
        hr_name = self.hr_name_var.get().strip()
        email = self.email_var.get().strip()
        position = self.position_var.get().strip()
        notes = self.notes_var.get().strip()
        
        if not company or not email:
            messagebox.showwarning("Warning", "Company and Email are required!")
            return
        
        if not self.is_valid_email(email):
            messagebox.showwarning("Warning", "Please enter a valid email address!")
            return
        
        # Check for duplicate email
        for recipient in self.recipients:
            if recipient['email'] == email:
                messagebox.showwarning("Warning", "Recipient with this email already exists!")
                return
        
        # Get max sends from settings
        try:
            max_sends = int(self.max_resends_var.get())
        except:
            max_sends = 1
        
        recipient = {
            "company": company,
            "hr_name": hr_name,
            "email": email,
            "position": position,
            "notes": notes,
            "status": "PENDING",
            "last_sent": "",
            "added_date": datetime.now().strftime("%Y-%m-%d"),
            "send_count": 0,
            "max_sends": max_sends,
            "stop_resend": False,  # NEW: Default to not stopped
            "email_template": "",  # NEW: Will store email template
            "email_subject": "",  # NEW: Will store email subject
            "attachments": []  # NEW: Will store attachments for this recipient
        }
        
        self.recipients.append(recipient)
        
        # Add to treeview WITHOUT auto-selecting (checkbox remains empty)
        status_display = f"PENDING (0/{max_sends})"
        self.recipient_tree.insert("", tk.END, values=(
            "",  # FIXED: Empty checkbox (not selected by default)
            company,
            hr_name,
            email,
            position,
            status_display,
            datetime.now().strftime("%Y-%m-%d"),
            ""  # Empty stop column
        ))
        
        # Clear form
        self.company_var.set("")
        self.hr_name_var.set("")
        self.email_var.set("")
        self.position_var.set("")
        self.notes_var.set("")
        
        self.log_message(f"Added recipient: {company} ({email}) - Max sends: {max_sends}")
        
        # Update selected listbox (should show 0 selected since checkbox is empty)
        self.update_selected_listbox()
        
        # Auto-save recipients data
        self.save_recipients_data()
        
        # Show success message
        messagebox.showinfo("Success", f"✅ Recipient added: {company}\n\nYou can now select this recipient by ticking the checkbox in the recipients list.\n\nMax resends: {max_sends} times")

    def auto_configure_gmail(self):
        """Auto-configure settings for Gmail"""
        self.smtp_server_var.set("smtp.gmail.com")
        self.smtp_port_var.set("587")
        
        # Check if email looks like Gmail
        email = self.sender_email_var.get()
        if email and "@gmail.com" in email.lower():
            messagebox.showinfo("Auto-configured", 
                              "✅ Gmail settings configured:\n\n"
                              "SMTP Server: smtp.gmail.com\n"
                              "SMTP Port: 587\n\n"
                              "Remember to use an App Password, not your regular password!")
            self.log_message("Auto-configured Gmail settings")
        else:
            messagebox.showinfo("Auto-configured", 
                              "✅ Gmail settings configured:\n\n"
                              "SMTP Server: smtp.gmail.com\n"
                              "SMTP Port: 587\n\n"
                              "These are standard Gmail settings.")
            self.log_message("Configured Gmail SMTP settings")

    def toggle_password_visibility(self):
        """Toggle password visibility"""
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")

    def update_email_preview(self):
        """Update email preview with recipient data - NOW INCLUDES ATTACHMENTS WITHOUT DUPLICATION"""
        if not self.selected_recipients:
            self.send_preview_text.config(state='normal')
            self.send_preview_text.delete(1.0, tk.END)
            self.send_preview_text.insert(1.0, "No recipients selected. Please select recipients from the Recipients tab.")
            self.send_preview_text.config(state='disabled')
            return
        
        # Get first recipient for preview
        recipient = self.selected_recipients[0]
        
        subject = self.subject_var.get()
        body = self.body_text.get("1.0", tk.END)
        
        # Replace tags with actual data
        replacements = {
            "{company}": recipient['company'],
            "{hr_name}": recipient['hr_name'] or "Hiring Manager",
            "{position}": recipient['position'] or "the position",
            "{date}": datetime.now().strftime("%B %d, %Y"),
            "{your_name}": self.your_name_var.get() or "[Your Name]",
            "{your_title}": self.your_title_var.get() or "[Your Title]"
        }
        
        for tag, value in replacements.items():
            subject = subject.replace(tag, value)
            body = body.replace(tag, value)
        
        # Get all attachments that will be sent - FIXED: No duplication
        all_attachments = []
        
        # Check which files will actually be sent (remove duplicates)
        files_to_send = self.get_attachments_for_sending()
        
        for file_path in files_to_send:
            if os.path.exists(file_path):
                all_attachments.append(os.path.basename(file_path))
        
        # Create attachments list text
        attachments_text = ""
        if all_attachments:
            attachments_text = "\n\n📎 Attachments:\n"
            for i, att in enumerate(all_attachments, 1):
                attachments_text += f"  {i}. {att}\n"
        else:
            attachments_text = "\n\n📎 Attachments: None"
        
        preview_content = f"""TO: {recipient['email']}
SUBJECT: {subject}

{body}{attachments_text}

---
📊 Email Details:
• Number of recipients selected: {len(self.selected_recipients)}
• Total attachments: {len(all_attachments)}
• This email will be sent to all selected recipients."""
        
        self.send_preview_text.config(state='normal')
        self.send_preview_text.delete(1.0, tk.END)
        self.send_preview_text.insert(1.0, preview_content)
        self.send_preview_text.config(state='disabled')

    def get_attachments_for_sending(self):
        """Get all attachments for sending without duplication - FIXED"""
        # Use a set to avoid duplicates
        attachments_set = set()
        
        # Add all attachments from the list
        for attachment in self.attachments:
            if os.path.exists(attachment):
                attachments_set.add(os.path.abspath(attachment))
        
        # Convert back to list
        return list(attachments_set)

    def send_selected_emails(self):
        """Send emails to selected recipients"""
        if not self.selected_recipients:
            messagebox.showwarning("No Recipients", "Please select at least one recipient to send emails to.")
            return
        
        # Validate email settings
        email = self.sender_email_var.get()
        password = self.sender_pass_var.get()
        
        # Try to get password from vault if not entered
        if not password and email:
            password = self.get_credentials_from_vault(email)
            if password:
                self.sender_pass_var.set(password)
                self.log_message(f"Auto-loaded password for {email} from secure storage")
        
        if not email or not password:
            messagebox.showwarning("Settings Required", "Please configure your email settings first.")
            self.notebook.select(3)  # Go to settings tab
            return
        
        # Get attachments without duplication
        all_attachments = self.get_attachments_for_sending()
        
        if not all_attachments:
            response = messagebox.askyesno("No Attachments", 
                                         "No attachments are selected. Send email without attachments?")
            if not response:
                return
        
        send_mode = self.send_mode_var.get()
        total = len(self.selected_recipients)
        
        if send_mode == "send":
            response = messagebox.askyesno("Confirm Send", 
                                         f"Send email to {total} selected recipient(s) with {len(all_attachments)} attachment(s)?")
            if not response:
                return
            
            success_count = 0
            fail_count = 0
            
            for i, recipient in enumerate(self.selected_recipients, 1):
                self.send_status_var.set(f"Sending {i}/{total} to {recipient['company']}")
                self.send_status_label.config(fg=self.colors["warning"])
                self.root.update()
                
                success = self.send_single_email(recipient, all_attachments)
                
                if success:
                    success_count += 1
                    # Update treeview status
                    for item in self.recipient_tree.get_children():
                        values = self.recipient_tree.item(item, "values")
                        if values and values[3] == recipient['email']:
                            # Find the actual recipient to get current counts
                            for r in self.recipients:
                                if r['email'] == recipient['email']:
                                    if r['stop_resend']:
                                        status_display = f"⏹️ STOPPED ({r['send_count']}/{r['max_sends']})"
                                    else:
                                        status_display = f"SENT ({r['send_count']}/{r['max_sends']})"
                                    self.recipient_tree.set(item, column="Status", value=status_display)
                                    self.recipient_tree.set(item, column="✓", value="")  # Deselect after sending
                                    break
                else:
                    fail_count += 1
                    for item in self.recipient_tree.get_children():
                        values = self.recipient_tree.item(item, "values")
                        if values and values[3] == recipient['email']:
                            # Find the actual recipient to get current counts
                            for r in self.recipients:
                                if r['email'] == recipient['email']:
                                    status_display = f"FAILED ({r['send_count']}/{r['max_sends']})"
                                    self.recipient_tree.set(item, column="Status", value=status_display)
                                    break
                
                time.sleep(1)  # Small delay between emails
            
            # Update selected listbox
            self.update_selected_listbox()
            
            # Save recipients data after sending
            self.save_recipients_data()
            
            # Show results
            result_message = f"✅ Successfully sent: {success_count} emails\n"
            result_message += f"📎 Attachments sent: {len(all_attachments)} files\n"
            if fail_count > 0:
                result_message += f"❌ Failed to send: {fail_count} emails\n\n"
                result_message += "Check the Activity Log for details."
            
            messagebox.showinfo("Send Complete", result_message)
            
            self.send_status_var.set(f"Sent {success_count} of {total} emails")
            self.send_status_label.config(fg=self.colors["success"])
            
            self.log_message(f"Sent emails: {success_count} success, {fail_count} failed, {len(all_attachments)} attachments")
            
        else:  # Save as draft
            for recipient in self.selected_recipients:
                self.save_as_draft(recipient)
            
            messagebox.showinfo("Drafts Saved", 
                              f"✅ Saved {total} email(s) as drafts.")
            self.send_status_var.set(f"Saved {total} emails as drafts")
            self.log_message(f"Saved {total} emails as drafts")

    def send_single_email(self, recipient, attachments_list=None):
        """Send email to a single recipient - FIXED: No duplicate attachments"""
        try:
            msg = MIMEMultipart()
            msg['From'] = f"{self.your_name_var.get()} <{self.sender_email_var.get()}>"
            msg['To'] = recipient['email']
            
            # Check if recipient has saved email template and subject
            use_saved_template = False
            for r in self.recipients:
                if r['email'] == recipient['email'] and r.get('email_template') and r.get('email_subject'):
                    use_saved_template = True
                    saved_subject = r['email_subject']
                    saved_body = r['email_template']
                    saved_attachments = r.get('attachments', [])
                    break
            
            if use_saved_template and recipient.get('send_count', 0) > 0:
                # Use saved template for resend
                subject = saved_subject
                body = saved_body
                attachments = saved_attachments if saved_attachments else attachments_list
            else:
                # Use current template for first send
                subject = self.subject_var.get()
                body = self.body_text.get("1.0", tk.END)
                attachments = attachments_list if attachments_list else self.get_attachments_for_sending()
            
            # Replace tags with recipient data
            replacements = {
                "{company}": recipient['company'],
                "{hr_name}": recipient['hr_name'] or "Hiring Manager",
                "{position}": recipient['position'] or "the position",
                "{date}": datetime.now().strftime("%B %d, %Y"),
                "{your_name}": self.your_name_var.get(),
                "{your_title}": self.your_title_var.get()
            }
            
            for tag, value in replacements.items():
                subject = subject.replace(tag, value)
                body = body.replace(tag, value)
            
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            
            # Log attachment info
            attachment_names = [os.path.basename(att) for att in attachments if os.path.exists(att)]
            self.log_message(f"Attaching {len(attachment_names)} files: {', '.join(attachment_names)}")
            
            # Actually attach all files
            attached_count = 0
            for file_path in attachments:
                if os.path.exists(file_path):
                    try:
                        with open(file_path, 'rb') as f:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', 
                                          f'attachment; filename="{os.path.basename(file_path)}"')
                            msg.attach(part)
                        attached_count += 1
                    except Exception as e:
                        self.log_message(f"❌ Failed to attach {os.path.basename(file_path)}: {str(e)}")
                        continue
            
            # Send email
            server = smtplib.SMTP(self.smtp_server_var.get(), int(self.smtp_port_var.get()))
            server.starttls()
            server.login(self.sender_email_var.get(), self.sender_pass_var.get())
            server.send_message(msg)
            server.quit()
            
            # Find the recipient in our list to update counts
            for r in self.recipients:
                if r['email'] == recipient['email']:
                    # Increment send count and update last sent
                    r['send_count'] += 1
                    r['last_sent'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Save email template, subject and attachments for future resends
                    if r['send_count'] == 1:  # Save on first send
                        r['email_template'] = body
                        r['email_subject'] = self.subject_var.get()  # Save template with tags
                        r['attachments'] = [os.path.abspath(att) for att in attachments]
                    
                    # Update status
                    if r['send_count'] >= r['max_sends']:
                        r['status'] = "COMPLETED"
                    elif r['stop_resend']:
                        r['status'] = "STOPPED"
                    else:
                        r['status'] = "SENT"
                    break
            
            # Log sent email
            sent_email = {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "time": datetime.now().strftime("%H:%M:%S"),
                "to": recipient['email'],
                "company": recipient['company'],
                "subject": subject,
                "status": "SENT",
                "attachments": attached_count,
                "send_count": r['send_count'] if 'r' in locals() else 1
            }
            self.sent_emails.append(sent_email)
            
            # Update recipient status in treeview
            for r in self.recipients:
                if r['email'] == recipient['email']:
                    # Find and update the treeview item
                    for item in self.recipient_tree.get_children():
                        values = self.recipient_tree.item(item, "values")
                        if values and values[3] == recipient['email']:
                            if r['stop_resend']:
                                status_display = f"⏹️ STOPPED ({r['send_count']}/{r['max_sends']})"
                            else:
                                status_display = f"SENT ({r['send_count']}/{r['max_sends']})"
                            self.recipient_tree.set(item, column="Status", value=status_display)
                            break
                    break
            
            # Add to history tree
            send_info = f"✅ SENT ({attached_count} files) - {r['send_count']}/{r['max_sends']}" if 'r' in locals() else f"✅ SENT ({attached_count} files)"
            self.history_tree.insert("", tk.END, values=(
                sent_email["date"],
                sent_email["time"],
                sent_email["to"],
                sent_email["company"],
                sent_email["subject"][:50] + "..." if len(sent_email["subject"]) > 50 else sent_email["subject"],
                send_info
            ))
            
            # Log to file
            self.log_sent_email(sent_email)
            
            # Update stats
            self.update_stats()
            
            self.log_message(f"✅ Email sent to {recipient['company']} ({recipient['email']}) with {attached_count} attachments (Send #{r['send_count'] if 'r' in locals() else 1})")
            return True
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"❌ Failed to send to {recipient['email']}: {error_msg}")
            
            # Add to history as failed
            failed_email = {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "time": datetime.now().strftime("%H:%M:%S"),
                "to": recipient['email'],
                "company": recipient['company'],
                "subject": subject,
                "status": f"FAILED: {error_msg[:50]}"
            }
            self.history_tree.insert("", tk.END, values=(
                failed_email["date"],
                failed_email["time"],
                failed_email["to"],
                failed_email["company"],
                failed_email["subject"][:50] + "..." if len(failed_email["subject"]) > 50 else failed_email["subject"],
                "❌ FAILED"
            ))
            
            return False

    def test_connection(self):
        """Test SMTP connection with better feedback"""
        email = self.sender_email_var.get().strip()
        password = self.sender_pass_var.get().strip()
        
        # Try to get password from vault if not entered
        if not password and email:
            password = self.get_credentials_from_vault(email)
            if password:
                self.sender_pass_var.set(password)
                self.log_message(f"Auto-loaded password for {email} for connection test")
        
        if not email or not password:
            messagebox.showwarning("Settings Required", 
                                 "Please enter your email and App Password first.\n\n"
                                 "For Gmail: Use 16-character App Password, not regular password.")
            self.connection_status_var.set("❌ Enter credentials")
            self.connection_status_label.config(fg=self.colors["error"])
            return
        
        try:
            self.log_message("Testing SMTP connection...")
            self.connection_status_var.set("🔄 Testing...")
            self.connection_status_label.config(fg=self.colors["warning"])
            self.root.update()
            
            # Validate email format
            if not self.is_valid_email(email):
                messagebox.showwarning("Invalid Email", 
                                     "Please enter a valid email address.")
                self.connection_status_var.set("❌ Invalid email")
                self.connection_status_label.config(fg=self.colors["error"])
                return
            
            # Test connection
            server = smtplib.SMTP(self.smtp_server_var.get(), 
                                int(self.smtp_port_var.get()), 
                                timeout=10)
            server.ehlo()
            server.starttls()
            server.ehlo()
            
            # Try to login
            server.login(email, password)
            server.quit()
            
            # Success
            self.connection_status_var.set("✅ Connected")
            self.connection_status_label.config(fg=self.colors["success"])
            
            messagebox.showinfo("Success", 
                              f"✅ SMTP connection successful!\n\n"
                              f"Server: {self.smtp_server_var.get()}:{self.smtp_port_var.get()}\n"
                              f"Account: {email}\n\n"
                              f"You can now send emails.")
            self.log_message(f"SMTP connection successful for {email}")
            
            # Auto-save settings on successful connection
            self.save_settings()
            
        except smtplib.SMTPAuthenticationError:
            self.connection_status_var.set("❌ Authentication failed")
            self.connection_status_label.config(fg=self.colors["error"])
            
            error_msg = ("❌ Authentication failed!\n\n"
                        "For Gmail accounts:\n"
                        "1. Make sure you're using an App Password, not your regular password\n"
                        "2. Enable 2-Step Verification first\n"
                        "3. Generate a new App Password at:\n"
                        "   https://myaccount.google.com/apppasswords\n\n"
                        "For other email providers:\n"
                        "Check your SMTP settings and password.")
            
            messagebox.showerror("Authentication Failed", error_msg)
            self.log_message("SMTP authentication failed")
            
            # Offer to open Google App Password page
            if "@gmail.com" in email.lower():
                if messagebox.askyesno("Need App Password?", 
                                      "Open Google App Passwords page to generate one?"):
                    webbrowser.open("https://myaccount.google.com/apppasswords")
        
        except smtplib.SMTPException as e:
            self.connection_status_var.set("❌ SMTP error")
            self.connection_status_label.config(fg=self.colors["error"])
            
            error_msg = f"❌ SMTP Error: {str(e)}\n\n"
            if "SSL" in str(e):
                error_msg += "Try using port 465 with SSL instead of 587 with TLS."
            
            messagebox.showerror("Connection Error", error_msg)
            self.log_message(f"SMTP connection error: {str(e)}")
        
        except Exception as e:
            self.connection_status_var.set("❌ Connection failed")
            self.connection_status_label.config(fg=self.colors["error"])
            
            error_msg = f"❌ Connection failed: {str(e)}\n\n"
            error_msg += "Check your internet connection and SMTP settings."
            
            messagebox.showerror("Connection Failed", error_msg)
            self.log_message(f"Connection failed: {str(e)}")

    def send_test_email(self):
        """Send test email to yourself with better feedback"""
        email = self.sender_email_var.get().strip()
        password = self.sender_pass_var.get().strip()
        
        # Try to get password from vault if not entered
        if not password and email:
            password = self.get_credentials_from_vault(email)
            if password:
                self.sender_pass_var.set(password)
                self.log_message(f"Auto-loaded password for {email} for test email")
        
        if not email or not password:
            messagebox.showwarning("Settings Required", 
                                 "Please enter your email and password first.")
            return
        
        # First test connection
        try:
            self.log_message("Testing connection before sending test email...")
            server = smtplib.SMTP(self.smtp_server_var.get(), 
                                int(self.smtp_port_var.get()), 
                                timeout=10)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(email, password)
            server.quit()
        except Exception as e:
            messagebox.showerror("Connection Error", 
                               f"❌ Cannot send test email:\n{str(e)}\n\n"
                               "Please fix your connection settings first.")
            return
        
        # Send test email
        test_recipient = {
            "company": "Test Company",
            "hr_name": "Test Recipient",
            "email": email,  # Send to self
            "position": "Test Position",
            "status": "TEST",
            "send_count": 0,
            "max_sends": 1,
            "stop_resend": False
        }
        
        try:
            success = self.send_single_email(test_recipient, [])
            
            if success:
                messagebox.showinfo("Test Email", 
                                  f"✅ Test email sent successfully!\n\n"
                                  f"To: {email}\n"
                                  f"Subject: Application for Test Position at Test Company\n\n"
                                  "Check your inbox (and spam folder).")
                self.log_message(f"Test email sent to {email}")
            else:
                messagebox.showerror("Test Failed", 
                                   "❌ Failed to send test email.\n"
                                   "Check the Activity Log for details.")
        except Exception as e:
            messagebox.showerror("Test Failed", 
                               f"❌ Failed to send test email:\n{str(e)}")
            self.log_message(f"Test email failed: {str(e)}")

    def add_attachment(self, file_type=None):
        file_path = filedialog.askopenfilename(
            title="Select Document",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Word documents", "*.docx *.doc"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            filename = os.path.basename(file_path)
            file_abs_path = os.path.abspath(file_path)
            
            # Check if file is already in attachments (by absolute path)
            if file_abs_path in [os.path.abspath(att) for att in self.attachments]:
                messagebox.showinfo("Already Added", f"File '{filename}' is already in the attachments list.")
                return
            
            # Add to attachments list
            self.attachments.append(file_path)
            
            # Add icon based on file type
            icon = "📎"
            if file_path.lower().endswith('.pdf'):
                icon = "📕"
            elif file_path.lower().endswith(('.doc', '.docx')):
                icon = "📘"
            elif file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
                icon = "🖼️"
            elif 'resume' in filename.lower() or file_type == "resume":
                icon = "📄"
            elif 'cover' in filename.lower() or 'letter' in filename.lower() or file_type == "cover":
                icon = "📝"
            
            self.attach_listbox.insert(tk.END, f"{icon} {filename}")
            self.log_message(f"Added attachment: {filename}")
            
            # Update email preview to show new attachments
            self.update_email_preview()

    def show_file_preview(self, event):
        """Show preview of selected file"""
        selection = self.attach_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.attachments):
                file_path = self.attachments[index]
                filename = os.path.basename(file_path)
                
                self.preview_text.config(state='normal')
                self.preview_text.delete(1.0, tk.END)
                
                try:
                    if file_path.lower().endswith('.txt'):
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                        self.preview_text.insert(1.0, f"📄 File: {filename}\n\n{content}")
                    else:
                        # For non-text files, show file info
                        file_size = os.path.getsize(file_path)
                        file_date = datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M")
                        file_info = f"📄 File: {filename}\n📁 Type: {os.path.splitext(filename)[1]}\n📏 Size: {file_size:,} bytes\n📅 Created: {file_date}\n\n⚠️ Preview not available for this file type."
                        self.preview_text.insert(1.0, file_info)
                except Exception as e:
                    self.preview_text.insert(1.0, f"❌ Error reading file: {str(e)}")
                
                self.preview_text.config(state='disabled')

    def preview_selected_file(self):
        """Preview the selected file"""
        selection = self.attach_listbox.curselection()
        if not selection:
            messagebox.showinfo("No Selection", "Please select a file to preview.")
            return
        self.show_file_preview(None)

    def remove_attachment(self):
        selected = self.attach_listbox.curselection()
        if selected:
            index = selected[0]
            filename = self.attach_listbox.get(index)
            self.attach_listbox.delete(index)
            if index < len(self.attachments):
                removed_file = self.attachments.pop(index)
                self.log_message(f"Removed attachment: {os.path.basename(removed_file)}")
                
                # Clear preview
                self.preview_text.config(state='normal')
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(1.0, "Select a file to preview...")
                self.preview_text.config(state='disabled')
                
                # Update email preview
                self.update_email_preview()

    def clear_attachments(self):
        if messagebox.askyesno("Clear All", "Remove all attachments?"):
            self.attach_listbox.delete(0, tk.END)
            self.attachments.clear()
            
            # Clear preview
            self.preview_text.config(state='normal')
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "Select a file to preview...")
            self.preview_text.config(state='disabled')
            
            # Update email preview
            self.update_email_preview()
            
            self.log_message("Cleared all attachments")

    def save_as_draft(self, recipient=None):
        """Save email as draft"""
        drafts = self.load_drafts()
        
        draft = {
            "subject": self.subject_var.get(),
            "body": self.body_text.get("1.0", tk.END),
            "recipient": recipient,
            "attachments": self.attachments.copy(),
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        drafts.append(draft)
        
        with open(self.drafts_file, 'w') as f:
            json.dump(drafts, f, indent=2)
        
        self.log_message(f"Saved draft for {recipient['company'] if recipient else 'unknown'}")

    def load_drafts(self):
        """Load saved drafts"""
        try:
            if os.path.exists(self.drafts_file):
                with open(self.drafts_file, 'r') as f:
                    return json.load(f)
            else:
                return []
        except Exception as e:
            print(f"Error loading drafts: {e}")
            return []

    def load_draft(self):
        """Load a saved draft"""
        drafts = self.load_drafts()
        if not drafts:
            messagebox.showinfo("No Drafts", "No saved drafts found.")
            return
        
        # Create draft selection dialog
        draft_window = tk.Toplevel(self.root)
        draft_window.title("Load Draft")
        draft_window.geometry("500x400")
        draft_window.configure(bg='white')
        
        tk.Label(draft_window,
                text="Select a draft to load:",
                font=self.fonts["heading"],
                bg='white',
                fg=self.colors["primary"]).pack(pady=20)
        
        # Listbox for drafts
        listbox = tk.Listbox(draft_window,
                           bg='white',
                           fg=self.colors["text"],
                           font=self.fonts["normal"],
                           height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        for i, draft in enumerate(drafts):
            recipient_info = draft['recipient']['company'] if draft['recipient'] else "No recipient"
            listbox.insert(tk.END, f"{draft['created']} - {recipient_info}")
        
        def load_selected_draft():
            selection = listbox.curselection()
            if selection:
                draft = drafts[selection[0]]
                self.subject_var.set(draft['subject'])
                self.body_text.delete(1.0, tk.END)
                self.body_text.insert(1.0, draft['body'])
                self.attachments = draft['attachments'].copy()
                
                # Update attachments listbox
                self.attach_listbox.delete(0, tk.END)
                for att in self.attachments:
                    self.attach_listbox.insert(tk.END, os.path.basename(att))
                
                draft_window.destroy()
                messagebox.showinfo("Draft Loaded", "✅ Draft loaded successfully!")
                self.log_message("Loaded saved draft")
                
                # Update email preview
                self.update_email_preview()
        
        tk.Button(draft_window,
                 text="Load Draft",
                 command=load_selected_draft,
                 bg=self.colors["secondary"],
                 fg='white',
                 font=self.fonts["normal"]).pack(pady=20)

    def log_message(self, message):
        """Add message to activity log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def update_stats(self):
        """Update statistics display"""
        total = len(self.sent_emails)
        successful = len([e for e in self.sent_emails if e['status'] == 'SENT'])
        pending = len([r for r in self.recipients if r['status'] == 'PENDING'])
        failed = len([e for e in self.sent_emails if 'FAILED' in e['status']])
        
        # Calculate resends (total sends - unique recipients sent to)
        unique_recipients = set()
        for email in self.sent_emails:
            unique_recipients.add(email['to'])
        resends = total - len(unique_recipients)
        
        success_rate = (successful / total * 100) if total > 0 else 0
        
        self.stats_vars["total_sent"].set(str(total))
        self.stats_vars["success_rate"].set(f"{success_rate:.1f}%")
        self.stats_vars["pending"].set(str(pending))
        self.stats_vars["failed"].set(str(failed))
        self.stats_vars["resends"].set(str(resends))  # NEW: Update resends count

    def create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Create tooltip window
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            label = tk.Label(self.tooltip, text=text, bg="white", relief="solid", borderwidth=1)
            label.pack()
        
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
        
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def open_email(self, email):
        """Open default email client"""
        webbrowser.open(f"mailto:{email}")

    def save_settings(self):
        """Save settings to file - WITHOUT SAVING PASSWORD"""
        email = self.sender_email_var.get().strip()
        password = self.sender_pass_var.get().strip()
        
        if not email:
            messagebox.showwarning("Settings Required", "Please enter your email address.")
            return
        
        if not password:
            messagebox.showwarning("Settings Required", "Please enter your App Password.")
            return
        
        try:
            # Save credentials to Windows Credential Manager
            success = self.save_credentials_to_vault(email, password)
            
            if not success:
                messagebox.showerror("Error", "Could not save credentials securely.")
                return
            
            # Save other settings (without password) to JSON file
            settings = {
                'sender_email': email,
                'sender_name': self.your_name_var.get(),
                'sender_title': self.your_title_var.get(),
                'smtp_server': self.smtp_server_var.get(),
                'smtp_port': self.smtp_port_var.get(),
                'send_time': self.send_time_var.get(),
                'interval_days': self.interval_days_var.get(),
                'max_resends': self.max_resends_var.get(),
                'show_password': self.show_password_var.get()
            }
            
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
            
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            
            # Update connection status
            self.connection_status_var.set("⚙️ Configured & Saved")
            self.connection_status_label.config(fg=self.colors["success"])
            
            # Clear password field after saving (for security)
            self.sender_pass_var.set("")
            
            messagebox.showinfo("Success", 
                              f"✅ Settings saved securely!\n\n"
                              f"Email: {email}\n"
                              f"Password: 🔒 Saved to Windows Credential Manager\n\n"
                              f"Resend Settings:\n"
                              f"• Interval: Every {self.interval_days_var.get()} day(s)\n"
                              f"• Maximum: {self.max_resends_var.get()} time(s)\n\n"
                              f"Use 'Load Saved Credentials' button to retrieve your password.")
            self.log_message(f"Settings saved securely for {email}")
            
        except Exception as e:
            print(f"Error saving settings: {e}")
            messagebox.showerror("Error", f"Could not save settings: {str(e)}")

    def load_settings(self):
        """Load settings from file - WITHOUT LOADING PASSWORD FROM FILE"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                
                # Set the values in tkinter variables (NO PASSWORD LOADED FROM FILE)
                self.sender_email_var.set(settings.get('sender_email', ''))
                self.your_name_var.set(settings.get('sender_name', ''))
                self.your_title_var.set(settings.get('sender_title', ''))
                self.smtp_server_var.set(settings.get('smtp_server', 'smtp.gmail.com'))
                self.smtp_port_var.set(settings.get('smtp_port', '587'))
                self.send_time_var.set(settings.get('send_time', '09:00'))
                self.interval_days_var.set(settings.get('interval_days', '1'))
                self.max_resends_var.set(settings.get('max_resends', '1'))
                self.show_password_var.set(settings.get('show_password', False))
                
                # Apply password visibility
                if self.show_password_var.get():
                    if hasattr(self, 'password_entry'):
                        self.password_entry.config(show="")
                else:
                    if hasattr(self, 'password_entry'):
                        self.password_entry.config(show="*")
                
                # Initialize connection status
                email = settings.get('sender_email', '')
                if email:
                    self.connection_status_var.set("⚙️ Configured (Load Creds)")
                    if hasattr(self, 'connection_status_label'):
                        self.connection_status_label.config(fg=self.colors["warning"])
                else:
                    self.connection_status_var.set("❌ Not configured")
                    if hasattr(self, 'connection_status_label'):
                        self.connection_status_label.config(fg=self.colors["error"])
                        
                # Update email preview after loading settings
                if hasattr(self, 'update_email_preview'):
                    self.update_email_preview()
                        
        except Exception as e:
            print(f"Error loading settings: {e}")
            # Set default values
            self.sender_email_var.set('')
            self.sender_pass_var.set('')
            self.your_name_var.set('')
            self.your_title_var.set('')
            self.smtp_server_var.set('smtp.gmail.com')
            self.smtp_port_var.set('587')
            self.send_time_var.set('09:00')
            self.interval_days_var.set('1')
            self.max_resends_var.set('1')
            self.show_password_var.set(False)
            if hasattr(self, 'password_entry'):
                self.password_entry.config(show="*")

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None

    def run_scheduler(self):
        """Improved scheduler that sends based on last_sent date and max_resends"""
        last_sent_date = None
        
        while self.scheduler_running:
            try:
                now = datetime.now()
                current_time = now.strftime("%H:%M")
                current_date = now.strftime("%Y-%m-%d")
                
                # Check if it's time to send (within 1 minute of scheduled time)
                scheduled_time = self.send_time_var.get()
                if scheduled_time:
                    scheduled_hour, scheduled_minute = map(int, scheduled_time.split(':'))
                    
                    # Check if current time matches scheduled time (within 1 minute)
                    time_diff = abs((now.hour - scheduled_hour) * 60 + (now.minute - scheduled_minute))
                    
                    if time_diff <= 1:  # Within 1 minute of scheduled time
                        # Only send once per day
                        if current_date != last_sent_date:
                            self.log_message(f"Scheduler triggered at {current_time}")
                            self.send_scheduled_emails()
                            last_sent_date = current_date
                            time.sleep(65)  # Wait 65 seconds to avoid duplicate sends
                
                time.sleep(30)  # Check every 30 seconds
                
            except Exception as e:
                print(f"Scheduler error: {e}")
                time.sleep(60)  # Wait a minute on error

    def send_scheduled_emails(self):
        """Send scheduled emails based on last_sent date and max_resends"""
        try:
            # Get settings
            interval_days = int(self.interval_days_var.get())
            now = datetime.now()
            
            # Find recipients that should receive emails today
            recipients_to_send = []
            
            for recipient in self.recipients:
                # Skip if stopped or already sent maximum times
                if recipient.get('stop_resend', False):
                    continue
                    
                if recipient['send_count'] >= recipient['max_sends']:
                    continue
                
                # If never sent before, send now
                if not recipient['last_sent']:
                    recipients_to_send.append(recipient)
                    continue
                
                # Parse last sent date
                try:
                    last_sent_str = recipient['last_sent'].split()[0]  # Get date part only
                    last_sent_date = datetime.strptime(last_sent_str, "%Y-%m-%d")
                    
                    # Calculate days since last sent
                    days_since_last = (now - last_sent_date).days
                    
                    # If enough days have passed, send again
                    if days_since_last >= interval_days:
                        recipients_to_send.append(recipient)
                except Exception as e:
                    # If there's an error parsing date, send anyway
                    self.log_message(f"Error parsing last_sent for {recipient['email']}: {e}")
                    recipients_to_send.append(recipient)
            
            if not recipients_to_send:
                self.log_message("No recipients due for scheduled sending today")
                return
            
            self.log_message(f"Starting scheduled send for {len(recipients_to_send)} recipients")
            
            # Get credentials from vault
            email = self.sender_email_var.get()
            password = self.get_credentials_from_vault(email)
            
            if not password:
                self.log_message("❌ Cannot send scheduled emails: No saved credentials found")
                return
            
            # Temporarily set password
            original_password = self.sender_pass_var.get()
            self.sender_pass_var.set(password)
            
            # Send emails
            success_count = 0
            for recipient in recipients_to_send:
                try:
                    self.log_message(f"Scheduled send to {recipient['company']} ({recipient['email']})")
                    
                    # Use saved attachments for this recipient if available
                    saved_attachments = recipient.get('attachments', [])
                    attachments_to_use = saved_attachments if saved_attachments else self.get_attachments_for_sending()
                    
                    success = self.send_single_email(recipient, attachments_to_use)
                    
                    if success:
                        success_count += 1
                    else:
                        self.log_message(f"❌ Failed scheduled send to {recipient['email']}")
                    
                    time.sleep(2)  # Delay between emails
                    
                except Exception as e:
                    self.log_message(f"❌ Error in scheduled send to {recipient['email']}: {e}")
            
            # Restore password
            self.sender_pass_var.set("")
            
            # Save recipients data after sending
            self.save_recipients_data()
            
            self.log_message(f"✅ Scheduled send completed: {success_count}/{len(recipients_to_send)} sent successfully")
            
        except Exception as e:
            self.log_message(f"❌ Error in scheduled send: {e}")
            print(f"Scheduled send error: {e}")

    def preview_email(self):
        """Preview email in a separate window"""
        preview = tk.Toplevel(self.root)
        preview.title("Email Preview")
        preview.geometry("600x500")
        preview.configure(bg='white')
        
        text = scrolledtext.ScrolledText(preview,
                                       bg='white',
                                       fg=self.colors["text"],
                                       font=self.fonts["normal"],
                                       wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Get email content
        subject = self.subject_var.get()
        body = self.body_text.get("1.0", tk.END)
        
        # Replace tags with sample data
        sample_data = {
            "{company}": "Sample Company",
            "{hr_name}": "Sample HR Manager",
            "{position}": "Sample Position",
            "{date}": datetime.now().strftime("%B %d, %Y"),
            "{your_name}": self.your_name_var.get() or "Your Name",
            "{your_title}": self.your_title_var.get() or "Your Title"
        }
        
        for tag, value in sample_data.items():
            subject = subject.replace(tag, value)
            body = body.replace(tag, value)
        
        preview_content = f"""SUBJECT: {subject}

{body}

---
This is a preview. Actual emails will use recipient-specific data."""
        
        text.insert("1.0", preview_content)
        text.config(state='disabled')

    def save_template(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                               filetypes=[("Text files", "*.txt")])
        if file_path:
            content = self.body_text.get("1.0", tk.END)
            with open(file_path, 'w') as f:
                f.write(content)
            messagebox.showinfo("Success", "✅ Template saved!")
            self.log_message("Template saved")

    def load_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as f:
                content = f.read()
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", content)
            messagebox.showinfo("Success", "✅ Template loaded!")
            self.log_message("Template loaded")

    def log_sent_email(self, email_data):
        """Log sent email to history file"""
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.history_file), exist_ok=True)
            
            # Include send count in the log
            send_info = f"{email_data['send_count']} of ?" if 'send_count' in email_data else "1"
            
            with open(self.history_file, 'a') as f:
                f.write(f"{email_data['date']},{email_data['time']},{email_data['to']},{email_data['company']},{email_data['subject']},{email_data['status']},{send_info}\n")
                
            # Also add to sent_emails list for immediate access
            self.sent_emails.append(email_data)
            
        except Exception as e:
            print(f"Error logging email: {e}")

    def load_history(self):
        """Load email history from CSV file and populate sent_emails list"""
        try:
            self.history_tree.delete(*self.history_tree.get_children())
            self.sent_emails.clear()  # Clear existing sent emails
            
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    for line in f:
                        parts = line.strip().split(',')
                        if len(parts) >= 6:
                            # Add to sent_emails list
                            email_data = {
                                'date': parts[0],
                                'time': parts[1],
                                'to': parts[2],
                                'company': parts[3],
                                'subject': parts[4],
                                'status': parts[5]
                            }
                            self.sent_emails.append(email_data)
                            
                            # Add to history treeview
                            status_display = "✅ SENT" if parts[5] == "SENT" else f"❌ {parts[5]}"
                            if len(parts) > 6:
                                status_display += f" ({parts[6]})"
                                
                            self.history_tree.insert("", tk.END, values=(
                                parts[0], parts[1], parts[2], parts[3], 
                                parts[4][:50] + "..." if len(parts[4]) > 50 else parts[4],
                                status_display
                            ))
                self.log_message(f"Loaded {len(self.sent_emails)} sent emails from history")
            else:
                self.log_message("No history file found. Starting fresh.")
                
            self.update_stats()
            
        except Exception as e:
            print(f"Error loading history: {e}")
            self.log_message(f"Error loading history: {str(e)}")

    def clear_history(self):
        if messagebox.askyesno("Clear History", "Delete all history records?"):
            self.history_tree.delete(*self.history_tree.get_children())
            try:
                if os.path.exists(self.history_file):
                    os.remove(self.history_file)
                self.sent_emails.clear()
                self.update_stats()
                self.log_message("Cleared all history")
            except Exception as e:
                print(f"Error clearing history: {e}")

    def export_history(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                               filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                import shutil
                if os.path.exists(self.history_file):
                    shutil.copy(self.history_file, file_path)
                    messagebox.showinfo("Success", "✅ History exported successfully!")
                    self.log_message(f"Exported history to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"❌ Export failed: {str(e)}")

    def import_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                df = pd.read_csv(file_path)
                imported_count = 0
                for _, row in df.iterrows():
                    # Check if recipient already exists
                    email = str(row.get('email', ''))
                    exists = False
                    for recipient in self.recipients:
                        if recipient['email'] == email:
                            exists = True
                            break
                    
                    if not exists and email:
                        # Get max sends from settings for new recipients
                        max_sends = int(self.max_resends_var.get())
                        
                        recipient = {
                            "company": str(row.get('company', '')),
                            "hr_name": str(row.get('hr_name', '')),
                            "email": email,
                            "position": str(row.get('position', '')),
                            "notes": str(row.get('notes', '')),
                            "status": "PENDING",
                            "last_sent": "",
                            "added_date": datetime.now().strftime("%Y-%m-%d"),
                            "send_count": 0,
                            "max_sends": max_sends,
                            "stop_resend": False,
                            "email_template": "",
                            "email_subject": "",
                            "attachments": []
                        }
                        self.recipients.append(recipient)
                        
                        status_display = f"PENDING (0/{max_sends})"
                        self.recipient_tree.insert("", tk.END, values=(
                            "",  # Empty checkbox (not selected)
                            recipient['company'],
                            recipient['hr_name'],
                            recipient['email'],
                            recipient['position'],
                            status_display,
                            recipient['added_date'],
                            ""  # Empty stop column
                        ))
                        imported_count += 1
                
                messagebox.showinfo("Success", f"✅ Imported {imported_count} new records")
                self.log_message(f"Imported {imported_count} recipients from CSV")
                
                # Save recipients data after import
                self.save_recipients_data()
                
                # Update the selected listbox
                self.update_selected_listbox()
            except Exception as e:
                messagebox.showerror("Error", f"❌ Import failed: {str(e)}")

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                               filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                data = []
                for recipient in self.recipients:
                    data.append({
                        'company': recipient['company'],
                        'hr_name': recipient['hr_name'],
                        'email': recipient['email'],
                        'position': recipient['position'],
                        'notes': recipient['notes'],
                        'status': recipient['status'],
                        'last_sent': recipient['last_sent'],
                        'added_date': recipient['added_date'],
                        'send_count': recipient['send_count'],
                        'max_sends': recipient['max_sends'],
                        'stop_resend': recipient.get('stop_resend', False)
                    })
                df = pd.DataFrame(data)
                df.to_csv(file_path, index=False)
                messagebox.showinfo("Success", "✅ Data exported successfully!")
                self.log_message(f"Exported {len(data)} recipients to CSV")
            except Exception as e:
                messagebox.showerror("Error", f"❌ Export failed: {str(e)}")

    def remove_recipient(self):
        selected = self.recipient_tree.selection()
        if selected:
            item = self.recipient_tree.item(selected[0])
            email = item['values'][3]
            self.recipient_tree.delete(selected[0])
            self.recipients = [r for r in self.recipients if r['email'] != email]
            
            # Save recipients data after removal
            self.save_recipients_data()
            
            self.update_selected_listbox()
            messagebox.showinfo("Success", "✅ Recipient removed")
            self.log_message(f"Removed recipient: {email}")

    def clear_recipients(self):
        if messagebox.askyesno("Clear All", "Remove all recipients?"):
            for item in self.recipient_tree.get_children():
                self.recipient_tree.delete(item)
            self.recipients.clear()
            
            # Save empty recipients data
            self.save_recipients_data()
            
            self.update_selected_listbox()
            self.log_message("Cleared all recipients")

if __name__ == "__main__":
    # Check if keyring is available
    try:
        import keyring
        print("✅ Secure credential storage available")
    except ImportError:
        print("❌ keyring module not installed. Installing...")
        import subprocess
        import sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "keyring"])
        import keyring
    
    root = tk.Tk()
    
    # Center window
    root.update_idletasks()
    width = 1200
    height = 800
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Status bar
    status_var = tk.StringVar()
    status_var.set("Ready to automate job applications - Secure Mode with Resend Feature")
    status_bar = tk.Frame(root, bg="#bdc3c7", height=2)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    status_text_frame = tk.Frame(root, bg="#f5f7fa", height=30)
    status_text_frame.pack(side=tk.BOTTOM, fill=tk.X)
    status_text_frame.pack_propagate(False)
    
    status_label = tk.Label(status_text_frame,
                          textvariable=status_var,
                          bg="#f5f7fa",
                          fg="#2c3e50",
                          font=("Segoe UI", 9))
    status_label.pack(side=tk.LEFT, padx=15)
    
    app = JobApplicationSender(root)
    
    # Link status variable
    app.status_var = status_var
    
    root.mainloop()