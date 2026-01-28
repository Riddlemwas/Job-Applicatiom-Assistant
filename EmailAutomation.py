"""
Job Application Assistant - Python Automation Tool
Version: Public Edition
Description: A GUI application for automating job application emails with scheduling and tracking features.
"""

# Basic structure overview - Simplified for public sharing
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import smtplib
import schedule
import threading
import time
from datetime import datetime
import os
import json

class JobApplicationSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Application Assistant")
        self.root.geometry("1200x800")
        
        # Data storage
        self.recipients = []
        self.attachments = []
        self.sent_emails = []
        
        # Setup directories for user data
        self.setup_data_directory()
        
        # Initialize GUI
        self.create_widgets()
        
        # Load existing data
        self.load_data()
        
        # Start background scheduler
        self.start_scheduler()

    def setup_data_directory(self):
        """Setup directory for storing user data"""
        home_dir = os.path.expanduser("~")
        self.data_dir = os.path.join(home_dir, "JobApplicationAssistant")
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        # Define data file paths
        self.recipients_file = os.path.join(self.data_dir, "recipients.json")
        self.history_file = os.path.join(self.data_dir, "history.csv")
        self.settings_file = os.path.join(self.data_dir, "settings.json")

    def create_widgets(self):
        """Create main application interface"""
        # Main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create application tabs
        self.create_recipients_tab()
        self.create_email_tab()
        self.create_attachments_tab()
        self.create_settings_tab()
        self.create_send_tab()
        self.create_history_tab()
        self.create_about_tab()

    def create_recipients_tab(self):
        """Tab for managing job application recipients"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📋 Recipients")
        
        # Recipient entry form
        form_frame = tk.Frame(tab, bg='white')
        form_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Form fields
        fields = ["Company", "Contact", "Email", "Position", "Notes"]
        self.recipient_vars = {}
        
        for i, field in enumerate(fields):
            tk.Label(form_frame, text=f"{field}:", bg='white').grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar()
            entry = tk.Entry(form_frame, textvariable=var, width=40)
            entry.grid(row=i, column=1, sticky=tk.W, pady=5, padx=10)
            self.recipient_vars[field.lower()] = var
        
        # Add recipient button
        add_btn = tk.Button(form_frame, text="Add Recipient", 
                          command=self.add_recipient, bg="#4CAF50", fg='white')
        add_btn.grid(row=len(fields), column=1, sticky=tk.E, pady=10)
        
        # Recipients list
        list_frame = tk.Frame(tab, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Treeview for displaying recipients
        columns = ("Company", "Contact", "Email", "Position", "Status")
        self.recipient_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.recipient_tree.heading(col, text=col)
            self.recipient_tree.column(col, width=150)
        
        self.recipient_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, command=self.recipient_tree.yview)
        self.recipient_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_email_tab(self):
        """Tab for composing email templates"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📝 Email")
        
        # Subject field
        tk.Label(tab, text="Subject:", bg='white').pack(anchor=tk.W, padx=20, pady=(10, 0))
        self.subject_var = tk.StringVar(value="Application for {position} position")
        tk.Entry(tab, textvariable=self.subject_var, width=80).pack(padx=20, pady=5)
        
        # Email body
        tk.Label(tab, text="Email Body:", bg='white').pack(anchor=tk.W, padx=20, pady=(10, 0))
        self.body_text = tk.Text(tab, height=20, width=80)
        self.body_text.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        # Default template
        default_template = """Dear {contact},

I am writing to apply for the {position} position at {company}.

[Your application content here]

Best regards,
[Your Name]
[Your Title]"""
        
        self.body_text.insert("1.0", default_template)

    def create_attachments_tab(self):
        """Tab for managing file attachments"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📎 Attachments")
        
        # Attachment list
        tk.Label(tab, text="Attached Files:", bg='white').pack(anchor=tk.W, padx=20, pady=(10, 0))
        
        self.attach_listbox = tk.Listbox(tab, height=10)
        self.attach_listbox.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        # Buttons frame
        btn_frame = tk.Frame(tab, bg='white')
        btn_frame.pack(padx=20, pady=10)
        
        tk.Button(btn_frame, text="Add File", command=self.add_attachment).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remove Selected", command=self.remove_attachment).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Clear All", command=self.clear_attachments).pack(side=tk.LEFT, padx=5)

    def create_settings_tab(self):
        """Tab for email account settings"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="⚙️ Settings")
        
        # Email settings frame
        settings_frame = tk.Frame(tab, bg='white')
        settings_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Email configuration fields
        configs = [
            ("Email:", "email_var"),
            ("Password:", "pass_var"),
            ("SMTP Server:", "smtp_var"),
            ("SMTP Port:", "port_var"),
            ("Your Name:", "name_var"),
            ("Your Title:", "title_var")
        ]
        
        self.settings_vars = {}
        
        for i, (label, var_name) in enumerate(configs):
            tk.Label(settings_frame, text=label, bg='white').grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar()
            if "password" in label.lower():
                entry = tk.Entry(settings_frame, textvariable=var, show="*", width=40)
            else:
                entry = tk.Entry(settings_frame, textvariable=var, width=40)
            
            entry.grid(row=i, column=1, sticky=tk.W, pady=5, padx=10)
            self.settings_vars[var_name] = var
        
        # Set default values for SMTP
        self.settings_vars["smtp_var"].set("smtp.gmail.com")
        self.settings_vars["port_var"].set("587")
        
        # Save and test buttons
        btn_frame = tk.Frame(tab, bg='white')
        btn_frame.pack(padx=20, pady=20)
        
        tk.Button(btn_frame, text="Save Settings", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Test Connection", command=self.test_connection).pack(side=tk.LEFT, padx=5)

    def create_send_tab(self):
        """Tab for sending emails"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="🚀 Send")
        
        # Selected recipients display
        tk.Label(tab, text="Selected Recipients:", bg='white').pack(anchor=tk.W, padx=20, pady=(10, 0))
        
        self.selected_listbox = tk.Listbox(tab, height=8)
        self.selected_listbox.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        # Send controls
        control_frame = tk.Frame(tab, bg='white')
        control_frame.pack(padx=20, pady=20)
        
        tk.Button(control_frame, text="Send Emails", command=self.send_emails, 
                 bg="#2196F3", fg='white', padx=20).pack(side=tk.LEFT, padx=5)
        
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(control_frame, textvariable=self.status_var, bg='white').pack(side=tk.LEFT, padx=20)

    def create_history_tab(self):
        """Tab for viewing sent email history"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="📊 History")
        
        # History treeview
        columns = ("Date", "Time", "To", "Company", "Status")
        self.history_tree = ttk.Treeview(tab, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=120)
        
        self.history_tree.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        # History controls
        btn_frame = tk.Frame(tab, bg='white')
        btn_frame.pack(padx=20, pady=(0, 20))
        
        tk.Button(btn_frame, text="Refresh", command=self.load_history).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Export", command=self.export_history).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Clear", command=self.clear_history).pack(side=tk.LEFT, padx=5)

    def create_about_tab(self):
        """Tab with application information"""
        tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(tab, text="ℹ️ About")
        
        # Application info
        info_text = """
        Job Application Assistant
        
        Version: Public Edition
        
        Features:
        • Manage job application recipients
        • Create email templates with placeholders
        • Attach resumes and cover letters
        • Schedule automatic sending
        • Track sent applications
        
        This is a simplified version for demonstration purposes.
        """
        
        tk.Label(tab, text=info_text, bg='white', justify=tk.LEFT).pack(padx=40, pady=40)

    # ====== Core Functionality Methods ======
    
    def add_recipient(self):
        """Add a new recipient to the list"""
        # Get values from form fields
        company = self.recipient_vars.get('company', tk.StringVar()).get().strip()
        email = self.recipient_vars.get('email', tk.StringVar()).get().strip()
        
        if not company or not email:
            messagebox.showwarning("Input Required", "Company and Email are required fields.")
            return
        
        # Add to recipients list
        recipient = {
            "company": company,
            "contact": self.recipient_vars.get('contact', tk.StringVar()).get().strip(),
            "email": email,
            "position": self.recipient_vars.get('position', tk.StringVar()).get().strip(),
            "notes": self.recipient_vars.get('notes', tk.StringVar()).get().strip(),
            "status": "Pending",
            "added_date": datetime.now().strftime("%Y-%m-%d")
        }
        
        self.recipients.append(recipient)
        
        # Add to treeview
        self.recipient_tree.insert("", tk.END, values=(
            recipient["company"],
            recipient["contact"],
            recipient["email"],
            recipient["position"],
            recipient["status"]
        ))
        
        # Clear form fields
        for var in self.recipient_vars.values():
            var.set("")
        
        # Save data
        self.save_data()

    def add_attachment(self):
        """Add file attachment"""
        file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("All files", "*.*"), ("PDF files", "*.pdf"), ("Word documents", "*.docx")]
        )
        
        if file_path:
            filename = os.path.basename(file_path)
            self.attachments.append(file_path)
            self.attach_listbox.insert(tk.END, filename)

    def remove_attachment(self):
        """Remove selected attachment"""
        selection = self.attach_listbox.curselection()
        if selection:
            index = selection[0]
            self.attach_listbox.delete(index)
            if index < len(self.attachments):
                self.attachments.pop(index)

    def clear_attachments(self):
        """Remove all attachments"""
        if messagebox.askyesno("Confirm", "Remove all attachments?"):
            self.attach_listbox.delete(0, tk.END)
            self.attachments.clear()

    def send_emails(self):
        """Send emails to selected recipients"""
        if not self.recipients:
            messagebox.showwarning("No Recipients", "No recipients to send emails to.")
            return
        
        # Check email settings
        email = self.settings_vars.get("email_var", tk.StringVar()).get()
        password = self.settings_vars.get("pass_var", tk.StringVar()).get()
        
        if not email or not password:
            messagebox.showwarning("Settings Required", "Please configure email settings first.")
            return
        
        # Simulate sending (in actual app, this would connect to SMTP server)
        self.status_var.set("Sending emails...")
        time.sleep(2)  # Simulate sending delay
        
        # Update status
        self.status_var.set(f"Sent to {len(self.recipients)} recipients")
        messagebox.showinfo("Success", f"Emails sent to {len(self.recipients)} recipients")

    def test_connection(self):
        """Test email connection"""
        messagebox.showinfo("Connection Test", "Connection test would run here.")

    def save_settings(self):
        """Save application settings"""
        messagebox.showinfo("Settings Saved", "Settings saved successfully.")

    def load_history(self):
        """Load email history"""
        # This would load from history file
        pass

    def export_history(self):
        """Export history to file"""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", 
                                                 filetypes=[("CSV files", "*.csv")])
        if file_path:
            messagebox.showinfo("Export", f"History exported to {file_path}")

    def clear_history(self):
        """Clear email history"""
        if messagebox.askyesno("Confirm", "Clear all history?"):
            self.history_tree.delete(*self.history_tree.get_children())

    def save_data(self):
        """Save application data"""
        try:
            # Save recipients
            with open(self.recipients_file, 'w') as f:
                json.dump(self.recipients, f, indent=2)
        except Exception as e:
            print(f"Error saving data: {e}")

    def load_data(self):
        """Load application data"""
        try:
            if os.path.exists(self.recipients_file):
                with open(self.recipients_file, 'r') as f:
                    self.recipients = json.load(f)
                
                # Populate treeview
                for recipient in self.recipients:
                    self.recipient_tree.insert("", tk.END, values=(
                        recipient.get("company", ""),
                        recipient.get("contact", ""),
                        recipient.get("email", ""),
                        recipient.get("position", ""),
                        recipient.get("status", "Pending")
                    ))
        except Exception as e:
            print(f"Error loading data: {e}")

    def start_scheduler(self):
        """Start background scheduler for automated sending"""
        def scheduler_loop():
            while True:
                # Check for scheduled sends
                # This would check current time against scheduled times
                time.sleep(60)  # Check every minute
        
        scheduler_thread = threading.Thread(target=scheduler_loop, daemon=True)
        scheduler_thread.start()

def main():
    """Main application entry point"""
    root = tk.Tk()
    
    # Center window on screen
    root.update_idletasks()
    width = 1000
    height = 700
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create application
    app = JobApplicationSender(root)
    
    # Run application
    root.mainloop()

if __name__ == "__main__":
    main()
