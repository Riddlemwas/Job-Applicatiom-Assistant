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
        

        self.app_id = self.get_or_create_app_id()
        
     
        self.set_application_icon()
        
     
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
        
     
        home_dir = os.path.expanduser("~")
        app_data_dir = os.path.join(home_dir, "JobApplicationAssistant")
        
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
      
        self.settings_file = os.path.join(app_data_dir, "job_app_settings.json")
        self.history_file = os.path.join(app_data_dir, "sent_history.csv")
        self.drafts_file = os.path.join(app_data_dir, "email_drafts.json")
        self.recipients_file = os.path.join(app_data_dir, "recipients_data.json") 
        
        # Connection status
        self.connection_status_var = None
        self.connection_status_label = None
        
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
