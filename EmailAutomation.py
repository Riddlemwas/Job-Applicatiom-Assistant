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
 
       
    

