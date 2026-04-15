#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
import win32com.client
from dotenv import load_dotenv

# ============================================================================
# CONFIGURATION & ENVIRONMENT VARIABLES
# ============================================================================
load_dotenv()

EMAIL_SUBJECT = os.getenv("titulo", "ServiceNow Automated Report")
EMAIL_TO = os.getenv("destinatario", "")
EMAIL_CC = os.getenv("copia", "")
EMAIL_IMPORTANCE = os.getenv("importance", 2)

DOWNLOAD_DIR = "path_temp"
HTML_REPORT_PATH = "index_v1.2.html"

# TODO: Replace with your actual ServiceNow export URLs
# Example: "https://[instance].service-now.com/incident.do?EXCEL&sysparm_query=active=true"
FILES_TO_DOWNLOAD = {
    "incident.xls": "URL_FOR_INCIDENT_XLS",
    "problem_rca.xls": "URL_FOR_PROBLEM_RCA_XLS",
    "sc_task.xls": "URL_FOR_SC_TASK_XLS"
}

# ServiceNow Credentials (if basic auth is needed for direct download)
SNOW_USER = os.getenv("SNOW_USER", "")
SNOW_PASS = os.getenv("SNOW_PASS", "")

# ============================================================================
# FUNCTIONS
# ============================================================================

def ensure_dir(path: str):
    """Ensures that the directory exists."""
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"Directory created: {path}")

def download_file(url: str, dest_path: str) -> bool:
    """Downloads a file from a URL to the specified destination."""
    print(f"Downloading {dest_path}...")
    try:
        # If your ServiceNow instance requires authentication, pass auth=(SNOW_USER, SNOW_PASS)
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        
        with open(dest_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                
        print(f"Successfully downloaded: {dest_path}")
        return True
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def send_email_with_attachments(body_html: str, attachments: list):
    """Sends an email via Outlook with the given HTML body and attachments."""
    print("\n--- Preparing Outlook Email ---")
    
    if not EMAIL_TO:
        print("Warning: 'destinatario' is empty in .env. Email might fail.")

    try:
        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        
        mail.To = EMAIL_TO
        if EMAIL_CC:
            mail.CC = EMAIL_CC
            
        mail.Subject = EMAIL_SUBJECT
        try:
            mail.Importance = int(EMAIL_IMPORTANCE)
        except ValueError:
            mail.Importance = 2
            
        mail.HTMLBody = body_html

        # Attach all valid files
        attached_count = 0
        for file_path in attachments:
            if os.path.exists(file_path):
                abs_path = os.path.abspath(file_path)
                mail.Attachments.Add(abs_path)
                attached_count += 1
                print(f"Attached: {file_path}")
            else:
                print(f"Warning: Attachment not found - {file_path}")

        mail.Send()
        print(f"Email sent successfully with {attached_count} attachments!")
        
    except Exception as e:
        print(f"Critical error sending email: {e}")

# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    print("=" * 80)
    print("STARTING DOWNLOAD AND EMAIL PROCESS")
    print("=" * 80)

    ensure_dir(DOWNLOAD_DIR)
    downloaded_files = []

    # 1. Download the Excel files
    for filename, url in FILES_TO_DOWNLOAD.items():
        file_path = os.path.join(DOWNLOAD_DIR, filename)
        success = download_file(url, file_path)
        if success:
            downloaded_files.append(file_path)

    # 2. Add the HTML index to the attachments list
    if os.path.exists(HTML_REPORT_PATH):
        downloaded_files.append(HTML_REPORT_PATH)
    else:
        print(f"Warning: {HTML_REPORT_PATH} not found. It will not be attached.")

    # 3. Read the HTML file to use as the email body (optional, or you can use a simple message)
    email_body = ""
    if os.path.exists(HTML_REPORT_PATH):
        with open(HTML_REPORT_PATH, 'r', encoding='utf-8') as f:
            email_body = f.read()
    else:
        email_body = "<p>Please find the requested reports and the dashboard attached.</p>"

    # 4. Send the email
    if downloaded_files:
        send_email_with_attachments(body_html=email_body, attachments=downloaded_files)
    else:
        print("No files were prepared. Email process aborted.")

if __name__ == "__main__":
    main()
