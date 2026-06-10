#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard Emailer Automation
Downloads ticket data, reads the local dashboard, and sends it via Outlook.
"""

import os
import sys
import platform
import subprocess
import urllib.request
from datetime import datetime
from dotenv import load_dotenv

# Path configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DASHBOARD_PATH = os.path.join(BASE_DIR, "index.html")
ENV_PATH = os.path.join(os.path.dirname(BASE_DIR), ".env")

# Ticket Files Configuration (Names as used in the project)
TICKET_FILES = {
    "URL_INCIDENT": "incident.xlsx",
    "URL_PROBLEM": "problem_rca.xlsx",
    "URL_TASK": "sc_task.xlsx"
}

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def download_file(url, target_name):
    """Downloads a file from a URL to the local directory."""
    target_path = os.path.join(BASE_DIR, target_name)
    try:
        log(f"Baixando {target_name} de {url[:40]}...")
        # Use a User-Agent to avoid blocks from some servers (like ServiceNow)
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response, open(target_path, 'wb') as out_file:
            data = response.read()
            out_file.write(data)
        log(f"✓ {target_name} baixado com sucesso.")
        return target_path
    except Exception as e:
        log(f"❌ Erro ao baixar {target_name}: {e}")
        return None

def send_outlook_windows(to, cc, subject, html_body, attachments):
    """Sends email via Outlook COM interface on Windows."""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        # Priority mapping (1=Low, 2=Normal, 3=High)
        try:
            importance = int(os.getenv("importance", 2))
            mail.Importance = importance
        except:
            mail.Importance = 2

        for path in attachments:
            if path and os.path.exists(path):
                mail.Attachments.Add(os.path.abspath(path))
        
        mail.Send()
        return True, "E-mail enviado com sucesso via Outlook (Windows)."
    except Exception as e:
        return False, f"Erro no Windows/Outlook: {e}"

def send_outlook_mac(to, cc, subject, html_body, attachments):
    """Fallback for Mac using AppleScript to control Outlook."""
    try:
        # Note: HTML via AppleScript is challenging. 
        # For Mac, we often create a draft or use a specialized library.
        # This implementation creates a draft with attachments and the HTML body as plain text fallback
        # or simplified rich text if possible.
        
        attachment_scripts = []
        for path in attachments:
            if path and os.path.exists(path):
                abs_path = os.path.abspath(path)
                attachment_scripts.append(f'make new attachment with properties {{file name:"{abs_path}"}} at end of attachments')

        attachments_str = "\n            ".join(attachment_scripts)
        
        # Simplified AppleScript for Mac Outlook
        as_script = f'''
        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{html_body}"}}
            tell newMessage
                make new recipient at end of recipients with properties {{email address:{{address:"{to}"}}}}
                {f'make new recipient at end of cc recipients with properties {{email address:{{address:"{cc}"}}}}' if cc else ""}
                {attachments_str}
            end tell
            send newMessage
        end tell
        '''
        subprocess.run(['osascript', '-e', as_script], check=True)
        return True, "Mensagem enviada via Outlook for Mac."
    except Exception as e:
        return False, f"Erro no Mac/Outlook: {e}"

def main():
    # 1. Load Configuration
    if os.path.exists(ENV_PATH):
        load_dotenv(ENV_PATH)
        log("Arquivo .env carregado.")
    else:
        log("⚠️ Arquivo .env não encontrado no diretório raiz.")

    # 2. Download Ticket Files
    downloaded_paths = []
    for env_key, file_name in TICKET_FILES.items():
        url = os.getenv(env_key)
        if url:
            path = download_file(url, file_name)
            if path:
                downloaded_paths.append(path)
        else:
            log(f"⚠️ URL não definida para {env_key} no .env. Pulando download.")
            # Verify if local file exists to attach anyway
            local_path = os.path.join(BASE_DIR, file_name)
            if os.path.exists(local_path):
                downloaded_paths.append(local_path)

    # 3. Read Dashboard HTML
    if not os.path.exists(DASHBOARD_PATH):
        log("❌ Dashboard index.html não encontrado. Encerrando.")
        return

    with open(DASHBOARD_PATH, 'r', encoding='utf-8') as f:
        html_body = f.read()
    log("✓ Dashbord HTML lido com sucesso.")

    # 4. Prepare Email Metadata
    to = os.getenv("destinatario", "")
    cc = os.getenv("copia", "")
    subject_prefix = os.getenv("titulo", "Dashboard ITSM Report")
    subject = f"{subject_prefix} - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    if not to:
        log("❌ Variável 'destinatario' não definida no .env. E-mail não enviado.")
        return

    # 5. Delivery
    system = platform.system()
    log(f"Iniciando envio no sistema: {system}")
    
    if system == "Windows":
        success, msg = send_outlook_windows(to, cc, subject, html_body, downloaded_paths)
    elif system == "Darwin": # Mac
        success, msg = send_outlook_mac(to, cc, subject, html_body, downloaded_paths)
    else:
        success, msg = False, f"Sistema {system} não suportado para disparos automáticos de Outlook."

    if success:
        log(f"✨ {msg}")
    else:
        log(f"🚨 {msg}")

if __name__ == "__main__":
    main()
