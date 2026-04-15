#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
import pandas as pd
import win32com.client
from dotenv import load_dotenv

# ============================================================================
# CONFIGURATIONS & ENVIRONMENT VARIABLES
# ============================================================================
load_dotenv()

EMAIL_SUBJECT = os.getenv("titulo", "ServiceNow Integrated Report")
EMAIL_TO = os.getenv("destinatario", "")
EMAIL_CC = os.getenv("copia", "")
EMAIL_IMPORTANCE = os.getenv("importance", 2)
MY_OWN_EMAIL = os.getenv("meu_email", "").strip().lower()

DOWNLOAD_DIR = "path_temp"
HTML_REPORT_PATH = "index_v1.2.html"

# URLs de exportação do ServiceNow (Substituir pelos links reais)
FILES_TO_DOWNLOAD = {
    "incident.xls": "URL_FOR_INCIDENT_XLS",
    "problem_rca.xls": "URL_FOR_PROBLEM_RCA_XLS",
    "sc_task.xls": "URL_FOR_SC_TASK_XLS"
}

# ============================================================================
# UTILITIES
# ============================================================================
def ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"Directory created: {path}")

def download_file(url: str, dest_path: str) -> bool:
    print(f"Downloading {dest_path}...")
    try:
        response = requests.get(url, stream=True, timeout=30) # Adicionar auth=() se necessário
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"Successfully downloaded: {dest_path}")
        return True
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

# ============================================================================
# EXCEL PARSING & EMAIL SCANNING LOGIC
# ============================================================================
def extract_tickets_from_excel(file_paths: list) -> set:
    """Varre todas as colunas de identificação e retorna um set com os números dos tickets (SCTASK, RITM)."""
    ticket_list = set()
    print("\n--- Scanning Excel Files for Tickets ---")
    
    for file in file_paths:
        if not os.path.exists(file):
            continue
        try:
            df = pd.read_excel(file)
            # Tenta encontrar a coluna principal de tickets
            col_candidates = ["Number", "Número", "Ticket", "ID"]
            target_col = next((c for c in col_candidates if c in df.columns), None)
            
            if target_col:
                tickets = df[target_col].dropna().astype(str).str.strip().tolist()
                for t in tickets:
                    if t.startswith("SCTASK") or t.startswith("RITM"):
                        ticket_list.add(t)
            print(f"Extracted tickets from {os.path.basename(file)}")
        except Exception as e:
            print(f"Failed to read {file}: {e}")
            
    return ticket_list

def scan_outlook_for_third_party_emails(tickets_to_search: set) -> dict:
    """Varre o Outlook buscando interações sobre SCTASK e RITM, excluindo o próprio e-mail rigorosamente."""
    print("\n--- Scanning Outlook for Third-Party Interactions ---")
    if not MY_OWN_EMAIL:
        print("WARNING: 'meu_email' not set in .env. Self-exclusion will not work properly.")
        
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True) # Do mais recente para o mais antigo

    found_interactions = {ticket: [] for ticket in tickets_to_search}
    scan_limit = 500 # Limita aos últimos 500 e-mails para performance
    count = 0

    for msg in messages:
        if count >= scan_limit:
            break
        try:
            # Tratamento para extrair o e-mail real do remetente
            if msg.SenderEmailType == "EX":
                sender_email = msg.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
            else:
                sender_email = msg.SenderEmailAddress.lower()
                
            # EXCLUSÃO RIGOROSA: Pula qualquer e-mail enviado por mim
            if MY_OWN_EMAIL in sender_email:
                continue

            subject = str(msg.Subject)
            body = str(msg.Body)
            
            # Verifica apenas se tem relação com os tickets mapeados
            for ticket in tickets_to_search:
                if ticket in subject or ticket in body:
                    found_interactions[ticket].append({
                        "Sender": sender_email,
                        "ReceivedTime": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                        "Subject": subject
                    })
            count += 1
        except Exception:
            # Ignora convites de calendário ou mensagens corrompidas
            continue

    # Remove tickets sem interações do log final
    filtered_interactions = {k: v for k, v in found_interactions.items() if v}
    print(f"Found third-party interactions for {len(filtered_interactions)} tickets.")
    return filtered_interactions

# ============================================================================
# MAILER
# ============================================================================
def send_final_report(attachments: list, interactions: dict):
    print("\n--- Preparing to Send Email ---")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        if EMAIL_CC:
            mail.CC = EMAIL_CC
        mail.Subject = EMAIL_SUBJECT
        
        try:
            mail.Importance = int(EMAIL_IMPORTANCE)
        except (ValueError, TypeError):
            mail.Importance = 2

        # Monta o corpo do e-mail com o sumário das interações de terceiros
        html_body = "<h3>Relatório Consolidado de Tickets</h3>"
        html_body += "<p>Segue em anexo o dashboard mais recente (index_v1.2.html) e as bases brutas atualizadas.</p>"
        
        if interactions:
            html_body += "<h4>Últimas Interações de Terceiros (SCTASK/RITM):</h4><ul>"
            for ticket, logs in interactions.items():
                html_body += f"<li><b>{ticket}</b>: {len(logs)} e-mail(s) recebido(s) de terceiros. Último remetente: {logs[0]['Sender']}</li>"
            html_body += "</ul>"
            
        mail.HTMLBody = html_body

        attached_count = 0
        for file_path in attachments:
            if os.path.exists(file_path):
                abs_path = os.path.abspath(file_path)
                mail.Attachments.Add(abs_path)
                attached_count += 1

        mail.Send()
        print(f"Success! Email sent with {attached_count} attachments.")
        
    except Exception as e:
        print(f"Critical error sending email: {e}")

# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================
def main():
    ensure_dir(DOWNLOAD_DIR)
    downloaded_files = []

    # 1. Download
    for filename, url in FILES_TO_DOWNLOAD.items():
        file_path = os.path.join(DOWNLOAD_DIR, filename)
        if download_file(url, file_path):
            downloaded_files.append(file_path)

    # 2. Varredura no Excel (Apenas se baixou arquivos)
    tickets_to_search = extract_tickets_from_excel(downloaded_files)

    # 3. Varredura rigorosa no Outlook
    interactions_log = {}
    if tickets_to_search:
        interactions_log = scan_outlook_for_third_party_emails(tickets_to_search)

    # 4. Anexa o HTML principal
    if os.path.exists(HTML_REPORT_PATH):
        downloaded_files.append(HTML_REPORT_PATH)
        print(f"Included dashboard HTML: {HTML_REPORT_PATH}")
    else:
        print(f"Warning: Dashboard HTML '{HTML_REPORT_PATH}' not found in current directory.")

    # 5. Disparo do E-mail
    if downloaded_files:
        send_final_report(downloaded_files, interactions_log)
    else:
        print("Abortando: Nenhum arquivo disponível para envio.")

if __name__ == "__main__":
    main()
