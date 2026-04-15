import os
import json
import base64
import platform
import subprocess
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Allow requests from local file:// or browser

# Detect OS
IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

def send_outlook_windows(to, cc, subject, html_body, screenshot_base64=None):
    import win32com.client
    try:
        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        # Handle Screenshot Attachment
        if screenshot_base64:
            # Save temporary file
            temp_path = os.path.join(os.getcwd(), "temp_screenshot_mail.png")
            img_data = base64.b64decode(screenshot_base64.split(",")[1])
            with open(temp_path, "wb") as f:
                f.write(img_data)
            
            # Add attachment
            mail.Attachments.Add(os.path.abspath(temp_path))
            # Delete temp file after some time or just overwrite next time
            # os.remove(temp_path) # wait until send?

        mail.Display() # Open the window for the user
        # mail.Send()  # Or use mail.Send() for immediate background send
        return True, "E-mail preparado no Outlook (Windows)."
    except Exception as e:
        return False, str(e)

def send_outlook_mac(to, cc, subject, html_body):
    """Fallback for Mac using AppleScript to control Outlook."""
    try:
        # Simple AppleScript to create a draft in Outlook for Mac
        # Note: HTML via AppleScript is limited, usually plain text or rich text conversion
        as_script = f'''
        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{html_body}"}}
            make new recipient at newMessage with properties {{email address:{{address:"{to}"}}}}
            open newMessage
        end tell
        '''
        subprocess.run(['osascript', '-e', as_script], check=True)
        return True, "Draft criado no Outlook for Mac (Nota: HTML limitado em AppleScript)."
    except Exception as e:
        return False, str(e)

@app.route('/status', methods=['GET'])
def get_status():
    return jsonify({
        "status": "online",
        "platform": platform.system(),
        "ready": True
    })

@app.route('/rpa/logs', methods=['GET'])
def get_rpa_logs():
    """Lê os logs da base do robô de dados em RPA/tickets_processados.db"""
    ticket_filter = request.args.get('ticket')
    try:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        RPA_DIR = os.path.join(os.path.dirname(BASE_DIR), "RPA")
        db_path = os.path.join(RPA_DIR, "tickets_processados.db")
        
        if not os.path.exists(db_path):
            return jsonify({"status": "error", "message": f"Banco de dados RPA não criado em: {db_path}", "data": []})

        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            if ticket_filter:
                terms = [t.strip() for t in ticket_filter.split(',') if t.strip()]
                if terms:
                    # Construir query dinâmica para suportar múltiplos termos com LIKE (__contains__)
                    query = "SELECT * FROM processados WHERE "
                    conditions = []
                    params = []
                    for term in terms:
                        conditions.append("(ticket_number LIKE ? OR assunto LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    
                    query += " OR ".join(conditions)
                    query += " ORDER BY data_processamento DESC LIMIT 100"
                    cursor.execute(query, params)
                else:
                    cursor.execute("SELECT * FROM processados ORDER BY data_processamento DESC LIMIT 100")
            else:
                cursor.execute("SELECT * FROM processados ORDER BY data_processamento DESC LIMIT 100")
                
            rows = cursor.fetchall()
            data = [dict(row) for row in rows]
            return jsonify({"status": "success", "data": data})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "data": []})

@app.route('/rpa/email', methods=['GET'])
def get_rpa_email():
    """Lê o arquivo físico do e-mail do disco"""
    filepath = request.args.get('filepath')
    if not filepath or not os.path.exists(filepath):
        return jsonify({"status": "error", "message": "Caminho do arquivo não fornecido ou e-mail excluído.", "content": ""})

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        return jsonify({"status": "success", "content": content})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Erro de leitura: {str(e)}", "content": ""})

@app.route('/send-email', methods=['POST'])
def send_email():
    data = request.json
    to = data.get('to', '')
    cc = data.get('cc', '')
    subject = data.get('subject', 'Relatório Dashboard')
    html_body = data.get('body', '')
    screenshot = data.get('screenshot') # base64

    if IS_WINDOWS:
        success, msg = send_outlook_windows(to, cc, subject, html_body, screenshot)
    elif IS_MAC:
        success, msg = send_outlook_mac(to, cc, subject, html_body)
    else:
        success, msg = False, f"Sistema Operacional {platform.system()} não suportado para Outlook COM."

    return jsonify({"success": success, "message": msg})

if __name__ == "__main__":
    print(f"🚀 Dashboard Email Bridge rodando em http://localhost:5000")
    print(f"Detected OS: {platform.system()}")
    app.run(port=5000, debug=False)
