import os
import json
import base64
import platform
import subprocess
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify
from flask_cors import CORS
from critical_sla_logic import analyze_critical_slas, format_email_body

app = Flask(__name__)
CORS(app)  # Allow requests from local file:// or browser

# Detect OS
IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

def send_outlook_windows(to, cc, subject, html_body, screenshot_base64=None):
    """
    Envia e-mail de forma silenciosa no Windows via Outlook COM.
    Segue o modelo robusto solicitado pelo usuário.
    """
    import win32com.client
    try:
        # Tratamento de segurança para os parâmetros
        safe_to = str(to) if to else ""
        safe_cc = str(cc) if cc else ""
        safe_subject = str(subject) if subject else "Relatório Dashboard - " + datetime.now().strftime('%d/%m/%Y')
        
        print(f"--- Preparando envio de E-mail Silencioso via Outlook (Windows) ---")
        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        
        if safe_to:
            mail.To = safe_to
        else:
            print("⚠️ Aviso: Destinatário vazio. O envio pode falhar.")
            
        mail.Subject = safe_subject
        mail.HTMLBody = html_body
        mail.Importance = 2 # Normal
        
        if safe_cc:
            mail.CC = safe_cc
            
        # Handle Screenshot Attachment
        if screenshot_base64:
            temp_path = os.path.join(os.getcwd(), "temp_screenshot_mail.png")
            try:
                img_data = base64.b64decode(screenshot_base64.split(",")[1])
                with open(temp_path, "wb") as f:
                    f.write(img_data)
                mail.Attachments.Add(os.path.abspath(temp_path))
            except Exception as e_img:
                print(f"⚠️ Erro ao anexar screenshot: {e_img}")

        # DISPARO SILENCIOSO
        mail.Send()
        print(f"✅ E-mail enviado com sucesso para {safe_to}!")
        return True, f"E-mail enviado silenciosamente via Outlook (Windows) para {safe_to}."
    except Exception as e:
        print(f"❌ Erro crítico ao enviar e-mail (Windows): {e}")
        return False, str(e)

def send_outlook_mac(to, cc, subject, html_body):
    """
    Envia e-mail de forma silenciosa no Mac via AppleScript.
    """
    try:
        safe_to = str(to) if to else ""
        safe_cc = str(cc) if cc else ""
        
        # AppleScript para envio silencioso usando 'html content' para renderizar formatação
        # Importante: Protegemos aspas duplas no corpo HTML
        escaped_body = html_body.replace('"', '\\"')
        
        as_script = f'''
        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", html content:"{escaped_body}"}}
            make new recipient at newMessage with properties {{email address:{{address:"{safe_to}"}}}}
        '''
        
        if safe_cc:
            as_script += f'\n            make new recipient at newMessage with properties {{email address:{{address:"{safe_cc}"}}, type:cc recipient}}'
        
        as_script += '\n            send newMessage\n        end tell'
        
        subprocess.run(['osascript', '-e', as_script], check=True)
        print(f"✅ E-mail HTML enviado com sucesso (Mac) para {safe_to}!")
        return True, f"E-mail enviado silenciosamente via Outlook (Mac) para {safe_to}."
    except Exception as e:
        print(f"❌ Erro crítico ao enviar e-mail (Mac): {e}")
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

@app.route('/analyze-sla', methods=['POST'])
def analyze_sla_endpoint():
    data = request.json or {}
    config = {
        'sla_incs': int(data.get('slaIncs', 7)),
        'sla_reqs': int(data.get('slaReqs', 3)),
        'sla_ptask': int(data.get('slaPtask', 3)),
        'sla_rca': int(data.get('slaRca', 5))
    }
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    inc_path = os.path.join(base_dir, "incident.xlsx")
    task_path = os.path.join(base_dir, "sc_task.xlsx")
    
    # 1. Analyze
    analysis_results = analyze_critical_slas(inc_path, task_path, config)
    
    # 2. Send Emails & Log
    dispatched = []
    
    # Get DB path
    rpa_dir = os.path.join(os.path.dirname(base_dir), "RPA")
    db_path = os.path.join(rpa_dir, "tickets_processados.db")
    if not os.path.exists(rpa_dir): os.makedirs(rpa_dir)

    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sla_dispatches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                analyst TEXT,
                tickets TEXT,
                email_status TEXT,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        for analyst_name, res_data in analysis_results.items():
            tickets = res_data['tickets']
            analyst_email = res_data['email']
            
            subject = f"ALERTA: Chamados Críticos - SLA Próximo do Limite"
            body = format_email_body(analyst_name, tickets)
            
            # Use email from column, or fallback if empty
            to_email = str(analyst_email).strip()
            if not to_email or to_email.lower() == 'nan' or '@' not in to_email:
                # Log do aviso de e-mail não encontrado na planilha
                print(f"⚠️ Aviso: E-mail para '{analyst_name}' não encontrado ou inválido na coluna 'Email'. Usando fallback.")
                to_email = f"{analyst_name.replace(' ', '.').lower()}@empresa.com"
            
            cc_email = data.get('emailCc', '')
            
            # Send using existing logic
            if IS_WINDOWS:
                success, msg = send_outlook_windows(to_email, cc_email, subject, body)
            elif IS_MAC:
                success, msg = send_outlook_mac(to_email, cc_email, subject, body)
            else:
                success, msg = False, "OS Not Supported"
            
            # Log
            cursor.execute(
                "INSERT INTO sla_dispatches (analyst, tickets, email_status) VALUES (?, ?, ?)",
                (analyst_name, json.dumps(tickets), "Success" if success else f"Error: {msg}")
            )
            dispatched.append({"analyst": analyst_name, "tickets_count": len(tickets), "status": "Sent" if success else "Failed"})
    
    return jsonify({"status": "success", "dispatched": dispatched})

@app.route('/sla-logs', methods=['GET'])
def get_sla_logs():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        rpa_dir = os.path.join(os.path.dirname(base_dir), "RPA")
        db_path = os.path.join(rpa_dir, "tickets_processados.db")
        
        if not os.path.exists(db_path):
            return jsonify({"status": "success", "data": []})

        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM sla_dispatches ORDER BY sent_at DESC LIMIT 50")
            rows = cursor.fetchall()
            data = [dict(row) for row in rows]
            return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == "__main__":
    print(f"🚀 Dashboard Email Bridge rodando em http://localhost:5001")
    print(f"Detected OS: {platform.system()}")
    app.run(port=5001, debug=False)
