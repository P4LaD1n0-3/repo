import os
import json
import base64
import platform
import subprocess
import sqlite3
import math
from datetime import datetime
from flask import Flask, request, jsonify
from flask_cors import CORS
from critical_sla_logic import format_email_body

app = Flask(__name__)
CORS(app)  # Allow requests from local file:// or browser

# Detect OS
IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

def sanitize_for_json(data):
    """
    Remove recursivamente valores float('nan') que quebram o JSON no frontend,
    substituindo-os por None (que se torna 'null' no JSON).
    """
    if isinstance(data, dict):
        return {k: sanitize_for_json(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [sanitize_for_json(v) for v in data]
    elif isinstance(data, float) and math.isnan(data):
        return None
    return data

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
    Envia e-mail via AppleScript (Outlook Mac).
    Usa a propriedade 'content' do Outlook, que aceita HTML.
    HTML é compactado em linha única para evitar problemas no literal de string do AppleScript.
    """
    try:
        safe_to  = str(to).strip() if to else ""
        safe_cc  = str(cc).strip() if cc else ""
        safe_subj = (str(subject) if subject else "Relatório Dashboard") \
            .replace('\\', '\\\\').replace('"', '\\"')

        # Compactar HTML em linha única + escapar para string AppleScript
        compact = html_body.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        compact = compact.replace('\\', '\\\\').replace('"', '\\"')

        safe_to_esc = safe_to.replace('\\', '\\\\').replace('"', '\\"')

        as_script = (
            'tell application "Microsoft Outlook"\n'
            f'    set newMsg to make new outgoing message with properties'
            f' {{subject:"{safe_subj}", content:"{compact}"}}\n'
            f'    make new to recipient at newMsg with properties'
            f' {{email address:{{address:"{safe_to_esc}"}}}}\n'
        )

        if safe_cc:
            safe_cc_esc = safe_cc.replace('\\', '\\\\').replace('"', '\\"')
            as_script += (
                f'    make new cc recipient at newMsg with properties'
                f' {{email address:{{address:"{safe_cc_esc}"}}}}\n'
            )

        as_script += '    send newMsg\nend tell'

        subprocess.run(['osascript', '-e', as_script], check=True, capture_output=True)
        print(f"✅ E-mail HTML enviado com sucesso (Mac) para {safe_to}!")
        return True, f"E-mail enviado silenciosamente via Outlook (Mac) para {safe_to}."
    except subprocess.CalledProcessError as e:
        err = e.stderr.decode('utf-8', errors='replace') if e.stderr else str(e)
        print(f"❌ Erro AppleScript ao enviar e-mail (Mac): {err}")
        return False, err
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
                    query = "SELECT * FROM processados WHERE "
                    conditions = []
                    params = []
                    for term in terms:
                        conditions.append("(ticket_number LIKE ? OR assunto LIKE ?)")
                        params.extend([f"%{term}%", f"%{term}%"])
                    query += " OR ".join(conditions)
                    query += " ORDER BY data_processamento DESC"
                    cursor.execute(query, params)
                else:
                    cursor.execute("SELECT * FROM processados ORDER BY data_processamento DESC")
            else:
                cursor.execute("SELECT * FROM processados ORDER BY data_processamento DESC")
                
            rows = cursor.fetchall()
            data = [dict(row) for row in rows]
            data = sanitize_for_json(data)
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

def find_excel_by_pattern(base_dir, patterns):
    """
    Descobre o primeiro arquivo .xlsx cujo nome contém um dos padrões.
    Testa os padrões em ordem de especificidade (mais específico primeiro),
    para que 'incident.xlsx' seja preferido sobre 'change_inc.xlsx'.
    Espelha a lógica de detecção de arquivo do index.html.
    """
    try:
        files = [f for f in os.listdir(base_dir) if f.lower().endswith('.xlsx')]
    except OSError:
        return None

    for pat in patterns:  # padrão mais específico tem prioridade
        for fname in files:
            if pat in fname.lower():
                return os.path.join(base_dir, fname)
    return None


@app.route('/analyze-sla', methods=['POST'])
def analyze_sla_endpoint():
    import pandas as pd
    from critical_sla_logic import process_df

    data = request.json or {}
    config = {
        'sla_incs': int(data.get('slaIncs', 7)),
        'sla_reqs': int(data.get('slaReqs', 3)),
        'sla_ptask': int(data.get('slaPtask', 3)),
        'sla_rca': int(data.get('slaRca', 5))
    }

    base_dir = os.path.dirname(os.path.abspath(__file__))
    now = __import__('datetime').datetime.now()
    results = {}

    # Prioridade 1: dados enviados pelo browser (RawData já carregado no dashboard)
    inc_rows = data.get('incData')
    req_rows = data.get('reqData')

    if inc_rows:
        print(f"[SLA] Usando dados do browser: {len(inc_rows)} linhas INC")
        df_inc = pd.DataFrame(inc_rows)
        process_df(df_inc, config['sla_incs'], "INC", results, now)
    else:
        inc_path = (
            find_excel_by_pattern(base_dir, ['incident', 'inc']) or
            find_excel_by_pattern(os.path.dirname(base_dir), ['incident', 'inc'])
        )
        print(f"[SLA] Fallback arquivo INC: {inc_path}")
        if inc_path:
            df_inc = pd.read_excel(inc_path)
            process_df(df_inc, config['sla_incs'], "INC", results, now)
        else:
            print("[SLA] ❌ INC: nenhum dado recebido e nenhum arquivo encontrado.")

    if req_rows:
        print(f"[SLA] Usando dados do browser: {len(req_rows)} linhas TASK/REQ")
        df_req = pd.DataFrame(req_rows)
        process_df(df_req, config['sla_reqs'], "TASK", results, now)
    else:
        task_path = (
            find_excel_by_pattern(base_dir, ['sc_task', 'sctask']) or
            find_excel_by_pattern(os.path.dirname(base_dir), ['sc_task', 'sctask'])
        )
        print(f"[SLA] Fallback arquivo TASK: {task_path}")
        if task_path:
            df_task = pd.read_excel(task_path)
            process_df(df_task, config['sla_reqs'], "TASK", results, now)
        else:
            print("[SLA] ❌ TASK: nenhum dado recebido e nenhum arquivo encontrado.")

    analysis_results = results

    # Checar se encontrou algo
    if not analysis_results:
        return jsonify({
            "status": "error",
            "message": "Nenhum chamado crítico encontrado. Verifique se os arquivos foram carregados no dashboard (arraste incident.xlsx e sc_task.xlsx) e tente novamente."
        })
    
    # 2. Send Emails & Log
    dispatched = []
    
    # Get DB path
    rpa_dir = os.path.join(os.path.dirname(base_dir), "RPA")
    db_path = os.path.join(rpa_dir, "tickets_processados.db")
    if not os.path.exists(rpa_dir): os.makedirs(rpa_dir)

    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        # Legacy table (kept for backward compat)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sla_dispatches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                analyst TEXT,
                tickets TEXT,
                email_status TEXT,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Per-ticket table — one row per ticket for full subject/detail visibility
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sla_ticket_dispatches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_number TEXT,
                subject TEXT,
                type TEXT,
                analyst TEXT,
                analyst_email TEXT,
                aging_days REAL,
                sla_pct REAL,
                remaining_hours REAL,
                reason TEXT,
                assignment_group TEXT,
                email_status TEXT,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        for analyst_name, res_data in analysis_results.items():
            tickets = res_data['tickets']
            analyst_email = res_data['email']

            sanitized_tickets = sanitize_for_json(tickets)

            subject = f"ALERTA: Chamados Críticos - SLA Próximo do Limite"
            body = format_email_body(analyst_name, tickets)

            to_email = str(analyst_email).strip()
            if not to_email or to_email.lower() == 'nan' or '@' not in to_email:
                print(f"⚠️ Aviso: E-mail para '{analyst_name}' não encontrado ou inválido na coluna 'Email'. Usando fallback.")
                to_email = f"{analyst_name.replace(' ', '.').lower()}@empresa.com"

            cc_email = data.get('emailCc', '')

            if IS_WINDOWS:
                success, msg = send_outlook_windows(to_email, cc_email, subject, body)
            elif IS_MAC:
                success, msg = send_outlook_mac(to_email, cc_email, subject, body)
            else:
                success, msg = False, "OS Not Supported"

            email_status_str = "Success" if success else f"Error: {msg}"

            # Legacy insert (per analyst)
            cursor.execute(
                "INSERT INTO sla_dispatches (analyst, tickets, email_status) VALUES (?, ?, ?)",
                (analyst_name, json.dumps(sanitized_tickets), email_status_str)
            )
            # Per-ticket insert
            for t in sanitized_tickets:
                cursor.execute(
                    """INSERT INTO sla_ticket_dispatches
                       (ticket_number, subject, type, analyst, analyst_email,
                        aging_days, sla_pct, remaining_hours, reason, assignment_group, email_status)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        t.get('number', 'N/A'),
                        t.get('subject', ''),
                        t.get('type', ''),
                        analyst_name,
                        to_email,
                        t.get('aging_days'),
                        t.get('sla_pct'),
                        t.get('remaining_hours'),
                        t.get('reason', ''),
                        t.get('group', ''),
                        email_status_str
                    )
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
            # Prefer per-ticket table; fall back to legacy if it doesn't exist yet
            try:
                cursor.execute(
                    "SELECT * FROM sla_ticket_dispatches ORDER BY sent_at DESC"
                )
            except Exception:
                cursor.execute("SELECT * FROM sla_dispatches ORDER BY sent_at DESC")
            rows = cursor.fetchall()
            data = [dict(row) for row in rows]
            return jsonify({"status": "success", "data": sanitize_for_json(data)})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == "__main__":
    print(f"🚀 Dashboard Email Bridge rodando em http://localhost:5001")
    print(f"Detected OS: {platform.system()}")
    app.run(port=5001, debug=False)
