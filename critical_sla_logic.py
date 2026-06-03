import os
import json
import pandas as pd
from datetime import datetime, timedelta

def analyze_critical_slas(incident_path, sc_task_path, config=None):
    if config is None:
        config = {
            'sla_incs': 7,
            'sla_reqs': 3,
            'sla_ptask': 3,
            'sla_rca': 5
        }
    
    now = datetime.now()
    results = {} # analyst_name -> { 'email': str, 'tickets': [] }
    
    # Process Incidents
    if os.path.exists(incident_path):
        df_inc = pd.read_excel(incident_path)
        process_df(df_inc, config['sla_incs'], "INC", results, now)

    # Process Tasks
    if os.path.exists(sc_task_path):
        df_task = pd.read_excel(sc_task_path)
        process_df(df_task, config['sla_reqs'], "TASK", results, now)
            
    return results

def process_df(df, sla_days, type_label, results, now):
    print(f"\\n--- INICIANDO PROCESSAMENTO: {type_label} ---")
    print(f"[{type_label}] Total de linhas originais: {len(df)}")
    
    if 'Opened' not in df.columns or 'Assigned to' not in df.columns:
        print(f"[{type_label}] ❌ ERRO: Colunas 'Opened' ou 'Assigned to' não encontradas!")
        print(f"[{type_label}] Colunas disponíveis: {df.columns.tolist()}")
        return

    # Ensure dates are datetime
    df['Opened'] = pd.to_datetime(df['Opened'], errors='coerce')
    
    # Filter only open tickets (State != Closed/Resolved/Canceled)
    closed_states = ['Closed', 'Resolved', 'Closed Complete', 'Closed Incomplete', 'Canceled', 'Cancelled']
    if 'State' in df.columns:
        df = df[~df['State'].isin(closed_states)]
    
    print(f"[{type_label}] Total de linhas após remover status fechado/cancelado: {len(df)}")
    
    skipped_null = 0
    not_critical = 0
    critical_found = 0
    
    for _, row in df.iterrows():
        if pd.isna(row['Opened']) or pd.isna(row['Assigned to']):
            skipped_null += 1
            continue
            
        analyst = str(row['Assigned to']).strip()
        
        # Procura coluna de Email de forma robusta (case-insensitive)
        email_col = next((c for c in df.columns if str(c).lower() == 'email'), None)
        analyst_email = str(row[email_col]).strip() if email_col and not pd.isna(row[email_col]) else ""
        opened_date = row['Opened']
        
        # Identify Priority or Impact
        priority_col = next((c for c in df.columns if str(c).lower() in ['priority', 'impact']), None)
        priority_val = str(row[priority_col]).lower() if priority_col and not pd.isna(row[priority_col]) else ""
        
        p = 4 # Default priority
        if '1' in priority_val or 'p1' in priority_val or 'critical' in priority_val:
            p = 1
        elif '2' in priority_val or 'p2' in priority_val or 'high' in priority_val:
            p = 2
        elif '3' in priority_val or 'p3' in priority_val or 'moderate' in priority_val:
            p = 3
        elif '4' in priority_val or 'p4' in priority_val or 'low' in priority_val:
            p = 4
            
        # Determine SLA Hours based on Priority and Type
        total_sla_hours = sla_days * 24 # Fallback
        
        if type_label == "INC":
            if p == 1:
                total_sla_hours = 4
            elif p == 2:
                total_sla_hours = 8
            elif p == 3:
                total_sla_hours = 5 * 24
            elif p == 4:
                total_sla_hours = 7 * 24
        elif type_label in ["TASK", "REQ", "RITM"]:
            if p == 1:
                total_sla_hours = 10 * 24
            
        # Calculate SLA
        aging_days = (now - opened_date).total_seconds() / (24 * 3600)
        used_hours = aging_days * 24
        
        sla_pct = (used_hours / total_sla_hours) * 100 if total_sla_hours > 0 else 100
        remaining_hours = total_sla_hours - used_hours
        
        is_critical = False
        reason = ""
        
        # Rule 1: Above 90%
        if sla_pct >= 90:
            is_critical = True
            reason = f"SLA at {sla_pct:.1f}%"
        
        # Rule 2: Weekend breach
        weekday = now.weekday()
        if not is_critical:
            days_to_monday = (7 - weekday) % 7
            if days_to_monday == 0: days_to_monday = 7
            
            if weekday == 4 and remaining_hours < 72:
                is_critical = True
                reason = "Will breach over the weekend"
            elif weekday >= 4 and remaining_hours < (days_to_monday * 24 + 8):
                is_critical = True
                reason = "Will breach over the weekend"

        if is_critical:
            critical_found += 1
            if analyst not in results:
                results[analyst] = {
                    'email': analyst_email,
                    'tickets': []
                }
            
            # Update email if it was empty and now we found it
            if not results[analyst]['email'] and analyst_email:
                results[analyst]['email'] = analyst_email

            results[analyst]['tickets'].append({
                'number': row['Number'],
                'type': type_label,
                'subject': row.get('Short description', 'No description'),
                'opened': opened_date.strftime('%Y-%m-%d %H:%M'),
                'aging_days': round(aging_days, 1),
                'sla_pct': round(sla_pct, 1),
                'remaining_hours': round(remaining_hours, 1),
                'reason': reason,
                'group': row.get('Assignment group', 'N/A')
            })
        else:
            not_critical += 1

    print(f"[{type_label}] RESULTADO: {skipped_null} ignorados (Sem Opened/Assigned), {not_critical} não críticos, {critical_found} críticos encontrados.")
    print(f"--- FIM PROCESSAMENTO: {type_label} ---\\n")

def format_email_body(analyst, tickets):
    # Fix encoding with meta tag and explicit style
    body = """
    <html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <style>
            body { font-family: 'Segoe UI', Tahoma, sans-serif; color: #333; line-height: 1.5; }
            table { border-collapse: collapse; width: 100%; border: 1px solid #e2e8f0; margin-top: 20px; }
            th { background-color: #001871; color: white; padding: 12px; text-align: left; font-size: 12px; text-transform: uppercase; }
            td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; }
            tr:nth-child(even) { background-color: #f8fafc; }
            .sla-bar-bg { background: #e2e8f0; border-radius: 4px; width: 60px; height: 8px; display: inline-block; vertical-align: middle; margin-right: 8px; }
            .sla-bar-fill { height: 100%; border-radius: 4px; }
            .pct-text { font-weight: bold; font-size: 11px; }
            .text-breached { color: #ef4444; }
            .text-warning { color: #f59e0b; }
            .text-ok { color: #3b82f6; }
        </style>
    </head>
    <body>
    """
    body += f"<h2>Olá {analyst},</h2>"
    body += "<p>Identificamos chamados críticos sob sua responsabilidade que precisam de atenção imediata:</p>"
    body += "<table>"
    body += "<thead><tr><th>Ticket</th><th>Assunto</th><th>Aging</th><th>Restante</th><th>Motivo</th></tr></thead>"
    body += "<tbody>"
    
    for t in tickets:
        sla_pct = t['sla_pct']
        # Progress bar color
        bar_color = "#ef4444" if sla_pct >= 100 else ("#f59e0b" if sla_pct >= 80 else "#3b82f6")
        text_class = "text-breached" if sla_pct >= 100 else ("text-warning" if sla_pct >= 80 else "text-ok")
        fill_width = min(100, sla_pct)
        
        # Remaining time logic (Hide negatives)
        rem_h = t['remaining_hours']
        if rem_h > 0:
            h = int(rem_h)
            m = int((rem_h - h) * 60)
            time_display = f"<b class='text-ok'>{h}h {m}m</b>"
        else:
            time_display = "<b class='text-breached'>ESTOURADO</b>"
            
        body += f"<tr>"
        body += f"<td><b>{t['number']}</b></td>"
        body += f"<td>{t['subject']}</td>"
        body += f"<td><b style='color:#1e293b;'>{t['aging_days']} dias</b></td>"
        body += f"<td>{time_display}</td>"
        body += f"<td style='font-size: 11px; color: #64748b;'>{t['reason']}</td>"
        body += "</tr>"
    
    body += "</tbody></table>"
    body += "<p>Por favor, verifique o status desses chamados o quanto antes no sistema.</p>"
    body += "<br><p><i>Atenciosamente, <b>Equipe ITSM Dashboard</b></i></p>"
    body += "</body></html>"
    return body

if __name__ == "__main__":
    # Test
    inc_p = "incident.xlsx"
    task_p = "sc_task.xlsx"
    res = analyze_critical_slas(inc_p, task_p)
    print(json.dumps(res, indent=2))
