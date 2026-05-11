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
        try:
            df_inc = pd.read_excel(incident_path)
            process_df(df_inc, config['sla_incs'], "INC", results, now)
        except Exception as e:
            print(f"Error processing incidents: {e}")

    # Process Tasks
    if os.path.exists(sc_task_path):
        try:
            df_task = pd.read_excel(sc_task_path)
            process_df(df_task, config['sla_reqs'], "TASK", results, now)
        except Exception as e:
            print(f"Error processing tasks: {e}")
            
    return results

def process_df(df, sla_days, type_label, results, now):
    if 'Opened' not in df.columns or 'Assigned to' not in df.columns:
        return

    # Ensure dates are datetime
    df['Opened'] = pd.to_datetime(df['Opened'], errors='coerce')
    
    # Filter only open tickets (State != Closed/Resolved)
    closed_states = ['Closed', 'Resolved', 'Closed Complete', 'Closed Incomplete']
    if 'State' in df.columns:
        df = df[~df['State'].isin(closed_states)]
    
    for _, row in df.iterrows():
        if pd.isna(row['Opened']) or pd.isna(row['Assigned to']):
            continue
            
        analyst = str(row['Assigned to']).strip()
        analyst_email = str(row.get('Email', '')).strip()
        opened_date = row['Opened']
        
        # Calculate SLA
        aging_days = (now - opened_date).total_seconds() / (24 * 3600)
        sla_pct = (aging_days / sla_days) * 100
        
        # Calculate Remaining Hours
        total_sla_hours = sla_days * 24
        used_hours = aging_days * 24
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
                'sla_pct': round(sla_pct, 1),
                'remaining_hours': round(remaining_hours, 1),
                'reason': reason,
                'group': row.get('Assignment group', 'N/A')
            })

def format_email_body(analyst, tickets):
    # Fix encoding with meta tag
    body = "<html><head><meta charset='UTF-8'></head><body>"
    body += f"<h2>Olá {analyst},</h2>"
    body += "<p>Identificamos chamados críticos sob sua responsabilidade que precisam de atenção imediata:</p>"
    body += "<table border='1' style='border-collapse: collapse; width: 100%; font-family: sans-serif;'>"
    body += "<tr style='background-color: #f2f2f2;'><th>Ticket</th><th>Assunto</th><th>SLA %</th><th>Restante (h)</th><th>Motivo</th></tr>"
    
    for t in tickets:
        color = "red" if t['sla_pct'] >= 100 else "#f59e0b"
        body += f"<tr>"
        body += f"<td style='padding: 8px;'>{t['number']}</td>"
        body += f"<td style='padding: 8px;'>{t['subject']}</td>"
        body += f"<td style='padding: 8px; color: {color}; font-weight: bold;'>{t['sla_pct']}%</td>"
        body += f"<td style='padding: 8px;'>{t['remaining_hours']}h</td>"
        body += f"<td style='padding: 8px;'>{t['reason']}</td>"
        body += "</tr>"
    
    body += "</table>"
    body += "<p>Por favor, verifique o status desses chamados o quanto antes.</p>"
    body += "<br><p><i>Atenciosamente, ITSM Dashboard Automático</i></p>"
    body += "</body></html>"
    return body

if __name__ == "__main__":
    # Test
    inc_p = "incident.xlsx"
    task_p = "sc_task.xlsx"
    res = analyze_critical_slas(inc_p, task_p)
    print(json.dumps(res, indent=2))
