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
    results = {} # analyst -> [tickets]
    
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
        # If today is Friday (4), or if remaining hours <= time until Monday morning
        # Let's simplify: if remaining_hours < 72 and it's Friday/Saturday/Sunday
        weekday = now.weekday()
        if not is_critical:
            # Time until Monday 8 AM
            days_to_monday = (7 - weekday) % 7
            if days_to_monday == 0: days_to_monday = 7 # If Monday, look at next Monday? No, usually dispatches happen on Friday.
            
            # If it breaches within the next 72 hours and today is Friday (4)
            if weekday == 4 and remaining_hours < 72:
                is_critical = True
                reason = "Will breach over the weekend"
            elif weekday >= 4 and remaining_hours < (days_to_monday * 24 + 8):
                is_critical = True
                reason = "Will breach over the weekend"

        if is_critical:
            if analyst not in results:
                results[analyst] = []
            
            results[analyst].append({
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
    body = f"<h2>Olá {analyst},</h2>"
    body += "<p>Identificamos chamados críticos sob sua responsabilidade que precisam de atenção imediata:</p>"
    body += "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    body += "<tr style='background-color: #f2f2f2;'><th>Ticket</th><th>Assunto</th><th>SLA %</th><th>Restante (h)</th><th>Motivo</th></tr>"
    
    for t in tickets:
        color = "red" if t['sla_pct'] >= 100 else "orange"
        body += f"<tr>"
        body += f"<td>{t['number']}</td>"
        body += f"<td>{t['subject']}</td>"
        body += f"<td style='color: {color}; font-weight: bold;'>{t['sla_pct']}%</td>"
        body += f"<td>{t['remaining_hours']}h</td>"
        body += f"<td>{t['reason']}</td>"
        body += "</tr>"
    
    body += "</table>"
    body += "<p>Por favor, verifique o status desses chamados o quanto antes.</p>"
    body += "<br><p><i>Atenciosamente, ITSM Dashboard Automático</i></p>"
    return body

if __name__ == "__main__":
    # Test
    inc_p = "incident.xlsx"
    task_p = "sc_task.xlsx"
    res = analyze_critical_slas(inc_p, task_p)
    print(json.dumps(res, indent=2))
