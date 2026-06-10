import os
import json
import pandas as pd
from datetime import datetime

# Column alias lists — mirrors the getCol() pattern from index.html
OPENED_ALIASES   = ['Opened', 'Aberto', 'opened', 'Open Date', 'Data de abertura', 'Created', 'Sys created on', 'Data Abertura']
ASSIGNED_ALIASES = ['Assigned to', 'Assigned To', 'Atribuído a', 'Atribuido a', 'Assignee', 'assigned_to', 'Responsável', 'Responsavel']
NUMBER_ALIASES   = ['Number', 'Número', 'Numero', 'Ticket', 'number', 'Ticket Number', 'ID', 'Call Number']
SUBJECT_ALIASES  = ['Short description', 'Descrição breve', 'Descricao breve', 'Description', 'Subject', 'short_description', 'Assunto', 'Titulo', 'Título']
EMAIL_ALIASES    = ['Email', 'email', 'E-mail', 'Analyst Email', 'Analyst email', 'analyst_email']
PRIORITY_ALIASES = ['Priority', 'Prioridade', 'Impact', 'Impacto', 'priority', 'Urgency']
GROUP_ALIASES    = ['Assignment group', 'Grupo de atribuição', 'Grupo de atribuicao', 'Group', 'assignment_group', 'Grupo', 'Support group']
STATE_ALIASES    = ['State', 'Estado', 'Status', 'state', 'status']


def find_col(df, aliases):
    """
    Case-insensitive column lookup with aliases. Returns the real column name or None.
    Mirrors the getCol() utility in index.html.
    """
    col_map = {c.strip().lower(): c for c in df.columns}
    for alias in aliases:
        match = col_map.get(alias.strip().lower())
        if match is not None:
            return match
    return None


def analyze_critical_slas(incident_path, sc_task_path, config=None):
    if config is None:
        config = {
            'sla_incs': 7,
            'sla_reqs': 3,
            'sla_ptask': 3,
            'sla_rca': 5
        }

    now = datetime.now()
    results = {}  # analyst_name -> { 'email': str, 'tickets': [] }

    # Process Incidents
    if incident_path and os.path.exists(incident_path):
        df_inc = pd.read_excel(incident_path)
        process_df(df_inc, config['sla_incs'], "INC", results, now)
    else:
        print(f"[INC] ⚠️ Arquivo não encontrado: {incident_path}")

    # Process Tasks
    if sc_task_path and os.path.exists(sc_task_path):
        df_task = pd.read_excel(sc_task_path)
        process_df(df_task, config['sla_reqs'], "TASK", results, now)
    else:
        print(f"[TASK] ⚠️ Arquivo não encontrado: {sc_task_path}")

    return results

def process_df(df, sla_days, type_label, results, now):
    print(f"\n--- INICIANDO PROCESSAMENTO: {type_label} ---")
    print(f"[{type_label}] Total de linhas originais: {len(df)}")
    print(f"[{type_label}] Colunas disponíveis: {df.columns.tolist()}")

    opened_col   = find_col(df, OPENED_ALIASES)
    assigned_col = find_col(df, ASSIGNED_ALIASES)

    if not opened_col or not assigned_col:
        print(f"[{type_label}] ❌ ERRO: Coluna de abertura ('{opened_col}') ou responsável ('{assigned_col}') não encontrada!")
        print(f"[{type_label}] Aliases tentados — Opened: {OPENED_ALIASES}")
        print(f"[{type_label}] Aliases tentados — Assigned: {ASSIGNED_ALIASES}")
        return

    print(f"[{type_label}] ✅ Coluna de data: '{opened_col}' | Responsável: '{assigned_col}'")

    number_col   = find_col(df, NUMBER_ALIASES)
    subject_col  = find_col(df, SUBJECT_ALIASES)
    email_col    = find_col(df, EMAIL_ALIASES)
    priority_col = find_col(df, PRIORITY_ALIASES)
    group_col    = find_col(df, GROUP_ALIASES)
    state_col    = find_col(df, STATE_ALIASES)

    # Ensure dates are datetime — strip timezone so arithmetic with tz-naive `now` works
    # (browser sends ISO-8601 strings with "Z"/offset; Excel files are usually tz-naive)
    _series = pd.to_datetime(df[opened_col], errors='coerce')
    if _series.dt.tz is not None:
        _series = _series.dt.tz_convert(None)
    df[opened_col] = _series

    # Filter only open tickets
    closed_states = ['closed', 'resolved', 'closed complete', 'closed incomplete', 'canceled', 'cancelled',
                     'fechado', 'resolvido', 'cancelado', 'concluído', 'concluido']
    if state_col:
        df = df[~df[state_col].astype(str).str.strip().str.lower().isin(closed_states)]

    print(f"[{type_label}] Total após remover fechados/cancelados: {len(df)}")

    skipped_null = 0
    not_critical = 0
    critical_found = 0

    for _, row in df.iterrows():
        if pd.isna(row[opened_col]) or pd.isna(row[assigned_col]):
            skipped_null += 1
            continue

        analyst = str(row[assigned_col]).strip()

        analyst_email = ""
        if email_col and not pd.isna(row[email_col]):
            analyst_email = str(row[email_col]).strip()

        opened_date = row[opened_col]

        # Identify Priority or Impact
        priority_val = ""
        if priority_col and not pd.isna(row[priority_col]):
            priority_val = str(row[priority_col]).lower()

        p = 4  # Default priority
        if '1' in priority_val or 'p1' in priority_val or 'critical' in priority_val or 'crítico' in priority_val:
            p = 1
        elif '2' in priority_val or 'p2' in priority_val or 'high' in priority_val or 'alta' in priority_val:
            p = 2
        elif '3' in priority_val or 'p3' in priority_val or 'moderate' in priority_val or 'média' in priority_val or 'media' in priority_val:
            p = 3
        elif '4' in priority_val or 'p4' in priority_val or 'low' in priority_val or 'baixa' in priority_val:
            p = 4

        # Determine SLA Hours based on Priority and Type
        total_sla_hours = sla_days * 24  # Fallback

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
            reason = f"SLA em {sla_pct:.1f}%"

        # Rule 2: Weekend breach
        weekday = now.weekday()
        if not is_critical:
            days_to_monday = (7 - weekday) % 7
            if days_to_monday == 0:
                days_to_monday = 7

            if weekday == 4 and remaining_hours < 72:
                is_critical = True
                reason = "Estourará no final de semana"
            elif weekday >= 4 and remaining_hours < (days_to_monday * 24 + 8):
                is_critical = True
                reason = "Estourará no final de semana"

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

            ticket_number = str(row[number_col]) if number_col and not pd.isna(row[number_col]) else "N/A"
            subject_val   = str(row[subject_col]) if subject_col and not pd.isna(row[subject_col]) else "Sem descrição"
            group_val     = str(row[group_col]) if group_col and not pd.isna(row[group_col]) else "N/A"

            results[analyst]['tickets'].append({
                'number': ticket_number,
                'type': type_label,
                'subject': subject_val,
                'opened': opened_date.strftime('%Y-%m-%d %H:%M'),
                'aging_days': round(aging_days, 1),
                'sla_pct': round(sla_pct, 1),
                'remaining_hours': round(remaining_hours, 1),
                'reason': reason,
                'group': group_val
            })
        else:
            not_critical += 1

    print(f"[{type_label}] RESULTADO: {skipped_null} ignorados (sem data/responsável), {not_critical} não críticos, {critical_found} críticos encontrados.")
    print(f"--- FIM PROCESSAMENTO: {type_label} ---\n")

def format_email_body(analyst, tickets):
    body = (
        '<html>'
        '<head>'
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
        '<style>'
        'body{font-family:\'Segoe UI\',Tahoma,sans-serif;color:#333;line-height:1.5;}'
        'table{border-collapse:collapse;width:100%;border:1px solid #e2e8f0;margin-top:20px;}'
        'th{background-color:#001871;color:white;padding:12px;text-align:left;font-size:12px;text-transform:uppercase;}'
        'td{padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;}'
        'tr:nth-child(even){background-color:#f8fafc;}'
        '.text-breached{color:#ef4444;}'
        '.text-warning{color:#f59e0b;}'
        '.text-ok{color:#3b82f6;}'
        '</style>'
        '</head>'
        '<body>'
    )
    body += f"<h2>Ol&#225; {analyst},</h2>"
    body += "<p>Identificamos chamados cr&#237;ticos sob sua responsabilidade que precisam de aten&#231;&#227;o imediata:</p>"
    body += "<table><thead><tr><th>Ticket</th><th>Assunto</th><th>Aging</th><th>Restante</th><th>Motivo</th></tr></thead><tbody>"

    for t in tickets:
        sla_pct = t['sla_pct']
        text_class = "text-breached" if sla_pct >= 100 else ("text-warning" if sla_pct >= 80 else "text-ok")

        rem_h = t['remaining_hours']
        if rem_h > 0:
            h = int(rem_h)
            m = int((rem_h - h) * 60)
            time_display = f"<b class='text-ok'>{h}h {m}m</b>"
        else:
            time_display = "<b class='text-breached'>ESTOURADO</b>"

        body += (
            f"<tr>"
            f"<td><b>{t['number']}</b></td>"
            f"<td>{t['subject']}</td>"
            f"<td><b style='color:#1e293b;'>{t['aging_days']} dias</b></td>"
            f"<td>{time_display}</td>"
            f"<td style='font-size:11px;color:#64748b;'>{t['reason']}</td>"
            f"</tr>"
        )

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
    print(json.dumps(res, indent=2, default=str))
