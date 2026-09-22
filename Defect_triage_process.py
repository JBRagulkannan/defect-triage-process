import streamlit as st
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from datetime import date, timedelta
import re
import html
import platform
import os
import tempfile
import uuid
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
JIRA_BASE_URL = st.secrets["JIRA_BASE_URL"]

# --------------------------------------------------
# Optional Outlook Integration
# --------------------------------------------------
try:
    import pythoncom
    import win32com.client as win32
    OUTLOOK_AVAILABLE = True
except Exception:
    OUTLOOK_AVAILABLE = False

# --------------------------------------------------
# Page Config
# --------------------------------------------------
st.set_page_config(
    page_title="Defect Triage Process Automation",
    page_icon="GVR.ico",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --------------------------------------------------
# Constants
# --------------------------------------------------
JIRA_BASE_URL = "https://jira.gilbarco.com"

SEVERITY_FIELD = "customfield_10128"
CUSTOMER_IMPACT_FIELD = "customfield_12003"
COST_TO_FIX_FIELD = "customfield_12004"
FREQUENCY_OCCURRENCE_FIELD = "customfield_12518"
STEPS_FIELD = "customfield_11124"
EXPECTED_RESULT_FIELD = "customfield_11121"

SEVERITY_OPTIONS = [
    "",
    "S1 : Blocker",
    "S2 : Critical",
    "S3 : Major",
    "S4 : Minor",
]

COMMON_PRIORITY_OPTIONS = [
    "",
    "Low",
    "Medium",
    "High",
]

PROJECT_OPTIONS = ["TIP", "GEP"]

OUTLOOK_SUBJECT = "Jira Defects QA Summary Report"
OUTLOOK_CHART_CID = "jira_status_chart_cid"

OLD_BACKLOG_CUTOFF = pd.Timestamp("2025-12-31")
NEW_BACKLOG_START = pd.Timestamp("2026-01-01")

# --------------------------------------------------
# Session Defaults
# --------------------------------------------------
defaults = {
    "logged_in": False,
    "jira_email": "",
    "jira_password": "",
    "jira_user": None,
    "issues_df": pd.DataFrame(),
    "search_text": "",
    "selected_jira_id": None,
    "from_date": date.today() - timedelta(days=30),
    "to_date": date.today(),
    "max_results": 100,
    "refresh_after_update": False,
    "selected_project": "TIP",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --------------------------------------------------
# Utility Functions
# --------------------------------------------------
def strip_html_tags(text):
    if text is None:
        return ""
    text = str(text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def safe_text(value, default="-"):
    if value is None:
        return default
    text = strip_html_tags(value)
    return text if text else default

def extract_value(field_value):
    if field_value is None:
        return ""
    if isinstance(field_value, dict):
        if "value" in field_value:
            return field_value["value"]
        if "name" in field_value:
            return field_value["name"]
        if "displayName" in field_value:
            return field_value["displayName"]
        if "id" in field_value:
            return field_value["id"]
        return str(field_value)
    if isinstance(field_value, list):
        values = []
        for item in field_value:
            if isinstance(item, dict):
                values.append(item.get("name") or item.get("value") or item.get("displayName") or str(item))
            else:
                values.append(str(item))
        return ", ".join(values)
    return str(field_value)

def extract_comments(comment_field):
    if not comment_field:
        return ""
    comments = comment_field.get("comments", [])
    formatted_comments = []
    for c in comments:
        author = c.get("author", {}).get("displayName", "Unknown")
        body = strip_html_tags(c.get("body", ""))
        created = c.get("created", "")
        if created:
            try:
                created = pd.to_datetime(created, errors="coerce").strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
        formatted_comments.append(f"{author} ({created}): {body}")
    return "\n\n".join(formatted_comments)

def filter_df(df, search_text):
    if df.empty or not search_text.strip():
        return df
    search = search_text.strip().lower()
    return df[
        df.apply(
            lambda row: any(search in str(val).lower() for val in row.values),
            axis=1
        )
    ]

def logout():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

# --------------------------------------------------
# Risk Factor Functions
# --------------------------------------------------
def map_severity_to_score(severity_value):
    val = str(severity_value).strip().upper()
    if val in ["S1 : BLOCKER", "S1", "1"]:
        return 5
    elif val in ["S2 : CRITICAL", "S2", "2"]:
        return 3
    elif val == "":
        return 0
    else:
        return 1

def map_priority_to_score(priority_value):
    val = str(priority_value).strip().upper()
    if val == "HIGH":
        return 5
    elif val == "MEDIUM":
        return 3
    elif val == "":
        return 0
    else:
        return 1

def calculate_defect_risk_factor(severity, customer_impact, cost_to_fix, frequency_occurrence):
    severity_score = map_severity_to_score(severity)
    customer_impact_score = map_priority_to_score(customer_impact)
    cost_to_fix_score = map_priority_to_score(cost_to_fix)
    frequency_score = map_priority_to_score(frequency_occurrence)

    if severity_score == 0:
        return 0.0

    risk_factor = (
        (severity_score * 0.15) +
        (customer_impact_score * 0.40) +
        (cost_to_fix_score * 0.25) +
        (frequency_score * 0.20)
    ) * 100

    return round(risk_factor, 2)

def update_risk_factor_session(issue_key):
    sev_key = f"sev_{issue_key}"
    ci_key = f"ci_{issue_key}"
    ctf_key = f"ctf_{issue_key}"
    freq_key = f"freq_{issue_key}"
    risk_key = f"risk_{issue_key}"

    st.session_state[risk_key] = calculate_defect_risk_factor(
        st.session_state.get(sev_key, ""),
        st.session_state.get(ci_key, ""),
        st.session_state.get(ctf_key, ""),
        st.session_state.get(freq_key, "")
    )

# --------------------------------------------------
# Status Helpers
# --------------------------------------------------
def normalize_status(status_value):
    val = str(status_value).strip().lower()
    if val in ["open", "to do", "new", "reopened"]:
        return "Open"
    if val in ["closed", "done", "resolved"]:
        return "Closed"
    if val in ["in progress", "in-progress", "progress"]:
        return "In Progress"
    if val in ["deferred"]:
        return "Deferred"
    if not val:
        return "Unknown"
    return str(status_value).strip().title()

def is_closed_status(status_value):
    return normalize_status(status_value) == "Closed"

def is_open_status(status_value):
    return normalize_status(status_value) in ["Open", "In Progress", "Unknown"]

def is_deferred_status(status_value):
    return str(status_value).strip().lower() == "deferred"

# --------------------------------------------------
# Outlook Summary Helpers
# --------------------------------------------------
def safe_percent(numerator, denominator):
    if denominator == 0:
        return 0.0
    return round((numerator / denominator) * 100, 2)

def build_backlog_summary_data(df):
    if df.empty:
        return {
            "old_total": 0, "old_burned": 0, "old_deferred": 0, "old_open": 0, "old_reduction_pct": 0.0,
            "new_total": 0, "new_burned": 0, "new_deferred": 0, "new_open": 0, "new_reduction_pct": 0.0,
            "overall_total": 0, "overall_burned": 0, "overall_deferred": 0, "overall_open": 0, "overall_reduction_pct": 0.0,
            "deferred_total": 0, "deferred_burned": 0, "deferred_open": 0, "deferred_reduction_pct": 0.0,
        }

    temp_df = df.copy()
    temp_df["created_dt"] = pd.to_datetime(temp_df["created"], errors="coerce") if "created" in temp_df.columns else pd.NaT
    temp_df["normalized_status"] = temp_df["status"].apply(normalize_status)

    old_df = temp_df[temp_df["created_dt"] <= OLD_BACKLOG_CUTOFF].copy()
    new_df = temp_df[temp_df["created_dt"] >= NEW_BACKLOG_START].copy()

    old_total = len(old_df)
    old_burned = len(old_df[old_df["normalized_status"] == "Closed"])
    old_deferred = len(old_df[old_df["normalized_status"] == "Deferred"])
    old_open = len(old_df[old_df["normalized_status"].isin(["Open", "In Progress", "Unknown"])])
    old_reduction_pct = safe_percent(old_burned, old_total)

    new_total = len(new_df)
    new_burned = len(new_df[new_df["normalized_status"] == "Closed"])
    new_deferred = len(new_df[new_df["normalized_status"] == "Deferred"])
    new_open = len(new_df[new_df["normalized_status"].isin(["Open", "In Progress", "Unknown"])])
    new_reduction_pct = safe_percent(new_burned, new_total)

    overall_total = len(temp_df)
    overall_burned = len(temp_df[temp_df["normalized_status"] == "Closed"])
    overall_deferred = len(temp_df[temp_df["normalized_status"] == "Deferred"])
    overall_open = len(temp_df[temp_df["normalized_status"].isin(["Open", "In Progress", "Unknown"])])
    overall_reduction_pct = safe_percent(overall_burned, overall_total)

    deferred_df = temp_df[temp_df["normalized_status"] == "Deferred"].copy()
    deferred_total = len(deferred_df)
    deferred_open = deferred_total
    deferred_burned = 0
    deferred_reduction_pct = safe_percent(deferred_burned, deferred_total)

    return {
        "old_total": old_total, "old_burned": old_burned, "old_deferred": old_deferred, "old_open": old_open, "old_reduction_pct": old_reduction_pct,
        "new_total": new_total, "new_burned": new_burned, "new_deferred": new_deferred, "new_open": new_open, "new_reduction_pct": new_reduction_pct,
        "overall_total": overall_total, "overall_burned": overall_burned, "overall_deferred": overall_deferred, "overall_open": overall_open, "overall_reduction_pct": overall_reduction_pct,
        "deferred_total": deferred_total, "deferred_burned": deferred_burned, "deferred_open": deferred_open, "deferred_reduction_pct": deferred_reduction_pct,
    }

def backlog_reduction_summary_html(df):
    s = build_backlog_summary_data(df)
    return f"""
    <table style="border-collapse:collapse;width:820px;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;margin-top:8px;margin-bottom:18px;">
        <tr>
            <th colspan="2" style="background-color:#8db4e2;border:2px solid #1f1f1f;padding:6px 8px;text-align:center;font-weight:bold;">Backlog Reduction Summary (Old Defects till Dec 2025)</th>
            <th colspan="2" style="background-color:#8db4e2;border:2px solid #1f1f1f;padding:6px 8px;text-align:center;font-weight:bold;">Backlog Reduction Summary (New Defects from Jan 2026)</th>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Defects backlog:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['old_total']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">New defects logged:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['new_total']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Burned defects count:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['old_burned']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">New Burned defects count:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['new_burned']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Deferred Defects Count</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['old_deferred']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Deferred Defects Count</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['new_deferred']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Current open defects</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['old_open']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Current open defects:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['new_open']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Backlog reduction %age</td>
            <td style="background-color:#00b050;color:#ffffff;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['old_reduction_pct']:.2f}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Backlog reduction %age:</td>
            <td style="background-color:#00b050;color:#ffffff;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['new_reduction_pct']:.2f}</td>
        </tr>
        <tr>
            <th colspan="2" style="background-color:#8db4e2;border:2px solid #1f1f1f;padding:6px 8px;text-align:center;font-weight:bold;">Over all Backlog Reduction Summary</th>
            <th colspan="2" style="background-color:#8db4e2;border:2px solid #1f1f1f;padding:6px 8px;text-align:center;font-weight:bold;">Deferred Backlog Reduction Summary</th>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Total Defects</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['overall_total']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Total No of defects :</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['deferred_total']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Total Defects Burned defects</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['overall_burned']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">No of Burned defects count:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['deferred_burned']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Deferred Defects Count</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['overall_deferred']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Current open defects:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['deferred_open']}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Current open defects</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['overall_open']}</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Backlog reduction %age:</td>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['deferred_reduction_pct']:.2f}</td>
        </tr>
        <tr>
            <td style="background-color:#c6e0b4;border:2px solid #1f1f1f;padding:5px 8px;font-weight:bold;">Total backlog reduction %age</td>
            <td style="background-color:#00b050;color:#ffffff;border:2px solid #1f1f1f;padding:5px 8px;text-align:center;font-weight:bold;">{s['overall_reduction_pct']:.2f}</td>
            <td style="background-color:#ffffff;border:2px solid #1f1f1f;padding:5px 8px;"></td>
            <td style="background-color:#ffffff;border:2px solid #1f1f1f;padding:5px 8px;"></td>
        </tr>
    </table>
    """

# --------------------------------------------------
# Outlook + Chart Functions
# --------------------------------------------------
def create_status_pie_chart(df):
    temp_df = df.copy()

    if "status" not in temp_df.columns or temp_df.empty:
        temp_df["status"] = ["Unknown"] * len(temp_df)

    temp_df["status"] = temp_df["status"].apply(normalize_status)
    status_counts = temp_df["status"].value_counts()

    if status_counts.empty:
        status_counts = pd.Series({"Unknown": 1})

    labels = status_counts.index.tolist()
    sizes = status_counts.values.tolist()

    color_map = {
        "Open": "#ef4444",
        "Closed": "#10b981",
        "In Progress": "#f59e0b",
        "Deferred": "#8b5cf6",
        "Unknown": "#94a3b8"
    }
    colors = [color_map.get(label, "#6366f1") for label in labels]

    temp_dir = tempfile.gettempdir()
    chart_file = os.path.join(temp_dir, f"jira_status_pie_{uuid.uuid4().hex}.png")

    plt.figure(figsize=(7.5, 5.5), facecolor="white")
    plt.pie(
        sizes,
        labels=labels,
        colors=colors,
        autopct="%1.1f%%",
        startangle=140,
        wedgeprops={"edgecolor": "white", "linewidth": 2},
        textprops={"fontsize": 11, "fontweight": "bold", "color": "#1f2937"}
    )
    plt.title("Defect Status Distribution", fontsize=16, fontweight="bold", color="#0f172a", pad=18)
    plt.tight_layout()
    plt.savefig(chart_file, dpi=160, bbox_inches="tight")
    plt.close()

    return chart_file, status_counts

def dataframe_to_html_table(df):
    if df.empty:
        return """
        <p style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">
            No defect data available.
        </p>
        """

    display_columns = [
        "jiraId", "summary", "status", "severity", "customerImpact",
        "costToFix", "frequencyOccurrence", "actionOwner", "steps",
        "expectedResult", "project", "created",
    ]

    temp_df = df.copy()
    available_columns = [c for c in display_columns if c in temp_df.columns]
    temp_df = temp_df[available_columns].fillna("")

    header_map = {
        "jiraId": "Jira ID",
        "summary": "Summary",
        "status": "Status",
        "severity": "Severity",
        "customerImpact": "Customer Impact",
        "costToFix": "Cost To Fix",
        "frequencyOccurrence": "Frequency Occurrence",
        "actionOwner": "Action Owner",
        "steps": "Steps",
        "expectedResult": "Expected Result",
        "project": "Project",
        "created": "Created",
    }

    table_html = """
    <table style="border-collapse:collapse; width:100%; font-family:Calibri,Arial,sans-serif; font-size:10.5pt;">
        <thead><tr>
    """

    for col in available_columns:
        table_html += f"""
            <th style="background-color:#1d4ed8;color:#ffffff;border:1px solid #cbd5e1;padding:8px;text-align:left;vertical-align:top;">
                {html.escape(header_map.get(col, col))}
            </th>
        """

    table_html += "</tr></thead><tbody>"

    for _, row in temp_df.iterrows():
        table_html += "<tr>"
        for col in available_columns:
            cell_value = "" if pd.isna(row[col]) else str(row[col])
            cell_value = html.escape(cell_value).replace("\n", "<br>")
            table_html += f"""
                <td style="border:1px solid #cbd5e1;padding:8px;text-align:left;vertical-align:top;background-color:#ffffff;">
                    {cell_value}
                </td>
            """
        table_html += "</tr>"

    table_html += "</tbody></table>"
    return table_html

def build_outlook_html_body(df, chart_cid, status_counts):
    total = int(status_counts.sum()) if len(status_counts) > 0 else 0
    status_summary_html = ""

    color_map = {
        "Open": "#ef4444",
        "Closed": "#10b981",
        "In Progress": "#f59e0b",
        "Deferred": "#8b5cf6",
        "Unknown": "#94a3b8"
    }

    for status_name, count in status_counts.items():
        percent = (count / total * 100) if total else 0
        pill_color = color_map.get(status_name, "#6366f1")
        status_summary_html += f"""
            <span style="display:inline-block;margin:4px 8px 4px 0;padding:7px 12px;border-radius:999px;background:{pill_color};color:#ffffff;font-size:10.5pt;font-weight:bold;">
                {html.escape(str(status_name))}: {int(count)} ({percent:.1f}%)
            </span>
        """

    backlog_summary_html = backlog_reduction_summary_html(df)
    defects_table_html = dataframe_to_html_table(df.head(10))

    html_body = f"""
    <html>
    <body style="font-family:Calibri,Arial,sans-serif; font-size:11pt; color:#000000;">
        <p style="margin:0 0 8px 0;">Hi Team</p>
        <p style="margin:0 0 8px 0; font-weight:bold; font-size:13pt;">Overall Defect summary:</p>
        {backlog_summary_html}
        <p style="margin:18px 0 10px 0; font-weight:bold; text-decoration:underline;">Please find the MOM for the triage meeting.</p>
        <p style="margin:0 0 18px 0; font-weight:bold;">Attendees: Thanima, Hemant, Poonam.</p>

        <div style="margin-bottom:12px;">{status_summary_html}</div>

        <div style="margin:16px 0 22px 0;padding:16px;border:1px solid #dbe4f0;border-radius:12px;background:#f8fbff;text-align:center;">
            <div style="font-size:13pt; font-weight:bold; color:#0f172a; margin-bottom:10px;">Defect Status Pie Chart</div>
            <img src="cid:{chart_cid}" style="max-width:700px; width:100%; height:auto; border-radius:10px;" />
        </div>

        <div style="font-size:13pt; font-weight:bold; color:#0f172a; margin:8px 0 10px 0;">Top 10 Jira Defects</div>
        {defects_table_html}

        <br>
        <p>Regards,<br>QA Team</p>
    </body>
    </html>
    """
    return html_body

def open_outlook_draft_with_defects(df):
    if platform.system().lower() != "windows":
        raise EnvironmentError("Outlook desktop integration works only on Windows.")

    if not OUTLOOK_AVAILABLE:
        raise ImportError("win32com.client is not available. Please install pywin32.")

    chart_file, status_counts = create_status_pie_chart(df)
    html_body = build_outlook_html_body(df, OUTLOOK_CHART_CID, status_counts)

    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = OUTLOOK_SUBJECT

    attachment = mail.Attachments.Add(chart_file)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        OUTLOOK_CHART_CID
    )
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
        True
    )

    mail.HTMLBody = html_body
    mail.Display()

# --------------------------------------------------
# Jira API Functions
# --------------------------------------------------
def authenticate_jira(email, password):
    try:
        url = f"{JIRA_BASE_URL}/rest/api/2/myself"
        response = requests.get(
            url,
            headers={"Accept": "application/json"},
            auth=HTTPBasicAuth(email.strip(), password.strip()),
            timeout=20
        )
        if response.status_code == 200:
            return True, response.json()
        elif response.status_code == 401:
            return False, "Invalid email/username or password."
        elif response.status_code == 403:
            return False, "Access forbidden. Jira may require SSO or another authentication method."
        return False, f"Login failed. Status {response.status_code}: {response.text}"
    except requests.exceptions.RequestException as e:
        return False, f"Connection error: {str(e)}"

def fetch_jira_issues(email, password, jql, max_results=100):
    url = f"{JIRA_BASE_URL}/rest/api/2/search"
    params = {
        "jql": jql,
        "maxResults": max_results,
        "fields": ",".join([
            "summary", "status", "project", "issuetype", "assignee", "versions",
            "created", "comment",
            SEVERITY_FIELD, CUSTOMER_IMPACT_FIELD, COST_TO_FIX_FIELD,
            FREQUENCY_OCCURRENCE_FIELD, STEPS_FIELD, EXPECTED_RESULT_FIELD
        ])
    }
    return requests.get(
        url,
        headers={"Accept": "application/json"},
        auth=HTTPBasicAuth(email.strip(), password.strip()),
        params=params,
        timeout=30
    )

def update_jira_issue(email, password, jira_id, payload_fields):
    url = f"{JIRA_BASE_URL}/rest/api/2/issue/{jira_id}"
    response = requests.put(
        url,
        headers={"Accept": "application/json", "Content-Type": "application/json"},
        auth=HTTPBasicAuth(email.strip(), password.strip()),
        json={"fields": payload_fields},
        timeout=30
    )
    return response

def add_jira_comment(email, password, jira_id, comment_text):
    url = f"{JIRA_BASE_URL}/rest/api/2/issue/{jira_id}/comment"
    response = requests.post(
        url,
        headers={"Accept": "application/json", "Content-Type": "application/json"},
        auth=HTTPBasicAuth(email.strip(), password.strip()),
        json={"body": comment_text},
        timeout=30
    )
    return response

def build_jql(project_key, from_date, to_date):
    return f'''
    project = "{project_key}"
    AND issuetype in ("Bug", "Defect")
    AND created >= "{from_date}"
    AND created <= "{to_date}"
    ORDER BY created DESC
    '''

def issues_to_df(issues):
    rows = []
    for issue in issues:
        f = issue.get("fields", {})
        jira_id = issue.get("key", "")
        severity_value = safe_text(extract_value(f.get(SEVERITY_FIELD)), "")
        customer_impact_value = safe_text(extract_value(f.get(CUSTOMER_IMPACT_FIELD)), "")
        cost_to_fix_value = safe_text(extract_value(f.get(COST_TO_FIX_FIELD)), "")
        frequency_occurrence_value = safe_text(extract_value(f.get(FREQUENCY_OCCURRENCE_FIELD)), "")

        risk_factor = calculate_defect_risk_factor(
            severity_value, customer_impact_value, cost_to_fix_value, frequency_occurrence_value
        )

        rows.append({
            "jiraId": safe_text(jira_id, ""),
            "summary": safe_text(f.get("summary")),
            "status": safe_text(extract_value(f.get("status")), ""),
            "severity": severity_value,
            "customerImpact": customer_impact_value,
            "costToFix": cost_to_fix_value,
            "frequencyOccurrence": frequency_occurrence_value,
            "defectRiskFactor": risk_factor,
            "actionOwner": safe_text(extract_value(f.get("assignee")), "Unassigned"),
            "steps": safe_text(extract_value(f.get(STEPS_FIELD)), ""),
            "expectedResult": safe_text(extract_value(f.get(EXPECTED_RESULT_FIELD)), ""),
            "comments": safe_text(extract_comments(f.get("comment")), ""),
            "project": safe_text(extract_value(f.get("project", {}).get("key") if isinstance(f.get("project"), dict) else f.get("project")), ""),
            "fixVersions": safe_text(extract_value(f.get("versions")), ""),
            "created": f.get("created", ""),
            "jiraLink": f"{JIRA_BASE_URL}/browse/{jira_id}" if jira_id else ""
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["created"] = pd.to_datetime(df["created"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    return df

def build_update_fields(summary, severity, customer_impact, cost_to_fix, frequency_occurrence,
                        action_owner, steps, expected_result, project):
    severity_option_map = {
        "S1 : Blocker": "40277",
        "S2 : Critical": "40278",
        "S3 : Major": "40279",
        "S4 : Minor": "40280",
    }

    fields_payload = {"summary": summary if summary is not None else ""}

    if severity:
        severity_id = severity_option_map.get(severity)
        if not severity_id:
            raise ValueError(f"Invalid severity value: {severity}")
        fields_payload[SEVERITY_FIELD] = {"id": severity_id}

    if customer_impact != "":
        fields_payload[CUSTOMER_IMPACT_FIELD] = {"value": customer_impact}

    if cost_to_fix != "":
        fields_payload[COST_TO_FIX_FIELD] = {"value": cost_to_fix}

    if frequency_occurrence != "":
        fields_payload[FREQUENCY_OCCURRENCE_FIELD] = {"value": frequency_occurrence}

    fields_payload[STEPS_FIELD] = steps if steps is not None else ""
    fields_payload[EXPECTED_RESULT_FIELD] = expected_result if expected_result is not None else ""

    if action_owner and action_owner.lower() != "unassigned":
        fields_payload["assignee"] = {"name": action_owner}
    else:
        fields_payload["assignee"] = None

    if project != "":
        fields_payload["versions"] = [{"name": project}]

    return fields_payload

def refresh_issues():
    jql = build_jql(
        st.session_state.selected_project,
        st.session_state.from_date.strftime("%Y-%m-%d"),
        st.session_state.to_date.strftime("%Y-%m-%d")
    )
    response = fetch_jira_issues(
        st.session_state.jira_email,
        st.session_state.jira_password,
        jql,
        int(st.session_state.max_results)
    )
    if response.status_code == 200:
        data = response.json()
        st.session_state.issues_df = issues_to_df(data.get("issues", []))
        return True, f"Fetched {len(st.session_state.issues_df)} defects successfully for project {st.session_state.selected_project}."
    return False, f"Failed to fetch issues: {response.status_code} | {response.text}"

def build_severity_summary(df):
    if df.empty:
        return {"total": 0, "s1": 0, "s2": 0, "s3": 0, "s4": 0}

    severity_series = df["severity"].astype(str).str.upper().fillna("")
    return {
        "total": int(len(df)),
        "s1": int(severity_series.str.contains("S1", na=False).sum()),
        "s2": int(severity_series.str.contains("S2", na=False).sum()),
        "s3": int(severity_series.str.contains("S3", na=False).sum()),
        "s4": int(severity_series.str.contains("S4", na=False).sum())
    }

# --------------------------------------------------
# Ultra Premium Simple Chart Helpers
# --------------------------------------------------
def get_severity_counts(df):
    if df.empty or "severity" not in df.columns:
        return pd.Series([0, 0, 0, 0], index=["S1 : Blocker", "S2 : Critical", "S3 : Major", "S4 : Minor"])
    severity_order = ["S1 : Blocker", "S2 : Critical", "S3 : Major", "S4 : Minor"]
    counts = df["severity"].fillna("").value_counts()
    return pd.Series([int(counts.get(k, 0)) for k in severity_order], index=severity_order)

def get_status_counts(df):
    if df.empty:
        return pd.Series([0], index=["Unknown"])
    temp = df.copy()
    temp["status_norm"] = temp["status"].apply(normalize_status)
    order = ["Open", "In Progress", "Closed", "Deferred", "Unknown"]
    counts = temp["status_norm"].value_counts()
    final = pd.Series([int(counts.get(k, 0)) for k in order], index=order)
    final = final[final > 0] if (final > 0).any() else pd.Series([0], index=["Unknown"])
    return final

def render_severity_bar_chart(df):
    counts = get_severity_counts(df)
    labels = counts.index.tolist()
    values = counts.values.tolist()
    colors = ["#ef4444", "#f97316", "#f59e0b", "#10b981"]

    fig, ax = plt.subplots(figsize=(7, 4.2), facecolor="white")
    bars = ax.bar(labels, values, color=colors, width=0.56, edgecolor="white", linewidth=1.5)

    ax.set_title("Severity Distribution", fontsize=15, fontweight="bold", color="#0f172a", pad=14)
    ax.set_ylabel("Issue Count", fontsize=11, color="#475569")
    ax.set_xlabel("")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#cbd5e1")
    ax.spines["bottom"].set_color("#cbd5e1")
    ax.tick_params(axis="x", labelrotation=8, colors="#334155")
    ax.tick_params(axis="y", colors="#334155")
    ax.grid(axis="y", linestyle="--", alpha=0.25)

    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height + 0.05,
            f"{int(height)}",
            ha="center",
            va="bottom",
            fontsize=11,
            fontweight="bold",
            color="#0f172a"
        )

    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

def render_status_donut_chart(df):
    counts = get_status_counts(df)
    labels = counts.index.tolist()
    values = counts.values.tolist()

    color_map = {
        "Open": "#ef4444",
        "In Progress": "#f59e0b",
        "Closed": "#10b981",
        "Deferred": "#8b5cf6",
        "Unknown": "#94a3b8"
    }
    colors = [color_map.get(label, "#6366f1") for label in labels]

    fig, ax = plt.subplots(figsize=(6.6, 4.2), facecolor="white")
    ax.pie(
        values,
        labels=labels,
        colors=colors,
        autopct="%1.1f%%",
        startangle=120,
        wedgeprops={"width": 0.42, "edgecolor": "white", "linewidth": 2},
        textprops={"fontsize": 10, "fontweight": "bold", "color": "#334155"}
    )
    ax.set_title("Status Distribution", fontsize=15, fontweight="bold", color="#0f172a", pad=14)
    centre_circle = plt.Circle((0, 0), 0.42, fc="white")
    ax.add_artist(centre_circle)
    ax.text(0, 0.05, f"{sum(values)}", ha="center", va="center", fontsize=18, fontweight="bold", color="#0f172a")
    ax.text(0, -0.12, "Issues", ha="center", va="center", fontsize=10, color="#64748b")
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

# --------------------------------------------------
# Popup Dialog
# --------------------------------------------------
@st.dialog("Edit Jira Defect")
def edit_issue_dialog(selected_row):
    st.markdown(f"""
    <div class="dialog-hero">
        <div class="dialog-topline">Structured Defect Editor</div>
        <div class="dialog-title">Edit Defect: {selected_row['jiraId']}</div>
        <div class="dialog-subtitle">
            Review defect details, update triage values, assign ownership, and add a Jira comment in one structured workspace.
        </div>
        <div class="dialog-meta-grid">
            <div class="meta-pill">
                <div class="meta-label">Jira ID</div>
                <div class="meta-value">{selected_row['jiraId']}</div>
            </div>
            <div class="meta-pill">
                <div class="meta-label">Created</div>
                <div class="meta-value">{selected_row['created']}</div>
            </div>
            <div class="meta-pill">
                <div class="meta-label">Project</div>
                <div class="meta-value">{selected_row['project']}</div>
            </div>
            <div class="meta-pill">
                <div class="meta-label">Current Owner</div>
                <div class="meta-value">{selected_row['actionOwner']}</div>
            </div>
            <div class="meta-pill">
                <div class="meta-label">Status</div>
                <div class="meta-value">{selected_row.get('status', '')}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"**Jira Link:** [Open in Jira]({selected_row['jiraLink']})")

    issue_key = selected_row["jiraId"]

    sev_key = f"sev_{issue_key}"
    ci_key = f"ci_{issue_key}"
    ctf_key = f"ctf_{issue_key}"
    freq_key = f"freq_{issue_key}"
    risk_key = f"risk_{issue_key}"

    summary_key = f"summary_{issue_key}"
    action_owner_key = f"action_owner_{issue_key}"
    project_key = f"project_{issue_key}"
    steps_key = f"steps_{issue_key}"
    expected_result_key = f"expected_result_{issue_key}"
    new_comment_key = f"new_comment_{issue_key}"

    if sev_key not in st.session_state:
        st.session_state[sev_key] = selected_row["severity"] if selected_row["severity"] in SEVERITY_OPTIONS else ""
    if ci_key not in st.session_state:
        st.session_state[ci_key] = selected_row["customerImpact"] if selected_row["customerImpact"] in COMMON_PRIORITY_OPTIONS else ""
    if ctf_key not in st.session_state:
        st.session_state[ctf_key] = selected_row["costToFix"] if selected_row["costToFix"] in COMMON_PRIORITY_OPTIONS else ""
    if freq_key not in st.session_state:
        st.session_state[freq_key] = selected_row["frequencyOccurrence"] if selected_row["frequencyOccurrence"] in COMMON_PRIORITY_OPTIONS else ""

    if summary_key not in st.session_state:
        st.session_state[summary_key] = selected_row["summary"]
    if action_owner_key not in st.session_state:
        st.session_state[action_owner_key] = "" if selected_row["actionOwner"] == "Unassigned" else selected_row["actionOwner"]
    if project_key not in st.session_state:
        st.session_state[project_key] = selected_row.get("fixVersions", "")
    if steps_key not in st.session_state:
        st.session_state[steps_key] = selected_row["steps"]
    if expected_result_key not in st.session_state:
        st.session_state[expected_result_key] = selected_row["expectedResult"]
    if new_comment_key not in st.session_state:
        st.session_state[new_comment_key] = ""

    if risk_key not in st.session_state:
        update_risk_factor_session(issue_key)

    st.markdown('<div class="dialog-section">', unsafe_allow_html=True)
    st.markdown('<div class="dialog-section-title">Core Information</div>', unsafe_allow_html=True)

    e1, e2 = st.columns(2)

    with e1:
        st.text_input("Summary", key=summary_key)
        st.selectbox("Severity", SEVERITY_OPTIONS, key=sev_key, on_change=update_risk_factor_session, args=(issue_key,))
        st.selectbox("Customer Impact", COMMON_PRIORITY_OPTIONS, key=ci_key, on_change=update_risk_factor_session, args=(issue_key,))
        st.selectbox("Cost To Fix", COMMON_PRIORITY_OPTIONS, key=ctf_key, on_change=update_risk_factor_session, args=(issue_key,))

    with e2:
        st.selectbox("Frequency Occurrence", COMMON_PRIORITY_OPTIONS, key=freq_key, on_change=update_risk_factor_session, args=(issue_key,))
        st.text_input("Action Owner", key=action_owner_key)
        st.text_input("Project", key=project_key)
        st.text_input("Created", value=selected_row["created"], disabled=True)

    st.number_input(
        "Defect Risk Factor",
        value=float(st.session_state.get(risk_key, 0.0)),
        disabled=True,
        help="Auto-calculated live based on Severity, Customer Impact, Cost To Fix, and Frequency Occurrence."
    )

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="dialog-section">', unsafe_allow_html=True)
    st.markdown('<div class="dialog-section-title">Defect Details</div>', unsafe_allow_html=True)
    st.text_area("Steps", key=steps_key, height=150)
    st.text_area("Expected Result", key=expected_result_key, height=150)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="dialog-section">', unsafe_allow_html=True)
    st.markdown('<div class="dialog-section-title">Comments</div>', unsafe_allow_html=True)
    st.text_area("Comments History", value=selected_row["comments"], height=190, disabled=True)
    st.text_area("Add New Comment", key=new_comment_key, height=130, placeholder="Type a new Jira comment here...")
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        submit_update = st.button("Update Jira Defect", use_container_width=True)

    with c2:
        clear_selection = st.button("Close", use_container_width=True)

    if clear_selection:
        st.session_state.selected_jira_id = None
        for k in [
            sev_key, ci_key, ctf_key, freq_key, risk_key,
            summary_key, action_owner_key, project_key,
            steps_key, expected_result_key, new_comment_key
        ]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    if submit_update:
        try:
            payload_fields = build_update_fields(
                summary=st.session_state[summary_key].strip(),
                severity=st.session_state[sev_key].strip(),
                customer_impact=st.session_state[ci_key].strip(),
                cost_to_fix=st.session_state[ctf_key].strip(),
                frequency_occurrence=st.session_state[freq_key].strip(),
                action_owner=st.session_state[action_owner_key].strip(),
                steps=st.session_state[steps_key].strip(),
                expected_result=st.session_state[expected_result_key].strip(),
                project=st.session_state[project_key].strip()
            )

            with st.spinner("Updating Jira defect..."):
                update_response = update_jira_issue(
                    st.session_state.jira_email,
                    st.session_state.jira_password,
                    selected_row["jiraId"],
                    payload_fields
                )

            if update_response.status_code not in [200, 204]:
                st.error(f"Update failed: {update_response.status_code} | {update_response.text}")
                return

            if st.session_state[new_comment_key].strip():
                with st.spinner("Adding Jira comment..."):
                    comment_response = add_jira_comment(
                        st.session_state.jira_email,
                        st.session_state.jira_password,
                        selected_row["jiraId"],
                        st.session_state[new_comment_key].strip()
                    )
                if comment_response.status_code not in [200, 201]:
                    st.warning(f"Defect updated, but comment failed: {comment_response.status_code} | {comment_response.text}")

            st.success(f"Jira defect {selected_row['jiraId']} updated successfully.")
            ok, msg = refresh_issues()
            st.session_state.selected_jira_id = None

            for k in [
                sev_key, ci_key, ctf_key, freq_key, risk_key,
                summary_key, action_owner_key, project_key,
                steps_key, expected_result_key, new_comment_key
            ]:
                if k in st.session_state:
                    del st.session_state[k]

            if not ok:
                st.warning(msg)
            st.rerun()

        except ValueError as ve:
            st.error(str(ve))
        except Exception as ex:
            st.error(f"Unexpected error: {str(ex)}")

# --------------------------------------------------
# Table Builder
# --------------------------------------------------
def render_clickable_table(export_df):
    summary_renderer = JsCode("""
    class SummaryCellRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            this.eGui.innerText = params.value || '';
            this.eGui.style.color = '#1e293b';
            this.eGui.style.fontWeight = '700';
            this.eGui.style.cursor = 'pointer';
            this.eGui.style.lineHeight = '1.45';
            this.eGui.style.whiteSpace = 'normal';
        }
        getGui() {
            return this.eGui;
        }
    }
    """)

    jira_renderer = JsCode("""
    class JiraLinkRenderer {
        init(params) {
            this.eGui = document.createElement('a');
            this.eGui.innerText = params.value || '';
            this.eGui.setAttribute('href', params.data.jiraLink);
            this.eGui.setAttribute('target', '_blank');
            this.eGui.style.color = '#2563eb';
            this.eGui.style.fontWeight = '900';
            this.eGui.style.textDecoration = 'none';
        }
        getGui() {
            return this.eGui;
        }
    }
    """)

    gb = GridOptionsBuilder.from_dataframe(export_df)

    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filter=True,
        editable=False,
        floatingFilter=False,
        wrapText=True,
        autoHeight=True,
        suppressMenu=True,
    )

    gb.configure_selection(
        selection_mode="single",
        use_checkbox=False,
        pre_selected_rows=[]
    )

    gb.configure_pagination(
        enabled=True,
        paginationAutoPageSize=False,
        paginationPageSize=10
    )

    gb.configure_grid_options(
        rowHeight=54,
        headerHeight=48,
        suppressRowClickSelection=False,
        animateRows=True,
        enableCellTextSelection=True,
        tooltipShowDelay=0
    )

    gb.configure_column("jiraId", headerName="Jira ID", width=130, pinned="left", cellRenderer=jira_renderer)
    gb.configure_column("summary", headerName="Summary", width=420, cellRenderer=summary_renderer)
    gb.configure_column("status", width=140)
    gb.configure_column("severity", width=150)
    gb.configure_column("customerImpact", headerName="Customer Impact", width=170)
    gb.configure_column("actionOwner", headerName="Owner", width=170)
    gb.configure_column("project", width=140)
    gb.configure_column("created", width=160)
    gb.configure_column("defectRiskFactor", headerName="Risk", width=110)
    gb.configure_column("costToFix", hide=True)
    gb.configure_column("frequencyOccurrence", hide=True)
    gb.configure_column("steps", hide=True)
    gb.configure_column("expectedResult", hide=True)
    gb.configure_column("comments", hide=True)
    gb.configure_column("jiraLink", hide=True)
    gb.configure_column("fixVersions", hide=True)

    grid_options = gb.build()

    grid_response = AgGrid(
        export_df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        allow_unsafe_jscode=True,
        enable_enterprise_modules=False,
        fit_columns_on_grid_load=False,
        height=560,
        theme="streamlit",
        reload_data=False,
        custom_css={
            ".ag-root-wrapper": {
                "border": "1px solid #e2e8f0",
                "border-radius": "18px",
                "overflow": "hidden",
                "box-shadow": "0 8px 26px rgba(15,23,42,0.06)",
                "background": "#ffffff"
            },
            ".ag-header": {
                "background": "#f8fafc !important",
                "border-bottom": "1px solid #e2e8f0 !important"
            },
            ".ag-header-cell": {
                "font-weight": "800 !important",
                "color": "#475569 !important",
                "font-size": "12px !important",
                "text-transform": "uppercase",
                "letter-spacing": "0.06em",
                "border-right": "1px solid #f1f5f9 !important"
            },
            ".ag-row": {
                "font-size": "13px !important",
                "color": "#334155 !important",
                "background": "#ffffff !important",
                "border-bottom": "1px solid #f5f7fb !important",
                "transition": "all 0.18s ease !important"
            },
            ".ag-row-hover": {
                "background-color": "#f8fbff !important"
            },
            ".ag-row-selected": {
                "background-color": "#eff6ff !important",
                "box-shadow": "inset 3px 0 0 #2563eb !important"
            },
            ".ag-cell": {
                "display": "flex",
                "align-items": "center",
                "line-height": "1.5 !important",
                "padding-top": "8px !important",
                "padding-bottom": "8px !important"
            },
            ".ag-paging-panel": {
                "border-top": "1px solid #eef2f7 !important",
                "padding": "12px 14px !important",
                "background": "#fbfdff !important",
                "color": "#64748b !important"
            },
            ".ag-icon": {
                "color": "#64748b !important"
            }
        }
    )

    return grid_response

# --------------------------------------------------
# GLOBAL UI CSS - Ultra Premium Simple
# --------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

:root{
    --bg:#f4f7fb;
    --panel:rgba(255,255,255,0.92);
    --line:#e2e8f0;
    --text:#0f172a;
    --muted:#64748b;
    --primary:#2563eb;
    --primary-soft:#eff6ff;
    --shadow-sm:0 8px 24px rgba(15,23,42,0.06);
    --shadow-md:0 18px 44px rgba(15,23,42,0.10);
    --radius:22px;
}

html, body, [class*="css"], .stApp{
    font-family:'Inter', sans-serif;
}

html, body, [data-testid="stAppViewContainer"], .stApp{
    background:
        radial-gradient(circle at top left, rgba(37,99,235,0.08), transparent 25%),
        radial-gradient(circle at top right, rgba(20,184,166,0.06), transparent 20%),
        var(--bg) !important;
}

#MainMenu, header, footer{
    visibility:hidden;
}

.block-container{
    max-width:1480px !important;
    padding-top:1rem !important;
    padding-bottom:1rem !important;
    padding-left:1rem !important;
    padding-right:1rem !important;
}

.hero-card,
.panel-card,
.stat-card,
.chart-card{
    background:var(--panel);
    backdrop-filter:blur(10px);
    border:1px solid rgba(226,232,240,0.95);
    border-radius:var(--radius);
    box-shadow:var(--shadow-sm);
}

.hero-card{
    padding:28px;
    min-height:165px;
    position:relative;
    overflow:hidden;
}

.hero-card::before{
    content:"";
    position:absolute;
    top:-60px;
    right:-60px;
    width:180px;
    height:180px;
    background:radial-gradient(circle, rgba(37,99,235,0.16), transparent 70%);
    border-radius:50%;
}

.hero-kicker{
    display:inline-flex;
    align-items:center;
    gap:8px;
    padding:7px 14px;
    border-radius:999px;
    background:var(--primary-soft);
    color:var(--primary);
    font-size:12px;
    font-weight:900;
    letter-spacing:.10em;
    text-transform:uppercase;
    margin-bottom:14px;
    border:1px solid #dbeafe;
}

.hero-title{
    font-size:2.15rem;
    font-weight:900;
    color:var(--text);
    line-height:1.02;
    margin-bottom:10px;
    letter-spacing:-.04em;
}

.hero-subtitle{
    color:var(--muted);
    font-size:.97rem;
    line-height:1.8;
    max-width:880px;
}

.user-card{
    padding:24px;
    min-height:165px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
}

.user-avatar{
    width:62px;
    height:62px;
    border-radius:50%;
    background:linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
    color:var(--primary);
    display:flex;
    align-items:center;
    justify-content:center;
    margin:0 auto 12px auto;
    font-weight:900;
    font-size:1rem;
    border:1px solid #dbeafe;
    box-shadow:0 10px 24px rgba(37,99,235,0.14);
}

.user-label{
    font-size:.72rem;
    color:#94a3b8;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:.08em;
}

.user-name{
    font-size:1rem;
    color:var(--text);
    font-weight:800;
    margin-top:5px;
}

.panel-card{
    padding:22px;
    margin-bottom:18px;
}

.section-kicker{
    color:var(--primary);
    font-size:.74rem;
    font-weight:900;
    letter-spacing:.10em;
    text-transform:uppercase;
    margin-bottom:7px;
}

.section-title{
    color:var(--text);
    font-size:1.18rem;
    font-weight:900;
    margin-bottom:7px;
    letter-spacing:-.02em;
}

.section-copy{
    color:var(--muted);
    font-size:.93rem;
    line-height:1.7;
}

.stat-card{
    padding:20px;
    min-height:138px;
    position:relative;
    overflow:hidden;
    transition:all .22s ease;
}

.stat-card:hover{
    transform:translateY(-4px);
    box-shadow:0 18px 36px rgba(15,23,42,0.10);
}

.stat-card::before{
    content:"";
    position:absolute;
    left:0;
    top:0;
    width:100%;
    height:5px;
    background:var(--accent, #2563eb);
}

.stat-top{
    display:flex;
    align-items:center;
    justify-content:space-between;
    margin-bottom:18px;
}

.stat-icon{
    width:50px;
    height:50px;
    border-radius:16px;
    display:flex;
    align-items:center;
    justify-content:center;
    background:var(--soft, #eff6ff);
    color:var(--accent, #2563eb);
    font-size:1.15rem;
    font-weight:800;
    box-shadow:inset 0 0 0 1px rgba(255,255,255,0.45);
}

.stat-chip{
    padding:6px 10px;
    border-radius:999px;
    background:var(--soft, #eff6ff);
    color:var(--accent, #2563eb);
    font-size:.72rem;
    font-weight:900;
    letter-spacing:.03em;
}

.stat-value{
    font-size:2.15rem;
    font-weight:900;
    line-height:1;
    color:var(--text);
    margin-bottom:8px;
    letter-spacing:-.05em;
}

.stat-label{
    font-size:.9rem;
    color:var(--muted);
    font-weight:700;
}

.chart-card{
    padding:20px;
    min-height:100%;
}

.chart-title{
    font-size:1rem;
    font-weight:900;
    color:var(--text);
    margin-bottom:5px;
}

.chart-copy{
    font-size:.86rem;
    color:var(--muted);
    margin-bottom:10px;
    line-height:1.6;
}

.dialog-hero{
    background:#f9fbff;
    border:1px solid var(--line);
    border-radius:18px;
    padding:18px;
    margin-bottom:14px;
}

.dialog-topline{
    color:var(--primary);
    font-size:.72rem;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.08em;
    margin-bottom:8px;
}

.dialog-title{
    color:var(--text);
    font-size:1.12rem;
    font-weight:900;
    margin-bottom:6px;
}

.dialog-subtitle{
    color:var(--muted);
    font-size:.88rem;
    line-height:1.6;
    font-weight:500;
}

.dialog-section{
    background:#ffffff;
    border:1px solid var(--line);
    border-radius:16px;
    padding:18px;
    margin-bottom:14px;
    box-shadow:var(--shadow-sm);
}

.dialog-section-title{
    color:var(--text);
    font-size:.96rem;
    font-weight:900;
    margin-bottom:12px;
}

.dialog-meta-grid{
    display:grid;
    grid-template-columns:repeat(2, minmax(0, 1fr));
    gap:10px;
    margin-top:12px;
}

.meta-pill{
    background:#ffffff;
    border:1px solid var(--line);
    border-radius:12px;
    padding:12px;
}

.meta-label{
    color:#94a3b8;
    font-size:.72rem;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:.08em;
    margin-bottom:4px;
}

.meta-value{
    color:var(--text);
    font-size:.92rem;
    font-weight:700;
}

.stButton > button,
.stDownloadButton > button,
.stFormSubmitButton > button{
    border:none !important;
    border-radius:14px !important;
    min-height:48px !important;
    font-weight:800 !important;
    color:white !important;
    background:linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
    box-shadow:0 12px 24px rgba(37,99,235,0.20) !important;
    transition:all .18s ease !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover,
.stFormSubmitButton > button:hover{
    transform:translateY(-2px) !important;
}

div[data-baseweb="input"] > div,
div[data-baseweb="base-input"] > div,
div[data-baseweb="select"] > div,
div[data-baseweb="textarea"] > div{
    border-radius:12px !important;
    background:#ffffff !important;
    border:1px solid #cbd5e1 !important;
    min-height:48px !important;
    box-shadow:none !important;
}

label, .stDateInput label, .stNumberInput label, .stTextInput label, .stTextArea label, .stSelectbox label{
    color:#334155 !important;
    font-weight:800 !important;
    font-size:.88rem !important;
}

.stAlert{
    border-radius:14px !important;
    border:1px solid var(--line) !important;
    background:#ffffff !important;
}

div[role="dialog"] > div{
    border-radius:24px !important;
    background:#ffffff !important;
    border:1px solid var(--line) !important;
    box-shadow:0 22px 44px rgba(15,23,42,0.12) !important;
}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# LOGIN PAGE CSS ONLY
# --------------------------------------------------
if not st.session_state.logged_in:
    st.markdown("""
    <style>
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > .main {
        min-height: 100vh !important;
        background: #f8fafc !important;
    }

    .main .block-container{
        max-width: 100% !important;
        min-height: 100vh !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }

    .login-logo{
        width:46px;
        height:46px;
        border-radius:14px;
        background:linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        display:flex;
        align-items:center;
        justify-content:center;
        color:#2563eb;
        font-size:1rem;
        font-weight:900;
        margin:0 auto 10px auto;
        border:1px solid #dbeafe;
        box-shadow:0 6px 16px rgba(37,99,235,0.08);
    }

    .login-mini-badge{
        display:flex;
        justify-content:center;
        margin-bottom:8px;
    }

    .login-mini-badge span{
        display:inline-flex;
        padding:5px 10px;
        border-radius:999px;
        background:#eff6ff;
        border:1px solid #dbeafe;
        color:#2563eb;
        font-size:.64rem;
        font-weight:900;
        letter-spacing:.08em;
        text-transform:uppercase;
    }

    .login-form-title{
        text-align:center;
        color:#1e293b;
        font-size:1.35rem;
        font-weight:900;
        margin-bottom:4px;
        letter-spacing:-.03em;
        line-height:1.1;
    }

    .login-form-copy{
        text-align:center;
        color:#64748b;
        font-size:.84rem;
        line-height:1.45;
        font-weight:500;
        margin-bottom:1rem;
    }

    .login-footnote{
        text-align:center;
        color:#94a3b8;
        font-size:.74rem;
        margin-top:12px;
        font-weight:600;
        line-height:1.35;
    }
    </style>
    """, unsafe_allow_html=True)

# --------------------------------------------------
# Login
# --------------------------------------------------
if not st.session_state.logged_in:

    st.markdown("""
        <div class="login-logo">DT</div>
        <div class="login-mini-badge"><span>Secure Enterprise Access</span></div>
        <div class="login-form-title">Secure Sign In</div>
        <div class="login-form-copy">
            Login with your Jira credentials to access the defect triage workspace.
        </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        email = st.text_input("Email / Username", placeholder="Enter your Jira email or username")
        password = st.text_input("Password", type="password", placeholder="Enter your password")
        submit = st.form_submit_button("Enter Dashboard")

    st.markdown("""
        <div class="login-footnote">
            Jira-authenticated access • No backend changes • Leadership-ready enterprise UI
        </div>
    """, unsafe_allow_html=True)

    if submit:
        if not email or not password:
            st.warning("Please enter both email/username and password.")
        else:
            with st.spinner("Authenticating with Jira..."):
                success, result = authenticate_jira(email, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.jira_email = email
                st.session_state.jira_password = password
                st.session_state.jira_user = result
                st.rerun()
            else:
                st.error(result)

# --------------------------------------------------
# Dashboard
# --------------------------------------------------
else:
    user = st.session_state.jira_user or {}
    user_name = user.get("displayName", "User")
    initials = "".join([part[0] for part in user_name.split()[:2]]).upper() if user_name else "U"

    with st.sidebar:
        st.markdown("""
        <div class="panel-card">
            <div class="section-kicker">Workspace</div>
            <div class="section-title">Navigation</div>
            <div class="section-copy">Use this workspace to fetch, review, edit, and export Jira defects.</div>
        </div>
        """, unsafe_allow_html=True)

        st.radio(
            "Go to",
            ["Defect Control Panel", "Reports", "TIP Overview"],
            index=0,
            label_visibility="collapsed"
        )

    _, main = st.columns([0.001, 1], gap="small")

    with main:
        top1, top2, top3 = st.columns([6.2, 2.0, 1.3])

        with top1:
            st.markdown(f"""
            <div class="hero-card">
                <div class="hero-kicker">Premium Executive Dashboard</div>
                <div class="hero-title">Defect Triage Process Automation</div>
                <div class="hero-subtitle">
                    Retrieve, monitor, review, and export Jira defects for project <b>{st.session_state.selected_project}</b>
                    through a cleaner and more premium triage workspace designed for leadership visibility.
                </div>
            </div>
            """, unsafe_allow_html=True)

        with top2:
            st.markdown(f"""
            <div class="hero-card user-card">
                <div>
                    <div class="user-avatar">{initials}</div>
                    <div class="user-label">Logged In User</div>
                    <div class="user-name">{user_name}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        with top3:
            if st.button("Logout", use_container_width=True):
                logout()

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        st.markdown("""
        <div class="panel-card">
            <div class="section-kicker">Control Panel</div>
            <div class="section-title">Fetch Jira Defects</div>
            <div class="section-copy">
                Select the project, date range, and result volume to refresh the defect dashboard instantly.
            </div>
        </div>
        """, unsafe_allow_html=True)

        f1, f2, f3, f4, f5 = st.columns([1.0, 1.1, 1.1, 0.9, 1.0])

        with f1:
            st.session_state.selected_project = st.selectbox(
                "Select Project",
                PROJECT_OPTIONS,
                index=PROJECT_OPTIONS.index(st.session_state.selected_project) if st.session_state.selected_project in PROJECT_OPTIONS else 0
            )

        with f2:
            st.session_state.from_date = st.date_input("From Date", value=st.session_state.from_date)

        with f3:
            st.session_state.to_date = st.date_input("To Date", value=st.session_state.to_date)

        with f4:
            st.session_state.max_results = int(st.number_input(
                "Max Results",
                min_value=10,
                max_value=1000,
                value=int(st.session_state.max_results),
                step=10
            ))

        with f5:
            st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
            fetch_btn = st.button("Fetch Defects", use_container_width=True)

        if fetch_btn:
            if st.session_state.from_date > st.session_state.to_date:
                st.error("From Date cannot be greater than To Date.")
            else:
                with st.spinner(f"Fetching {st.session_state.selected_project} Jira defects..."):
                    ok, msg = refresh_issues()
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

        df = st.session_state.issues_df.copy()

        if not df.empty:
            severity_summary = build_severity_summary(df)

            st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

            s1, s2, s3, s4, s5 = st.columns(5)

            with s1:
                st.markdown(f"""
                <div class="stat-card" style="--accent:#2563eb; --soft:#eff6ff;">
                    <div class="stat-top">
                        <div class="stat-icon">📋</div>
                        <div class="stat-chip">TOTAL</div>
                    </div>
                    <div class="stat-value">{severity_summary['total']}</div>
                    <div class="stat-label">Total Issues</div>
                </div>
                """, unsafe_allow_html=True)

            with s2:
                st.markdown(f"""
                <div class="stat-card" style="--accent:#ef4444; --soft:#fef2f2;">
                    <div class="stat-top">
                        <div class="stat-icon">🔴</div>
                        <div class="stat-chip">S1</div>
                    </div>
                    <div class="stat-value">{severity_summary['s1']}</div>
                    <div class="stat-label">Blocker Issues</div>
                </div>
                """, unsafe_allow_html=True)

            with s3:
                st.markdown(f"""
                <div class="stat-card" style="--accent:#f97316; --soft:#fff7ed;">
                    <div class="stat-top">
                        <div class="stat-icon">🟠</div>
                        <div class="stat-chip">S2</div>
                    </div>
                    <div class="stat-value">{severity_summary['s2']}</div>
                    <div class="stat-label">Critical Issues</div>
                </div>
                """, unsafe_allow_html=True)

            with s4:
                st.markdown(f"""
                <div class="stat-card" style="--accent:#f59e0b; --soft:#fffbeb;">
                    <div class="stat-top">
                        <div class="stat-icon">🟡</div>
                        <div class="stat-chip">S3</div>
                    </div>
                    <div class="stat-value">{severity_summary['s3']}</div>
                    <div class="stat-label">Major Issues</div>
                </div>
                """, unsafe_allow_html=True)

            with s5:
                st.markdown(f"""
                <div class="stat-card" style="--accent:#10b981; --soft:#ecfdf5;">
                    <div class="stat-top">
                        <div class="stat-icon">🟢</div>
                        <div class="stat-chip">S4</div>
                    </div>
                    <div class="stat-value">{severity_summary['s4']}</div>
                    <div class="stat-label">Minor Issues</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

            c1, c2 = st.columns(2)

            with c1:
                st.markdown("""
                <div class="chart-card">
                    <div class="chart-title">Severity Distribution</div>
                    <div class="chart-copy">Overview of S1, S2, S3, and S4 defects.</div>
                </div>
                """, unsafe_allow_html=True)
                render_severity_bar_chart(df)

            with c2:
                st.markdown("""
                <div class="chart-card">
                    <div class="chart-title">Status Distribution</div>
                    <div class="chart-copy">Overview of Open, In Progress, Closed, and Deferred defects.</div>
                </div>
                """, unsafe_allow_html=True)
                render_status_donut_chart(df)

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

            st.markdown(f"""
            <div class="panel-card">
                <div class="section-kicker">Search & Review</div>
                <div class="section-title">Search Defects - {st.session_state.selected_project}</div>
                <div class="section-copy">
                    Quickly find defects by Jira ID, summary, severity, owner, comments, reproduction steps, expected result, or project values.
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.session_state.search_text = st.text_input(
                "Search defects",
                value=st.session_state.search_text,
                placeholder=f"Search {st.session_state.selected_project} defects..."
            )

            filtered_df = filter_df(df, st.session_state.search_text)

            export_df = filtered_df[
                [
                    "jiraId",
                    "summary",
                    "status",
                    "severity",
                    "customerImpact",
                    "costToFix",
                    "frequencyOccurrence",
                    "defectRiskFactor",
                    "actionOwner",
                    "steps",
                    "expectedResult",
                    "comments",
                    "project",
                    "fixVersions",
                    "created",
                    "jiraLink"
                ]
            ].copy()

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

            st.markdown(f"""
            <div class="panel-card">
                <div class="section-kicker">Defect Table</div>
                <div class="section-title">Defect Review Table - {st.session_state.selected_project}</div>
                <div class="section-copy">
                    Review the fetched Jira defects below. Select a row to open the premium defect editor and update Jira directly.
                </div>
            </div>
            """, unsafe_allow_html=True)

            h1, h2, h3 = st.columns([4, 1.25, 1.6])

            with h1:
                st.markdown("")

            with h2:
                csv = export_df.drop(columns=["jiraLink"]).to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Export CSV",
                    data=csv,
                    file_name=f"{st.session_state.selected_project.lower()}_defect_data.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            with h3:
                if st.button("Open in Outlook App", use_container_width=True):
                    try:
                        if export_df.empty:
                            st.warning("No defect data available to open in Outlook.")
                        else:
                            outlook_df = export_df.drop(columns=["jiraLink"], errors="ignore").copy()
                            open_outlook_draft_with_defects(outlook_df)
                            st.success("Outlook draft opened successfully.")
                    except Exception as ex:
                        st.error(f"Unable to open Outlook draft: {str(ex)}")

            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

            grid_response = render_clickable_table(export_df)
            selected_rows = grid_response.get("selected_rows", [])

            if isinstance(selected_rows, pd.DataFrame):
                if not selected_rows.empty:
                    st.session_state.selected_jira_id = selected_rows.iloc[0]["jiraId"]
            elif isinstance(selected_rows, list):
                if len(selected_rows) > 0:
                    st.session_state.selected_jira_id = selected_rows[0]["jiraId"]

            if st.session_state.selected_jira_id:
                selected_matches = filtered_df[filtered_df["jiraId"] == st.session_state.selected_jira_id]
                if not selected_matches.empty:
                    selected_row = selected_matches.iloc[0]
                    edit_issue_dialog(selected_row)

        else:
            st.markdown(f"""
            <div class="panel-card">
                <div class="section-kicker">Ready</div>
                <div class="section-title">No Defects Loaded Yet</div>
                <div class="section-copy">
                    Select a project, choose the date range, and click <b>Fetch Defects</b> to load Jira data.
                </div>
            </div>
            """, unsafe_allow_html=True)
