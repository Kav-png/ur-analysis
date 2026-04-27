import re
import json
import io
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from st_copy import copy_button

st.set_page_config(page_title="Incident Pattern Analyser", layout="wide")

st.title("Incident Pattern Analyser")
st.caption("Surface recurring issues for senior management reporting")

# ── session state ──────────────────────────────────────────────────────────────
for key in ("df", "prompt", "groups"):
    if key not in st.session_state:
        st.session_state[key] = None


# ── helpers ────────────────────────────────────────────────────────────────────
# Get current month and year
now = datetime.now()
current_month = now.strftime("%B")  # e.g., "April"
current_year = now.year

# List of months and years (e.g., last 5 years)
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
years = [str(y) for y in range(current_year, current_year - 5, -1)]


def preprocess(text: str, maxlen: int = 200) -> str:
    text = str(text)
    # Remove ticket IDs (e.g., INC00123, UR123456)
    text = re.sub(r'\b[A-Z]{2,4}\d{5,}\b', '', text)
    # Remove emails
    text = re.sub(r'\S+@\S+', '', text)
    # Remove long numbers (likely IDs, phone numbers)
    text = re.sub(r'\b\d{5,}\b', '', text)
    # Remove bracketed tags and content
    text = re.sub(r'\[.*?\]', '', text)
    # Remove boilerplate/disclaimer lines
    text = re.sub(r'Client Identifying Data.*?not allowed.*?\.?', '', text, flags=re.IGNORECASE)
    text = re.sub(r'No CID Disclaimer accepted: TRUE', '', text, flags=re.IGNORECASE)
    text = re.sub(r'I have read and understood the disclaimer.*', '', text, flags=re.IGNORECASE)
    # Remove greetings and sign-offs (start/end of lines)
    text = re.sub(r'^(hi|hello|dear|greetings)[\s,]+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(thanks|thank you|regards|best regards|cheers)[\s,]*$', '', text, flags=re.IGNORECASE)
    # Remove phone numbers and user details
    text = re.sub(r'Tel\.?[\s:\-]*\d+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'User details:.*?(?=Topic:|$)', '', text, flags=re.IGNORECASE)
    # Remove (T12345) style tags
    text = re.sub(r'\(T\d+\)', '', text)
    # Remove long digit/space/hyphen combos (phone numbers)
    text = re.sub(r'\+?\d[\d\s\-]{7,}', '', text)
    # Remove repeated dashes/underscores
    text = re.sub(r'[-_]{2,}', '', text)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    # Truncate to maxlen
    return text

def build_prompt(rows: list[dict]) -> str:
    incidents = "\n".join(
        f"{i+1}. [{r['number']}] {r['description']}" for i, r in enumerate(rows)
    )
    return f"""You are analyzing {len(rows)} IT incidents from an Investment Bank.
Your output will be used by senior management to reduce recurring incidents.

Rules:
- Extract the application or system name from the description text.
- Group incidents together if they share the same root cause OR if they represent repeated requests for the same operational process or workflow (e.g., repeated requests to create Storm IDs, repeated manual report retriggers, etc.).
- Be specific: write "Users locked out after AD password reset" not "access issue".
- If only one incident exists for a particular issue, include it as its own group (do not omit unique incidents).
- **Every incident must appear in the output, either as part of a group or as its own group. Do not omit any incident for any reason.**
- If the application cannot be identified from the text, use "Unknown System".
- Return ONLY valid JSON — no markdown fences, no explanation, nothing else.

Output format:
{{
  "groups": [
    {{
      "application": "name of the system or application",
      "issue": "precise description of the problem or request",
      "count": <number>,
      "incident_numbers": ["INC001", "INC002"],
      "business_impact": "one sentence describing the business effect",
      "recommended_action": "one concrete action to prevent recurrence or address the issue"
    }}
  ]
}}


Incidents:
{incidents}
"""


def parse_response(raw: str) -> list[dict] | None:
    # strip markdown fences if present
    cleaned = re.sub(r"```json|```", "", raw).strip()
    try:
        data = json.loads(cleaned)
        return data.get("groups", [])
    except json.JSONDecodeError:
        return None


def build_excel(groups: list[dict]) -> bytes:
    wb = Workbook()

    # ── Sheet 1: Management Summary ──
    ws1 = wb.active
    ws1.title = "Management Summary"

    headers = ["Application", "Issue", "Count", "Business Impact", "Recommended Action"]
    ws1.append(headers)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F3864")
    alt_fill = PatternFill("solid", fgColor="DCE6F1")

    for col_idx, _ in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    ws1.freeze_panes = "A2"

    sorted_groups = sorted(groups, key=lambda g: g.get("count", 0), reverse=True)
    for row_idx, g in enumerate(sorted_groups, 2):
        ws1.append([
            g.get("application", ""),
            g.get("issue", ""),
            g.get("count", 0),
            g.get("business_impact", ""),
            g.get("recommended_action", ""),
        ])
        if row_idx % 2 == 0:
            for col_idx in range(1, len(headers) + 1):
                ws1.cell(row=row_idx, column=col_idx).fill = alt_fill
        for col_idx in range(1, len(headers) + 1):
            ws1.cell(row=row_idx, column=col_idx).alignment = Alignment(wrap_text=True)

    col_widths = [28, 48, 10, 48, 48]
    for i, width in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = width

    # ── Sheet 2: Incident Detail ──
    ws2 = wb.create_sheet("Incident Detail")
    detail_headers = ["Group", "Application", "Issue", "Incident Numbers"]
    ws2.append(detail_headers)

    for col_idx, _ in enumerate(detail_headers, 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    ws2.freeze_panes = "A2"

    for group_idx, g in enumerate(sorted_groups, 1):
        inc_numbers = ", ".join(g.get("incident_numbers", []))
        ws2.append([
            group_idx,
            g.get("application", ""),
            g.get("issue", ""),
            inc_numbers,
        ])
        ws2.cell(row=group_idx + 1, column=4).alignment = Alignment(wrap_text=True)

    for col_idx, width in enumerate([8, 28, 48, 60], 1):
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload Excel
# ══════════════════════════════════════════════════════════════════════════════
st.header("Step 1 — Upload Incident Excel")

uploaded = st.file_uploader("Upload your monthly incidents Excel file", type="xlsx")
if uploaded:
    df_raw = pd.read_excel(uploaded)
    cols = df_raw.columns.tolist()

    col1, col2 = st.columns(2)
    num_col = col1.selectbox(
        "Incident number column", cols,
        index=next((i for i, c in enumerate(cols) if "number" in c.lower() or "id" in c.lower() or "inc" in c.lower()), 0)
    )
    desc_col = col2.selectbox(
        "Description column",
        cols,
        index=(
            next((i for i, c in enumerate(cols) if c.lower() == "description"),
                next((i for i, c in enumerate(cols) if "description" in c.lower() or "summary" in c.lower()),
                    min(1, len(cols)-1)))
        )
    )

    df = df_raw[[num_col, desc_col]].dropna(subset=[desc_col]).copy()
    df.columns = ["number", "description_raw"]
    df["description"] = df["description_raw"].apply(preprocess)

    st.session_state.df = df

    st.success(f"{len(df)} incidents loaded")
    st.dataframe(df[["number", "description_raw"]].head(5), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Generate Prompt
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.df is not None:
    st.divider()
    st.header("Step 2 — Generate AI Prompt")

    if st.button("Generate Prompt", type="primary"):
        rows = st.session_state.df[["number", "description"]].to_dict("records")
        st.session_state.prompt = build_prompt(rows)

    if st.session_state.prompt:
        st.info("Copy the entire prompt below and paste it into Claude, ChatGPT, or your AI tool of choice.")
        preview_lines = "\n".join(st.session_state.prompt.splitlines()[:10])
        st.code(preview_lines + "\n...", language=None)
        with st.expander("Show full prompt"):
            st.code(st.session_state.prompt, language=None)
        copy_button(st.session_state.prompt, tooltip="Copy full prompt to clipboard", copied_label="Copied!", icon="st")
        st.markdown(
            """
            <a href="https://goto/red" target="_blank">
                <button style="background-color:red;color:white;padding:0.5em 1.5em;border:none;border-radius:4px;font-size:1em;cursor:pointer;">
                    Go to Red Portal
                </button>
            </a>
            """,
            unsafe_allow_html=True
        )

# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Paste AI Response
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.prompt:
    st.divider()
    st.header("Step 3 — Paste AI Response")

    raw_response = st.text_area(
        "Paste the AI response here",
        height=250,
        placeholder='{"groups": [{"application": "...", "issue": "...", ...}]}'
    )

    if st.button("Process Response", type="primary"):
        if not raw_response.strip():
            st.warning("Paste the AI response before processing.")
        else:
            groups = parse_response(raw_response)
            if groups is None:
                st.error("Could not parse the response as JSON. Make sure you copied the full response and try again.")
            elif len(groups) == 0:
                st.warning("No groups were found. The AI may not have identified recurring patterns — try re-running the prompt.")
            else:
                st.session_state.groups = groups
                st.success(f"Parsed {len(groups)} incident groups successfully.")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Results & Export
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.groups:
    st.divider()
    st.header("Step 4 — Results")

    col_month, col_year = st.columns([2, 1])
    selected_month = col_month.selectbox("Report Month", months, index=now.month - 1)
    selected_year = col_year.selectbox("Report Year", years, index=0)

    groups = st.session_state.groups
    total_covered = sum(g.get("count", 0) for g in groups)
    applications = sorted({g.get("application", "Unknown") for g in groups})

    m0, m1, m2, m3 = st.columns(4)
    m0.metric("Total Incidents Analysed", len(st.session_state.df))
    m1.metric("Incident Groups", len(groups))
    m2.metric("Incidents Covered", total_covered)
    m3.metric("Applications Affected", len(applications))

    st.subheader("Groups by Application")

    sorted_groups = sorted(groups, key=lambda g: g.get("count", 0), reverse=True)

    for app in sorted({g.get("application", "Unknown") for g in sorted_groups},
                      key=lambda a: sum(g["count"] for g in sorted_groups if g.get("application") == a),
                      reverse=True):
        app_groups = [g for g in sorted_groups if g.get("application") == app]
        app_total = sum(g.get("count", 0) for g in app_groups)

        with st.expander(f"{app}  —  {app_total} incident{'s' if app_total != 1 else ''}"):
            for g in app_groups:
                st.markdown(f"**Issue:** {g.get('issue', '')}")
                cols = st.columns([1, 3, 3])
                cols[0].metric("Count", g.get("count", 0))
                cols[1].markdown(f"**Business Impact**\n\n{g.get('business_impact', '')}")
                cols[2].markdown(f"**Recommended Action**\n\n{g.get('recommended_action', '')}")

                inc_nums = g.get("incident_numbers", [])
                if inc_nums:
                    st.caption("Incidents: " + " · ".join(inc_nums))
                st.divider()

    # ── Excel export ──
    st.subheader("Export")
    excel_bytes = build_excel(groups)
    excel_filename = f"universal_request_incident_patterns_{selected_year}-{selected_month}.xlsx"

    st.download_button(
        label="Download Excel Report",
        data=excel_bytes,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
