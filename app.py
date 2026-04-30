import re
import json
import io
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from st_copy import copy_button

st.set_page_config(page_title="Incident Pattern Analyser", layout="wide")

st.title("Incident Pattern Analyser")
st.caption("Surface recurring issues for senior management reporting")

# ── constants ──────────────────────────────────────────────────────────────────
DATA_DIR = Path("data")
now = datetime.now()
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
years = [str(y) for y in range(now.year, now.year - 5, -1)]

# ── session state ──────────────────────────────────────────────────────────────
for key in ("df", "batches", "batch_size", "all_groups", "missing_ids",
            "processing_mode", "cross_batch_done", "management_email_text"):
    if key not in st.session_state:
        st.session_state[key] = None


# ── helpers ────────────────────────────────────────────────────────────────────
def preprocess(text: str, maxlen: int = 200) -> str:
    text = str(text)
    text = re.sub(r'\b[A-Z]{2,4}\d{5,}\b', '', text)
    text = re.sub(r'\S+@\S+', '', text)
    text = re.sub(r'\b\d{5,}\b', '', text)
    text = re.sub(r'\[.*?\]', '', text)
    text = re.sub(r'Client Identifying Data.*?not allowed.*?\.?', '', text, flags=re.IGNORECASE)
    text = re.sub(r'No CID Disclaimer accepted: TRUE', '', text, flags=re.IGNORECASE)
    text = re.sub(r'I have read and understood the disclaimer.*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^(hi|hello|dear|greetings)[\s,]+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(thanks|thank you|regards|best regards|cheers)[\s,]*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'Tel\.?[\s:\-]*\d+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'User details:.*?(?=Topic:|$)', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\(T\d+\)', '', text)
    text = re.sub(r'\+?\d[\d\s\-]{7,}', '', text)
    text = re.sub(r'[-_]{2,}', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text[:maxlen]


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
- Every incident must appear in the output, either as part of a group or as its own group. Do not omit any incident for any reason.
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


def normalise_quotes(text: str) -> str:
    return (text
        .replace('“', '"').replace('”', '"')
        .replace('‘', "'").replace('’', "'")
        .replace('′', "'").replace('″', '"')
        .replace('﻿', '')
    )


def parse_response(raw: str) -> list[dict] | None:
    cleaned = re.sub(r"```json|```", "", raw).strip()
    cleaned = normalise_quotes(cleaned)
    json_match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if json_match:
        cleaned = json_match.group(0)
    try:
        data = json.loads(cleaned)
        raw_groups = data.get("groups") or data.get("g", [])
    except json.JSONDecodeError:
        return None

    groups = []
    seen: set[str] = set()
    for g in raw_groups:
        ids_raw = g.get("incident_numbers") or g.get("ids", [])
        unique_ids = []
        for n in ids_raw:
            key = str(n).strip()
            if key and key not in seen:
                seen.add(key)
                unique_ids.append(key)
        if not unique_ids:
            continue
        groups.append({
            "application":        g.get("application") or g.get("app", "Unknown System"),
            "issue":              g.get("issue") or g.get("iss", ""),
            "incident_numbers":   unique_ids,
            "count":              len(unique_ids),
            "business_impact":    g.get("business_impact") or g.get("imp", ""),
            "recommended_action": g.get("recommended_action") or g.get("act", ""),
        })
    return groups


def build_cross_batch_prompt(groups: list[dict]) -> str:
    groups_json = json.dumps({"groups": groups}, indent=2)
    return f"""You are a senior IT incident analyst reviewing grouped incident data from an Investment Bank.
The groups below were produced by analysing batches of incidents separately. Because they were processed
in separate batches, similar issues may have been described slightly differently and placed into distinct
groups when they should be one consolidated group.

Your task:
1. Review ALL groups below carefully.
2. Identify groups that represent the same underlying root cause or operational issue, even if the
   wording differs (e.g. "AD password reset locking users out" and "Users cannot log in after password
   change" are the same issue).
3. Consolidate such groups: merge their incident_numbers, sum their counts, and write a single precise
   issue description. Keep the most informative business_impact and recommended_action.
4. Groups that are genuinely distinct should remain separate.
5. Every incident_number from the input must appear exactly once in the output. Do not omit any.
6. Return ONLY valid JSON — no markdown fences, no explanation, nothing else.

Output format (same as input):
{{
  "groups": [
    {{
      "application": "name of the system or application",
      "issue": "precise description of the problem or request",
      "count": <number of incident_numbers in this group>,
      "incident_numbers": ["INC001", "INC002"],
      "business_impact": "one sentence describing the business effect",
      "recommended_action": "one concrete action to prevent recurrence or address the issue"
    }}
  ]
}}

Current merged groups ({len(groups)} groups, produced from batch processing):
{groups_json}
"""


def build_management_email_prompt(groups: list[dict], top_n: int) -> str:
    sorted_groups = sorted(groups, key=lambda g: g.get("count", 0), reverse=True)
    top_groups = sorted_groups[:top_n]
    numbered_list = "\n".join(
        f"{i+1}. Application: {g['application']}\n"
        f"   Issue: {g['issue']}\n"
        f"   Incident Count: {g['count']}\n"
        f"   Business Impact: {g.get('business_impact', '')}\n"
        f"   Recommended Action: {g.get('recommended_action', '')}"
        for i, g in enumerate(top_groups)
    )
    return f"""You are a senior IT communications specialist at an Investment Bank.
Write a professional executive email to senior management summarising the top {top_n} recurring
IT incidents from this month's analysis. The email should be clear, concise, and appropriate for
C-suite and VP-level readers who need to understand business impact and recommended actions.

Requirements:
- Start with a brief executive summary paragraph (2–3 sentences max).
- Include a formatted table with columns: Rank | Application | Issue | Count | Business Impact | Recommended Action
- Close with a short paragraph on overall risk and a recommended next step.
- Use plain text with markdown-style formatting (| for table columns, ** for bold headings).
- Do NOT use JSON. Return the email body as plain text only.
- Tone: professional, direct, no jargon, suitable for senior management.

Top {top_n} incident groups (sorted by count, highest first):
{numbered_list}
"""


def merge_all_batches(batches: list[dict]) -> list[dict]:
    seen: set[str] = set()
    merged: list[dict] = []
    for batch in batches:
        for g in batch.get("groups", []):
            unique_ids = [n for n in g["incident_numbers"] if n not in seen]
            seen.update(unique_ids)
            if unique_ids:
                existing = next(
                    (m for m in merged
                     if m["application"].lower() == g["application"].lower()
                     and m["issue"].lower() == g["issue"].lower()),
                    None
                )
                if existing:
                    existing["incident_numbers"].extend(unique_ids)
                    existing["count"] = len(existing["incident_numbers"])
                else:
                    merged.append({**g, "incident_numbers": unique_ids, "count": len(unique_ids)})
    return merged


def _autofit_columns(ws, min_width: int = 10, max_width: int = 60) -> None:
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        best = min_width
        for cell in col_cells:
            if cell.value is not None:
                longest_line = max(len(str(line)) for line in str(cell.value).splitlines() or [""])
                best = max(best, longest_line)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(best + 2, max_width)


def build_excel(groups: list[dict], unaccounted_df: pd.DataFrame | None = None) -> bytes:
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F3864")
    alt_fill    = PatternFill("solid", fgColor="DCE6F1")

    ws1 = wb.active
    ws1.title = "Management Summary"
    headers = ["Application", "Issue", "Count", "Business Impact", "Recommended Action"]
    ws1.append(headers)
    for col_idx, _ in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws1.freeze_panes = "A2"
    sorted_groups = sorted(groups, key=lambda g: g.get("count", 0), reverse=True)
    for row_idx, g in enumerate(sorted_groups, 2):
        ws1.append([g.get("application",""), g.get("issue",""), g.get("count",0),
                    g.get("business_impact",""), g.get("recommended_action","")])
        if row_idx % 2 == 0:
            for col_idx in range(1, len(headers)+1):
                ws1.cell(row=row_idx, column=col_idx).fill = alt_fill
    _autofit_columns(ws1)

    ws2 = wb.create_sheet("Incident Detail")
    detail_headers = ["Group", "Application", "Issue", "Incident Numbers"]
    ws2.append(detail_headers)
    for col_idx, _ in enumerate(detail_headers, 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    ws2.freeze_panes = "A2"
    for group_idx, g in enumerate(sorted_groups, 1):
        ws2.append([group_idx, g.get("application",""), g.get("issue",""),
                    ", ".join(g.get("incident_numbers",[]))])
    _autofit_columns(ws2)

    if unaccounted_df is not None and not unaccounted_df.empty:
        ws3 = wb.create_sheet("Unaccounted")
        ws3.append(["Incident Number", "Description"])
        for col_idx in range(1, 3):
            cell = ws3.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = PatternFill("solid", fgColor="C00000")
            cell.alignment = Alignment(horizontal="center")
        ws3.freeze_panes = "A2"
        for _, row in unaccounted_df.iterrows():
            ws3.append([str(row["number"]), str(row["description_raw"])])
        _autofit_columns(ws3)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save_monthly_data(df: pd.DataFrame, groups: list[dict], year: str, month: str) -> tuple[str, str]:
    DATA_DIR.mkdir(exist_ok=True)
    raw_path  = DATA_DIR / f"raw_data_{year}_{month}.csv"
    proc_path = DATA_DIR / f"processed_data_{year}_{month}.json"
    df[["number", "description_raw"]].to_csv(raw_path, index=False)
    proc_path.write_text(json.dumps({"year": year, "month": month, "groups": groups}, indent=2))
    return str(raw_path), str(proc_path)


def load_history() -> list[dict]:
    if not DATA_DIR.exists():
        return []
    records = []
    for f in sorted(DATA_DIR.glob("processed_data_*.json")):
        try:
            records.append(json.loads(f.read_text()))
        except (json.JSONDecodeError, OSError):
            pass
    return records


def red_portal_button() -> None:
    st.markdown(
        """<a href="https://goto/red" target="_blank">
            <button style="background-color:red;color:white;padding:0.5em 1.5em;
                border:none;border-radius:4px;font-size:1em;cursor:pointer;">
                Go to Red Portal
            </button>
        </a>""",
        unsafe_allow_html=True,
    )


def prompt_panel(prompt: str, copy_label: str, key: str, height: int = 200,
                 placeholder: str = '{"groups": [...]}') -> str:
    st.code("\n".join(prompt.splitlines()[:10]) + "\n...", language=None)
    with st.expander("Show full prompt"):
        st.code(prompt, language=None)
    copy_button(prompt, tooltip=copy_label, copied_label="Copied!", icon="st")
    red_portal_button()
    return st.text_area("Paste AI response here", height=height, key=key, placeholder=placeholder)


def show_parse_error(raw: str) -> None:
    st.error("Could not parse as JSON.")
    with st.expander("Debug — show received text"):
        st.text(f"First 300 chars:\n{raw[:300]}")
        st.text(f"Last 100 chars:\n{raw[-100:]}")


def update_coverage(groups: list[dict]) -> None:
    all_ids = set(st.session_state.df["number"].astype(str))
    covered = {n for g in groups for n in g["incident_numbers"]}
    st.session_state.missing_ids = sorted(all_ids - covered)


def highlight_recurring(row) -> list[str]:
    color = "background-color: #fff3cd" if row["Months Active"] >= 3 else ""
    return [color] * len(row)


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
        index=next((i for i, c in enumerate(cols)
                    if "number" in c.lower() or "id" in c.lower() or "inc" in c.lower()), 0)
    )
    desc_col = col2.selectbox(
        "Description column", cols,
        index=(next((i for i, c in enumerate(cols) if c.lower() == "description"),
               next((i for i, c in enumerate(cols)
                     if "description" in c.lower() or "summary" in c.lower()), min(1, len(cols)-1))))
    )

    df = df_raw[[num_col, desc_col]].dropna(subset=[desc_col]).copy()
    df.columns = ["number", "description_raw"]
    df["description"] = df["description_raw"].apply(preprocess)
    st.session_state.df = df

    st.success(f"{len(df)} incidents loaded")
    st.dataframe(df[["number", "description_raw"]].head(5), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Choose Processing Mode & Generate Prompts
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.df is not None:
    st.divider()
    st.header("Step 2 — Choose Processing Mode")

    mode = st.radio(
        "How do you want to analyse the incidents?",
        options=["batch", "all"],
        format_func=lambda x: (
            "Batch processing — split into smaller chunks (recommended for large datasets)"
            if x == "batch"
            else "Process all incidents together — single prompt (best for smaller datasets)"
        ),
        index=0 if st.session_state.processing_mode != "all" else 1,
        horizontal=True,
        key="mode_radio",
    )
    st.session_state.processing_mode = mode

    total = len(st.session_state.df)

    # ── Option A: All together ─────────────────────────────────────────────────
    if mode == "all":
        st.subheader("Step 2A — Process All Incidents Together")
        st.info(f"This generates a single prompt for all {total} incidents. Best for datasets under ~100 incidents.")

        rows = st.session_state.df[["number", "description"]].to_dict("records")
        full_prompt = build_prompt(rows)

        raw_all = prompt_panel(full_prompt, "Copy full prompt", "response_all")

        if st.button("Process All Incidents", type="primary", key="process_all"):
            if not raw_all.strip():
                st.warning("Paste the AI response before processing.")
            else:
                groups = parse_response(raw_all)
                if groups is None:
                    show_parse_error(raw_all)
                elif len(groups) == 0:
                    st.warning("No groups found — try re-running the prompt.")
                else:
                    st.session_state.batches = [{
                        "index": 0, "rows": rows, "prompt": full_prompt,
                        "groups": groups, "complete": True,
                    }]
                    st.session_state.batch_size = len(rows)
                    st.session_state.all_groups = groups
                    update_coverage(groups)
                    st.session_state.cross_batch_done = True
                    st.rerun()

        if st.session_state.all_groups and st.session_state.processing_mode == "all":
            st.success(f"Processed — {len(st.session_state.all_groups)} groups found. Continue to Step 3 below.")

    # ── Option B: Batch processing ─────────────────────────────────────────────
    else:
        st.subheader("Step 2B — Configure & Analyse Batches")

        batch_size = st.slider(
            "Incidents per batch", min_value=25, max_value=150, value=75, step=25,
            help="Smaller batches = more paste-backs but better AI coverage per batch. 50–75 is the sweet spot."
        )

        n_batches = -(-total // batch_size)
        st.caption(f"{total} incidents → **{n_batches} batch{'es' if n_batches != 1 else ''}** of up to {batch_size}")

        if st.button("Generate Batch Prompts", type="primary"):
            rows = st.session_state.df[["number", "description"]].to_dict("records")
            st.session_state.batches = [
                {"index": i, "rows": chunk, "prompt": build_prompt(chunk), "groups": None, "complete": False}
                for i, chunk in (
                    (i, rows[i * batch_size: (i + 1) * batch_size]) for i in range(n_batches)
                )
            ]
            st.session_state.batch_size    = batch_size
            st.session_state.all_groups    = None
            st.session_state.missing_ids   = None
            st.session_state.cross_batch_done = None

        if st.session_state.batches:
            batches = st.session_state.batches
            completed = sum(1 for b in batches if b["complete"])

            st.progress(completed / len(batches), text=f"{completed}/{len(batches)} batches complete")

            tabs = st.tabs([
                f"{'✅' if b['complete'] else '⏳'} Batch {b['index']+1} ({len(b['rows'])} incidents)"
                for b in batches
            ])

            for tab, batch in zip(tabs, batches):
                with tab:
                    start = batch['index'] * st.session_state.batch_size + 1
                    end   = start + len(batch['rows']) - 1
                    st.caption(f"Incidents {start}–{end} of {total}")

                    if batch["complete"]:
                        st.success(f"Complete — {sum(g['count'] for g in batch['groups'])} incidents grouped into {len(batch['groups'])} groups")
                    else:
                        st.info("Copy the prompt below, paste into your AI tool, then paste the response back here.")

                    raw_response = prompt_panel(
                        batch["prompt"],
                        f"Copy Batch {batch['index']+1} prompt",
                        f"response_batch_{batch['index']}",
                    )

                    if st.button(f"Process Batch {batch['index']+1}", type="primary",
                                 key=f"process_{batch['index']}"):
                        if not raw_response.strip():
                            st.warning("Paste the AI response before processing.")
                        else:
                            groups = parse_response(raw_response)
                            if groups is None:
                                show_parse_error(raw_response)
                            elif len(groups) == 0:
                                st.warning("No groups found — try re-running the prompt.")
                            else:
                                st.session_state.batches[batch["index"]]["groups"]   = groups
                                st.session_state.batches[batch["index"]]["complete"] = True
                                st.session_state.all_groups = merge_all_batches(st.session_state.batches)
                                update_coverage(st.session_state.all_groups)
                                st.session_state.cross_batch_done = False
                                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Cross-Batch Pattern Analyser
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.all_groups is not None and st.session_state.processing_mode == "batch":
    st.divider()
    st.header("Step 3 — Cross-Batch Pattern Analysis")

    if st.session_state.cross_batch_done:
        st.success(f"Cross-batch analysis complete — {len(st.session_state.all_groups)} final groups.")
        if st.button("Re-run cross-batch analysis", type="secondary", key="rerun_crossbatch"):
            st.session_state.cross_batch_done = False
            st.rerun()
    else:
        n_batches_done = sum(1 for b in (st.session_state.batches or []) if b["complete"])
        st.info(
            f"You have {n_batches_done} batch(es) of results merged into "
            f"**{len(st.session_state.all_groups)} groups**. This step asks Claude to consolidate "
            f"any groups that represent the same issue but were described differently across batches."
        )

        raw_cross = prompt_panel(
            build_cross_batch_prompt(st.session_state.all_groups),
            "Copy cross-batch prompt",
            "response_cross_batch",
            height=250,
        )

        if st.button("Process Cross-Batch Response", type="primary", key="process_cross_batch"):
            if not raw_cross.strip():
                st.warning("Paste the AI response before processing.")
            else:
                refined_groups = parse_response(raw_cross)
                if refined_groups is None:
                    show_parse_error(raw_cross)
                elif len(refined_groups) == 0:
                    st.warning("No groups returned — try re-running the prompt.")
                else:
                    st.session_state.all_groups = refined_groups
                    update_coverage(refined_groups)
                    st.session_state.cross_batch_done = True
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 / 3 — Results & Export
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.all_groups:
    st.divider()
    _s = 0 if st.session_state.processing_mode == "all" else 1
    st.header(f"Step {3 + _s} — Results")

    groups       = st.session_state.all_groups
    missing_ids  = st.session_state.missing_ids or []
    batches_done = sum(1 for b in (st.session_state.batches or []) if b["complete"])
    batches_total= len(st.session_state.batches or [])

    col_month, col_year = st.columns([2, 1])
    selected_month = col_month.selectbox("Report Month", months, index=now.month - 1)
    selected_year  = col_year.selectbox("Report Year", years, index=0)

    total_covered = sum(g.get("count", 0) for g in groups)
    applications  = sorted({g.get("application", "Unknown") for g in groups})

    m0, m1, m2, m3, m4 = st.columns(5)
    m0.metric("Total Incidents",    len(st.session_state.df))
    m1.metric("Batches Complete",   f"{batches_done}/{batches_total}")
    m2.metric("Incidents Covered",  total_covered)
    m3.metric("Groups Found",       len(groups))
    m4.metric("Applications",       len(applications))

    if missing_ids:
        st.warning(f"⚠ {len(missing_ids)} incidents not yet covered — complete remaining batches or check AI responses.")
        with st.expander(f"Show {len(missing_ids)} uncovered incidents"):
            missing_df = st.session_state.df[st.session_state.df["number"].astype(str).isin(missing_ids)]
            st.dataframe(missing_df[["number", "description_raw"]], use_container_width=True, hide_index=True)
    else:
        st.success(f"All {len(st.session_state.df)} incidents covered across {batches_done} batches.")

    st.subheader("Groups by Application")
    sorted_groups = sorted(groups, key=lambda g: g.get("count", 0), reverse=True)

    for app in sorted({g.get("application", "Unknown") for g in sorted_groups},
                      key=lambda a: sum(g["count"] for g in sorted_groups if g.get("application") == a),
                      reverse=True):
        app_groups = [g for g in sorted_groups if g.get("application") == app]
        app_total  = sum(g.get("count", 0) for g in app_groups)
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

    # ── Save ──
    st.subheader("Save Monthly Data")
    if st.button(f"Save {selected_month} {selected_year} Data", type="secondary"):
        raw_path, proc_path = save_monthly_data(st.session_state.df, groups, selected_year, selected_month)
        st.success(f"Saved:\n- `{raw_path}`\n- `{proc_path}`")

    # ── Excel Export ──
    st.subheader("Export")
    unaccounted_df = (
        st.session_state.df[st.session_state.df["number"].astype(str).isin(missing_ids)]
        if missing_ids else None
    )
    excel_bytes    = build_excel(groups, unaccounted_df)
    excel_filename = f"universal_request_incident_patterns_{selected_year}-{selected_month}.xlsx"
    st.download_button(
        label="Download Excel Report",
        data=excel_bytes,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    # ── History & Trends ──
    st.divider()
    st.subheader("Monthly Trends")
    history = load_history()

    if len(history) < 2:
        st.info("Save data for 2 or more months to see trend analysis.")
    else:
        month_idx = {m: i for i, m in enumerate(months)}
        sorted_history = sorted(history, key=lambda r: (int(r["year"]), month_idx.get(r["month"], 0)))

        rows_hist = [
            {"month_label": f"{r['month']} {r['year']}", "year": r["year"],
             "month": r["month"], "application": g.get("application", "Unknown"), "count": g.get("count", 0)}
            for r in sorted_history for g in r.get("groups", [])
        ]

        hist_df    = pd.DataFrame(rows_hist)
        month_order = [f"{r['month']} {r['year']}" for r in sorted_history]

        app_month_counts = (
            hist_df.groupby(["application", "month_label"])["count"]
            .sum().unstack(fill_value=0)
            .reindex(columns=[m for m in month_order if m in hist_df["month_label"].values])
        )
        app_month_counts["Total"]         = app_month_counts.sum(axis=1)
        app_month_counts["Months Active"] = (app_month_counts.drop(columns="Total") > 0).sum(axis=1)
        app_month_counts = app_month_counts.sort_values("Total", ascending=False)

        st.markdown("**Applications by Month (incident count)**")
        st.dataframe(app_month_counts.style.apply(highlight_recurring, axis=1), use_container_width=True)
        st.caption("Amber rows = application appeared in 3 or more months")

        st.markdown("**Monthly Incident Volume**")
        target_pct = st.slider(
            "Target reduction per month (%)", min_value=1, max_value=30, value=10, step=1
        )

        df_vol = pd.DataFrame([
            {"month": f"{r['month'][:3]} {r['year']}",
             "total": sum(g.get("count", 0) for g in r.get("groups", []))}
            for r in sorted_history
        ])
        df_vol["order"]      = range(len(df_vol))
        df_vol["moving_avg"] = df_vol["total"].rolling(window=3, min_periods=1).mean().round(1)
        baseline             = df_vol["total"].iloc[0]
        df_vol["target"]     = [round(baseline * ((1 - target_pct / 100) ** i), 1) for i in range(len(df_vol))]
        df_vol["status"]     = df_vol.apply(
            lambda r: "Above average" if r["total"] > r["moving_avg"] else "At or below average", axis=1
        )

        base = alt.Chart(df_vol).encode(
            x=alt.X("month:N", sort=None, title="Month", axis=alt.Axis(labelAngle=-30))
        )
        bars = base.mark_bar(opacity=0.85).encode(
            y=alt.Y("total:Q", title="Total Incidents"),
            color=alt.Color("status:N", scale=alt.Scale(
                domain=["Above average", "At or below average"],
                range=["#E07B39", "#1F3864"]
            ), legend=alt.Legend(title="vs Moving Average")),
            tooltip=[
                alt.Tooltip("month:N",      title="Month"),
                alt.Tooltip("total:Q",      title="Total Incidents"),
                alt.Tooltip("moving_avg:Q", title="3-Month Avg"),
                alt.Tooltip("target:Q",     title=f"Target (-{target_pct}%/mo)"),
            ]
        )
        ma_line = base.mark_line(color="#F4A900", strokeWidth=2.5, point=True).encode(y="moving_avg:Q")
        target_line = base.mark_line(color="#C00000", strokeWidth=2, strokeDash=[6, 3]).encode(y="target:Q")
        st.altair_chart(
            (bars + ma_line + target_line).properties(height=350).configure_axis(labelFontSize=12, titleFontSize=13),
            use_container_width=True
        )
        st.caption("Orange line = 3-month moving average  |  Red dashed = target reduction trajectory")

        top10 = app_month_counts.head(10)["Total"].sort_values(ascending=True)
        st.markdown("**Top 10 Applications — Total Incidents Across All Months**")
        st.bar_chart(top10)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5 / 4 — Management Email Generator
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.all_groups and st.session_state.cross_batch_done:
    st.divider()
    _s = 0 if st.session_state.processing_mode == "all" else 1
    st.header(f"Step {4 + _s} — Management Email")
    st.caption("Generate a professional executive email summarising the top recurring incidents for senior management.")

    groups = st.session_state.all_groups

    top_n = st.number_input(
        "Number of top incident groups to include in the email",
        min_value=1, max_value=len(groups), value=min(5, len(groups)), step=1,
        help="Groups are sorted by incident count. The top N by count will be included in the email table.",
        key="email_top_n",
    )

    raw_email = prompt_panel(
        build_management_email_prompt(groups, int(top_n)),
        "Copy email prompt",
        "response_email",
        height=300,
        placeholder="Paste the plain-text email from Claude here...",
    )

    if st.button("Process Email Response", type="primary", key="process_email"):
        if not raw_email.strip():
            st.warning("Paste the AI response before processing.")
        else:
            st.session_state.management_email_text = raw_email.strip()
            st.rerun()

    if st.session_state.management_email_text:
        st.subheader("Generated Management Email")
        st.markdown(st.session_state.management_email_text)
        copy_button(
            st.session_state.management_email_text,
            tooltip="Copy email to clipboard",
            copied_label="Copied!",
            icon="st",
        )
