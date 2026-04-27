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
for key in ("df", "grouping_prompt", "raw_groups", "enrichment_prompt", "groups", "missing_ids"):
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


def build_grouping_prompt(rows: list[dict]) -> str:
    n = len(rows)
    incidents = "\n".join(f"{i+1}. [{r['number']}] {r['description']}" for i, r in enumerate(rows))
    return f"""⚠ CRITICAL: Every incident number must appear in exactly one group. PREFER MORE GROUPS OVER FEWER.

<instructions>
You are a meticulous IT incident categoriser for an Investment Bank. Your job is precise, granular categorisation — not summarisation.

A group is only valid if ALL three conditions are true:
1. Every incident in the group involves the exact same application or system
2. Every incident describes the exact same specific failure mode, error, or request type
3. Resolving one incident would resolve all others in the group

GRANULARITY RULES:
- Same application but different issue = DIFFERENT groups
- Same issue type but different application = DIFFERENT groups
- "Bloomberg access issues" is NOT a valid group — it is too vague
- "Bloomberg Terminal — users locked out after AD password reset" is a valid group
- "Bloomberg Terminal — market data feed not refreshing" is a SEPARATE group from the above
- "Charles River — incorrect pricing data" and "Charles River — calendar sync failure" are TWO groups

Do NOT merge incidents just because they share the same application.
Do NOT merge incidents just because they are both "access" or "performance" issues.
When in doubt — create SEPARATE groups rather than merging.
Groups of 1 are correct and expected for unique issues.

⚠ If you create a group with more than 12 incidents, you have almost certainly been too broad. Split it into more specific sub-groups before returning.

Output format — plain text only, no JSON, no descriptions:
Group 1: UR001, UR004, UR012
Group 2: UR002
Group 3: UR003, UR007, UR099
</instructions>

⚠ CRITICAL: Before returning, verify two things:
1. Every incident number appears exactly once across all groups. You received {n}. Count must equal {n}.
2. No group has more than 12 incidents. If one does, split it.

<incidents>
{incidents}
</incidents>"""


def parse_grouping_response(raw: str) -> dict[str, list[str]] | None:
    groups: dict[str, list[str]] = {}
    seen: set[str] = set()
    for line in raw.strip().splitlines():
        m = re.match(r'[Gg]roup\s*\d+\s*[:\-]\s*(.+)', line.strip())
        if m:
            ids = []
            for x in m.group(1).split(','):
                key = x.strip()
                if key and key not in seen:
                    seen.add(key)
                    ids.append(key)
            if ids:
                groups[f"Group {len(groups)+1}"] = ids
    return groups if groups else None


def build_enrichment_prompt(raw_groups: dict[str, list[str]], df: pd.DataFrame) -> str:
    id_to_desc = dict(zip(df["number"].astype(str), df["description"]))
    summaries = []
    for label, ids in raw_groups.items():
        descs = [id_to_desc.get(i, "") for i in ids if id_to_desc.get(i)]
        # Pass up to 3 sample descriptions so the AI has enough context to label precisely
        samples = sorted(descs, key=len, reverse=True)[:3]
        sample_text = " | ".join(s[:180] for s in samples) if samples else "No description available"
        summaries.append(f"{label} ({len(ids)} incident{'s' if len(ids) != 1 else ''}): {sample_text}")
    groups_text = "\n".join(summaries)
    g = len(raw_groups)
    return f"""You are enriching {g} pre-formed incident groups for senior management reporting at an Investment Bank.

Each group has already been categorised — your job is to label it precisely and assess its business impact.

Example — use exactly this JSON structure for each group:
{{"app":"Bloomberg Terminal","iss":"Users locked out after overnight AD sync failure — authentication tokens not refreshed","imp":"Traders unable to access live pricing at market open, requiring manual workaround","act":"Schedule AD sync outside trading hours and add token refresh validation"}}

Now enrich all {g} groups. Return a single JSON object:
{{"g":[{{"app":"...","iss":"...","imp":"...","act":"..."}}]}}

Rules:
- "app": specific application name — include the module if known (e.g. "Charles River IMS" not just "Charles River")
- "iss": precise issue — include the specific failure mode, not a category (e.g. "calendar sync not reflecting market holidays" not "calendar issue")
- "imp": one sentence on the concrete business effect (who is affected and how)
- "act": one specific, actionable recommendation — not generic advice
- Return ONLY valid JSON — no markdown fences, no explanation

<groups>
{groups_text}
</groups>"""


def normalise_quotes(text: str) -> str:
    return (text
        .replace('“', '"').replace('”', '"')   # curly double quotes
        .replace('‘', "'").replace('’', "'")   # curly single quotes
        .replace('′', "'").replace('″', '"')   # prime characters
        .replace('﻿', '')                            # BOM
    )


def parse_enrichment_response(raw: str, raw_groups: dict[str, list[str]]) -> list[dict] | None:
    cleaned = re.sub(r"```json|```", "", raw).strip()
    cleaned = normalise_quotes(cleaned)
    # Extract the outermost JSON object even if the AI added surrounding text
    json_match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if json_match:
        cleaned = json_match.group(0)
    try:
        data = json.loads(cleaned)
        enriched = data.get("g") or data.get("groups", [])
    except json.JSONDecodeError:
        return None

    group_labels = list(raw_groups.keys())
    result = []
    for idx, e in enumerate(enriched):
        label = group_labels[idx] if idx < len(group_labels) else f"Group {idx+1}"
        ids = raw_groups.get(label, [])
        result.append({
            "application":        e.get("app") or e.get("application", "Unknown System"),
            "issue":              e.get("iss") or e.get("issue", ""),
            "incident_numbers":   ids,
            "count":              len(ids),
            "business_impact":    e.get("imp") or e.get("business_impact", ""),
            "recommended_action": e.get("act") or e.get("recommended_action", ""),
        })
    return result


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


def build_excel(groups: list[dict], unaccounted_df: pd.DataFrame | None = None) -> bytes:
    wb = Workbook()
    header_font  = Font(bold=True, color="FFFFFF")
    header_fill  = PatternFill("solid", fgColor="1F3864")
    alt_fill     = PatternFill("solid", fgColor="DCE6F1")

    # ── Sheet 1: Management Summary ──
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
        for col_idx in range(1, len(headers)+1):
            ws1.cell(row=row_idx, column=col_idx).alignment = Alignment(wrap_text=True)
    for i, width in enumerate([28, 48, 10, 48, 48], 1):
        ws1.column_dimensions[get_column_letter(i)].width = width

    # ── Sheet 2: Incident Detail ──
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
        ws2.cell(row=group_idx+1, column=4).alignment = Alignment(wrap_text=True)
    for col_idx, width in enumerate([8, 28, 48, 60], 1):
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Sheet 3: Unaccounted (only if gaps remain) ──
    if unaccounted_df is not None and not unaccounted_df.empty:
        ws3 = wb.create_sheet("Unaccounted")
        ua_headers = ["Incident Number", "Description"]
        ws3.append(ua_headers)
        for col_idx, _ in enumerate(ua_headers, 1):
            cell = ws3.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = PatternFill("solid", fgColor="C00000")
            cell.alignment = Alignment(horizontal="center")
        ws3.freeze_panes = "A2"
        for _, row in unaccounted_df.iterrows():
            ws3.append([str(row["number"]), str(row["description_raw"])])
        ws3.column_dimensions["A"].width = 20
        ws3.column_dimensions["B"].width = 80

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
# STEP 2 — Pass 1: Grouping
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.df is not None:
    st.divider()
    st.header("Step 2 — Group Incidents (Pass 1 of 2)")
    st.caption("This prompt asks the AI to assign every incident to a group. Output is plain text — always fits within AI response limits.")

    if st.button("Generate Grouping Prompt", type="primary"):
        rows = st.session_state.df[["number", "description"]].to_dict("records")
        st.session_state.grouping_prompt = build_grouping_prompt(rows)
        # Reset downstream state when regenerating
        st.session_state.raw_groups = None
        st.session_state.enrichment_prompt = None
        st.session_state.groups = None

    if st.session_state.grouping_prompt:
        st.info("Copy the prompt below and paste it into Claude, ChatGPT, or your AI tool. The AI will return a plain list of groups.")
        preview = "\n".join(st.session_state.grouping_prompt.splitlines()[:12])
        st.code(preview + "\n...", language=None)
        with st.expander("Show full grouping prompt"):
            st.code(st.session_state.grouping_prompt, language=None)
        copy_button(st.session_state.grouping_prompt, tooltip="Copy grouping prompt", copied_label="Copied!", icon="st")
        st.markdown(
            """<a href="https://goto/red" target="_blank">
                <button style="background-color:red;color:white;padding:0.5em 1.5em;border:none;border-radius:4px;font-size:1em;cursor:pointer;">
                    Go to Red Portal
                </button>
            </a>""",
            unsafe_allow_html=True
        )

        st.subheader("Paste Grouping Response")
        grouping_response = st.text_area(
            "Paste the AI response here (plain text groups)",
            height=200,
            placeholder="Group 1: UR001, UR004, UR012\nGroup 2: UR002\nGroup 3: UR003, UR007",
            key="grouping_textarea"
        )

        if st.button("Process Grouping", type="primary", key="process_grouping"):
            if not grouping_response.strip():
                st.warning("Paste the AI grouping response before processing.")
            else:
                raw_groups = parse_grouping_response(grouping_response)
                if not raw_groups:
                    st.error("Could not parse groups from the response. Make sure the AI returned lines like 'Group 1: UR001, UR002'.")
                else:
                    st.session_state.raw_groups = raw_groups

                    # Coverage check
                    all_ids = set(st.session_state.df["number"].astype(str))
                    covered_ids = {n for ids in raw_groups.values() for n in ids}
                    missing = sorted(all_ids - covered_ids)
                    st.session_state.missing_ids = missing

                    total_groups = len(raw_groups)
                    total_covered = len(covered_ids)

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Groups Found", total_groups)
                    c2.metric("Incidents Assigned", total_covered)
                    c3.metric("Missing", len(missing))

                    if not missing:
                        st.success(f"All {len(all_ids)} incidents assigned to groups.")
                    else:
                        st.warning(f"⚠ {len(missing)} incidents not found in the grouping response. They will be listed in the Unaccounted sheet.")
                        missing_df = st.session_state.df[st.session_state.df["number"].astype(str).isin(missing)]
                        with st.expander(f"Show {len(missing)} unassigned incidents"):
                            st.dataframe(missing_df[["number", "description_raw"]], use_container_width=True, hide_index=True)

                    # Auto-generate enrichment prompt
                    st.session_state.enrichment_prompt = build_enrichment_prompt(raw_groups, st.session_state.df)
                    st.success(f"Grouping complete. Enrichment prompt ready in Step 3.")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Pass 2: Enrichment
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.enrichment_prompt:
    st.divider()
    st.header("Step 3 — Enrich Groups (Pass 2 of 2)")
    st.caption("This prompt sends only the group summaries — not all 300 incidents — so the AI response is small and complete.")

    st.info("Copy the prompt below and paste it into your AI tool. The AI will return JSON with application names, issue descriptions, business impact, and recommended actions.")
    preview = "\n".join(st.session_state.enrichment_prompt.splitlines()[:12])
    st.code(preview + "\n...", language=None)
    with st.expander("Show full enrichment prompt"):
        st.code(st.session_state.enrichment_prompt, language=None)
    copy_button(st.session_state.enrichment_prompt, tooltip="Copy enrichment prompt", copied_label="Copied!", icon="st")

    enrichment_response = st.text_area(
        "Paste the AI enrichment response here (JSON)",
        height=250,
        placeholder='{"g":[{"app":"Bloomberg","iss":"...","imp":"...","act":"..."}]}',
        key="enrichment_textarea"
    )

    if st.button("Process Enrichment", type="primary", key="process_enrichment"):
        if not enrichment_response.strip():
            st.warning("Paste the AI enrichment response before processing.")
        else:
            groups = parse_enrichment_response(enrichment_response, st.session_state.raw_groups)
            print(enrichment_response)
            print(groups)
            if groups is None:
                st.error("Could not parse the response as JSON. Make sure you copied the full response.")
                with st.expander("Debug — show what was received"):
                    st.text(f"Length: {len(enrichment_response)} chars")
                    st.text(f"First 300 chars:\n{enrichment_response[:300]}")
                    st.text(f"Last 100 chars:\n{enrichment_response[-100:]}")
            elif len(groups) == 0:
                st.warning("No groups were parsed. Try re-running the enrichment prompt.")
            else:
                st.session_state.groups = groups
                st.success(f"Enrichment complete — {len(groups)} groups ready.")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Results, Save & History
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.groups:
    st.divider()
    st.header("Step 4 — Results")

    col_month, col_year = st.columns([2, 1])
    selected_month = col_month.selectbox("Report Month", months, index=now.month - 1)
    selected_year  = col_year.selectbox("Report Year", years, index=0)

    groups = st.session_state.groups
    total_covered = sum(g.get("count", 0) for g in groups)
    applications  = sorted({g.get("application", "Unknown") for g in groups})

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
    missing_ids = st.session_state.missing_ids or []
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
        # Build flat records: one row per group per month
        rows_hist = []
        for record in history:
            label = f"{record['month']} {record['year']}"
            for g in record.get("groups", []):
                rows_hist.append({
                    "month_label": label,
                    "year":        record["year"],
                    "month":       record["month"],
                    "application": g.get("application", "Unknown"),
                    "issue":       g.get("issue", ""),
                    "count":       g.get("count", 0),
                })

        hist_df = pd.DataFrame(rows_hist)
        month_order = [f"{m} {y}" for y in sorted(hist_df["year"].unique(), reverse=True)
                       for m in months if f"{m} {y}" in hist_df["month_label"].values]

        # Top recurring applications across months
        app_month_counts = (
            hist_df.groupby(["application", "month_label"])["count"]
            .sum()
            .unstack(fill_value=0)
            .reindex(columns=[m for m in month_order if m in hist_df["month_label"].values])
        )
        app_month_counts["Total"] = app_month_counts.sum(axis=1)
        app_month_counts["Months Active"] = (app_month_counts.drop(columns="Total") > 0).sum(axis=1)
        app_month_counts = app_month_counts.sort_values("Total", ascending=False)

        st.markdown("**Applications by Month (incident count)**")

        # Highlight apps recurring 3+ consecutive months
        def highlight_recurring(row):
            color = "background-color: #fff3cd" if row["Months Active"] >= 3 else ""
            return [color] * len(row)

        st.dataframe(
            app_month_counts.style.apply(highlight_recurring, axis=1),
            use_container_width=True
        )
        st.caption("Amber rows = application appeared in 3 or more months")

        # ── Monthly volume chart with moving average + target line ──
        st.markdown("**Monthly Incident Volume**")

        target_pct = st.slider(
            "Target reduction per month (%)", min_value=1, max_value=30, value=10, step=1,
            help="Red dashed line shows the target if incidents reduce by this % each month from the first saved month"
        )

        # Build monthly totals in chronological order
        month_idx = {m: i for i, m in enumerate(months)}
        sorted_history = sorted(history, key=lambda r: (int(r["year"]), month_idx.get(r["month"], 0)))
        monthly_rows = []
        for record in sorted_history:
            label = f"{record['month'][:3]} {record['year']}"
            total = sum(g.get("count", 0) for g in record.get("groups", []))
            monthly_rows.append({"month": label, "total": total})

        df_vol = pd.DataFrame(monthly_rows)
        df_vol["order"] = range(len(df_vol))

        # 3-month rolling moving average
        df_vol["moving_avg"] = df_vol["total"].rolling(window=3, min_periods=1).mean().round(1)

        # Target reduction line starting from first month
        baseline = df_vol["total"].iloc[0]
        df_vol["target"] = [round(baseline * ((1 - target_pct / 100) ** i), 1) for i in range(len(df_vol))]

        # Bar colour: amber if above moving average, steel blue if at or below
        df_vol["status"] = df_vol.apply(
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
                alt.Tooltip("month:N", title="Month"),
                alt.Tooltip("total:Q", title="Total Incidents"),
                alt.Tooltip("moving_avg:Q", title="3-Month Avg"),
                alt.Tooltip("target:Q", title=f"Target (-{target_pct}%/mo)"),
            ]
        )

        ma_line = base.mark_line(color="#F4A900", strokeWidth=2.5, point=True).encode(
            y=alt.Y("moving_avg:Q"),
            tooltip=[alt.Tooltip("moving_avg:Q", title="3-Month Moving Avg")]
        )

        target_line = base.mark_line(color="#C00000", strokeWidth=2, strokeDash=[6, 3]).encode(
            y=alt.Y("target:Q"),
            tooltip=[alt.Tooltip("target:Q", title=f"Target (-{target_pct}%/mo)")]
        )

        chart = (bars + ma_line + target_line).properties(height=350).configure_axis(
            labelFontSize=12, titleFontSize=13
        )
        st.altair_chart(chart, use_container_width=True)
        st.caption("Orange line = 3-month moving average  |  Red dashed = target reduction trajectory")

        # Top 10 applications by total incidents
        top10 = app_month_counts.head(10)["Total"].sort_values(ascending=True)
        st.markdown("**Top 10 Applications — Total Incidents Across All Months**")
        st.bar_chart(top10)
