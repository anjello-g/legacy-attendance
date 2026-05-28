import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher

st.set_page_config(
    page_title="Schedule & Timesheet Tool",
    page_icon="🔗",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=Space+Grotesk:wght@400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'Space Grotesk', sans-serif; }

.stApp { background: #0f1117; color: #e8eaf0; }
.block-container { padding-top: 2rem; max-width: 1400px; }

.metric-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    text-align: center;
}
.metric-val {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    line-height: 1;
}
.metric-lbl {
    font-size: 0.78rem;
    color: #7c7f8e;
    margin-top: 4px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}

.matched   { color: #4ade80; }
.unmatched { color: #f87171; }
.fuzzy     { color: #facc15; }
.neutral   { color: #a78bfa; }

div[data-testid="stFileUploader"] {
    background: #1a1d27;
    border: 2px dashed #2a2d3a;
    border-radius: 12px;
    padding: 1rem;
}
div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

.stButton > button {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 2rem;
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 600;
    font-size: 0.95rem;
    transition: opacity 0.2s;
    width: 100%;
}
.stButton > button:hover { opacity: 0.85; }

.section-header {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #6366f1;
    margin-bottom: 0.5rem;
}

div[data-testid="stTabs"] button {
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)


# ── constants ─────────────────────────────────────────────────────────────────
EXCLUDED_LOCATIONS = {"clark", "dsi", "zamboanga", "isabela"}
DAY_COLS           = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
FUZZY_THRESHOLD    = 0.82


# ── helpers ───────────────────────────────────────────────────────────────────
def normalize(name: str) -> str:
    if pd.isna(name):
        return ""
    name = str(name).lower().strip()
    name = re.sub(r"[^a-z\s]", "", name)
    return re.sub(r"\s+", " ", name)


def build_staff_lookup(staff_df: pd.DataFrame) -> dict:
    lookup: dict[str, list[int]] = {}
    for idx, row in staff_df.iterrows():
        key = normalize(row.get("Name", ""))
        if key:
            lookup.setdefault(key, []).append(idx)
        first = normalize(row.get("FirstName", ""))
        last  = normalize(row.get("LastName",  ""))
        if first and last:
            lookup.setdefault(f"{first} {last}", []).append(idx)
    return lookup


def match_name(sched_name: str, lookup: dict, staff_df: pd.DataFrame):
    norm = normalize(sched_name)
    if not norm:
        return None, "no_name"
    if norm in lookup:
        return staff_df.loc[lookup[norm][0]], "exact"
    best_score, best_key = 0.0, None
    for key in lookup:
        sc = SequenceMatcher(None, norm, key).ratio()
        if sc > best_score:
            best_score, best_key = sc, key
    if best_score >= FUZZY_THRESHOLD and best_key:
        return staff_df.loc[lookup[best_key][0]], f"fuzzy({best_score:.0%})"
    return None, "unmatched"


def normalize_schedule(val) -> str:
    if pd.isna(val):
        return "Rest Day"
    s = str(val).strip()
    if s == "" or s == "0000 - 0000":
        return "Rest Day"
    return s


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Result")
    return buf.getvalue()


def read_timesheet(file) -> pd.DataFrame:
    """
    Timesheet has 2 junk rows then row 2 (0-indexed) is the real header.
    Returns a clean df with renamed columns.
    """
    raw = pd.read_excel(file, sheet_name=0, header=None)
    # Find the header row (contains "Agent Email")
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("Agent Email", case=False).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find header row with 'Agent Email' in timesheet.")
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    # Keep only relevant columns
    keep = ["Agent", "Agent Email", "Active", "Break", "First Login", "Last Logout"]
    df = df[[c for c in keep if c in df.columns]].copy()
    # Drop rows where Agent Email is blank (totals / blank rows)
    df = df[df["Agent Email"].notna() & (df["Agent Email"].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    return df


def build_email_lookup(staff_df: pd.DataFrame) -> dict:
    """email (lowercase) → row index"""
    lookup = {}
    for idx, row in staff_df.iterrows():
        email = str(row.get("HBOSEmail", "") or "").strip().lower()
        if email:
            lookup[email] = idx
        # also try CSLoginName if it looks like an email
        cs = str(row.get("CSLoginName", "") or "").strip().lower()
        if cs and "@" in cs:
            lookup.setdefault(cs, idx)
    return lookup


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("## 🔗 Schedule & Timesheet Tool")
st.markdown('<p style="color:#7c7f8e;margin-top:-0.5rem;">Match schedules and timesheets against the Staff roster.</p>', unsafe_allow_html=True)

main_tab1, main_tab2 = st.tabs(["📅  Schedule Matcher", "⏱️  Timesheet Login Checker"])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — SCHEDULE MATCHER
# ════════════════════════════════════════════════════════════════════════════════
with main_tab1:
    st.markdown("")
    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown('<div class="section-header">📅 Schedule File</div>', unsafe_allow_html=True)
        sched_file = st.file_uploader("Schedule", type=["xlsx", "xls"],
                                      key="sched", label_visibility="collapsed")
    with col_r:
        st.markdown('<div class="section-header">👥 Staff Roster File</div>', unsafe_allow_html=True)
        staff_file_sched = st.file_uploader("Staff (schedule tab)", type=["xlsx", "xls"],
                                             key="staff_sched", label_visibility="collapsed")

    st.markdown("")
    run_sched = st.button("▶  Run Schedule Match", key="btn_sched")

    if run_sched:
        if not sched_file or not staff_file_sched:
            st.warning("Please upload **both** files before running.", icon="⚠️")
            st.stop()

        with st.spinner("Reading files…"):
            sched_df     = pd.read_excel(sched_file, sheet_name=0)
            staff_df_s   = pd.read_excel(staff_file_sched, sheet_name=0)

        if "Name" not in sched_df.columns:
            st.error("Schedule file must have a **Name** column.")
            st.stop()
        missing = {"Name", "CSLoginName", "EmployeeNumber"} - set(staff_df_s.columns)
        if missing:
            st.error(f"Staff file is missing columns: {missing}")
            st.stop()

        with st.spinner("Matching names…"):
            lookup = build_staff_lookup(staff_df_s)
            records = []
            for _, row in sched_df.iterrows():
                matched_row, match_type = match_name(row["Name"], lookup, staff_df_s)
                out = row.to_dict()
                for day in DAY_COLS:
                    if day in out:
                        out[day] = normalize_schedule(out[day])
                if matched_row is not None:
                    loc_raw = matched_row.get("On Site Location", "")
                    loc = loc_raw if (pd.notna(loc_raw) and str(loc_raw).strip()) else "WFH"
                    out["FullName_Staff"] = matched_row.get("Name", "")
                    out["CSLoginName"]    = matched_row.get("CSLoginName", "")
                    out["EmployeeNumber"] = matched_row.get("EmployeeNumber", "")
                    out["OnSiteLocation"] = loc
                    out["MatchType"]      = match_type
                else:
                    out["FullName_Staff"] = ""
                    out["CSLoginName"]    = ""
                    out["EmployeeNumber"] = ""
                    out["OnSiteLocation"] = ""
                    out["MatchType"]      = match_type
                records.append(out)

        result_df = pd.DataFrame(records)
        before_filter = len(result_df)
        result_df = result_df[
            ~result_df["OnSiteLocation"].str.strip().str.lower().isin(EXCLUDED_LOCATIONS)
        ]
        excluded_count = before_filter - len(result_df)

        total     = len(result_df)
        exact_n   = result_df["MatchType"].str.startswith("exact").sum()
        fuzzy_n   = result_df["MatchType"].str.startswith("fuzzy").sum()
        unmatched = result_df["MatchType"].isin(["unmatched", "no_name"]).sum()

        st.markdown("---")
        m1, m2, m3, m4, m5 = st.columns(5)
        for col, val, lbl, cls in [
            (m1, total,          "Rows After Filter",   ""),
            (m2, exact_n,        "Exact Matches",       "matched"),
            (m3, fuzzy_n,        "Fuzzy Matches",       "fuzzy"),
            (m4, unmatched,      "Unmatched",           "unmatched"),
            (m5, excluded_count, "Excluded (Location)", "unmatched"),
        ]:
            with col:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                    f'<div class="metric-lbl">{lbl}</div></div>',
                    unsafe_allow_html=True
                )

        st.markdown("")
        t_all, t_exact, t_fuzzy, t_unmatched = st.tabs(
            ["All Results", "✅ Exact", "🟡 Fuzzy", "❌ Unmatched"]
        )
        display_cols = ["Name", "FullName_Staff", "CSLoginName", "EmployeeNumber",
                        "OnSiteLocation", "HireStatus", "Position", "MatchType"]
        display_cols = [c for c in display_cols if c in result_df.columns]

        def show_table(df_sub):
            st.dataframe(df_sub[display_cols], use_container_width=True, height=400)

        with t_all:      show_table(result_df)
        with t_exact:    show_table(result_df[result_df["MatchType"].str.startswith("exact")])
        with t_fuzzy:    show_table(result_df[result_df["MatchType"].str.startswith("fuzzy")])
        with t_unmatched: show_table(result_df[result_df["MatchType"].isin(["unmatched","no_name"])])

        st.markdown("---")
        dl1, dl2 = st.columns([2, 1])
        with dl1:
            st.download_button(
                "⬇️  Download as Excel", to_excel_bytes(result_df),
                "schedule_matched.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "⬇️  Download as CSV", result_df.to_csv(index=False).encode(),
                "schedule_matched.csv", "text/csv", use_container_width=True,
            )

    elif not run_sched:
        st.markdown("""
        <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                    padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
            <div style="font-size:2.5rem;margin-bottom:0.5rem;">📋</div>
            <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;color:#e8eaf0;margin-bottom:0.4rem;">
                Upload Schedule + Staff Roster to begin
            </div>
            <div style="font-size:0.85rem;">
                Returns <strong style="color:#4ade80">CSLoginName</strong>,
                <strong style="color:#4ade80">EmployeeNumber</strong>,
                <strong style="color:#4ade80">OnSiteLocation</strong>.<br>
                Excludes Clark · DSI · Zamboanga · Isabela.<br>
                Blank / <em>0000-0000</em> slots → <strong style="color:#a78bfa">Rest Day</strong>.
            </div>
        </div>
        """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — TIMESHEET LOGIN CHECKER
# ════════════════════════════════════════════════════════════════════════════════
with main_tab2:
    st.markdown("")
    tc_l, tc_r = st.columns(2)

    with tc_l:
        st.markdown('<div class="section-header">⏱️ Timesheet File(s)</div>', unsafe_allow_html=True)
        ts_files = st.file_uploader(
            "Upload one or more timesheet .xlsx files (filename = date, e.g. 5-17-2026.xlsx)",
            type=["xlsx", "xls"], accept_multiple_files=True,
            key="ts", label_visibility="collapsed"
        )
    with tc_r:
        st.markdown('<div class="section-header">👥 Staff Roster File</div>', unsafe_allow_html=True)
        staff_file_ts = st.file_uploader("Staff (timesheet tab)", type=["xlsx", "xls"],
                                          key="staff_ts", label_visibility="collapsed")

    st.markdown("")
    run_ts = st.button("▶  Run Timesheet Match", key="btn_ts")

    if run_ts:
        if not ts_files or not staff_file_ts:
            st.warning("Please upload **at least one timesheet** and the **Staff Roster**.", icon="⚠️")
            st.stop()

        with st.spinner("Reading staff roster…"):
            staff_df_t   = pd.read_excel(staff_file_ts, sheet_name=0)

        missing = {"HBOSEmail", "CSLoginName"} - set(staff_df_t.columns)
        if missing:
            st.error(f"Staff file is missing columns: {missing}")
            st.stop()

        email_lookup = build_email_lookup(staff_df_t)

        all_records = []
        errors = []

        for ts_file in ts_files:
            # Derive date from filename
            date_str = ts_file.name.replace(".xlsx","").replace(".xls","")

            try:
                ts_df = read_timesheet(ts_file)
            except Exception as e:
                errors.append(f"**{ts_file.name}**: {e}")
                continue

            for _, row in ts_df.iterrows():
                agent_email = str(row.get("Agent Email","")).strip().lower()
                out = {
                    "Date":          date_str,
                    "Agent":         row.get("Agent", ""),
                    "AgentEmail":    row.get("Agent Email", ""),
                    "Active":        row.get("Active", ""),
                    "Break":         row.get("Break", ""),
                    "FirstLogin":    row.get("First Login", ""),
                    "LastLogout":    row.get("Last Logout", ""),
                    "HasLogin":      "Yes" if pd.notna(row.get("First Login")) else "No",
                }

                # Match to staff via email
                if agent_email in email_lookup:
                    idx = email_lookup[agent_email]
                    sr  = staff_df_t.loc[idx]
                    out["CSLoginName"]    = sr.get("CSLoginName", "")
                    out["EmployeeNumber"] = sr.get("EmployeeNumber", "")
                    out["FullName_Staff"] = sr.get("Name", "")
                    out["StaffEmail"]     = sr.get("HBOSEmail", "")
                    out["EmailMatch"]     = "matched"
                else:
                    out["CSLoginName"]    = ""
                    out["EmployeeNumber"] = ""
                    out["FullName_Staff"] = ""
                    out["StaffEmail"]     = ""
                    out["EmailMatch"]     = "unmatched"

                all_records.append(out)

        if errors:
            for e in errors:
                st.error(e)

        if not all_records:
            st.warning("No records could be processed.")
            st.stop()

        ts_result = pd.DataFrame(all_records)

        # ── metrics ──────────────────────────────────────────────────────────
        total_rows    = len(ts_result)
        matched_email = (ts_result["EmailMatch"] == "matched").sum()
        unmatched_em  = (ts_result["EmailMatch"] == "unmatched").sum()
        has_login     = (ts_result["HasLogin"] == "Yes").sum()
        no_login      = (ts_result["HasLogin"] == "No").sum()

        st.markdown("---")
        tm1, tm2, tm3, tm4, tm5 = st.columns(5)
        for col, val, lbl, cls in [
            (tm1, total_rows,    "Total Agents",       ""),
            (tm2, matched_email, "Email Matched",      "matched"),
            (tm3, unmatched_em,  "Email Unmatched",    "unmatched"),
            (tm4, has_login,     "Has Login",          "matched"),
            (tm5, no_login,      "No Login",           "unmatched"),
        ]:
            with col:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                    f'<div class="metric-lbl">{lbl}</div></div>',
                    unsafe_allow_html=True
                )

        st.markdown("")

        ts_display = ["Date", "Agent", "AgentEmail", "CSLoginName", "EmployeeNumber",
                      "FirstLogin", "LastLogout", "Active", "Break", "HasLogin", "EmailMatch"]
        ts_display = [c for c in ts_display if c in ts_result.columns]

        tt_all, tt_login, tt_nologin, tt_unmatched = st.tabs(
            ["All", "✅ Has Login", "❌ No Login", "⚠️ Email Unmatched"]
        )
        with tt_all:
            st.dataframe(ts_result[ts_display], use_container_width=True, height=420)
        with tt_login:
            st.dataframe(ts_result[ts_result["HasLogin"]=="Yes"][ts_display],
                         use_container_width=True, height=420)
        with tt_nologin:
            st.dataframe(ts_result[ts_result["HasLogin"]=="No"][ts_display],
                         use_container_width=True, height=420)
        with tt_unmatched:
            st.dataframe(ts_result[ts_result["EmailMatch"]=="unmatched"][ts_display],
                         use_container_width=True, height=420)

        st.markdown("---")
        tdl1, tdl2 = st.columns([2, 1])
        with tdl1:
            st.download_button(
                "⬇️  Download Timesheet Result as Excel", to_excel_bytes(ts_result),
                "timesheet_matched.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with tdl2:
            st.download_button(
                "⬇️  Download as CSV", ts_result.to_csv(index=False).encode(),
                "timesheet_matched.csv", "text/csv", use_container_width=True,
            )

    elif not run_ts:
        st.markdown("""
        <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                    padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
            <div style="font-size:2.5rem;margin-bottom:0.5rem;">⏱️</div>
            <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;color:#e8eaf0;margin-bottom:0.4rem;">
                Upload Timesheet(s) + Staff Roster to begin
            </div>
            <div style="font-size:0.85rem;">
                Name timesheet files after their date (e.g. <em>5-17-2026.xlsx</em>).<br>
                You can upload <strong style="color:#a78bfa">multiple timesheets</strong> at once.<br>
                Matches by <strong style="color:#4ade80">Agent Email → HBOSEmail</strong> and returns<br>
                <strong style="color:#4ade80">CSLoginName · EmployeeNumber · First Login · Last Logout · Active · Break</strong>.
            </div>
        </div>
        """, unsafe_allow_html=True)
