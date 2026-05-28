import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher

st.set_page_config(
    page_title="Attendance Builder",
    page_icon="🗂️",
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
    background: #1a1d27; border: 1px solid #2a2d3a;
    border-radius: 12px; padding: 1.2rem 1.5rem; text-align: center;
}
.metric-val {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 2rem; font-weight: 700; line-height: 1;
}
.metric-lbl {
    font-size: 0.78rem; color: #7c7f8e; margin-top: 4px;
    text-transform: uppercase; letter-spacing: 0.08em;
}

.matched   { color: #4ade80; }
.unmatched { color: #f87171; }
.fuzzy     { color: #facc15; }
.neutral   { color: #a78bfa; }
.review    { color: #fb923c; }

.step-badge {
    display: inline-block;
    background: #1e1b4b; border: 1px solid #4338ca;
    color: #a5b4fc; border-radius: 8px;
    padding: 0.3rem 0.8rem; font-size: 0.72rem;
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 600; letter-spacing: 0.08em;
    text-transform: uppercase; margin-bottom: 0.8rem;
}
.step-badge.done { background: #14532d; border-color: #166534; color: #4ade80; }
.step-badge.warn { background: #422006; border-color: #92400e; color: #fb923c; }

div[data-testid="stFileUploader"] {
    background: #1a1d27; border: 2px dashed #2a2d3a;
    border-radius: 12px; padding: 1rem;
}
div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
.stButton > button {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: white; border: none; border-radius: 8px;
    padding: 0.6rem 2rem;
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 600; font-size: 0.95rem;
    transition: opacity 0.2s; width: 100%;
}
.stButton > button:hover { opacity: 0.85; }
.section-header {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 0.7rem; font-weight: 600;
    letter-spacing: 0.12em; text-transform: uppercase;
    color: #6366f1; margin-bottom: 0.5rem;
}
</style>
""", unsafe_allow_html=True)

# ── session state init ────────────────────────────────────────────────────────
if "roster" not in st.session_state:
    st.session_state.roster = None   # enriched schedule+staff DataFrame

# ── constants ─────────────────────────────────────────────────────────────────
EXCLUDED_LOCATIONS = {"clark", "dsi", "zamboanga", "isabela"}
DAY_COLS           = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
FUZZY_THRESHOLD    = 0.82

# ── helpers ───────────────────────────────────────────────────────────────────
def normalize(name: str) -> str:
    if pd.isna(name): return ""
    name = str(name).lower().strip()
    name = re.sub(r"[^a-z\s]", "", name)
    return re.sub(r"\s+", " ", name)

def build_staff_lookup(staff_df: pd.DataFrame) -> dict:
    lookup: dict[str, list[int]] = {}
    for idx, row in staff_df.iterrows():
        key = normalize(row.get("Name", ""))
        if key: lookup.setdefault(key, []).append(idx)
        first = normalize(row.get("FirstName", ""))
        last  = normalize(row.get("LastName",  ""))
        if first and last:
            lookup.setdefault(f"{first} {last}", []).append(idx)
    return lookup

def match_name(sched_name: str, lookup: dict, staff_df: pd.DataFrame):
    norm = normalize(sched_name)
    if not norm: return None, "no_name"
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
    if pd.isna(val): return "Rest Day"
    s = str(val).strip()
    if s in ("", "0000 - 0000"): return "Rest Day"
    return s

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Result")
    return buf.getvalue()

def read_timesheet(file) -> pd.DataFrame:
    raw = pd.read_excel(file, sheet_name=0, header=None)
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("Agent Email", case=False).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find 'Agent Email' header in timesheet.")
    df = pd.read_excel(file, sheet_name=0, header=header_row)
    keep = ["Agent", "Agent Email", "Active", "Break", "First Login", "Last Logout"]
    df = df[[c for c in keep if c in df.columns]].copy()
    df = df[df["Agent Email"].notna() & (df["Agent Email"].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)

def attendance_status(first_login, active) -> str:
    """
    Present     : has First Login AND Active >= 0.1
    For Review  : has First Login but Active < 0.1
                  OR no First Login but Active > 0
    blank       : no First Login and no Active (leave blank)
    """
    has_login  = pd.notna(first_login) and str(first_login).strip() not in ("", "NaT")
    try:
        active_val = float(active) if pd.notna(active) else 0.0
    except (ValueError, TypeError):
        active_val = 0.0

    if has_login and active_val >= 0.1:
        return "Present"
    if has_login and active_val < 0.1:
        return "For Review"
    if not has_login and active_val > 0:
        return "For Review"
    return ""   # no login, no active → leave blank


# ── UI header ─────────────────────────────────────────────────────────────────
st.markdown("## 🗂️ Attendance Builder")
st.markdown('<p style="color:#7c7f8e;margin-top:-0.5rem;">Step 1: build the roster → Step 2: process timesheets against it.</p>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📋  Step 1 — Schedule + Staff", "⏱️  Step 2 — Timesheet Attendance"])


# ════════════════════════════════════════════════════════════════════════════════
# STEP 1 — Schedule + Staff → enriched roster
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("")

    roster_ready = st.session_state.roster is not None
    badge_html = (
        '<span class="step-badge done">✔ Roster ready — proceed to Step 2</span>'
        if roster_ready else
        '<span class="step-badge">Roster not built yet</span>'
    )
    st.markdown(badge_html, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">📅 Schedule File</div>', unsafe_allow_html=True)
        sched_file = st.file_uploader("Schedule", type=["xlsx","xls"],
                                      key="sched", label_visibility="collapsed")
    with c2:
        st.markdown('<div class="section-header">👥 Staff Roster File</div>', unsafe_allow_html=True)
        staff_file = st.file_uploader("Staff", type=["xlsx","xls"],
                                      key="staff", label_visibility="collapsed")

    st.markdown("")
    run_step1 = st.button("▶  Build Roster", key="btn_step1")

    if run_step1:
        if not sched_file or not staff_file:
            st.warning("Upload both Schedule and Staff files first.", icon="⚠️")
            st.stop()

        with st.spinner("Reading files…"):
            sched_df = pd.read_excel(sched_file, sheet_name=0)
            staff_df = pd.read_excel(staff_file, sheet_name=0)

        if "Name" not in sched_df.columns:
            st.error("Schedule must have a **Name** column.")
            st.stop()
        missing = {"Name", "CSLoginName", "EmployeeNumber"} - set(staff_df.columns)
        if missing:
            st.error(f"Staff file missing: {missing}")
            st.stop()

        with st.spinner("Matching names…"):
            lookup  = build_staff_lookup(staff_df)
            records = []
            for _, row in sched_df.iterrows():
                matched, mtype = match_name(row["Name"], lookup, staff_df)
                out = row.to_dict()
                for day in DAY_COLS:
                    if day in out:
                        out[day] = normalize_schedule(out[day])
                if matched is not None:
                    loc_raw = matched.get("On Site Location", "")
                    loc = loc_raw if (pd.notna(loc_raw) and str(loc_raw).strip()) else "WFH"
                    out["FullName_Staff"] = matched.get("Name", "")
                    out["CSLoginName"]    = matched.get("CSLoginName", "")
                    out["EmployeeNumber"] = matched.get("EmployeeNumber", "")
                    out["OnSiteLocation"] = loc
                    out["MatchType"]      = mtype
                else:
                    out["FullName_Staff"] = ""
                    out["CSLoginName"]    = ""
                    out["EmployeeNumber"] = ""
                    out["OnSiteLocation"] = ""
                    out["MatchType"]      = mtype
                records.append(out)

        roster = pd.DataFrame(records)

        # filter excluded locations
        before = len(roster)
        roster = roster[
            ~roster["OnSiteLocation"].str.strip().str.lower().isin(EXCLUDED_LOCATIONS)
        ]
        excluded_count = before - len(roster)

        st.session_state.roster = roster

        # metrics
        total     = len(roster)
        exact_n   = roster["MatchType"].str.startswith("exact").sum()
        fuzzy_n   = roster["MatchType"].str.startswith("fuzzy").sum()
        unmatched = roster["MatchType"].isin(["unmatched","no_name"]).sum()

        st.markdown("---")
        m1, m2, m3, m4, m5 = st.columns(5)
        for col, val, lbl, cls in [
            (m1, total,          "Roster Size",         ""),
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
        st.success("✔ Roster built! Switch to **Step 2** to process timesheets.")

        disp = ["Name","FullName_Staff","CSLoginName","EmployeeNumber",
                "OnSiteLocation","HireStatus","Position","MatchType"]
        disp = [c for c in disp if c in roster.columns]
        st.dataframe(roster[disp], use_container_width=True, height=380)

        st.markdown("")
        d1, d2 = st.columns([2,1])
        with d1:
            st.download_button("⬇️ Download Roster as Excel", to_excel_bytes(roster),
                "roster.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with d2:
            st.download_button("⬇️ Download as CSV", roster.to_csv(index=False).encode(),
                "roster.csv", "text/csv", use_container_width=True)

    elif not run_step1 and not roster_ready:
        st.markdown("""
        <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                    padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
            <div style="font-size:2.5rem;margin-bottom:0.5rem;">📋</div>
            <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;
                        color:#e8eaf0;margin-bottom:0.4rem;">Upload Schedule + Staff to build the roster</div>
            <div style="font-size:0.85rem;">
                Enriches each schedule row with <strong style="color:#4ade80">CSLoginName · EmployeeNumber · OnSiteLocation</strong>.<br>
                Excludes Clark · DSI · Zamboanga · Isabela &nbsp;·&nbsp;
                Blank/<em>0000-0000</em> → <strong style="color:#a78bfa">Rest Day</strong>.
            </div>
        </div>
        """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# STEP 2 — Timesheet matched against roster
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("")

    roster = st.session_state.roster

    if roster is None:
        st.markdown("""
        <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                    padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
            <div style="font-size:2.5rem;margin-bottom:0.5rem;">⏳</div>
            <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;
                        color:#e8eaf0;margin-bottom:0.4rem;">Complete Step 1 first</div>
            <div style="font-size:0.85rem;">
                Build the roster in <strong>Step 1</strong> before uploading timesheets.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown('<span class="step-badge done">✔ Roster loaded — upload timesheets below</span>',
                    unsafe_allow_html=True)
        st.markdown(f'<span style="color:#7c7f8e;font-size:0.82rem;">'
                    f'{len(roster):,} roster rows · join key: CSLoginName</span>',
                    unsafe_allow_html=True)

        # Build CSLoginName lookup from roster (lowercase → roster row index list)
        cs_lookup: dict[str, list[int]] = {}
        for idx, row in roster.iterrows():
            cs = str(row.get("CSLoginName","") or "").strip().lower()
            if cs:
                cs_lookup.setdefault(cs, []).append(idx)

        st.markdown("")
        st.markdown('<div class="section-header">⏱️ Timesheet File(s)</div>', unsafe_allow_html=True)
        ts_files = st.file_uploader(
            "Upload timesheet(s) — filename = date (e.g. 5-17-2026.xlsx). Multi-upload supported.",
            type=["xlsx","xls"], accept_multiple_files=True,
            key="ts", label_visibility="collapsed"
        )

        st.markdown("")
        run_step2 = st.button("▶  Build Attendance", key="btn_step2")

        if run_step2:
            if not ts_files:
                st.warning("Upload at least one timesheet file.", icon="⚠️")
                st.stop()

            all_records = []
            errors      = []

            for ts_file in ts_files:
                date_str = ts_file.name.replace(".xlsx","").replace(".xls","")
                try:
                    ts_df = read_timesheet(ts_file)
                except Exception as e:
                    errors.append(f"**{ts_file.name}**: {e}")
                    continue

                # Build a lookup: CSLoginName (agent email) → timesheet row
                # Agent Email in timesheet IS the CSLoginName value
                ts_by_cs: dict[str, dict] = {}
                for _, ts_row in ts_df.iterrows():
                    agent_email = str(ts_row.get("Agent Email","")).strip().lower()
                    if agent_email:
                        ts_by_cs[agent_email] = ts_row.to_dict()

                # Iterate over every roster row and join timesheet data
                for _, r_row in roster.iterrows():
                    cs_key = str(r_row.get("CSLoginName","") or "").strip().lower()
                    ts_row = ts_by_cs.get(cs_key)

                    out = {
                        "Date":          date_str,
                        "Name":          r_row.get("Name",""),
                        "CSLoginName":   r_row.get("CSLoginName",""),
                        "EmployeeNumber":r_row.get("EmployeeNumber",""),
                        "OnSiteLocation":r_row.get("OnSiteLocation",""),
                        "HireStatus":    r_row.get("HireStatus",""),
                        "Position":      r_row.get("Position",""),
                    }

                    # Day schedule columns
                    for day in DAY_COLS:
                        if day in r_row:
                            out[day] = r_row[day]

                    if ts_row is not None:
                        first_login = ts_row.get("First Login")
                        active      = ts_row.get("Active")
                        out["AgentEmail"]  = ts_row.get("Agent Email","")
                        out["FirstLogin"]  = first_login
                        out["LastLogout"]  = ts_row.get("Last Logout","")
                        out["Active"]      = active
                        out["Break"]       = ts_row.get("Break","")
                        out["Attendance"]  = attendance_status(first_login, active)
                    else:
                        # no timesheet record for this roster employee
                        out["AgentEmail"]  = ""
                        out["FirstLogin"]  = ""
                        out["LastLogout"]  = ""
                        out["Active"]      = ""
                        out["Break"]       = ""
                        out["Attendance"]  = ""   # leave blank per spec

                    all_records.append(out)

            if errors:
                for e in errors:
                    st.error(e)

            if not all_records:
                st.warning("No records processed.")
                st.stop()

            att_df = pd.DataFrame(all_records)

            # ── metrics ──────────────────────────────────────────────────────
            total_rows  = len(att_df)
            present_n   = (att_df["Attendance"] == "Present").sum()
            review_n    = (att_df["Attendance"] == "For Review").sum()
            blank_n     = (att_df["Attendance"] == "").sum()

            st.markdown("---")
            ma1, ma2, ma3, ma4 = st.columns(4)
            for col, val, lbl, cls in [
                (ma1, total_rows, "Total Rows",   ""),
                (ma2, present_n,  "Present",       "matched"),
                (ma3, review_n,   "For Review",    "review"),
                (ma4, blank_n,    "No Record",     "unmatched"),
            ]:
                with col:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                        f'<div class="metric-lbl">{lbl}</div></div>',
                        unsafe_allow_html=True
                    )

            st.markdown("")

            disp = (["Date","Name","CSLoginName","EmployeeNumber","OnSiteLocation",
                     "Attendance","FirstLogin","LastLogout","Active","Break",
                     "HireStatus","Position"] + DAY_COLS)
            disp = [c for c in disp if c in att_df.columns]

            ta_all, ta_present, ta_review, ta_blank = st.tabs(
                ["All", "✅ Present", "🟠 For Review", "⬜ No Record"]
            )
            with ta_all:
                st.dataframe(att_df[disp], use_container_width=True, height=420)
            with ta_present:
                st.dataframe(att_df[att_df["Attendance"]=="Present"][disp],
                             use_container_width=True, height=420)
            with ta_review:
                st.dataframe(att_df[att_df["Attendance"]=="For Review"][disp],
                             use_container_width=True, height=420)
            with ta_blank:
                st.dataframe(att_df[att_df["Attendance"]==""][disp],
                             use_container_width=True, height=420)

            st.markdown("---")
            ad1, ad2 = st.columns([2,1])
            with ad1:
                st.download_button("⬇️ Download Attendance as Excel",
                    to_excel_bytes(att_df), "attendance.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ad2:
                st.download_button("⬇️ Download as CSV",
                    att_df.to_csv(index=False).encode(),
                    "attendance.csv", "text/csv", use_container_width=True)
