import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher
from datetime import datetime

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
if "attendance" not in st.session_state:
    st.session_state.attendance = None # step 2 output
if "merged_headcount" not in st.session_state:
    st.session_state.merged_headcount = None # step 3 result

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

def read_leave_file(file) -> pd.DataFrame:
    """Reads leave file, skipping a duplicate header row if detected."""
    df = pd.read_excel(file, sheet_name=0, header=0)
    if len(df) > 0:
        first_row_vals = [str(x).lower().strip() for x in df.iloc[0].values]
        if "name" in first_row_vals and "position" in first_row_vals:
            df = pd.read_excel(file, sheet_name=0, header=1)
    return df

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

def parse_date_str(date_str: str):
    """Try common formats; return datetime or None."""
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def get_day_column(date_str: str) -> str:
    """Map a date string to its 3-letter day column name."""
    dt = parse_date_str(date_str)
    if dt is None:
        return None
    # strftime %a gives Mon, Tue, Wed, Thu, Fri, Sat, Sun
    return dt.strftime("%a")


def clean_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Convert problematic columns to strings to avoid Arrow serialization errors."""
    df = df.copy()
    for col in df.columns:
        # Convert object columns with mixed types to string
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).replace('nan', '').replace('None', '')
        # Convert nullable Int64 etc to regular int or object
        elif pd.api.types.is_integer_dtype(df[col]) and df[col].isna().any():
            df[col] = df[col].astype('Int64').astype(str).replace('<NA>', '')
    return df


# ── UI header ─────────────────────────────────────────────────────────────────
st.markdown("## 🗂️ Attendance Builder")
st.markdown('<p style="color:#7c7f8e;margin-top:-0.5rem;">Step 1: build the roster → Step 2: process timesheets + leave files → Step 3: merge headcount data.</p>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📋  Step 1 — Schedule + Staff", "⏱️  Step 2 — Timesheet + Leave Attendance", "👤  Step 3 — Headcount Merge"])


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
        st.dataframe(clean_for_display(roster[disp]), width="stretch", height=380)

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
# STEP 2 — Timesheet + Leave matched against roster
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
        st.markdown('<span class="step-badge done">✔ Roster loaded — upload files below</span>',
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
        st.markdown('<div class="section-header">🌴 Leave File(s)</div>', unsafe_allow_html=True)
        leave_files = st.file_uploader(
            "Upload leave file(s) — filename = date (e.g. 5-13-2026.xlsx). Multi-upload supported.",
            type=["xlsx","xls"], accept_multiple_files=True,
            key="leave", label_visibility="collapsed"
        )

        st.markdown("")
        run_step2 = st.button("▶  Build Attendance", key="btn_step2")

        if run_step2:
            if not ts_files:
                st.warning("Upload at least one timesheet file.", icon="⚠️")
                st.stop()

            # ── index leave files by date ────────────────────────────────────
            leave_by_date: dict[str, pd.DataFrame] = {}
            if leave_files:
                for lf in leave_files:
                    ldate = lf.name.replace(".xlsx","").replace(".xls","")
                    try:
                        leave_by_date[ldate] = read_leave_file(lf)
                    except Exception as e:
                        st.error(f"Error reading leave file **{lf.name}**: {e}")

            all_records = []
            errors      = []
            output_filename = "attendance.xlsx"

            for i, ts_file in enumerate(ts_files):
                date_str = ts_file.name.replace(".xlsx","").replace(".xls","")
                if i == 0:
                    output_filename = f"{date_str}.xlsx"

                try:
                    ts_df = read_timesheet(ts_file)
                except Exception as e:
                    errors.append(f"**{ts_file.name}**: {e}")
                    continue

                # Build a lookup: CSLoginName (agent email) → timesheet row
                ts_by_cs: dict[str, dict] = {}
                for _, ts_row in ts_df.iterrows():
                    agent_email = str(ts_row.get("Agent Email","")).strip().lower()
                    if agent_email:
                        ts_by_cs[agent_email] = ts_row.to_dict()

                # Build leave lookup for this date (match by normalized Name)
                leave_lookup: dict[str, dict] = {}
                leave_df = leave_by_date.get(date_str)
                if leave_df is not None:
                    for _, lrow in leave_df.iterrows():
                        name = str(lrow.get("name", lrow.get("Name", ""))).strip()
                        if name:
                            leave_lookup[normalize(name)] = lrow.to_dict()

                # Determine which day column applies for this date
                day_col = get_day_column(date_str)

                # Iterate over every roster row and join timesheet + leave data
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

                    # ── Is Scheduled ───────────────────────────────────────
                    # Using the Date, determine the day-of-week column.
                    # If that column is not "Rest Day", mark 1, else 0.
                    is_scheduled = 0
                    if day_col and day_col in r_row:
                        is_scheduled = 1 if str(r_row[day_col]).strip() != "Rest Day" else 0
                    out["Is Scheduled"] = is_scheduled

                    # ── timesheet fields ─────────────────────────────────────
                    first_login = None
                    active      = None
                    if ts_row is not None:
                        first_login = ts_row.get("First Login")
                        active      = ts_row.get("Active")
                        out["AgentEmail"]  = ts_row.get("Agent Email","")
                        out["FirstLogin"]  = first_login
                        out["LastLogout"]  = ts_row.get("Last Logout","")
                        out["Active"]      = active
                        out["Break"]       = ts_row.get("Break","")
                    else:
                        out["AgentEmail"]  = ""
                        out["FirstLogin"]  = ""
                        out["LastLogout"]  = ""
                        out["Active"]      = ""
                        out["Break"]       = ""

                    # ── determine timesheet status ─────────────────────────
                    ts_status = ""
                    if ts_row is not None:
                        ts_status = attendance_status(first_login, active)

                    # ── determine leave status ─────────────────────────────
                    leave_status = None
                    norm_name = normalize(str(r_row.get("Name", "")))
                    leave_row = leave_lookup.get(norm_name)

                    if leave_row is None:
                        # fuzzy fallback
                        best_score, best_key = 0.0, None
                        for key in leave_lookup:
                            sc = SequenceMatcher(None, norm_name, key).ratio()
                            if sc > best_score:
                                best_score, best_key = sc, key
                        if best_score >= FUZZY_THRESHOLD and best_key:
                            leave_row = leave_lookup[best_key]

                    if leave_row is not None:
                        # SL / UPL → Absent
                        for col in ["SL", "UPL"]:
                            val = leave_row.get(col, 0)
                            if pd.notna(val):
                                try:
                                    if float(val) > 0:
                                        leave_status = "Absent"
                                        break
                                except (ValueError, TypeError):
                                    if str(val).strip().lower() in ("1", "yes", "true", "y"):
                                        leave_status = "Absent"
                                        break
                        # Other leave types → Leave
                        if leave_status is None:
                            for col in ["VL", "ML", "BL", "SPL", "PL", "MWL", "BDL"]:
                                val = leave_row.get(col, 0)
                                if pd.notna(val):
                                    try:
                                        if float(val) > 0:
                                            leave_status = "Leave"
                                            break
                                    except (ValueError, TypeError):
                                        if str(val).strip().lower() in ("1", "yes", "true", "y"):
                                            leave_status = "Leave"
                                            break

                    # ── apply priority: Present → For Review → Leave → Absent
                    present    = 1 if ts_status == "Present"     else 0
                    for_review = 1 if ts_status == "For Review"  else 0
                    leave      = 0
                    absent     = 0

                    if present == 1:
                        pass
                    elif for_review == 1:
                        pass
                    elif leave_status == "Leave":
                        leave = 1
                    else:
                        # Covers: SL/UPL leave, no leave, no timesheet
                        absent = 1

                    out["Present"]    = present
                    out["For Review"] = for_review
                    out["Leave"]      = leave
                    out["Absent"]     = absent

                    all_records.append(out)

            if errors:
                for e in errors:
                    st.error(e)

            if not all_records:
                st.warning("No records processed.")
                st.stop()

            att_df = pd.DataFrame(all_records)
            st.session_state.attendance = att_df

            # ── metrics ──────────────────────────────────────────────────────
            total_rows = len(att_df)
            present_n  = att_df["Present"].sum()
            review_n   = att_df["For Review"].sum()
            leave_n    = att_df["Leave"].sum()
            absent_n   = att_df["Absent"].sum()

            st.markdown("---")
            ma1, ma2, ma3, ma4, ma5 = st.columns(5)
            for col, val, lbl, cls in [
                (ma1, total_rows, "Total Rows",   ""),
                (ma2, present_n,  "Present",       "matched"),
                (ma3, review_n,   "For Review",    "review"),
                (ma4, leave_n,    "On Leave",      "neutral"),
                (ma5, absent_n,   "Absent",        "unmatched"),
            ]:
                with col:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                        f'<div class="metric-lbl">{lbl}</div></div>',
                        unsafe_allow_html=True
                    )

            st.markdown("")

            disp = (["Date","Name","CSLoginName","EmployeeNumber","OnSiteLocation",
                     "Is Scheduled","Present","Absent","Leave","For Review",
                     "FirstLogin","LastLogout","Active","Break",
                     "HireStatus","Position"] + DAY_COLS)
            disp = [c for c in disp if c in att_df.columns]

            ta_all, ta_present, ta_review, ta_leave, ta_absent = st.tabs(
                ["All", "✅ Present", "🟠 For Review", "🌴 Leave", "❌ Absent"]
            )
            with ta_all:
                st.dataframe(clean_for_display(att_df[disp]), width="stretch", height=420)
            with ta_present:
                st.dataframe(clean_for_display(att_df[att_df["Present"]==1][disp]),
                             width="stretch", height=420)
            with ta_review:
                st.dataframe(clean_for_display(att_df[att_df["For Review"]==1][disp]),
                             width="stretch", height=420)
            with ta_leave:
                st.dataframe(clean_for_display(att_df[att_df["Leave"]==1][disp]),
                             width="stretch", height=420)
            with ta_absent:
                st.dataframe(clean_for_display(att_df[att_df["Absent"]==1][disp]),
                             width="stretch", height=420)

            st.markdown("---")
            ad1, ad2 = st.columns([2,1])
            with ad1:
                st.download_button("⬇️ Download Attendance as Excel",
                    to_excel_bytes(att_df), output_filename,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ad2:
                csv_name = output_filename.replace(".xlsx", ".csv").replace(".xls", ".csv")
                st.download_button("⬇️ Download as CSV",
                    att_df.to_csv(index=False).encode(),
                    csv_name, "text/csv", use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════════
# STEP 3 — Headcount Merge (ECN ↔ EmployeeNumber)
# ════════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("")

    attendance = st.session_state.attendance

    if attendance is None:
        st.markdown("""
        <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                    padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
            <div style="font-size:2.5rem;margin-bottom:0.5rem;">⏳</div>
            <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;
                        color:#e8eaf0;margin-bottom:0.4rem;">Complete Step 2 first</div>
            <div style="font-size:0.85rem;">
                Build the attendance in <strong>Step 2</strong> before uploading the headcount file.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        step3_badge = '<span class="step-badge done">✔ Attendance loaded — upload headcount below</span>'
        st.markdown(step3_badge, unsafe_allow_html=True)
        st.markdown(f'<span style="color:#7c7f8e;font-size:0.82rem;">'
                    f'{len(attendance):,} attendance rows · join key: EmployeeNumber → ECN</span>',
                    unsafe_allow_html=True)

        st.markdown("")
        st.markdown('<div class="section-header">👤 Headcount File</div>', unsafe_allow_html=True)
        hc_file = st.file_uploader(
            "Upload headcount export (e.g. Knack RCM Headcount_05.13.2026.xlsx). "
            "Must contain ECN column.",
            type=["xlsx","xls"],
            key="hc_uploader", label_visibility="collapsed"
        )

        st.markdown("")
        run_step3 = st.button("▶  Merge Headcount", key="btn_step3")

        if run_step3:
            if not hc_file:
                st.warning("Upload the headcount file first.", icon="⚠️")
                st.stop()

            with st.spinner("Reading headcount file…"):
                hc_df = pd.read_excel(hc_file, sheet_name=0)

            # Validate required columns
            required_hc = {"ECN", "Employee", "Project", "Sub-Process", "Supervisor",
                           "Role", "Manager", "DOJ Knack", "Date of Separation",
                           "Billable/Buffer", "Active/Inactive"}
            missing_hc = required_hc - set(hc_df.columns)
            if missing_hc:
                st.error(f"Headcount file missing required columns: {missing_hc}")
                st.stop()

            # Ensure ECN is numeric for clean merge
            hc_df["ECN"] = pd.to_numeric(hc_df["ECN"], errors="coerce")
            hc_df = hc_df.dropna(subset=["ECN"])
            hc_df["ECN"] = hc_df["ECN"].astype(int)

            # Prepare attendance EmployeeNumber for merge
            att_merge = attendance.copy()
            att_merge["_merge_key"] = pd.to_numeric(att_merge["EmployeeNumber"], errors="coerce")
            att_merge = att_merge.dropna(subset=["_merge_key"])
            att_merge["_merge_key"] = att_merge["_merge_key"].astype(int)

            # Select only the columns we want from headcount
            hc_cols = ["ECN", "Employee", "Project", "Sub-Process", "Supervisor",
                       "Role", "Manager", "DOJ Knack", "Date of Separation",
                       "Billable/Buffer", "Active/Inactive"]
            hc_subset = hc_df[hc_cols].copy()

            # Merge: attendance.EmployeeNumber == headcount.ECN
            merged = att_merge.merge(
                hc_subset,
                left_on="_merge_key",
                right_on="ECN",
                how="left"
            )

            # Drop helper column
            merged = merged.drop(columns=["_merge_key", "ECN"], errors="ignore")

            # Track match status
            merged["HeadcountMatch"] = merged["Employee"].notna().map({True: "✅ Matched", False: "❌ Unmatched"})

            st.session_state.merged_headcount = merged

            # ── metrics ──────────────────────────────────────────────────────
            total_m    = len(merged)
            matched_n  = (merged["HeadcountMatch"] == "✅ Matched").sum()
            unmatched_n = (merged["HeadcountMatch"] == "❌ Unmatched").sum()

            st.markdown("---")
            mm1, mm2, mm3 = st.columns(3)
            for col, val, lbl, cls in [
                (mm1, total_m,     "Total Rows",      ""),
                (mm2, matched_n,   "Matched",         "matched"),
                (mm3, unmatched_n, "Unmatched",        "unmatched"),
            ]:
                with col:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-val {cls}">{val:,}</div>'
                        f'<div class="metric-lbl">{lbl}</div></div>',
                        unsafe_allow_html=True
                    )

            st.markdown("")
            st.success("✔ Headcount merged! Review and download below.")

            # Display columns: original attendance + new headcount columns
            display_cols = (["Date", "Name", "CSLoginName", "EmployeeNumber",
                            "OnSiteLocation", "Is Scheduled",
                            "Present", "Absent", "Leave", "For Review",
                            "Employee", "Project", "Sub-Process", "Supervisor",
                            "Role", "Manager", "DOJ Knack", "Date of Separation",
                            "Billable/Buffer", "Active/Inactive",
                            "HeadcountMatch"] + DAY_COLS)
            display_cols = [c for c in display_cols if c in merged.columns]

            # Tabs for matched vs unmatched
            t_all, t_matched, t_unmatched = st.tabs(
                ["All Records", "✅ Matched", "❌ Unmatched"]
            )
            with t_all:
                st.dataframe(clean_for_display(merged[display_cols]), width="stretch", height=420)
            with t_matched:
                st.dataframe(clean_for_display(merged[merged["HeadcountMatch"] == "✅ Matched"][display_cols]),
                             width="stretch", height=420)
            with t_unmatched:
                st.dataframe(clean_for_display(merged[merged["HeadcountMatch"] == "❌ Unmatched"][display_cols]),
                             width="stretch", height=420)

            st.markdown("---")
            md1, md2 = st.columns([2,1])
            with md1:
                st.download_button("⬇️ Download Final as Excel",
                    to_excel_bytes(merged), "final_attendance.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with md2:
                st.download_button("⬇️ Download as CSV",
                    merged.to_csv(index=False).encode(),
                    "final_attendance.csv", "text/csv", use_container_width=True)
