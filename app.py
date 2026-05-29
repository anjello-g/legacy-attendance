import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher
from datetime import datetime, time, date
import dateutil.parser  # robust date/time parsing

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
.not-hired { color: #60a5fa; }

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
    st.session_state.roster = None
if "attendance" not in st.session_state:
    st.session_state.attendance = None
if "merged_headcount" not in st.session_state:
    st.session_state.merged_headcount = None
if "headcount_df" not in st.session_state:
    st.session_state.headcount_df = None

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
    return ""

def parse_date_str(date_str: str):
    """Try common formats; return datetime or None."""
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def get_day_column(date_str: str) -> str:
    dt = parse_date_str(date_str)
    if dt is None:
        return None
    return dt.strftime("%a")

def clean_for_display(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).replace('nan', '').replace('None', '')
        elif pd.api.types.is_integer_dtype(df[col]) and df[col].isna().any():
            df[col] = df[col].astype('Int64').astype(str).replace('<NA>', '')
    return df

# ── new helpers for shift & late calculation ───────────────────────────────────
def parse_shift_start(shift_str: str) -> time:
    """Extract start time from '0900 - 1800' format. Returns time object or None."""
    if pd.isna(shift_str) or shift_str == "Rest Day":
        return None
    parts = shift_str.strip().split("-")
    if len(parts) < 2:
        return None
    start_str = parts[0].strip()
    # expect HHMM
    if len(start_str) == 4 and start_str.isdigit():
        hour = int(start_str[:2])
        minute = int(start_str[2:])
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return time(hour, minute)
    return None

def safe_parse_datetime(value):
    """Convert a value (string, datetime, etc.) to datetime. Returns None if impossible."""
    if pd.isna(value) or str(value).strip() in ("", "NaT"):
        return None
    if isinstance(value, datetime):
        return value
    try:
        return dateutil.parser.parse(str(value))
    except:
        return None

def compute_late(shift_start: time, first_login_dt, date_obj: date) -> tuple[int, float]:
    """
    Returns (Late flag 0/1, late_minutes decimal).
    If first_login is after shift_start, returns 1 and minutes late.
    Otherwise returns 0, 0.0.
    """
    if shift_start is None or first_login_dt is None or date_obj is None:
        return 0, 0.0
    # Combine date with shift start time
    threshold = datetime.combine(date_obj, shift_start)
    diff = (first_login_dt - threshold).total_seconds() / 60.0
    if diff > 0:
        return 1, round(diff, 2)
    return 0, 0.0

# ── DOJ Knack helper ──────────────────────────────────────────────────────────
def parse_doj_knack(value) -> date:
    """Parse DOJ Knack value to date object. Returns None if unparseable."""
    if pd.isna(value) or str(value).strip() in ("", "NaT"):
        return None
    dt = safe_parse_datetime(value)
    if dt is not None:
        return dt.date() if hasattr(dt, 'date') else dt
    return None

# ── Filename date normalizer ──────────────────────────────────────────────────
def normalize_filename_date(filename: str) -> str:
    """Parse a date from a filename (e.g. '5-1-2026.xlsx') and return YYYY-MM-DD string."""
    base = filename.replace(".xlsx", "").replace(".xls", "").strip()
    dt = parse_date_str(base)
    if dt is not None:
        return dt.strftime("%Y-%m-%d")
    return base  # fallback to raw string if unparseable

# ── Headcount date extractor ────────────────────────────────────────────────────
def extract_headcount_date(filename: str, hc_df: pd.DataFrame) -> date:
    """Try to find the effective date of a headcount file.
    1. Look for 'Date Exported' column in the dataframe
    2. Parse date from filename (e.g. Headcount_05.13.2026.xlsx)
    3. Return None if neither works
    """
    # Check for Date Exported column
    for col in ["Date Exported", "DateExported", "Export Date", "ExportDate", "Date", "As Of"]:
        if col in hc_df.columns:
            val = hc_df[col].dropna().iloc[0] if len(hc_df) > 0 else None
            if val is not None:
                dt = safe_parse_datetime(val)
                if dt is not None:
                    return dt.date() if hasattr(dt, 'date') else dt

    # Try filename
    base = filename.replace(".xlsx", "").replace(".xls", "").strip()
    # Look for patterns like _05.13.2026 or _05-13-2026 or _5-1-2026
    dt = parse_date_str(base)
    if dt is not None:
        return dt.date()

    return None
    """Parse a date from a filename (e.g. '5-1-2026.xlsx') and return YYYY-MM-DD string."""
    base = filename.replace(".xlsx", "").replace(".xls", "").strip()
    dt = parse_date_str(base)
    if dt is not None:
        return dt.strftime("%Y-%m-%d")
    return base  # fallback to raw string if unparseable

# ── Override date helper ──────────────────────────────────────────────────────
def parse_override_date(value) -> date:
    """Parse an override file date value. Returns None if unparseable."""
    if pd.isna(value) or str(value).strip() in ("", "NaT", "None"):
        return None
    dt = safe_parse_datetime(value)
    if dt is not None:
        return dt.date() if hasattr(dt, 'date') else dt
    # Try common raw formats
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y",
                "%m-%d-%y", "%d-%m-%y", "%y-%m-%d"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    return None

def build_override_template() -> bytes:
    """Build a 3-sheet Excel template for overrides."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    buf = BytesIO()
    wb = Workbook()

    # Common styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4338ca", end_color="4338ca", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # ── Sheet 1: Schedule ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Schedule"
    sched_headers = ["ECN", "Effective Date", "Effective Until", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    ws1.append(sched_headers)
    for col in range(1, len(sched_headers) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws1.append([12345, "05-01-2026", "05-31-2026", "0900 - 1800", "Rest Day", "0900 - 1800", "0900 - 1800", "0900 - 1800", "0900 - 1800", "Rest Day"])
    ws1.append([67890, "05-15-2026", "05-15-2026", "Rest Day", "Rest Day", "Rest Day", "Rest Day", "1200 - 2100", "Rest Day", "Rest Day"])
    for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
        for cell in row:
            cell.border = thin_border
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 16
    ws1.column_dimensions['C'].width = 16
    for c in ['D','E','F','G','H','I','J']:
        ws1.column_dimensions[c].width = 14

    # ── Sheet 2: Leave ────────────────────────────────────────────────
    ws2 = wb.create_sheet(title="Leave")
    leave_headers = ["ECN", "Date", "Leave"]
    ws2.append(leave_headers)
    for col in range(1, len(leave_headers) + 1):
        cell = ws2.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws2.append([12345, "05-20-2026", 1])
    ws2.append([67890, "05-25-2026", 1])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            cell.border = thin_border
    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 10

    # ── Sheet 3: Holidays ─────────────────────────────────────────────
    ws3 = wb.create_sheet(title="Holidays")
    holiday_headers = ["Project", "Sub-Process", "Date"]
    ws3.append(holiday_headers)
    for col in range(1, len(holiday_headers) + 1):
        cell = ws3.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws3.append(["Project Alpha", "", "05-01-2026"])
    ws3.append(["Project Beta", "Sub-Process X", "05-12-2026"])
    ws3.append(["Project Gamma", "", "05-30-2026"])
    for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row):
        for cell in row:
            cell.border = thin_border
    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 16

    # ── Sheet 4: Exemptions ───────────────────────────────────────────
    ws4 = wb.create_sheet(title="Exemptions")
    exempt_headers = ["ECN", "Exemption Date", "Effective Until", "Active/Inactive"]
    ws4.append(exempt_headers)
    for col in range(1, len(exempt_headers) + 1):
        cell = ws4.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws4.append([12345, "05-10-2026", "05-20-2026", "Inactive"])
    ws4.append([67890, "05-15-2026", "05-25-2026", "Inactive"])
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
        for cell in row:
            cell.border = thin_border
    ws4.column_dimensions['A'].width = 12
    ws4.column_dimensions['B'].width = 18
    ws4.column_dimensions['C'].width = 18
    ws4.column_dimensions['D'].width = 18

    # ── Sheet 4: Exemptions ────────────────────────────────────────────
    ws4 = wb.create_sheet(title="Exemptions")
    exemp_headers = ["ECN", "Effective Date", "Effective Until", "Status"]
    ws4.append(exemp_headers)
    for col in range(1, len(exemp_headers) + 1):
        cell = ws4.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws4.append([12345, "05-01-2026", "05-15-2026", "RTWO"])
    ws4.append([67890, "05-20-2026", "05-25-2026", "Pending Deactivation"])
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
        for cell in row:
            cell.border = thin_border
    ws4.column_dimensions['A'].width = 12
    ws4.column_dimensions['B'].width = 16
    ws4.column_dimensions['C'].width = 16
    ws4.column_dimensions['D'].width = 22

    wb.save(buf)
    return buf.getvalue()

# ── UI header ─────────────────────────────────────────────────────────────────
st.markdown("## 🗂️ Attendance Builder")
st.markdown('<p style="color:#7c7f8e;margin-top:-0.5rem;">Step 1: build the roster → Step 2: process timesheets + leave + headcount merge + overrides → final attendance.</p>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📋  Step 1 — Schedule + Staff", "⏱️  Step 2 — Build Final Attendance"])


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
# STEP 2 — Timesheet + Leave matched against roster + DOJ Knack check
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

        cs_lookup: dict[str, list[int]] = {}
        for idx, row in roster.iterrows():
            cs = str(row.get("CSLoginName","") or "").strip().lower()
            if cs:
                cs_lookup.setdefault(cs, []).append(idx)

        st.markdown("")
        st.markdown('<div class="section-header">⏱️ Timesheet File(s) *</div>', unsafe_allow_html=True)
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
        st.markdown('<div class="section-header">👤 Headcount File(s) *</div>', unsafe_allow_html=True)
        st.markdown(
            '<span style="color:#7c7f8e;font-size:0.82rem;">'
            'Upload one or more running headcount exports. The app picks the best file per timesheet date '
            'based on the export date in the filename (e.g. Headcount_05.13.2026.xlsx).</span>',
            unsafe_allow_html=True
        )
        hc_files = st.file_uploader(
            "Upload headcount export(s).",
            type=["xlsx","xls"], accept_multiple_files=True,
            key="hc_step2", label_visibility="collapsed"
        )

        st.markdown("")
        st.markdown('<div class="section-header">📋 Override File (Optional)</div>', unsafe_allow_html=True)
        st.markdown(
            '<span style="color:#7c7f8e;font-size:0.82rem;">'
            '4 sheets: Schedule, Leave, Holidays, Exemptions. Overrides automated rules.</span>',
            unsafe_allow_html=True
        )
        override_file = st.file_uploader(
            "Override file",
            type=["xlsx"],
            key="override_uploader", label_visibility="collapsed"
        )

        st.markdown("")
        tmpl1, tmpl2 = st.columns([1,1])
        with tmpl1:
            st.download_button(
                "⬇️ Download Override Template",
                build_override_template(),
                "override_template.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with tmpl2:
            st.markdown(
                '<div style="color:#7c7f8e;font-size:0.78rem;padding-top:0.4rem;">'
                'Template includes sample rows for all 4 sheets.</div>',
                unsafe_allow_html=True
            )

        st.markdown("")
        run_step2 = st.button("▶  Build Final Attendance", key="btn_step2")

        if run_step2:
            if not ts_files:
                st.warning("Upload at least one timesheet file.", icon="⚠️")
                st.stop()
            if not hc_files:
                st.warning("Upload at least one headcount file.", icon="⚠️")
                st.stop()

            # ── Pre-read all headcount files ─────────────────────────────────
            hc_cache: dict[str, pd.DataFrame] = {}
            hc_dates: list[date] = []

            with st.spinner(f"Reading {len(hc_files)} headcount file(s)…"):
                for f in hc_files:
                    export_dt = extract_headcount_date(f.name, df)
                    dt_key = export_dt.strftime("%Y-%m-%d") if export_dt else f.name
                    try:
                        df = pd.read_excel(f, sheet_name=0)
                        if "ECN" not in df.columns:
                            st.error(f"Headcount file **{f.name}** missing **ECN** column.")
                            st.stop()
                        if "DOJ Knack" not in df.columns:
                            st.error(f"Headcount file **{f.name}** missing **DOJ Knack** column.")
                            st.stop()
                        df["ECN"] = pd.to_numeric(df["ECN"], errors="coerce")
                        df = df.dropna(subset=["ECN"])
                        df["ECN"] = df["ECN"].astype(int)
                        hc_cache[dt_key] = df
                        if export_dt:
                            hc_dates.append(export_dt)
                    except Exception as e:
                        st.error(f"Error reading headcount **{f.name}**: {e}")
                        st.stop()

            if hc_dates:
                hc_dates.sort()
                st.info(f"Headcount export dates detected: {', '.join(d.strftime('%m-%d-%Y') for d in hc_dates)}")
            else:
                st.warning("Could not parse dates from headcount filenames. Using first file for all dates.", icon="⚠️")

            leave_by_date: dict[str, pd.DataFrame] = {}
            if leave_files:
                for lf in leave_files:
                    ldate = normalize_filename_date(lf.name)
                    try:
                        leave_by_date[ldate] = read_leave_file(lf)
                    except Exception as e:
                        st.error(f"Error reading leave file **{lf.name}**: {e}")

            all_merged_records: list[pd.DataFrame] = []
            errors = []
            output_filename = "attendance.xlsx"

            # Warn about stale headcount
            if hc_date:
                stale_files = []
                for ts_file in ts_files:
                    ts_dt = parse_date_str(ts_file.name.replace(".xlsx","").replace(".xls",""))
                    if ts_dt and ts_dt.date() > hc_date:
                        stale_files.append(ts_file.name)
                if stale_files:
                    st.warning(
                        f"⚠️ Headcount snapshot ({hc_date.strftime('%m/%d/%Y')}) is older than "
                        f"some timesheet dates: {', '.join(stale_files)}. "
                        f"Consider using a newer headcount export.",
                        icon="📅"
                    )

            for i, ts_file in enumerate(ts_files):
                date_str_raw = ts_file.name.replace(".xlsx","").replace(".xls","")
                date_str = normalize_filename_date(ts_file.name)
                if i == 0:
                    output_filename = f"{date_str_raw}.xlsx"

                parsed_date = parse_date_str(date_str_raw)
                date_obj = parsed_date.date() if parsed_date else None

                # ── Pick the best headcount for this timesheet date ──────────────
                hc_df = None
                hc_export_dt = None
                if date_obj is not None and hc_dates:
                    best_dt = None
                    for dt in sorted(hc_dates):
                        if dt <= date_obj:
                            best_dt = dt
                    if best_dt is None:
                        best_dt = min(hc_dates)
                    hc_export_dt = best_dt
                    hc_df = hc_cache[best_dt.strftime("%Y-%m-%d")]
                else:
                    hc_df = next(iter(hc_cache.values()))

                # Build DOJ lookup for this specific headcount
                doj_lookup: dict[int, date] = {}
                for _, row in hc_df.iterrows():
                    ecn = int(row["ECN"])
                    doj = parse_doj_knack(row.get("DOJ Knack"))
                    if doj is not None:
                        doj_lookup[ecn] = doj

                if hc_export_dt:
                    st.write(f"📅 **{date_str_raw}** → headcount exported **{hc_export_dt.strftime('%m-%d-%Y')}** ({len(doj_lookup):,} DOJ records)")
                else:
                    st.write(f"📅 **{date_str_raw}** → headcount file ({len(doj_lookup):,} DOJ records)")

                try:
                    ts_df = read_timesheet(ts_file)
                except Exception as e:
                    errors.append(f"**{ts_file.name}**: {e}")
                    continue

                ts_by_cs: dict[str, dict] = {}
                for _, ts_row in ts_df.iterrows():
                    agent_email = str(ts_row.get("Agent Email","")).strip().lower()
                    if agent_email:
                        ts_by_cs[agent_email] = ts_row.to_dict()

                leave_lookup: dict[str, dict] = {}
                leave_df = leave_by_date.get(date_str)
                if leave_df is not None:
                    for _, lrow in leave_df.iterrows():
                        name = ""
                        for col in ["Name", "name", "Employee Name", "EmployeeName", "Full Name", "FullName"]:
                            if col in lrow and pd.notna(lrow[col]):
                                name = str(lrow[col]).strip()
                                if name:
                                    break
                        if name:
                            clean_name = re.sub(r"\s+", " ", name.lower().strip())
                            clean_name = re.sub(r"[^a-z\s]", "", clean_name)
                            leave_lookup[clean_name] = lrow.to_dict()

                day_col = get_day_column(date_str_raw)
                all_records = []

                for _, r_row in roster.iterrows():
                    cs_key = str(r_row.get("CSLoginName","") or "").strip().lower()
                    ts_row = ts_by_cs.get(cs_key)

                    out = {
                        "Date":          date_str_raw,
                        "Name":          r_row.get("Name",""),
                        "CSLoginName":   r_row.get("CSLoginName",""),
                        "EmployeeNumber":r_row.get("EmployeeNumber",""),
                        "OnSiteLocation":r_row.get("OnSiteLocation",""),
                        "HireStatus":    r_row.get("HireStatus",""),
                        "Position":      r_row.get("Position",""),
                    }

                    for day in DAY_COLS:
                        if day in r_row:
                            out[day] = r_row[day]

                    is_scheduled = 0
                    shift_str = "Rest Day"
                    if day_col and day_col in r_row:
                        shift_str = normalize_schedule(r_row[day_col])
                        if shift_str != "Rest Day":
                            is_scheduled = 1
                    out["Is Scheduled"] = is_scheduled

                    out["Day"] = day_col if day_col else ""
                    out["Shift"] = shift_str if shift_str else ""

                    not_yet_hired = False
                    emp_num_raw = r_row.get("EmployeeNumber", "")
                    try:
                        emp_num = int(float(emp_num_raw)) if pd.notna(emp_num_raw) and str(emp_num_raw).strip() != "" else None
                    except (ValueError, TypeError):
                        emp_num = None

                    if emp_num is not None and emp_num in doj_lookup and date_obj is not None:
                        doj_date = doj_lookup[emp_num]
                        if doj_date > date_obj:
                            not_yet_hired = True
                            out["Shift"] = "Not yet hired"
                            out["Is Scheduled"] = 0
                            if day_col and day_col in out:
                                out[day_col] = "Not yet hired"

                    first_login_raw = None
                    active = None
                    if ts_row is not None:
                        first_login_raw = ts_row.get("First Login")
                        active      = ts_row.get("Active")
                        out["AgentEmail"]  = ts_row.get("Agent Email","")
                        out["FirstLogin"]  = first_login_raw
                        out["LastLogout"]  = ts_row.get("Last Logout","")
                        out["Active"]      = active
                        out["Break"]       = ts_row.get("Break","")
                    else:
                        out["AgentEmail"]  = ""
                        out["FirstLogin"]  = ""
                        out["LastLogout"]  = ""
                        out["Active"]      = ""
                        out["Break"]       = ""

                    ts_status = ""
                    if ts_row is not None:
                        ts_status = attendance_status(first_login_raw, active)

                    leave_status = None
                    roster_name = str(r_row.get("Name", ""))
                    norm_name = re.sub(r"\s+", " ", roster_name.lower().strip())
                    norm_name = re.sub(r"[^a-z\s]", "", norm_name)
                    leave_row = leave_lookup.get(norm_name)
                    if leave_row is None:
                        best_score, best_key = 0.0, None
                        for key in leave_lookup:
                            sc = SequenceMatcher(None, norm_name, key).ratio()
                            if sc > best_score:
                                best_score, best_key = sc, key
                        if best_score >= FUZZY_THRESHOLD and best_key:
                            leave_row = leave_lookup[best_key]

                    if leave_row is not None:
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

                    present    = 1 if ts_status == "Present"     else 0
                    for_review = 1 if ts_status == "For Review"  else 0
                    leave      = 0
                    absent     = 0

                    if not_yet_hired:
                        present = 0
                        for_review = 0
                        leave = 0
                        absent = 0
                    elif present == 1:
                        pass
                    elif for_review == 1:
                        if shift_str == "Rest Day":
                            absent = 0
                            is_scheduled = 0
                        else:
                            absent = 1
                            is_scheduled = 1
                    elif leave_status == "Leave":
                        leave = 1
                        absent = 0
                    else:
                        absent = 1

                    out["Present"]    = present
                    out["For Review"] = for_review
                    out["Leave"]      = leave
                    out["Absent"]     = absent

                    if not_yet_hired:
                        out["Late"] = 0
                        out["Late Minutes"] = 0.0
                    else:
                        shift_start_time = parse_shift_start(shift_str)
                        first_login_dt = safe_parse_datetime(first_login_raw) if ts_row else None
                        late_flag, late_minutes = compute_late(shift_start_time, first_login_dt, date_obj)
                        out["Late"] = late_flag
                        out["Late Minutes"] = late_minutes

                    all_records.append(out)

                if not all_records:
                    continue

                att_df = pd.DataFrame(all_records)

                # ── Merge Headcount for this date ──────────────────────────────
                required_hc = {"ECN", "Employee", "Project", "Sub-Process", "Supervisor",
                               "Role", "Manager", "DOJ Knack", "Date of Separation",
                               "Billable/Buffer", "Active/Inactive"}
                missing_hc = required_hc - set(hc_df.columns)
                if missing_hc:
                    st.error(f"Headcount file for {date_str_raw} missing required columns: {missing_hc}")
                    continue

                hc_df["ECN"] = pd.to_numeric(hc_df["ECN"], errors="coerce")
                hc_df = hc_df.dropna(subset=["ECN"])
                hc_df["ECN"] = hc_df["ECN"].astype(int)

                att_merge = att_df.copy()
                att_merge["_merge_key"] = pd.to_numeric(att_merge["EmployeeNumber"], errors="coerce")
                att_merge = att_merge.dropna(subset=["_merge_key"])
                att_merge["_merge_key"] = att_merge["_merge_key"].astype(int)

                hc_cols = ["ECN", "Employee", "Project", "Sub-Process", "Supervisor",
                           "Role", "Manager", "DOJ Knack", "Date of Separation",
                           "Billable/Buffer", "Active/Inactive"]
                hc_subset = hc_df[hc_cols].copy()

                merged = att_merge.merge(
                    hc_subset,
                    left_on="_merge_key",
                    right_on="ECN",
                    how="left"
                )

                merged = merged.drop(columns=["_merge_key", "ECN"], errors="ignore")
                merged["HeadcountMatch"] = merged["Employee"].notna().map({True: "✅ Matched", False: "❌ Unmatched"})

                # ── separation date rule ─────────────────────────────────────
                sep_col = "Date of Separation"
                if sep_col in merged.columns:
                    has_sep = merged[sep_col].apply(
                        lambda x: pd.notna(x) and str(x).strip() != "" and str(x).strip().lower() != "nat"
                    )
                    merged.loc[has_sep, "Absent"] = 0
                    merged.loc[has_sep, "Is Scheduled"] = 0

                # ── Maternity / Suspended rule ────────────────────────────────
                if "Role" in merged.columns:
                    role_lower = merged["Role"].astype(str).str.lower()
                    is_mat_or_susp = role_lower.str.contains("maternity|suspended", case=False, na=False)

                    has_login = merged["Present"] == 1
                    mat_susp_with_login = is_mat_or_susp & has_login
                    merged.loc[mat_susp_with_login, "Is Scheduled"] = 1
                    merged.loc[mat_susp_with_login, "Present"] = 1
                    merged.loc[mat_susp_with_login, "Absent"] = 0
                    merged.loc[mat_susp_with_login, "Leave"] = 0
                    if "Active/Inactive" in merged.columns:
                        merged.loc[mat_susp_with_login, "Active/Inactive"] = "Active"

                    no_login = merged["Present"] == 0
                    mat_susp_no_login = is_mat_or_susp & no_login
                    merged.loc[mat_susp_no_login, "Is Scheduled"] = 0
                    merged.loc[mat_susp_no_login, "Absent"] = 0
                    merged.loc[mat_susp_no_login, "Leave"] = 0

                # ── DOJ Knack re-check after merge ────────────────────────────
                if "DOJ Knack" in merged.columns:
                    for idx, row in merged.iterrows():
                        doj_val = row.get("DOJ Knack")
                        date_str = row.get("Date", "")
                        doj_date = parse_doj_knack(doj_val)
                        parsed_date = parse_date_str(str(date_str))
                        date_obj = parsed_date.date() if parsed_date else None

                        if doj_date is not None and date_obj is not None and doj_date > date_obj:
                            merged.at[idx, "Shift"] = "Not yet hired"
                            merged.at[idx, "Is Scheduled"] = 0
                            merged.at[idx, "Absent"] = 0
                            merged.at[idx, "Present"] = 0
                            merged.at[idx, "For Review"] = 0
                            merged.at[idx, "Leave"] = 0
                            merged.at[idx, "Late"] = 0
                            merged.at[idx, "Late Minutes"] = 0.0
                            day_col = row.get("Day", "")
                            if day_col and day_col in DAY_COLS and day_col in merged.columns:
                                merged.at[idx, day_col] = "Not yet hired"

                all_merged_records.append(merged)

            if errors:
                for e in errors:
                    st.error(e)

            if not all_merged_records:
                st.warning("No records processed.")
                st.stop()

            # Concatenate all per-date merged dataframes
            merged = pd.concat(all_merged_records, ignore_index=True)

            # ── Override processing (applied once across all dates) ──────────
            if override_file is not None:
                with st.spinner("Applying overrides…"):
                    xl = pd.ExcelFile(override_file)

                    # 1. Schedule Override
                    if "Schedule" in xl.sheet_names:
                        sched_ov = pd.read_excel(override_file, sheet_name="Schedule")
                        required_sched = {"ECN", "Effective Date", "Effective Until"}
                        day_cols_found = [d for d in DAY_COLS if d in sched_ov.columns]
                        if required_sched.issubset(set(sched_ov.columns)) and day_cols_found:
                            sched_ov["ECN"] = pd.to_numeric(sched_ov["ECN"], errors="coerce")
                            sched_ov = sched_ov.dropna(subset=["ECN"])
                            sched_ov["ECN"] = sched_ov["ECN"].astype(int)

                            for _, ov_row in sched_ov.iterrows():
                                ecn = int(ov_row["ECN"])
                                eff_start = parse_override_date(ov_row.get("Effective Date"))
                                eff_end   = parse_override_date(ov_row.get("Effective Until"))

                                if eff_start is None or eff_end is None:
                                    continue

                                date_mask = []
                                for _, mrow in merged.iterrows():
                                    mdate = parse_override_date(str(mrow.get("Date", "")))
                                    match = (
                                        str(mrow.get("EmployeeNumber", "")).strip() == str(ecn) and
                                        mdate is not None and eff_start <= mdate <= eff_end
                                    )
                                    date_mask.append(match)

                                if not any(date_mask):
                                    continue

                                for idx in merged[date_mask].index:
                                    mdate = parse_override_date(str(merged.at[idx, "Date"]))
                                    if mdate is None:
                                        continue
                                    day_name = mdate.strftime("%a")
                                    if day_name in ov_row and day_name in DAY_COLS:
                                        new_shift = normalize_schedule(ov_row[day_name])
                                        merged.at[idx, day_name] = new_shift
                                        merged.at[idx, "Shift"] = new_shift
                                        if new_shift == "Rest Day":
                                            merged.at[idx, "Is Scheduled"] = 0
                                        else:
                                            merged.at[idx, "Is Scheduled"] = 1

                    # 2. Leave Override
                    if "Leave" in xl.sheet_names:
                        leave_ov = pd.read_excel(override_file, sheet_name="Leave")
                        if {"ECN", "Date", "Leave"}.issubset(set(leave_ov.columns)):
                            leave_ov["ECN"] = pd.to_numeric(leave_ov["ECN"], errors="coerce")
                            leave_ov = leave_ov.dropna(subset=["ECN"])
                            leave_ov["ECN"] = leave_ov["ECN"].astype(int)

                            for _, ov_row in leave_ov.iterrows():
                                ecn = int(ov_row["ECN"])
                                ov_date = parse_override_date(ov_row.get("Date"))
                                leave_flag = ov_row.get("Leave", 0)

                                try:
                                    leave_flag = int(float(leave_flag))
                                except (ValueError, TypeError):
                                    leave_flag = 0

                                if ov_date is None or leave_flag != 1:
                                    continue

                                date_mask = []
                                for _, mrow in merged.iterrows():
                                    mdate = parse_override_date(str(mrow.get("Date", "")))
                                    match = (
                                        str(mrow.get("EmployeeNumber", "")).strip() == str(ecn) and
                                        mdate is not None and mdate == ov_date
                                    )
                                    date_mask.append(match)

                                if not any(date_mask):
                                    continue

                                for idx in merged[date_mask].index:
                                    merged.at[idx, "Leave"] = 1
                                    merged.at[idx, "Absent"] = 0
                                    merged.at[idx, "Present"] = 0
                                    merged.at[idx, "For Review"] = 0
                                    if merged.at[idx, "Shift"] == "Not yet hired":
                                        day_name = ov_date.strftime("%a")
                                        if day_name in merged.columns:
                                            merged.at[idx, day_name] = "Leave"
                                        merged.at[idx, "Shift"] = "Leave"

                    # 3. Special Holidays
                    if "Holidays" in xl.sheet_names:
                        hol_ov = pd.read_excel(override_file, sheet_name="Holidays")
                        if {"Project", "Date"}.issubset(set(hol_ov.columns)):
                            for _, ov_row in hol_ov.iterrows():
                                proj = str(ov_row.get("Project", "")).strip().lower()
                                subproc = str(ov_row.get("Sub-Process", "")).strip().lower()
                                hol_date = parse_override_date(ov_row.get("Date"))

                                if hol_date is None or not proj:
                                    continue

                                date_mask = []
                                for _, mrow in merged.iterrows():
                                    mdate = parse_override_date(str(mrow.get("Date", "")))
                                    m_proj = str(mrow.get("Project", "")).strip().lower()
                                    m_subproc = str(mrow.get("Sub-Process", "")).strip().lower()

                                    proj_match = (m_proj == proj)
                                    subproc_match = (subproc == "" or m_subproc == subproc)
                                    date_match = (mdate is not None and mdate == hol_date)

                                    date_mask.append(proj_match and subproc_match and date_match)

                                if not any(date_mask):
                                    continue

                                for idx in merged[date_mask].index:
                                    merged.at[idx, "Is Scheduled"] = 0
                                    merged.at[idx, "Absent"] = 0
                                    merged.at[idx, "Leave"] = 0
                                    if merged.at[idx, "Present"] == 1:
                                        merged.at[idx, "Is Scheduled"] = 1
                                    if merged.at[idx, "Shift"] not in ("Not yet hired", "Leave"):
                                        merged.at[idx, "Shift"] = "Holiday"
                                        day_name = hol_date.strftime("%a")
                                        if day_name in merged.columns:
                                            merged.at[idx, day_name] = "Holiday"

                    # 4. Exemptions
                    if "Exemptions" in xl.sheet_names:
                        exempt_ov = pd.read_excel(override_file, sheet_name="Exemptions")
                        if {"ECN", "Exemption Date", "Effective Until", "Active/Inactive"}.issubset(set(exempt_ov.columns)):
                            exempt_ov["ECN"] = pd.to_numeric(exempt_ov["ECN"], errors="coerce")
                            exempt_ov = exempt_ov.dropna(subset=["ECN"])
                            exempt_ov["ECN"] = exempt_ov["ECN"].astype(int)

                            for _, ov_row in exempt_ov.iterrows():
                                ecn = int(ov_row["ECN"])
                                eff_start = parse_override_date(ov_row.get("Exemption Date"))
                                eff_end   = parse_override_date(ov_row.get("Effective Until"))
                                new_status = str(ov_row.get("Active/Inactive", "")).strip()

                                if eff_start is None or eff_end is None:
                                    continue

                                date_mask = []
                                for _, mrow in merged.iterrows():
                                    mdate = parse_override_date(str(mrow.get("Date", "")))
                                    match = (
                                        str(mrow.get("EmployeeNumber", "")).strip() == str(ecn) and
                                        mdate is not None and eff_start <= mdate <= eff_end
                                    )
                                    date_mask.append(match)

                                if not any(date_mask):
                                    continue

                                for idx in merged[date_mask].index:
                                    merged.at[idx, "Is Scheduled"] = 0
                                    merged.at[idx, "Absent"] = 0
                                    if "Active/Inactive" in merged.columns and new_status:
                                        merged.at[idx, "Active/Inactive"] = new_status

                st.success("✔ Overrides applied!")

            # ── Final consistency rules ────────────────────────────────────────
            unmatched_mask = merged["HeadcountMatch"] == "❌ Unmatched"
            merged.loc[unmatched_mask, "Is Scheduled"] = 0
            merged.loc[unmatched_mask, "Absent"] = 0

            not_scheduled = merged["Is Scheduled"] == 0
            merged.loc[not_scheduled, "Absent"] = 0
            merged.loc[not_scheduled, "Leave"] = 0

            present_but_not_scheduled = (merged["Is Scheduled"] == 0) & (merged["Present"] == 1)
            merged.loc[present_but_not_scheduled, "Is Scheduled"] = 1

            for_review_mask = (merged["For Review"] == 1) & (merged["Is Scheduled"] == 1)
            merged.loc[for_review_mask, "Absent"] = 1

            leave_mask = merged["Leave"] == 1
            merged.loc[leave_mask, "Absent"] = 0

            st.session_state.attendance = merged
            st.session_state.merged_headcount = merged

            # ── Metrics ───────────────────────────────────────────────────────
            total_rows = len(merged)
            present_n  = merged["Present"].sum()
            review_n   = merged["For Review"].sum()
            leave_n    = merged["Leave"].sum()
            absent_n   = merged["Absent"].sum()
            not_hired_n = (merged["Shift"] == "Not yet hired").sum()
            matched_n  = (merged["HeadcountMatch"] == "✅ Matched").sum()
            unmatched_n = (merged["HeadcountMatch"] == "❌ Unmatched").sum()

            st.markdown("---")
            ma1, ma2, ma3, ma4, ma5, ma6, ma7 = st.columns(7)
            for col, val, lbl, cls in [
                (ma1, total_rows, "Total Rows",   ""),
                (ma2, present_n,  "Present",       "matched"),
                (ma3, review_n,   "For Review",    "review"),
                (ma4, leave_n,    "On Leave",      "neutral"),
                (ma5, absent_n,   "Absent",        "unmatched"),
                (ma6, not_hired_n, "Not Yet Hired", "not-hired"),
                (ma7, matched_n,  "HC Matched",    "matched"),
            ]:
                with col:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                        f'<div class="metric-lbl">{lbl}</div></div>',
                        unsafe_allow_html=True
                    )

            if unmatched_n > 0:
                st.markdown(
                    f'<div style="text-align:center;color:#f87171;font-size:0.85rem;margin-top:0.5rem;">'
                    f'⚠️ {unmatched_n:,} rows unmatched with headcount</div>',
                    unsafe_allow_html=True
                )

            st.markdown("")
            disp = (["Date","Name","CSLoginName","EmployeeNumber","OnSiteLocation",
                     "Day","Shift","Is Scheduled","Present","Absent","Leave","For Review",
                     "FirstLogin","LastLogout","Active","Break",
                     "Late","Late Minutes","HireStatus","Position",
                     "Employee", "Project", "Sub-Process", "Supervisor",
                     "Role", "Manager", "DOJ Knack", "Date of Separation",
                     "Billable/Buffer", "Active/Inactive",
                     "HeadcountMatch"] + DAY_COLS)
            disp = [c for c in disp if c in merged.columns]

            ta_all, ta_present, ta_review, ta_leave, ta_absent, ta_not_hired, ta_matched, ta_unmatched = st.tabs(
                ["All", "✅ Present", "🟠 For Review", "🌴 Leave", "❌ Absent", "🔵 Not Yet Hired", "✅ HC Matched", "❌ HC Unmatched"]
            )
            with ta_all:
                st.dataframe(clean_for_display(merged[disp]), width="stretch", height=420)
            with ta_present:
                st.dataframe(clean_for_display(merged[merged["Present"]==1][disp]),
                             width="stretch", height=420)
            with ta_review:
                st.dataframe(clean_for_display(merged[merged["For Review"]==1][disp]),
                             width="stretch", height=420)
            with ta_leave:
                st.dataframe(clean_for_display(merged[merged["Leave"]==1][disp]),
                             width="stretch", height=420)
            with ta_absent:
                st.dataframe(clean_for_display(merged[merged["Absent"]==1][disp]),
                             width="stretch", height=420)
            with ta_not_hired:
                st.dataframe(clean_for_display(merged[merged["Shift"]=="Not yet hired"][disp]),
                             width="stretch", height=420)
            with ta_matched:
                st.dataframe(clean_for_display(merged[merged["HeadcountMatch"] == "✅ Matched"][disp]),
                             width="stretch", height=420)
            with ta_unmatched:
                st.dataframe(clean_for_display(merged[merged["HeadcountMatch"] == "❌ Unmatched"][disp]),
                             width="stretch", height=420)

            st.markdown("---")
            ad1, ad2 = st.columns([2,1])
            with ad1:
                st.download_button("⬇️ Download Final Attendance as Excel",
                    to_excel_bytes(merged), output_filename,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ad2:
                csv_name = output_filename.replace(".xlsx", ".csv").replace(".xls", ".csv")
                st.download_button("⬇️ Download as CSV",
                    merged.to_csv(index=False).encode(),
                    csv_name, "text/csv", use_container_width=True)
