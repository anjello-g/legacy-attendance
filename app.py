import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher

st.set_page_config(
    page_title="Schedule → Staff Matcher",
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
</style>
""", unsafe_allow_html=True)


# ── constants ─────────────────────────────────────────────────────────────────
EXCLUDED_LOCATIONS = {"clark", "dsi", "zamboanga", "isabela"}
DAY_COLS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
FUZZY_THRESHOLD = 0.82


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
        df.to_excel(writer, index=False, sheet_name="Matched")
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("## 🔗 Schedule → Staff Matcher")
st.markdown('<p style="color:#7c7f8e;margin-top:-0.5rem;">Upload both files, run the match, download the enriched result.</p>', unsafe_allow_html=True)

st.markdown("---")

col_l, col_r = st.columns(2)

with col_l:
    st.markdown('<div class="section-header">📅 Schedule File</div>', unsafe_allow_html=True)
    sched_file = st.file_uploader("Upload Schedule .xlsx", type=["xlsx", "xls"],
                                  key="sched", label_visibility="collapsed")

with col_r:
    st.markdown('<div class="section-header">👥 Staff Roster File</div>', unsafe_allow_html=True)
    staff_file = st.file_uploader("Upload Staff Roster .xlsx", type=["xlsx", "xls"],
                                  key="staff", label_visibility="collapsed")

st.markdown("")
run = st.button("▶  Run Match", use_container_width=False)

if run:
    if not sched_file or not staff_file:
        st.warning("Please upload **both** files before running.", icon="⚠️")
        st.stop()

    with st.spinner("Reading files…"):
        sched_df = pd.read_excel(sched_file, sheet_name=0)
        staff_df = pd.read_excel(staff_file, sheet_name=0)

    if "Name" not in sched_df.columns:
        st.error("Schedule file must have a **Name** column.")
        st.stop()

    missing = {"Name", "CSLoginName", "EmployeeNumber"} - set(staff_df.columns)
    if missing:
        st.error(f"Staff file is missing columns: {missing}")
        st.stop()

    with st.spinner("Matching names…"):
        lookup = build_staff_lookup(staff_df)
        records = []
        for _, row in sched_df.iterrows():
            matched_row, match_type = match_name(row["Name"], lookup, staff_df)
            out = row.to_dict()

            # Normalize day columns → Rest Day where blank/null/0000-0000
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

    # ── exclude specified locations ───────────────────────────────────────────
    before_filter = len(result_df)
    result_df = result_df[
        ~result_df["OnSiteLocation"].str.strip().str.lower().isin(EXCLUDED_LOCATIONS)
    ]
    excluded_count = before_filter - len(result_df)

    # ── metrics ───────────────────────────────────────────────────────────────
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

    # ── tabs ──────────────────────────────────────────────────────────────────
    tab_all, tab_exact, tab_fuzzy, tab_unmatched = st.tabs(
        ["All Results", "✅ Exact", "🟡 Fuzzy", "❌ Unmatched"]
    )

    display_cols = ["Name", "FullName_Staff", "CSLoginName", "EmployeeNumber",
                    "OnSiteLocation", "HireStatus", "Position", "MatchType"]
    display_cols = [c for c in display_cols if c in result_df.columns]

    def show_table(df_sub):
        st.dataframe(df_sub[display_cols], use_container_width=True, height=420)

    with tab_all:
        show_table(result_df)
    with tab_exact:
        show_table(result_df[result_df["MatchType"].str.startswith("exact")])
    with tab_fuzzy:
        show_table(result_df[result_df["MatchType"].str.startswith("fuzzy")])
    with tab_unmatched:
        show_table(result_df[result_df["MatchType"].isin(["unmatched", "no_name"])])

    # ── download ──────────────────────────────────────────────────────────────
    st.markdown("---")
    dl1, dl2 = st.columns([2, 1])
    with dl1:
        st.download_button(
            label="⬇️  Download Full Result as Excel",
            data=to_excel_bytes(result_df),
            file_name="schedule_matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            label="⬇️  Download as CSV",
            data=result_df.to_csv(index=False).encode(),
            file_name="schedule_matched.csv",
            mime="text/csv",
            use_container_width=True,
        )

else:
    st.markdown("""
    <div style="background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
                padding:2rem;text-align:center;color:#7c7f8e;margin-top:1rem;">
        <div style="font-size:2.5rem;margin-bottom:0.5rem;">📋</div>
        <div style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;color:#e8eaf0;margin-bottom:0.4rem;">
            Upload both files to get started
        </div>
        <div style="font-size:0.85rem;">
            The app will match names from the Schedule to the Staff roster<br>
            and return <strong style="color:#4ade80">CSLoginName</strong>,
            <strong style="color:#4ade80">EmployeeNumber</strong>, and
            <strong style="color:#4ade80">OnSiteLocation</strong> for each row.<br><br>
            Rows where location is <em>Clark, DSI, Zamboanga, or Isabela</em> are excluded.<br>
            Schedule slots that are blank or <em>0000 - 0000</em> are marked as <strong style="color:#a78bfa">Rest Day</strong>.
        </div>
    </div>
    """, unsafe_allow_html=True)
