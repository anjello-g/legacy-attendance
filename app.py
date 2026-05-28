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

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}
h1, h2, h3 {
    font-family: 'Space Grotesk', sans-serif;
}

.stApp {
    background: #0f1117;
    color: #e8eaf0;
}

.block-container {
    padding-top: 2rem;
    max-width: 1400px;
}

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

.matched { color: #4ade80; }
.unmatched { color: #f87171; }
.fuzzy { color: #facc15; }

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 99px;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.04em;
}
.badge-exact { background: #14532d; color: #4ade80; }
.badge-fuzzy { background: #422006; color: #facc15; }
.badge-none  { background: #3f1515; color: #f87171; }

div[data-testid="stFileUploader"] {
    background: #1a1d27;
    border: 2px dashed #2a2d3a;
    border-radius: 12px;
    padding: 1rem;
}

div[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
}

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


# ── helpers ──────────────────────────────────────────────────────────────────

def normalize(name: str) -> str:
    """Lowercase, collapse whitespace, strip punctuation."""
    if pd.isna(name):
        return ""
    name = str(name).lower().strip()
    name = re.sub(r"[^a-z\s]", "", name)
    name = re.sub(r"\s+", " ", name)
    return name


def build_staff_lookup(staff_df: pd.DataFrame):
    """Return two dicts keyed by normalised name → row index list."""
    exact: dict[str, list[int]] = {}
    for idx, row in staff_df.iterrows():
        # index by the 'Name' column (full name as stored in staff)
        key = normalize(row.get("Name", ""))
        if key:
            exact.setdefault(key, []).append(idx)
        # also index by first+last concatenation as fallback
        first = normalize(row.get("FirstName", ""))
        last  = normalize(row.get("LastName",  ""))
        if first and last:
            exact.setdefault(f"{first} {last}", []).append(idx)
    return exact


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def match_name(sched_name: str, lookup: dict, staff_df: pd.DataFrame,
               fuzzy_threshold: float = 0.82):
    norm = normalize(sched_name)
    if not norm:
        return None, "no_name"

    # 1) exact match
    if norm in lookup:
        idx = lookup[norm][0]
        return staff_df.loc[idx], "exact"

    # 2) fuzzy match
    best_score, best_key = 0.0, None
    for key in lookup:
        sc = similarity(norm, key)
        if sc > best_score:
            best_score, best_key = sc, key

    if best_score >= fuzzy_threshold and best_key:
        idx = lookup[best_key][0]
        return staff_df.loc[idx], f"fuzzy({best_score:.0%})"

    return None, "unmatched"


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Matched")
    return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────

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

fuzzy_threshold = st.slider(
    "Fuzzy match sensitivity (higher = stricter)",
    min_value=0.70, max_value=1.00, value=0.82, step=0.01,
    help="Names scoring above this threshold are considered fuzzy matches."
)

run = st.button("▶  Run Match", use_container_width=False)

if run:
    if not sched_file or not staff_file:
        st.warning("Please upload **both** files before running.", icon="⚠️")
        st.stop()

    with st.spinner("Reading files…"):
        sched_df = pd.read_excel(sched_file,  sheet_name=0)
        staff_df = pd.read_excel(staff_file, sheet_name=0)

    if "Name" not in sched_df.columns:
        st.error("Schedule file must have a **Name** column.")
        st.stop()

    required = {"Name", "CSLoginName", "EmployeeNumber"}
    missing  = required - set(staff_df.columns)
    if missing:
        st.error(f"Staff file is missing columns: {missing}")
        st.stop()

    with st.spinner("Matching names…"):
        lookup = build_staff_lookup(staff_df)
        records = []
        for _, row in sched_df.iterrows():
            matched_row, match_type = match_name(
                row["Name"], lookup, staff_df, fuzzy_threshold
            )
            out = row.to_dict()
            if matched_row is not None:
                loc = matched_row.get("On Site Location", "")
                out["FullName_Staff"]   = matched_row.get("Name", "")
                out["CSLoginName"]      = matched_row.get("CSLoginName", "")
                out["EmployeeNumber"]   = matched_row.get("EmployeeNumber", "")
                out["OnSiteLocation"]   = loc if (pd.notna(loc) and str(loc).strip()) else "WFH"
                out["MatchType"]        = match_type
            else:
                out["FullName_Staff"]   = ""
                out["CSLoginName"]      = ""
                out["EmployeeNumber"]   = ""
                out["OnSiteLocation"]   = ""
                out["MatchType"]        = match_type
            records.append(out)

    result_df = pd.DataFrame(records)

    # ── metrics ──────────────────────────────────────────────────────────────
    total     = len(result_df)
    exact_n   = result_df["MatchType"].str.startswith("exact").sum()
    fuzzy_n   = result_df["MatchType"].str.startswith("fuzzy").sum()
    unmatched = result_df["MatchType"].isin(["unmatched", "no_name"]).sum()

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    for col, val, lbl, cls in [
        (m1, total,     "Total Rows",        ""),
        (m2, exact_n,   "Exact Matches",     "matched"),
        (m3, fuzzy_n,   "Fuzzy Matches",     "fuzzy"),
        (m4, unmatched, "Unmatched",         "unmatched"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-card"><div class="metric-val {cls}">{val}</div>'
                f'<div class="metric-lbl">{lbl}</div></div>',
                unsafe_allow_html=True
            )

    st.markdown("")

    # ── tabs ─────────────────────────────────────────────────────────────────
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

    # ── download ─────────────────────────────────────────────────────────────
    st.markdown("---")
    dl_col1, dl_col2 = st.columns([2, 1])
    with dl_col1:
        excel_bytes = to_excel_bytes(result_df)
        st.download_button(
            label="⬇️  Download Full Result as Excel",
            data=excel_bytes,
            file_name="schedule_matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    csv_bytes = result_df.to_csv(index=False).encode()
    with dl_col2:
        st.download_button(
            label="⬇️  Download as CSV",
            data=csv_bytes,
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
            and return <strong style="color:#4ade80">CSLoginName</strong> and
            <strong style="color:#4ade80">EmployeeNumber</strong> for each row.
        </div>
    </div>
    """, unsafe_allow_html=True)
