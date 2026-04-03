"""
ASC 842 Lease Automation — Streamlit App
Run: streamlit run src/app.py  (from project2-lease-automation/)
"""

import io
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent))
from lease_journal_entries import (
    load_harbor,
    je_amortization, je_interest, je_payment,
    je_new_lease, je_termination,
    JE_COLS, build_report, GL,
    _je_counter,
)

# ---------------------------------------------------------------------------
# Dark fintech CSS
# ---------------------------------------------------------------------------
DARK_CSS = """
<style>
#MainMenu, footer, header { visibility: hidden; }

.stApp { background-color: #0D1B2A; color: #E0E8F0; }

[data-testid="stSidebar"] {
    background-color: #0F2035;
    border-right: 1px solid #1E3A5F;
}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div { color: #C8D8E8 !important; }

[data-testid="stFileUploader"] {
    background-color: #12273D;
    border: 1px dashed #2E5FA3;
    border-radius: 8px;
    padding: 6px;
}

.stTabs [data-baseweb="tab-list"] {
    background-color: #12273D;
    border-radius: 8px 8px 0 0;
    gap: 2px;
    padding: 4px 4px 0 4px;
}
.stTabs [data-baseweb="tab"] {
    background-color: #0D1B2A;
    color: #8BA9C8;
    border-radius: 6px 6px 0 0;
    padding: 8px 14px;
}
.stTabs [aria-selected="true"] {
    background-color: #1E3A5F !important;
    color: #4FC3F7 !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background-color: #0D1B2A;
    padding-top: 12px;
}

[data-testid="stMetricValue"] { color: #4FC3F7 !important; font-size: 1.3rem !important; }
[data-testid="stMetricLabel"] { color: #8BA9C8 !important; }

.stButton > button {
    background-color: #2E5FA3; color: #FFFFFF;
    border: none; border-radius: 6px; font-weight: 600;
}
.stButton > button:hover { background-color: #4FC3F7; color: #0D1B2A; }

.stDownloadButton > button {
    background-color: #1A6B3A; color: #FFFFFF;
    border: none; border-radius: 6px; font-weight: 600;
    width: 100%;
}
.stDownloadButton > button:hover { background-color: #2A9D5A; }

h1 { color: #4FC3F7 !important; }
h2, h3 { color: #8BA9C8 !important; }
hr { border-color: #1E3A5F; margin: 8px 0; }
.step-card {
    background: #12273D; border-left: 3px solid #2E5FA3;
    border-radius: 4px; padding: 8px 12px; margin: 4px 0;
    font-size: 0.9rem; color: #C8D8E8;
}
.step-done { border-left-color: #2A9D5A; color: #81C784; }
.step-active { border-left-color: #4FC3F7; color: #4FC3F7; }
</style>
"""

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
QUARTERS = ["Q1", "Q2", "Q3", "Q4"]
FISCAL_YEARS = ["FY24", "FY25", "FY26", "FY27"]
ANNUAL_RATE = 0.05   # incremental borrowing rate

FINANCIAL_COLS = [
    "ROU_Asset_Cost", "Lease_Liability_Balance",
    "Remaining_Cash_Balance", "ROU_Asset_NBV",
    "Accumulated_Amortization",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def color_variance(val):
    if not isinstance(val, (int, float)):
        return ""
    if val < 0:
        return "background-color: #0B2D17; color: #81C784"
    if val > 0:
        return "background-color: #2D0B0B; color: #FF6B6B"
    return ""


def compute_variance(df_prev: pd.DataFrame, df_curr: pd.DataFrame) -> pd.DataFrame:
    """Merge prev and curr quarters; compute variance for each financial column."""
    merged = df_prev.merge(
        df_curr,
        on="Capital_Lease_ID",
        suffixes=("_prev", "_curr"),
        how="inner",
    )
    meta = ["Capital_Lease_ID", "Portfolio_curr", "Currency_curr", "Company_Code_curr"]
    result = merged[meta].rename(columns={
        "Portfolio_curr": "Portfolio",
        "Currency_curr": "Currency",
        "Company_Code_curr": "Company_Code",
    })
    for col in FINANCIAL_COLS:
        result[f"{col}_Prev"] = merged[f"{col}_prev"]
        result[f"{col}_Curr"] = merged[f"{col}_curr"]
        result[f"{col}_Var"]  = (merged[f"{col}_curr"] - merged[f"{col}_prev"]).round(2)
    return result


def generate_je_bytes(all_lines, summary, period) -> bytes:
    """Run build_report to a temp file and return as bytes."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        build_report(all_lines, summary, period, tmp_path)
        return tmp_path.read_bytes()
    finally:
        tmp_path.unlink(missing_ok=True)


def run_processing(prev_file, curr_file, prev_q, curr_q, fy) -> dict:
    """Full processing pipeline mirroring lease_journal_entries.py main()."""
    period  = f"{fy}-{curr_q}"
    je_date = {"Q1": "03-31", "Q2": "06-30", "Q3": "09-30", "Q4": "12-31"}[curr_q]
    year    = int("20" + fy[2:])
    je_date = f"{year}-{je_date}"

    df_prev = load_harbor(prev_file)
    df_curr = load_harbor(curr_file)

    prev_ids = set(df_prev["Capital_Lease_ID"])
    curr_ids = set(df_curr["Capital_Lease_ID"])

    continuing_ids = prev_ids & curr_ids
    terminated_ids = prev_ids - curr_ids
    new_ids        = curr_ids - prev_ids

    prev_idx = {lid: row for lid, row in df_prev.set_index("Capital_Lease_ID").iterrows()}
    curr_idx = {lid: row for lid, row in df_curr.set_index("Capital_Lease_ID").iterrows()}
    for lid, row in prev_idx.items():
        row["Capital_Lease_ID"] = lid
    for lid, row in curr_idx.items():
        row["Capital_Lease_ID"] = lid

    # Reset JE counter each run
    _je_counter[0] = 0
    all_lines = []

    for lid in sorted(continuing_ids):
        q_prev = prev_idx[lid]
        q_curr = curr_idx[lid]
        amort  = round(q_curr["Accumulated_Amortization"] - q_prev["Accumulated_Amortization"], 2)
        if amort > 0:
            all_lines += je_amortization(q_curr, amort, period, je_date)
        all_lines += je_interest(q_curr, q_prev["Lease_Liability_Balance"],
                                 ANNUAL_RATE, period, je_date)
        all_lines += je_payment(q_curr, q_prev["Lease_Liability_Balance"],
                                q_curr["Lease_Liability_Balance"],
                                ANNUAL_RATE, period, je_date)

    for lid in sorted(terminated_ids):
        all_lines += je_termination(prev_idx[lid], period, je_date)

    for lid in sorted(new_ids):
        all_lines += je_new_lease(curr_idx[lid], period, je_date)

    df_je = pd.DataFrame(all_lines, columns=JE_COLS)

    # Summaries for build_report
    by_type = {}
    for et in df_je["Entry_Type"].unique():
        sub = df_je[df_je["Entry_Type"] == et]
        by_type[et] = {
            "count":   len(sub),
            "debits":  round(sub["Debit"].sum(), 2),
            "credits": round(sub["Credit"].sum(), 2),
            "net":     round(sub["Debit"].sum() - sub["Credit"].sum(), 2),
        }
    summary = {
        "total_leases": len(continuing_ids) + len(terminated_ids) + len(new_ids),
        "continuing":   len(continuing_ids),
        "terminated":   len(terminated_ids),
        "new":          len(new_ids),
        "total_lines":  len(df_je),
        "grand_debits": round(df_je["Debit"].sum(), 2),
        "grand_credits":round(df_je["Credit"].sum(), 2),
        "by_type":      by_type,
    }

    return {
        "df_prev":      df_prev,
        "df_curr":      df_curr,
        "df_variance":  compute_variance(df_prev, df_curr),
        "df_new":       df_curr[df_curr["Capital_Lease_ID"].isin(new_ids)].reset_index(drop=True),
        "df_terminated":df_prev[df_prev["Capital_Lease_ID"].isin(terminated_ids)].reset_index(drop=True),
        "df_je":        df_je,
        "summary":      summary,
        "period":       period,
        "prev_q":       prev_q,
        "curr_q":       curr_q,
        "fy":           fy,
    }


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="ASC 842 Lease Automation",
    page_icon="📋",
    layout="wide",
)
st.markdown(DARK_CSS, unsafe_allow_html=True)

# Session state
if "results" not in st.session_state:
    st.session_state.results = None

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📋 ASC 842 Lease Automation")
    st.markdown("---")

    st.markdown("### 1. Upload Files")
    prev_file = st.file_uploader(
        "Previous Quarter Report",
        type=["xlsx"],
        help="Lease Harbor export for the prior quarter (e.g. Q3)",
    )
    curr_file = st.file_uploader(
        "Current Quarter Report",
        type=["xlsx"],
        help="Lease Harbor export for the current quarter (e.g. Q4)",
    )

    st.markdown("### 2. Period")
    col1, col2 = st.columns(2)
    with col1:
        prev_q = st.selectbox("Prev Quarter", QUARTERS, index=2, key="prev_q")
    with col2:
        curr_q = st.selectbox("Curr Quarter", QUARTERS, index=3, key="curr_q")
    fy = st.selectbox("Fiscal Year", FISCAL_YEARS, index=1)

    st.markdown("---")
    run_clicked = st.button(
        "Run Analysis",
        type="primary",
        use_container_width=True,
        disabled=not (prev_file and curr_file),
    )

    if not (prev_file and curr_file):
        st.caption("Upload both files to enable Run.")

# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
if run_clicked:
    st.session_state.results = None
    steps = [
        "Load and validate both reports",
        "Match leases by Capital_Lease_ID",
        "Calculate variances",
        "Identify new & terminated leases",
        "Generate Journal Entries",
    ]
    progress_bar = st.progress(0)
    status_box   = st.empty()

    try:
        for i, step in enumerate(steps):
            status_box.markdown(
                f'<div class="step-card step-active">&#9654; {step}...</div>',
                unsafe_allow_html=True,
            )
            if i == 0:
                results = run_processing(
                    prev_file, curr_file, prev_q, curr_q, fy
                )
            progress_bar.progress(int((i + 1) / len(steps) * 100))

        status_box.empty()
        progress_bar.empty()
        st.session_state.results = results
        st.success(
            f"Done — {results['summary']['total_lines']:,} JE lines generated "
            f"| {results['summary']['total_leases']} leases processed"
        )
    except Exception as e:
        status_box.empty()
        progress_bar.empty()
        st.error(f"Processing failed: {e}")

# ---------------------------------------------------------------------------
# Results display
# ---------------------------------------------------------------------------
if st.session_state.results:
    r    = st.session_state.results
    pq   = r["prev_q"]
    cq   = r["curr_q"]
    fy_  = r["fy"]
    summ = r["summary"]

    label = f"{pq} vs {cq} {fy_}"

    # KPI bar
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Leases Processed", summ["total_leases"])
    k2.metric("Continuing",       summ["continuing"])
    k3.metric("New Leases",       summ["new"])
    k4.metric("Terminated",       summ["terminated"])
    k5.metric("JE Lines",         summ["total_lines"])

    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs([
        f"Variance Summary — {label}",
        f"New Leases — {cq} {fy_}",
        f"Terminated Leases — {pq} {fy_}",
        "Journal Entry",
    ])

    # ---- Tab 1: Variance Summary ----------------------------------------
    with tab1:
        df_var = r["df_variance"]
        var_cols = [c for c in df_var.columns if c.endswith("_Var")]

        styled = df_var.style.map(color_variance, subset=var_cols).format(
            {c: "{:,.2f}" for c in df_var.select_dtypes("number").columns}
        )
        st.dataframe(styled, use_container_width=True, height=420)

    # ---- Tab 2: New Leases -----------------------------------------------
    with tab2:
        df_new = r["df_new"]
        if df_new.empty:
            st.info(f"No new leases commenced in {cq} {fy_}.")
        else:
            st.markdown(f"**{len(df_new)} new leases** commenced in {cq} {fy_}")
            st.dataframe(
                df_new.style.format(
                    {c: "{:,.2f}" for c in df_new.select_dtypes("number").columns}
                ),
                use_container_width=True,
                height=400,
            )

    # ---- Tab 3: Terminated Leases ----------------------------------------
    with tab3:
        df_term = r["df_terminated"]
        if df_term.empty:
            st.info(f"No leases terminated during {cq} {fy_}.")
        else:
            st.markdown(f"**{len(df_term)} leases** terminated during {cq} {fy_}")
            st.dataframe(
                df_term.style.format(
                    {c: "{:,.2f}" for c in df_term.select_dtypes("number").columns}
                ),
                use_container_width=True,
                height=400,
            )

    # ---- Tab 4: Journal Entry --------------------------------------------
    with tab4:
        df_je = r["df_je"]

        # Entry type breakdown
        et_counts = df_je.groupby("Entry_Type").agg(
            Lines  = ("JE_ID",  "count"),
            Debits = ("Debit",  "sum"),
            Credits= ("Credit", "sum"),
        ).round(2)

        c_left, c_right = st.columns([1, 2])
        with c_left:
            st.dataframe(et_counts.style.format("{:,.2f}", subset=["Debits","Credits"]),
                         use_container_width=True)
            balanced = abs(df_je["Debit"].sum() - df_je["Credit"].sum()) < 0.05
            if balanced:
                st.success("Debits = Credits — BALANCED")
            else:
                st.error("OUT OF BALANCE — review JE lines")
        with c_right:
            st.dataframe(
                df_je.style.format(
                    {c: "{:,.2f}" for c in ["Debit","Credit"]}
                ),
                use_container_width=True,
                height=380,
            )

        st.markdown("---")
        file_name = f"lease_JE_{pq}vs{cq}_{fy_}.xlsx"
        with st.spinner("Building download..."):
            je_bytes = generate_je_bytes(
                r["df_je"].to_dict("records"),
                summ,
                r["period"],
            )
        st.download_button(
            label=f"Download {file_name}",
            data=je_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.markdown("## ASC 842 Lease Automation")
    st.markdown(
        "Upload the **Previous Quarter** and **Current Quarter** Lease Harbor reports "
        "in the sidebar, select the reporting period, then click **Run Analysis**."
    )
    with st.expander("What this app does"):
        st.markdown("""
| Step | Description |
|------|-------------|
| Match | Aligns leases across quarters by `Capital_Lease_ID` |
| Variance | Calculates balance changes for all financial columns |
| New Leases | Flags leases appearing only in the current quarter |
| Terminated | Flags leases appearing only in the previous quarter |
| Journal Entry | Generates SAP-ready JE lines (Amortization, Interest, Payment, Initial Recognition, Termination) |
        """)
