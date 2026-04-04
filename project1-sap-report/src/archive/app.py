"""
SAP Report Automation — Streamlit App
Run: streamlit run src/app.py  (from project1-sap-report/)
"""

import io
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent))
from generate_report import (
    load_data,
    ACCOUNT_CATEGORIES,
    DOCUMENT_TYPE_LABELS,
    build_executive_summary,
    build_cost_center_sheet,
    build_company_code_sheet,
    build_gl_detail_sheet,
    build_variance_flags_sheet,
)

# ---------------------------------------------------------------------------
# Dark fintech CSS  (consistent with Projects 2 / 3 / 4)
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
</style>
"""

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
QUARTERS     = ["Q1", "Q2", "Q3", "Q4"]
FISCAL_YEARS = ["FY24", "FY25", "FY26", "FY27"]

QUARTER_MONTHS = {
    "Q1": (1, 3),
    "Q2": (4, 6),
    "Q3": (7, 9),
    "Q4": (10, 12),
}

# GL account number → human-readable name  (VLOOKUP reference table)
GL_ACCOUNT_MAP = {
    "100000": "Cash and Cash Equivalents",
    "110000": "Accounts Receivable - Trade",
    "120000": "Accounts Receivable - Interco",
    "150000": "Prepaid Expenses",
    "160000": "Inventory - Finished Goods",
    "200000": "Accounts Payable - Trade",
    "210000": "Accounts Payable - Interco",
    "220000": "Accrued Liabilities",
    "230000": "Deferred Revenue",
    "300000": "Common Stock",
    "310000": "Retained Earnings",
    "400000": "Revenue - Product Sales",
    "410000": "Revenue - Services",
    "420000": "Revenue - Intercompany",
    "500000": "Cost of Goods Sold",
    "510000": "Direct Labor",
    "520000": "Manufacturing Overhead",
    "600000": "Salaries and Wages",
    "610000": "Employee Benefits",
    "620000": "Payroll Tax Expense",
    "630000": "Travel and Entertainment",
    "640000": "Office Supplies",
    "650000": "Rent Expense",
    "660000": "Utilities Expense",
    "670000": "Depreciation Expense",
    "680000": "Insurance Expense",
    "690000": "Professional Fees",
    "700000": "Marketing and Advertising",
    "710000": "Software Subscriptions",
    "720000": "Bank Charges",
    "730000": "Interest Expense",
    "740000": "Foreign Exchange Loss",
    "800000": "Income Tax Expense",
}

# Plotly dark theme settings
CHART_BG    = "#12273D"
CHART_PAPER = "#0D1B2A"
AXIS_COLOR  = "#8BA9C8"
GRID_COLOR  = "#1E3A5F"
COLORS      = ["#4FC3F7", "#81C784", "#FFB74D", "#CE93D8",
               "#FF6B6B", "#A8D8EA", "#F7DC6F", "#BB8FCE"]

# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def load_sap_file(uploaded_file) -> pd.DataFrame:
    """Load xlsx or csv upload through the existing load_data pipeline."""
    if uploaded_file.name.endswith(".csv"):
        tmp_df = pd.read_csv(uploaded_file, dtype=str)
        buf = io.BytesIO()
        tmp_df.to_excel(buf, index=False)
        buf.seek(0)
        return load_data(buf)
    return load_data(uploaded_file)


def quarter_date_range(quarter: str, fy: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Convert e.g. Q3 + FY25 → (2025-07-01, 2025-09-30)."""
    year = int("20" + fy[2:])
    m_start, m_end = QUARTER_MONTHS[quarter]
    start = pd.Timestamp(year=year, month=m_start, day=1)
    end   = pd.Timestamp(year=year, month=m_end, day=1) + pd.offsets.MonthEnd(0)
    return start, end


def filter_quarter(df: pd.DataFrame, quarter: str, fy: str) -> pd.DataFrame:
    start, end = quarter_date_range(quarter, fy)
    mask = (df["Posting_Date_dt"] >= start) & (df["Posting_Date_dt"] <= end)
    return df[mask].copy()


# ---------------------------------------------------------------------------
# Processing steps
# ---------------------------------------------------------------------------

def step1_load_validate(uploaded_file) -> pd.DataFrame:
    """Load and enrich the SAP GL export."""
    df = load_sap_file(uploaded_file)
    required = {"Document_Number", "Posting_Date", "GL_Account", "Amount"}
    missing  = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    return df


def step2_clean(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise types, strip whitespace, drop fully empty rows."""
    df = df.copy()
    str_cols = ["Document_Number", "Document_Type", "Company_Code",
                "GL_Account", "Cost_Center", "Currency", "Vendor_ID", "Description"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace("nan", "")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["GL_Account"]).reset_index(drop=True)
    return df


def step3_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot by GL_Account + Account_Category:
    rows = accounts, columns = Total Debits / Credits / Net / Count.
    """
    pivot = (
        df.groupby(["GL_Account", "Account_Category"])["Amount"]
        .agg(
            Total_Debits  = lambda x: x[x > 0].sum(),
            Total_Credits = lambda x: x[x < 0].sum(),
            Net_Amount    = "sum",
            Txn_Count     = "count",
        )
        .round(2)
        .reset_index()
        .sort_values(["Account_Category", "GL_Account"])
    )
    return pivot


def step4_vlookup(df_pivot: pd.DataFrame) -> pd.DataFrame:
    """Map GL account codes to descriptions (replaces manual VLOOKUP)."""
    df_pivot = df_pivot.copy()
    df_pivot.insert(
        1, "GL_Account_Name",
        df_pivot["GL_Account"].map(GL_ACCOUNT_MAP).fillna("Unknown Account"),
    )
    return df_pivot


def step5_build_gl_summary(df_prev_pivot: pd.DataFrame,
                            df_curr_pivot: pd.DataFrame) -> pd.DataFrame:
    """
    Merge previous and current quarter pivots.
    Compute variance = Curr Net - Prev Net.
    """
    merged = df_prev_pivot.merge(
        df_curr_pivot,
        on=["GL_Account", "GL_Account_Name", "Account_Category"],
        suffixes=("_Prev", "_Curr"),
        how="outer",
    ).fillna(0)

    merged["Net_Variance"] = (merged["Net_Amount_Curr"] - merged["Net_Amount_Prev"]).round(2)
    merged["Variance_Pct"] = merged.apply(
        lambda r: round(r["Net_Variance"] / abs(r["Net_Amount_Prev"]) * 100, 1)
        if r["Net_Amount_Prev"] != 0 else 0.0,
        axis=1,
    )
    merged["Txn_Count_Total"] = (
        merged["Txn_Count_Prev"].fillna(0).astype(int)
        + merged["Txn_Count_Curr"].fillna(0).astype(int)
    )
    return merged[[
        "Account_Category", "GL_Account", "GL_Account_Name",
        "Net_Amount_Prev", "Net_Amount_Curr",
        "Net_Variance", "Variance_Pct", "Txn_Count_Total",
    ]]


def build_cost_center_summary(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.groupby("Cost_Center")["Amount"]
        .agg(
            Total_Debits  = lambda x: x[x > 0].sum(),
            Total_Credits = lambda x: x[x < 0].sum(),
            Net_Amount    = "sum",
            Txn_Count     = "count",
        )
        .round(2)
        .reset_index()
        .sort_values("Net_Amount", ascending=False)
    )


def build_email_preview(df_curr: pd.DataFrame,
                        df_prev: pd.DataFrame,
                        curr_q: str, prev_q: str, fy: str) -> pd.DataFrame:
    """
    Executive-level summary table formatted for email confirmation.
    Rows = Account Category; columns = key financials + QoQ change.
    """
    def summarise(df):
        return (
            df.groupby("Account_Category")["Amount"]
            .agg(
                Transactions  = "count",
                Total_Debits  = lambda x: x[x > 0].sum(),
                Total_Credits = lambda x: x[x < 0].sum(),
                Net_Amount    = "sum",
            )
            .round(2)
        )

    curr_s = summarise(df_curr)
    prev_s = summarise(df_prev)

    CAT_ORDER = ["Revenue", "Cost of Goods Sold", "Operating Expenses",
                 "Tax & Other", "Assets", "Liabilities", "Equity"]
    all_cats = [c for c in CAT_ORDER if c in curr_s.index or c in prev_s.index]

    rows = []
    for cat in all_cats:
        curr_net = curr_s.loc[cat, "Net_Amount"] if cat in curr_s.index else 0.0
        prev_net = prev_s.loc[cat, "Net_Amount"] if cat in prev_s.index else 0.0
        txn      = int(curr_s.loc[cat, "Transactions"]) if cat in curr_s.index else 0
        chg      = curr_net - prev_net
        chg_pct  = round(chg / abs(prev_net) * 100, 1) if prev_net != 0 else 0.0
        rows.append({
            "Account Category":    cat,
            f"{prev_q} Net ($)":   round(prev_net, 2),
            f"{curr_q} Net ($)":   round(curr_net, 2),
            "QoQ Change ($)":      round(chg, 2),
            "QoQ Change (%)":      chg_pct,
            "Transactions":        txn,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Plotly chart helpers
# ---------------------------------------------------------------------------

def _base_layout(title=""):
    return dict(
        template="plotly_dark",
        plot_bgcolor=CHART_BG,
        paper_bgcolor=CHART_PAPER,
        font=dict(family="Calibri", color=AXIS_COLOR, size=11),
        title=dict(text=title, font=dict(color="#4FC3F7", size=13)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    font=dict(color=AXIS_COLOR)),
        xaxis=dict(gridcolor=GRID_COLOR, tickfont=dict(color=AXIS_COLOR)),
        yaxis=dict(gridcolor=GRID_COLOR, tickfont=dict(color=AXIS_COLOR),
                   tickprefix="$"),
        margin=dict(l=60, r=20, t=60, b=80),
        height=370,
    )


def chart_variance_by_category(df_summary: pd.DataFrame,
                                prev_q: str, curr_q: str) -> go.Figure:
    cat_grp = (
        df_summary.groupby("Account_Category")[
            ["Net_Amount_Prev", "Net_Amount_Curr", "Net_Variance"]
        ]
        .sum()
        .reset_index()
    )
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=cat_grp["Account_Category"], y=cat_grp["Net_Amount_Prev"],
        name=f"{prev_q} Net", marker_color=COLORS[0], marker_line_width=0,
    ))
    fig.add_trace(go.Bar(
        x=cat_grp["Account_Category"], y=cat_grp["Net_Amount_Curr"],
        name=f"{curr_q} Net", marker_color=COLORS[1], marker_line_width=0,
    ))
    layout = _base_layout(f"Net Amount by Account Category: {prev_q} vs {curr_q}")
    layout["barmode"] = "group"
    fig.update_layout(**layout)
    return fig


def chart_cost_center(df_cc: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_cc["Cost_Center"], y=df_cc["Total_Debits"],
        name="Total Debits", marker_color=COLORS[2], marker_line_width=0,
    ))
    fig.add_trace(go.Bar(
        x=df_cc["Cost_Center"], y=df_cc["Total_Credits"].abs(),
        name="Total Credits", marker_color=COLORS[0], marker_line_width=0,
    ))
    layout = _base_layout("Debits vs Credits by Cost Center")
    layout["barmode"] = "group"
    layout["xaxis"]["tickangle"] = -35
    fig.update_layout(**layout)
    return fig


def chart_email_qoq(df_email: pd.DataFrame, prev_q: str, curr_q: str) -> go.Figure:
    fig = go.Figure()
    colors = [
        "#81C784" if v <= 0 else "#FF6B6B"
        for v in df_email["QoQ Change ($)"]
    ]
    fig.add_trace(go.Bar(
        x=df_email["Account Category"],
        y=df_email["QoQ Change ($)"],
        name="QoQ Change",
        marker_color=colors,
        marker_line_width=0,
        text=df_email["QoQ Change (%)"].apply(lambda v: f"{v:+.1f}%"),
        textposition="outside",
        textfont=dict(color=AXIS_COLOR, size=10),
    ))
    layout = _base_layout(f"Quarter-on-Quarter Change: {prev_q} vs {curr_q} ($)")
    layout["showlegend"] = False
    layout["yaxis"]["tickprefix"] = "$"
    fig.update_layout(**layout)
    return fig


# ---------------------------------------------------------------------------
# Download report bytes (reuses generate_report.py sheet builders)
# ---------------------------------------------------------------------------

def generate_report_bytes(df: pd.DataFrame, threshold: float = 50_000) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    build_executive_summary(wb, df)
    build_cost_center_sheet(wb, df)
    build_company_code_sheet(wb, df)
    build_gl_detail_sheet(wb, df)
    build_variance_flags_sheet(wb, df, threshold)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Variance cell styler
# ---------------------------------------------------------------------------

def color_variance(val):
    if not isinstance(val, (int, float)):
        return ""
    if val < 0:
        return "background-color: #0B2D17; color: #81C784"
    if val > 0:
        return "background-color: #2D0B0B; color: #FF6B6B"
    return ""


# ---------------------------------------------------------------------------
# Full processing pipeline
# ---------------------------------------------------------------------------

def run_processing(uploaded_file, prev_q, curr_q, fy,
                   status_fn, progress_fn) -> dict:

    status_fn("Step 1 / 5 — Load and validate SAP GL export...")
    df_full = step1_load_validate(uploaded_file)
    progress_fn(20)

    status_fn("Step 2 / 5 — Clean and standardize data...")
    df_full = step2_clean(df_full)
    df_prev = filter_quarter(df_full, prev_q, fy)
    df_curr = filter_quarter(df_full, curr_q, fy)
    progress_fn(40)

    status_fn("Step 3 / 5 — Pivot by GL Account + Cost Center...")
    pivot_prev = step3_pivot(df_prev)
    pivot_curr = step3_pivot(df_curr)
    progress_fn(60)

    status_fn("Step 4 / 5 — VLOOKUP: map GL account descriptions...")
    pivot_prev = step4_vlookup(pivot_prev)
    pivot_curr = step4_vlookup(pivot_curr)
    df_gl_summary = step5_build_gl_summary(pivot_prev, pivot_curr)
    progress_fn(80)

    status_fn("Step 5 / 5 — Format for email confirmation layout...")
    df_cc    = build_cost_center_summary(df_curr)
    df_email = build_email_preview(df_curr, df_prev, curr_q, prev_q, fy)
    progress_fn(100)

    return {
        "df_full":      df_full,
        "df_prev":      df_prev,
        "df_curr":      df_curr,
        "df_gl_summary":df_gl_summary,
        "df_cc":        df_cc,
        "df_email":     df_email,
        "prev_q":       prev_q,
        "curr_q":       curr_q,
        "fy":           fy,
    }


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SAP Report Automation",
    page_icon="🗂️",
    layout="wide",
)
st.markdown(DARK_CSS, unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 🗂️ SAP Report Automation")
    st.markdown("---")

    st.markdown("### 1. Upload File")
    uploaded_file = st.file_uploader(
        "SAP GL Export",
        type=["xlsx", "csv"],
        help="Upload sap_export.xlsx (or .csv) with columns: "
             "Document_Number, Posting_Date, GL_Account, Amount, etc.",
    )

    st.markdown("### 2. Period")
    col1, col2 = st.columns(2)
    with col1:
        prev_q = st.selectbox("Prev Quarter", QUARTERS, index=2, key="prev_q")
    with col2:
        curr_q = st.selectbox("Curr Quarter", QUARTERS, index=3, key="curr_q")
    fy = st.selectbox("Fiscal Year", FISCAL_YEARS, index=1)

    st.markdown("### 3. Options")
    threshold = st.number_input(
        "Variance flag threshold ($)",
        min_value=0, max_value=10_000_000,
        value=50_000, step=5_000,
        help="Transactions above this amount are flagged in the download report.",
    )

    st.markdown("---")
    run_clicked = st.button(
        "Run Analysis",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_file,
    )
    if not uploaded_file:
        st.caption("Upload a file to enable Run.")

# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
if run_clicked:
    st.session_state.results = None
    progress_bar = st.progress(0)
    status_box   = st.empty()

    try:
        results = run_processing(
            uploaded_file,
            prev_q, curr_q, fy,
            status_fn   = lambda msg: status_box.info(msg),
            progress_fn = lambda pct: progress_bar.progress(pct),
        )
        status_box.empty()
        progress_bar.empty()

        r = results
        prev_rows = len(r["df_prev"])
        curr_rows = len(r["df_curr"])
        total_rows = len(r["df_full"])

        if prev_rows == 0 and curr_rows == 0:
            st.warning(
                f"No transactions found for {prev_q} or {curr_q} {fy}. "
                f"The file contains {total_rows:,} rows — check your quarter/FY selection."
            )
        else:
            st.session_state.results = results
            st.success(
                f"Done — {curr_rows:,} transactions in {curr_q} {fy} | "
                f"{prev_rows:,} in {prev_q} {fy} | "
                f"{r['df_gl_summary']['GL_Account'].nunique()} GL accounts"
            )
    except Exception as e:
        status_box.empty()
        progress_bar.empty()
        st.error(f"Processing failed: {e}")

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if st.session_state.results:
    r   = st.session_state.results
    pq  = r["prev_q"]
    cq  = r["curr_q"]
    fy_ = r["fy"]

    df_curr      = r["df_curr"]
    df_prev      = r["df_prev"]
    df_gl        = r["df_gl_summary"]
    df_cc        = r["df_cc"]
    df_email     = r["df_email"]

    label = f"{pq} vs {cq} {fy_}"

    # KPI row
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(f"{cq} Transactions",  f"{len(df_curr):,}")
    k2.metric(f"{pq} Transactions",  f"{len(df_prev):,}")
    k3.metric("GL Accounts",         df_gl["GL_Account"].nunique())
    k4.metric("Cost Centers",        df_cc["Cost_Center"].nunique())
    k5.metric(f"{cq} Net Amount",
              f"${df_curr['Amount'].sum():,.0f}")

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([
        f"GL Summary — {label}",
        f"Cost Center — {cq} {fy_}",
        f"Email Preview — {label}",
    ])

    # ---- Tab 1: GL Summary ----------------------------------------------
    with tab1:
        var_cols = ["Net_Amount_Prev", "Net_Amount_Curr",
                    "Net_Variance", "Variance_Pct"]

        col_left, col_right = st.columns([3, 2])

        with col_left:
            st.markdown(f"**{len(df_gl)} GL accounts** across {df_gl['Account_Category'].nunique()} categories")
            styled = (
                df_gl.style
                .map(color_variance, subset=["Net_Variance"])
                .format({
                    "Net_Amount_Prev": "${:,.2f}",
                    "Net_Amount_Curr": "${:,.2f}",
                    "Net_Variance":    "${:,.2f}",
                    "Variance_Pct":    "{:+.1f}%",
                })
            )
            st.dataframe(styled, use_container_width=True, height=420)

        with col_right:
            st.plotly_chart(
                chart_variance_by_category(df_gl, pq, cq),
                use_container_width=True,
            )

        # Category-level roll-up below the main table
        st.markdown("#### Category Roll-up")
        cat_rollup = (
            df_gl.groupby("Account_Category")[
                ["Net_Amount_Prev", "Net_Amount_Curr", "Net_Variance", "Txn_Count_Total"]
            ]
            .sum()
            .round(2)
            .reset_index()
        )
        st.dataframe(
            cat_rollup.style
            .map(color_variance, subset=["Net_Variance"])
            .format({
                "Net_Amount_Prev": "${:,.2f}",
                "Net_Amount_Curr": "${:,.2f}",
                "Net_Variance":    "${:,.2f}",
            }),
            use_container_width=True,
            height=280,
        )

    # ---- Tab 2: Cost Center Breakdown -----------------------------------
    with tab2:
        col_left, col_right = st.columns([2, 3])

        with col_left:
            st.markdown(f"**{len(df_cc)} cost centers** — {cq} {fy_}")
            st.dataframe(
                df_cc.style.format({
                    "Total_Debits":  "${:,.2f}",
                    "Total_Credits": "${:,.2f}",
                    "Net_Amount":    "${:,.2f}",
                }),
                use_container_width=True,
                height=420,
            )

        with col_right:
            st.plotly_chart(
                chart_cost_center(df_cc),
                use_container_width=True,
            )

    # ---- Tab 3: Email Report Preview ------------------------------------
    with tab3:
        # Header block
        start_curr, end_curr = quarter_date_range(cq, fy_)
        st.markdown(
            f"""
**ACME CORPORATION — SAP GL REPORT**
Period: **{cq} {fy_}** ({start_curr.strftime('%b %d, %Y')} — {end_curr.strftime('%b %d, %Y')})
Comparison: vs **{pq} {fy_}**
Generated: {date.today().strftime('%B %d, %Y')}
            """
        )
        st.markdown("---")

        col_left, col_right = st.columns([2, 3])

        with col_left:
            money_cols = [c for c in df_email.columns
                          if "$" in c or "Change ($)" in c]
            pct_cols   = [c for c in df_email.columns if "%" in c]

            styled_email = df_email.style.map(
                color_variance, subset=["QoQ Change ($)", "QoQ Change (%)"]
            ).format(
                {c: "${:,.2f}" for c in money_cols}
            ).format(
                {c: "{:+.1f}%" for c in pct_cols}
            )
            st.dataframe(styled_email, use_container_width=True, height=340)

            # Grand total row
            total_prev = df_email[f"{pq} Net ($)"].sum()
            total_curr = df_email[f"{cq} Net ($)"].sum()
            total_chg  = total_curr - total_prev
            st.markdown(
                f"**Grand Total:** {pq} `${total_prev:,.2f}` → "
                f"{cq} `${total_curr:,.2f}` | "
                f"Change: `${total_chg:+,.2f}`"
            )

        with col_right:
            st.plotly_chart(
                chart_email_qoq(df_email, pq, cq),
                use_container_width=True,
            )

        st.markdown("---")

        # Download button
        file_name = f"sap_report_{pq}vs{cq}_{fy_}.xlsx"
        with st.spinner("Building download..."):
            report_bytes = generate_report_bytes(r["df_curr"], threshold)

        st.download_button(
            label=f"Download {file_name}",
            data=report_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            f"Report covers {cq} {fy_} only | "
            f"5 sheets: Executive Summary, By Cost Center, By Company Code, "
            f"GL Detail, Variance Flags (>${threshold:,.0f})"
        )

# ---------------------------------------------------------------------------
# Empty state
# ---------------------------------------------------------------------------
else:
    st.markdown("## SAP Report Automation")
    st.markdown(
        "Upload a **SAP GL export** in the sidebar, select the reporting period, "
        "then click **Run Analysis**."
    )
    with st.expander("Processing steps explained"):
        st.markdown("""
| Step | What it does |
|------|-------------|
| **1. Load & Validate** | Read xlsx/csv, check required columns, enrich with account categories |
| **2. Clean & Standardize** | Strip whitespace, parse amounts, filter to selected quarters |
| **3. Pivot** | Group by GL Account + Cost Center — compute debits, credits, net, count |
| **4. VLOOKUP** | Map 6-digit GL account codes to human-readable descriptions |
| **5. Email Format** | Roll up by account category with QoQ change for management reporting |
        """)
    with st.expander("Expected file format"):
        st.markdown("""
| Column | Format | Example |
|--------|--------|---------|
| `Document_Number` | SAP doc number | `1900000001` |
| `Posting_Date` | `YYYYMMDD` | `20251031` |
| `Document_Type` | 2-char code | `KR`, `SA` |
| `Company_Code` | 4-digit | `1000` |
| `GL_Account` | 6-digit | `600000` |
| `Cost_Center` | Text | `CC1000` |
| `Amount` | Decimal | `12500.00` |
| `Currency` | ISO | `USD` |
        """)
