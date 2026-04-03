"""
RSU/ESPP Equity Report Automation — Streamlit App
Run: streamlit run src/app.py  (from project4-equity-tracker/)
"""

import io
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent))
from equity_processor import (
    step1_load_raw,
    step2_rename_columns,
    step3_forward_fill,
    step4_join_reference,
    step5_split_rsu_espp,
    step6_summarize,
    _build_transactions_sheet,
    _build_employee_summary_sheet,
    _build_dept_summary_sheet,
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
.badge-rsu  { background:#0B2D17; color:#81C784; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
.badge-espp { background:#0B1D3A; color:#4FC3F7; padding:2px 8px; border-radius:4px; font-size:0.8rem; font-weight:600; }
</style>
"""

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
QUARTERS     = ["Q1", "Q2", "Q3", "Q4"]
FISCAL_YEARS = ["FY24", "FY25", "FY26", "FY27"]
COLORS       = ["#4FC3F7", "#81C784", "#FFB74D", "#CE93D8"]
CHART_BG     = "#12273D"
CHART_PAPER  = "#0D1B2A"
AXIS_COLOR   = "#8BA9C8"
GRID_COLOR   = "#1E3A5F"

# ---------------------------------------------------------------------------
# Excel bytes builder (multi-sheet report)
# ---------------------------------------------------------------------------

def build_report_bytes(df_flat: pd.DataFrame,
                       df_emp:  pd.DataFrame,
                       df_dept: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    _build_transactions_sheet(wb, df_flat)
    _build_employee_summary_sheet(wb, df_emp)
    _build_dept_summary_sheet(wb, df_dept)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

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
        margin=dict(l=60, r=20, t=60, b=40),
        height=350,
    )


def chart_dept_totals(df_dept: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_dept["Department"], y=df_dept["RSU_Total_Value"],
        name="RSU Total Value", marker_color=COLORS[0], marker_line_width=0,
    ))
    fig.add_trace(go.Bar(
        x=df_dept["Department"], y=df_dept["ESPP_Total_Value"],
        name="ESPP Total Value", marker_color=COLORS[1], marker_line_width=0,
    ))
    fig.update_layout(barmode="stack", **_base_layout("Equity Value by Department ($)"))
    return fig


def chart_emp_top10(df_emp: pd.DataFrame) -> go.Figure:
    top = df_emp.nlargest(10, "Combined_Total_Value")[
        ["Employee_Name", "RSU_Total_Value", "ESPP_Total_Value"]
    ]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=top["Employee_Name"], x=top["RSU_Total_Value"],
        name="RSU", orientation="h", marker_color=COLORS[0], marker_line_width=0,
    ))
    fig.add_trace(go.Bar(
        y=top["Employee_Name"], x=top["ESPP_Total_Value"],
        name="ESPP", orientation="h", marker_color=COLORS[1], marker_line_width=0,
    ))
    layout = _base_layout("Top 10 Employees by Combined Equity Value")
    layout["xaxis"]["tickprefix"] = "$"
    layout["yaxis"] = dict(gridcolor=GRID_COLOR, tickfont=dict(color=AXIS_COLOR, size=9))
    layout["barmode"] = "stack"
    fig.update_layout(**layout)
    return fig


# ---------------------------------------------------------------------------
# Processing pipeline
# ---------------------------------------------------------------------------

def run_processing(raw_file, ref_file) -> dict:
    """
    Wraps the 6 equity_processor steps.
    ref_file may be None (skip join step).
    """
    # Step 1: load
    df = step1_load_raw(raw_file)
    # Step 2: rename
    df = step2_rename_columns(df)
    # Step 3: forward fill
    df = step3_forward_fill(df)
    # Step 4: join reference (if provided)
    if ref_file is not None:
        df = step4_join_reference(df, ref_file)
    # Step 5: split
    df_rsu, df_espp = step5_split_rsu_espp(df)
    # Step 6: summarize
    df_emp, df_dept = step6_summarize(df)

    return {
        "df_flat": df,
        "df_rsu":  df_rsu,
        "df_espp": df_espp,
        "df_emp":  df_emp,
        "df_dept": df_dept,
    }

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Equity Report Automation",
    page_icon="📈",
    layout="wide",
)
st.markdown(DARK_CSS, unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📈 RSU/ESPP Equity Automation")
    st.markdown("---")

    st.markdown("### 1. Upload Files")
    raw_file = st.file_uploader(
        "Fidelity Transaction Report",
        type=["xlsx", "csv"],
        help="Fidelity-style equity export (sparse employee columns)",
    )
    ref_file = st.file_uploader(
        "Employee Reference Table *(optional)*",
        type=["xlsx"],
        help="HR reference with Manager, Location columns",
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
        disabled=not raw_file,
    )
    if not raw_file:
        st.caption("Upload the Fidelity report to enable Run.")

# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
if run_clicked:
    st.session_state.results = None
    progress_bar = st.progress(0)
    status_box   = st.empty()

    steps = [
        "Load & validate Fidelity export",
        "Forward-fill blank employee fields",
        "Join employee reference table",
        "Separate RSU vs ESPP transactions",
        "Build employee & department summaries",
    ]
    try:
        for i, step in enumerate(steps):
            status_box.info(f"Step {i+1} / {len(steps)} — {step}...")
            if i == 0:
                results_data = run_processing(raw_file, ref_file)
            progress_bar.progress(int((i + 1) / len(steps) * 100))

        status_box.empty()
        progress_bar.empty()

        r = results_data
        st.session_state.results = {
            **r,
            "curr_q": curr_q,
            "prev_q": prev_q,
            "fy":     fy,
        }
        st.success(
            f"Done — {len(r['df_flat']):,} transactions | "
            f"{r['df_flat']['Employee_ID'].nunique()} employees | "
            f"RSU: {len(r['df_rsu'])} | ESPP: {len(r['df_espp'])}"
        )
    except Exception as e:
        status_box.empty()
        progress_bar.empty()
        st.error(f"Processing failed: {e}")

# ---------------------------------------------------------------------------
# Results display
# ---------------------------------------------------------------------------
if st.session_state.results:
    r   = st.session_state.results
    cq  = r["curr_q"]
    fy_ = r["fy"]

    df_flat = r["df_flat"]
    df_emp  = r["df_emp"]
    df_dept = r["df_dept"]
    df_rsu  = r["df_rsu"]
    df_espp = r["df_espp"]

    # KPI row
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Employees",         df_flat["Employee_ID"].nunique())
    k2.metric("RSU Transactions",  len(df_rsu))
    k3.metric("ESPP Transactions", len(df_espp))
    k4.metric("RSU Total Value",   f"${df_emp['RSU_Total_Value'].sum():,.0f}")
    k5.metric("ESPP Total Value",  f"${df_emp['ESPP_Total_Value'].sum():,.0f}")

    st.markdown("---")

    MONEY_COLS = [c for c in df_flat.columns
                  if any(k in c for k in ["Value","Withheld","Price","Shares"])]
    EMP_MONEY  = [c for c in df_emp.columns
                  if any(k in c for k in ["Value","Withheld"])]
    DEPT_MONEY = [c for c in df_dept.columns
                  if any(k in c for k in ["Value","Net","Withheld"])]

    tab1, tab2, tab3, tab4 = st.tabs([
        f"{cq} {fy_} — All Transactions",
        f"{cq} {fy_} — RSU Summary",
        f"{cq} {fy_} — ESPP Summary",
        f"{cq} {fy_} — Department Summary",
    ])

    # ---- Tab 1: All Transactions -----------------------------------------
    with tab1:
        display_cols = [c for c in df_flat.columns if c not in
                        ("Manager_ID", "Employment_Status")]
        st.markdown(f"**{len(df_flat):,} transactions** after forward-fill and join")
        st.dataframe(
            df_flat[display_cols].style.format(
                {c: "{:,.2f}" for c in df_flat[display_cols].select_dtypes("number").columns}
            ),
            use_container_width=True,
            height=420,
        )

    # ---- Tab 2: RSU Summary ----------------------------------------------
    with tab2:
        rsu_emp_cols = ["Employee_ID", "Employee_Name", "Department",
                        "Company_Code", "RSU_Transactions", "RSU_Shares",
                        "RSU_Total_Value", "RSU_Tax_Withheld", "RSU_Net_Value"]
        rsu_disp = df_emp[[c for c in rsu_emp_cols if c in df_emp.columns]]

        col_a, col_b = st.columns([2, 3])
        with col_a:
            st.markdown(f"**{len(rsu_disp)} employees** with RSU vesting")
            rsu_money = [c for c in ["RSU_Total_Value","RSU_Tax_Withheld","RSU_Net_Value"]
                         if c in rsu_disp.columns]
            st.dataframe(
                rsu_disp.style.format(
                    {c: "${:,.2f}" for c in rsu_money}
                ),
                use_container_width=True,
                height=380,
            )
        with col_b:
            st.plotly_chart(chart_emp_top10(df_emp), use_container_width=True)

    # ---- Tab 3: ESPP Summary ---------------------------------------------
    with tab3:
        espp_emp_cols = ["Employee_ID", "Employee_Name", "Department",
                         "Company_Code", "ESPP_Transactions", "ESPP_Shares",
                         "ESPP_Total_Value", "ESPP_Net_Value"]
        espp_disp = df_emp[[c for c in espp_emp_cols if c in df_emp.columns]]
        espp_participants = espp_disp[espp_disp["ESPP_Transactions"] > 0]

        espp_money = [c for c in ["ESPP_Total_Value","ESPP_Net_Value"]
                      if c in espp_disp.columns]

        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.markdown(f"**{len(espp_participants)} employees** participated in ESPP")
            st.dataframe(
                espp_participants.style.format(
                    {c: "${:,.2f}" for c in espp_money}
                ),
                use_container_width=True,
                height=380,
            )
        with col_b:
            # ESPP share count by department
            espp_dept = (
                df_espp.groupby("Department")["Total_Value"]
                .sum()
                .reset_index()
                .sort_values("Total_Value", ascending=False)
            )
            fig_espp = go.Figure(go.Bar(
                x=espp_dept["Department"],
                y=espp_dept["Total_Value"],
                marker_color=COLORS[1],
                marker_line_width=0,
            ))
            layout = _base_layout("ESPP Purchase Value by Department ($)")
            fig_espp.update_layout(**layout)
            st.plotly_chart(fig_espp, use_container_width=True)

    # ---- Tab 4: Department Summary ----------------------------------------
    with tab4:
        col_a, col_b = st.columns([1, 1])
        with col_a:
            dept_money = [c for c in DEPT_MONEY if c in df_dept.columns]
            st.dataframe(
                df_dept.style.format(
                    {c: "${:,.2f}" for c in dept_money}
                ),
                use_container_width=True,
                height=300,
            )
        with col_b:
            st.plotly_chart(chart_dept_totals(df_dept), use_container_width=True)

    # ---- Download button -------------------------------------------------
    st.markdown("---")
    file_name = f"equity_report_{cq}_{fy_}.xlsx"
    with st.spinner("Building download..."):
        report_bytes = build_report_bytes(df_flat, df_emp, df_dept)
    st.download_button(
        label=f"Download {file_name}",
        data=report_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.markdown("## RSU / ESPP Equity Report Automation")
    st.markdown(
        "Upload the **Fidelity Transaction Report** (and optionally the "
        "**Employee Reference Table**) in the sidebar, select the period, "
        "then click **Run Analysis**."
    )
    with st.expander("What this app does — Alteryx workflow replaced"):
        st.markdown("""
| Step | Alteryx Tool | What it does |
|------|-------------|--------------|
| 1 | Input Data | Load Fidelity export, skip header rows |
| 2 | Select | Rename columns to snake_case |
| 3 | Multi-Row Formula | Forward-fill sparse Employee_ID / Name / Dept / Co. |
| 4 | Join | VLOOKUP against employee reference (Manager, Location) |
| 5 | Filter | Split into RSU and ESPP transaction streams |
| 6 | Summarize | GroupBy employee and department — sum values |
        """)
