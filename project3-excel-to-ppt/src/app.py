"""
Financial Report → PowerPoint Automation — Streamlit App
Run: streamlit run src/app.py  (from project3-excel-to-ppt/)
"""

import io
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from excel_to_ppt import (
    load_excel,
    new_presentation,
    build_cover_slide,
    build_revenue_slide,
    build_capex_slide,
    build_debt_tax_slide,
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
    background-color: #1A4B8B; color: #FFFFFF;
    border: none; border-radius: 6px; font-weight: 600;
    width: 100%;
}
.stDownloadButton > button:hover { background-color: #4FC3F7; color: #0D1B2A; }

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
COLORS       = ["#4FC3F7", "#81C784", "#FFB74D", "#CE93D8"]
CHART_BG     = "#12273D"
CHART_PAPER  = "#0D1B2A"
AXIS_COLOR   = "#8BA9C8"
GRID_COLOR   = "#1E3A5F"

# ---------------------------------------------------------------------------
# Plotly chart builders  (dark theme mirrors the PPT output)
# ---------------------------------------------------------------------------

def _base_layout(title=""):
    return dict(
        template="plotly_dark",
        plot_bgcolor=CHART_BG,
        paper_bgcolor=CHART_PAPER,
        font=dict(family="Calibri", color=AXIS_COLOR, size=11),
        title=dict(text=title, font=dict(color="#4FC3F7", size=13)),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            xanchor="left", x=0,
            font=dict(color=AXIS_COLOR),
        ),
        xaxis=dict(
            gridcolor=GRID_COLOR, linecolor=GRID_COLOR,
            tickfont=dict(color=AXIS_COLOR),
        ),
        yaxis=dict(
            gridcolor=GRID_COLOR, linecolor=GRID_COLOR,
            tickfont=dict(color=AXIS_COLOR),
            tickprefix="$", ticksuffix="M",
        ),
        margin=dict(l=50, r=20, t=60, b=40),
        height=380,
    )


def chart_revenue(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for i, col in enumerate(["Product_A", "Product_B", "Product_C"]):
        fig.add_trace(go.Bar(
            x=df["Quarter"], y=df[col],
            name=col.replace("_", " "),
            marker_color=COLORS[i],
            marker_line_width=0,
        ))
    fig.update_layout(
        barmode="group",
        **_base_layout("Revenue by Product Line ($M)"),
    )
    return fig


def chart_capex(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for i, col in enumerate(["Maintenance_CapEx", "Growth_CapEx"]):
        fig.add_trace(go.Bar(
            x=df["Quarter"], y=df[col],
            name=col.replace("_", " "),
            marker_color=COLORS[i],
            marker_line_width=0,
        ))
    fig.update_layout(
        barmode="stack",
        **_base_layout("CapEx — Maintenance vs Growth ($M)"),
    )
    return fig


def chart_debt_tax(df: pd.DataFrame) -> go.Figure:
    """Line chart with secondary Y for Effective Tax Rate %."""
    fig = go.Figure()

    # Primary axis: Interest + Tax Expense
    for i, col in enumerate(["Interest_Expense", "Tax_Expense"]):
        fig.add_trace(go.Scatter(
            x=df["Quarter"], y=df[col],
            name=col.replace("_", " "),
            mode="lines+markers",
            line=dict(color=COLORS[i], width=2.5),
            marker=dict(size=7),
        ))

    # Secondary axis: Effective Tax Rate
    fig.add_trace(go.Scatter(
        x=df["Quarter"], y=df["Effective_Tax_Rate"],
        name="Effective Tax Rate (%)",
        mode="lines+markers",
        line=dict(color=COLORS[2], width=2, dash="dot"),
        marker=dict(size=7),
        yaxis="y2",
    ))

    layout = _base_layout("Interest Expense, Tax Expense & Effective Rate")
    layout["yaxis2"] = dict(
        overlaying="y", side="right",
        tickfont=dict(color=AXIS_COLOR),
        ticksuffix="%",
        gridcolor=GRID_COLOR,
        linecolor=GRID_COLOR,
        showgrid=False,
    )
    layout["yaxis"]["ticksuffix"] = "M"
    fig.update_layout(**layout)
    return fig


# ---------------------------------------------------------------------------
# PPTX builder → bytes
# ---------------------------------------------------------------------------
def build_pptx_bytes(sheets: dict, company: str, period_label: str) -> bytes:
    prs = new_presentation()
    build_cover_slide(prs, company, period_label, "Quarterly Financial Review — Management Pack")
    build_revenue_slide(prs, sheets["Revenue_Trend"], period_label)
    build_capex_slide(prs, sheets["CapEx"], period_label)
    build_debt_tax_slide(prs, sheets["Debt_and_Tax"], period_label)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Financial Report → PowerPoint",
    page_icon="📊",
    layout="wide",
)
st.markdown(DARK_CSS, unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📊 Excel → PowerPoint Automation")
    st.markdown("---")

    st.markdown("### 1. Upload File")
    uploaded_file = st.file_uploader(
        "Financial Data (Excel)",
        type=["xlsx"],
        help="Upload financial_report.xlsx with Revenue_Trend, CapEx, Debt_and_Tax sheets.",
    )

    st.markdown("### 2. Period")
    curr_q = st.selectbox("Quarter", QUARTERS, index=3)
    fy     = st.selectbox("Fiscal Year", FISCAL_YEARS, index=1)

    st.markdown("### 3. Deck Settings")
    company = st.text_input("Company Name", value="Acme Corporation")

    st.markdown("---")
    run_clicked = st.button(
        "Generate Report",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_file,
    )
    if not uploaded_file:
        st.caption("Upload a file to enable Generate.")

# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
if run_clicked:
    st.session_state.results = None
    progress_bar = st.progress(0)
    status_box   = st.empty()

    try:
        status_box.info("Step 1 / 3 — Loading Excel sheets...")
        sheets = load_excel(uploaded_file)
        progress_bar.progress(33)

        status_box.info("Step 2 / 3 — Building charts...")
        fig_rev  = chart_revenue(sheets["Revenue_Trend"])
        fig_cap  = chart_capex(sheets["CapEx"])
        fig_debt = chart_debt_tax(sheets["Debt_and_Tax"])
        progress_bar.progress(66)

        status_box.info("Step 3 / 3 — Generating PowerPoint...")
        period_label = f"{curr_q} {fy} Financial Review"
        pptx_bytes   = build_pptx_bytes(sheets, company, period_label)
        progress_bar.progress(100)

        status_box.empty()
        progress_bar.empty()

        st.session_state.results = {
            "sheets":       sheets,
            "fig_rev":      fig_rev,
            "fig_cap":      fig_cap,
            "fig_debt":     fig_debt,
            "pptx_bytes":   pptx_bytes,
            "curr_q":       curr_q,
            "fy":           fy,
            "company":      company,
            "period_label": period_label,
        }
        st.success("Presentation built — preview below, then download.")
    except Exception as e:
        status_box.empty()
        progress_bar.empty()
        st.error(f"Processing failed: {e}")

# ---------------------------------------------------------------------------
# Results display
# ---------------------------------------------------------------------------
if st.session_state.results:
    r  = st.session_state.results
    cq = r["curr_q"]
    fy_ = r["fy"]
    sheets = r["sheets"]

    # KPI row
    rev_total  = sheets["Revenue_Trend"]["Total_Revenue"].iloc[-1]
    cap_total  = sheets["CapEx"]["Total_CapEx"].iloc[-1]
    debt_total = sheets["Debt_and_Tax"]["Total_Debt"].iloc[-1]
    tax_rate   = sheets["Debt_and_Tax"]["Effective_Tax_Rate"].iloc[-1]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Revenue (latest Q)",  f"${rev_total:.1f}M")
    k2.metric("Total CapEx (latest Q)", f"${cap_total:.1f}M")
    k3.metric("Total Debt (latest Q)",  f"${debt_total:.0f}M")
    k4.metric("Effective Tax Rate",  f"{tax_rate:.1f}%")

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([
        f"{cq} {fy_} — Revenue Trend",
        f"{cq} {fy_} — CapEx",
        f"{cq} {fy_} — Debt & Tax",
    ])

    def _show_tab(df, fig, money_cols):
        col_a, col_b = st.columns([2, 3])
        with col_a:
            st.dataframe(
                df.style.format(
                    {c: "${:,.1f}" for c in money_cols}
                ),
                use_container_width=True,
                height=340,
            )
        with col_b:
            st.plotly_chart(fig, use_container_width=True)

    with tab1:
        _show_tab(
            sheets["Revenue_Trend"],
            r["fig_rev"],
            ["Product_A", "Product_B", "Product_C", "Total_Revenue"],
        )

    with tab2:
        _show_tab(
            sheets["CapEx"],
            r["fig_cap"],
            ["Maintenance_CapEx", "Growth_CapEx", "Total_CapEx"],
        )

    with tab3:
        _show_tab(
            sheets["Debt_and_Tax"],
            r["fig_debt"],
            ["Interest_Expense", "Total_Debt", "Tax_Expense"],
        )

    st.markdown("---")
    file_name = f"financial_report_{cq}_{fy_}.pptx"
    st.download_button(
        label=f"Download {file_name}",
        data=r["pptx_bytes"],
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

else:
    st.markdown("## Financial Report → PowerPoint Automation")
    st.markdown(
        "Upload your **financial_report.xlsx** in the sidebar, select the period "
        "and company name, then click **Generate Report**."
    )
    with st.expander("Expected Excel structure"):
        st.markdown("""
| Sheet | Key Columns |
|-------|-------------|
| `Revenue_Trend` | Quarter, Product_A, Product_B, Product_C, Total_Revenue |
| `CapEx` | Quarter, Maintenance_CapEx, Growth_CapEx, Total_CapEx |
| `Debt_and_Tax` | Quarter, Interest_Expense, Total_Debt, Tax_Expense, Effective_Tax_Rate |
        """)
