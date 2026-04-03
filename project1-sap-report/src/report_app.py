"""
SAP GL Report Generator — Streamlit UI
=======================================
Interactive web app for generating the SAP GL management report.
Users can upload their SAP export, filter by fiscal year / quarter /
date range, set a variance threshold, and download the formatted report.

Run:
    streamlit run src/report_app.py
"""

import io
from pathlib import Path
import pandas as pd
import streamlit as st
from generate_report import (
    load_data,
    build_executive_summary,
    build_cost_center_sheet,
    build_company_code_sheet,
    build_gl_detail_sheet,
    build_variance_flags_sheet,
)
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SAP GL Report Generator",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
FISCAL_QUARTERS = {
    "Q1 (Jan – Mar)": [1, 2, 3],
    "Q2 (Apr – Jun)": [4, 5, 6],
    "Q3 (Jul – Sep)": [7, 8, 9],
    "Q4 (Oct – Dec)": [10, 11, 12],
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def filter_dataframe(
    df: pd.DataFrame,
    date_from: pd.Timestamp,
    date_to: pd.Timestamp,
    company_codes: list[str],
    doc_types: list[str],
) -> pd.DataFrame:
    mask = (
        (df["Posting_Date_dt"] >= date_from) &
        (df["Posting_Date_dt"] <= date_to)
    )
    if company_codes:
        mask &= df["Company_Code"].isin(company_codes)
    if doc_types:
        mask &= df["Document_Type"].isin(doc_types)
    return df[mask].copy()


def build_report_bytes(df: pd.DataFrame, threshold: float) -> bytes:
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
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Step 1 — Load data into session_state BEFORE the sidebar renders
# This ensures filter options are available on the very first pass after upload.
# ---------------------------------------------------------------------------
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
if "load_error" not in st.session_state:
    st.session_state.load_error = None

# File uploader and sample checkbox live outside the sidebar block so we can
# read them here and populate session_state before the sidebar filter widgets run.
st.title("📊 SAP GL Report Generator")
st.caption("Upload your SAP GL export, configure filters, and download a formatted management report.")
st.divider()

with st.sidebar:
    st.header("⚙️ Report Settings")
    st.subheader("1. Data File")

    uploaded_file = st.file_uploader(
        "Upload SAP GL export",
        type=["xlsx", "csv"],
        help="Accepted formats: .xlsx or .csv with the standard SAP GL column layout.",
    )
    use_sample = st.checkbox(
        "Use built-in sample data",
        value=False,
        help="Load the sample sap_export.xlsx from the data/ folder.",
    )

# Load / reload data whenever the source changes
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            tmp_df = pd.read_csv(uploaded_file, dtype=str)
            buf = io.BytesIO()
            tmp_df.to_excel(buf, index=False)
            buf.seek(0)
            st.session_state.df_raw = load_data(buf)
        else:
            st.session_state.df_raw = load_data(uploaded_file)
        st.session_state.load_error = None
    except Exception as e:
        st.session_state.df_raw = None
        st.session_state.load_error = str(e)

elif use_sample:
    sample_path = Path(__file__).parent.parent / "data" / "sap_export.xlsx"
    if sample_path.exists():
        st.session_state.df_raw = load_data(sample_path)
        st.session_state.load_error = None
    else:
        st.session_state.df_raw = None
        st.session_state.load_error = "Sample file not found. Run `data/generate_sample_data.py` first."

elif not use_sample and uploaded_file is None:
    st.session_state.df_raw = None
    st.session_state.load_error = None

df_raw = st.session_state.df_raw

# ---------------------------------------------------------------------------
# Step 2 — Sidebar: period, filters, threshold (all rendered in one pass)
# Filter options are now populated from df_raw which is already in session_state.
# ---------------------------------------------------------------------------
with st.sidebar:
    st.divider()

    # --- Fiscal year & quarter ---
    st.subheader("2. Fiscal Period")
    current_year = pd.Timestamp.today().year
    fiscal_year = st.selectbox(
        "Fiscal Year",
        options=list(range(current_year - 3, current_year + 2)),
        index=3,
    )

    quarter_options = list(FISCAL_QUARTERS.keys())
    selected_quarters = st.multiselect(
        "Quarter(s)",
        options=quarter_options,
        default=quarter_options,
        help="Select one or more quarters to include.",
    )

    selected_months: list[int] = []
    for q in selected_quarters:
        selected_months.extend(FISCAL_QUARTERS[q])
    selected_months = sorted(set(selected_months))

    fy_start   = pd.Timestamp(year=fiscal_year, month=selected_months[0] if selected_months else 1, day=1)
    last_month = selected_months[-1] if selected_months else 12
    fy_end     = pd.Timestamp(year=fiscal_year, month=last_month, day=1) + pd.offsets.MonthEnd(0)

    st.subheader("3. Custom Date Range")
    st.caption("Overrides the fiscal quarter selection above if changed.")
    date_from = st.date_input("From date", value=fy_start.date())
    date_to   = st.date_input("To date",   value=fy_end.date())

    if date_from > date_to:
        st.error("'From date' must be before 'To date'.")

    st.divider()

    # --- Filters — options populated from loaded data, or empty with hint ---
    st.subheader("4. Filters")

    if df_raw is not None:
        all_codes     = sorted(df_raw["Company_Code"].dropna().unique().tolist())
        all_doc_types = sorted(df_raw["Document_Type"].dropna().unique().tolist())

        company_code_filter = st.multiselect(
            "Company Code(s)",
            options=all_codes,
            default=all_codes,
            key="cc_filter",
        )
        doc_type_filter = st.multiselect(
            "Document Type(s)",
            options=all_doc_types,
            default=all_doc_types,
            key="dt_filter",
        )
    else:
        st.info("Upload a file to enable filters.", icon="ℹ️")
        company_code_filter = []
        doc_type_filter     = []

    st.divider()

    # --- Variance threshold ---
    st.subheader("5. Variance Threshold")
    threshold = st.number_input(
        "Flag transactions above ($)",
        min_value=0,
        max_value=10_000_000,
        value=50_000,
        step=5_000,
        help="Transactions with |Amount| ≥ this value appear in the Variance Flags sheet.",
    )

    st.divider()

    # --- Output filename ---
    st.subheader("6. Output")
    output_filename = st.text_input(
        "Report filename",
        value=f"management_report_FY{fiscal_year}.xlsx",
    )

# ---------------------------------------------------------------------------
# Step 3 — Main area: status banner → preview → generate
# ---------------------------------------------------------------------------

if st.session_state.load_error:
    st.error(st.session_state.load_error)

elif df_raw is not None:
    if uploaded_file is not None:
        st.success(f"Loaded **{uploaded_file.name}** — {len(df_raw):,} rows")
    else:
        st.info(f"Using built-in sample data — {len(df_raw):,} rows")

    # Apply all filters
    df_filtered = filter_dataframe(
        df_raw,
        date_from     = pd.Timestamp(date_from),
        date_to       = pd.Timestamp(date_to),
        company_codes = company_code_filter,
        doc_types     = doc_type_filter,
    )

    # Metric cards
    st.subheader("Data Preview")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Rows",       f"{len(df_filtered):,}")
    col2.metric("Company Codes",    df_filtered["Company_Code"].nunique())
    col3.metric("GL Accounts",      df_filtered["GL_Account"].nunique())
    col4.metric("Total Net Amount", f"${df_filtered['Amount'].sum():,.0f}")
    col5.metric("Variance Flags",   f"{(df_filtered['Abs_Amount'] >= threshold).sum():,}")

    st.divider()

    # Preview tabs
    tab1, tab2, tab3 = st.tabs(["Transaction Sample", "By Account Category", "By Month"])

    with tab1:
        st.dataframe(
            df_filtered[[
                "Document_Number", "Posting_Date", "Document_Type",
                "Company_Code", "GL_Account", "Account_Category",
                "Cost_Center", "Amount", "Currency", "Description",
            ]].head(100).style.format({"Amount": "{:,.2f}"}),
            use_container_width=True,
            height=360,
        )
        if len(df_filtered) > 100:
            st.caption(f"Showing first 100 of {len(df_filtered):,} rows.")

    with tab2:
        cat_summary = (
            df_filtered.groupby("Account_Category")["Amount"]
            .agg(
                Transactions  = "count",
                Total_Debits  = lambda x: x[x > 0].sum(),
                Total_Credits = lambda x: x[x < 0].sum(),
                Net_Amount    = "sum",
            )
            .round(2)
            .reset_index()
            .sort_values("Net_Amount", ascending=False)
        )
        st.dataframe(
            cat_summary.style.format({
                "Total_Debits":  "{:,.2f}",
                "Total_Credits": "{:,.2f}",
                "Net_Amount":    "{:,.2f}",
            }),
            use_container_width=True,
        )

    with tab3:
        monthly = (
            df_filtered.groupby("Month")["Amount"]
            .agg(
                Transactions  = "count",
                Total_Debits  = lambda x: x[x > 0].sum(),
                Total_Credits = lambda x: x[x < 0].sum(),
                Net_Amount    = "sum",
            )
            .round(2)
            .reset_index()
        )
        st.dataframe(
            monthly.style.format({
                "Total_Debits":  "{:,.2f}",
                "Total_Credits": "{:,.2f}",
                "Net_Amount":    "{:,.2f}",
            }),
            use_container_width=True,
        )

    st.divider()

    # Generate & download
    if len(df_filtered) == 0:
        st.warning("No data matches the selected filters. Adjust your settings and try again.")
    else:
        st.subheader("Generate Report")
        if st.button("Generate Excel Report", type="primary", use_container_width=True):
            with st.spinner("Building report..."):
                report_bytes = build_report_bytes(df_filtered, threshold)
            st.success("Report ready!")
            st.download_button(
                label        = "⬇️ Download Report",
                data         = report_bytes,
                file_name    = output_filename,
                mime         = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption(
                f"Report contains {len(df_filtered):,} transactions | "
                f"Variance flags: {(df_filtered['Abs_Amount'] >= threshold).sum():,} | "
                f"Threshold: ${threshold:,.0f}"
            )

else:
    st.info("Upload a SAP GL export file or check **Use built-in sample data** in the sidebar to get started.")
    with st.expander("Expected file format"):
        st.markdown("""
| Column | Format | Example |
|---|---|---|
| `Document_Number` | SAP doc number | `1900000001` |
| `Posting_Date` | `YYYYMMDD` | `20260131` |
| `Document_Type` | 2-char code | `KR`, `SA`, `DR` |
| `Company_Code` | 4-digit | `1000` |
| `GL_Account` | 6-digit | `600000` |
| `Cost_Center` | Text | `CC1000` |
| `Amount` | Decimal (+ debit / − credit) | `12500.00` |
| `Currency` | ISO code | `USD` |
| `Vendor_ID` | Text | `V100001` |
| `Description` | Text | `Monthly rent payment` |
        """)
