# Project 1: SAP GL Report Automation

## Problem

Finance team exported raw SAP GL data monthly and spent 3–4 hours manually reformatting it into a management report — applying filters, pivot tables, number formatting, and color-coded highlights by hand.

## Solution

A Python pipeline that reads the SAP GL export and Chart of Accounts, merges and pivots the data by account category, and outputs a fully formatted 3-sheet Excel report — ready to distribute.

A Streamlit app wraps the pipeline so the report can be generated through a browser UI without touching the command line.

**Stack:** Python, pandas, openpyxl, Streamlit

## How It Works

### Pipeline (`src/process.py` → `src/formatting.py`)

1. **Parse** — Split the `Hierarchy` column in the Chart of Accounts on `" - "` to extract `Category`
2. **Merge** — Left-join GL export to CoA on `GL_Account`
3. **Pivot** — `groupby("Category")["Amount"].sum()`
4. **Total row** — Append numeric sum row labelled `"Total"`
5. **Write Excel** — 3 sheets: Summary Pivot, SAP GL Data, Chart of Accounts
6. **Format** — Auto-fit column widths, accounting number format on `Amount` columns, styled title/header/total rows on Summary Pivot

### Excel Output Formatting

- **Column widths** — auto-fitted (`max_length + 5`, min 10 chars)
- **Number format** — accounting format applied to all `Amount` columns
- **Summary Pivot** — blue title row (merged, bold, centered); light-blue header row; bold double-underline Total row
- **Other sheets** — light-blue header row only

## Inputs

| File | Columns |
|------|---------|
| `sap_export.xlsx` | `GL_Account`, `Amount` |
| `SAP_Chart_of_Accounts.xlsx` | `Account Number`, `Hierarchy` (`Numbering - Category`), `Description` |

## Output

`Processed_JE_Summary_{FY}_{Quarter}.xlsx` — 3 sheets:
- **Summary Pivot** — amounts grouped by account category, styled
- **SAP GL Data** — full merged GL detail with category lookup
- **Chart of Accounts** — parsed CoA reference

## Impact

- Reduced report turnaround from ~4 hours to under 5 minutes
- Eliminated manual formatting errors
- Reproducible and version-controlled output

## Usage

### Streamlit app (recommended)

```bash
cd project1-sap-report
streamlit run app.py
```

Upload both files in the sidebar, select the Quarter and Fiscal Year, and click **Process & Format** to download the formatted Excel report.

### Scripts (standalone)

```bash
# Step 1 — process: merge, pivot, write raw Excel
python src/process.py

# Step 2 — format: apply styles and number formatting
python src/formatting.py
```

> Both scripts resolve file paths relative to `project1-sap-report/` (e.g. `../data/sap_export.xlsx`). Run them from the `src/` directory or adjust paths as needed.

## Structure

```
project1-sap-report/
├── app.py                          # Streamlit portfolio app (Overview, Demo, How It Works)
├── assets/
│   └── sap_data_transform_diagram.html
├── data/
│   ├── sap_export.xlsx             # Sample SAP GL export (anonymized)
│   └── SAP_Chart_of_Accounts.xlsx  # Sample Chart of Accounts
├── output/
│   └── formatted_JE_summary.xlsx   # Generated output
├── src/
│   ├── process.py                  # Data pipeline: merge, pivot, write Excel
│   ├── formatting.py               # Formatting: widths, styles, number formats
│   └── je_summary_app.py           # Standalone Streamlit JE Summary tool
└── requirements.txt
```
