# Project 1: SAP Report Automation

## Problem

Finance team exported raw SAP GL data monthly and spent 3–4 hours manually reformatting it into a management report — applying filters, pivot tables, number formatting, and color-coded variance highlights by hand.

## Solution

A Python script that reads the SAP Excel export, applies business logic (cost center groupings, variance thresholds), and outputs a fully formatted Excel report with conditional formatting and summary tabs — ready to distribute.

**Stack:** Python 3.14.3, pandas, openpyxl

## Impact

- Reduced report turnaround from ~4 hours to under 5 minutes
- Eliminated manual formatting errors
- Report is now reproducible and version-controlled

## Usage

```bash
python src/generate_report.py --input data/sap_export.xlsx --output data/management_report.xlsx
```

## Structure

```
project1-sap-report/
├── data/
│   ├── sap_export.xlsx        # Sample SAP GL export (anonymized)
│   └── management_report.xlsx # Generated output
└── src/
    └── generate_report.py    # Main script
```
