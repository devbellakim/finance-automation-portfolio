# Project 3: Excel to PowerPoint Automation

## Problem

Every reporting cycle, an analyst spent 1–2 hours copying financial tables and charts from Excel into a PowerPoint deck — reformatting fonts, resizing tables, and realigning charts slide by slide. Any last-minute number changes meant repeating the process.

## Solution

A Python script that reads named ranges and charts from an Excel workbook and injects them into a PowerPoint template at defined placeholder positions — preserving formatting and allowing one-command refresh when data changes.

**Stack:** Python 3.14.3, openpyxl, python-pptx

## Impact

- Deck refresh time reduced from ~2 hours to under 1 minute
- Eliminated manual copy-paste errors in presented figures
- Reusable template works for any recurring reporting deck

## Usage

```bash
python src/excel_to_ppt.py --workbook data/financials.xlsx --template data/deck_template.pptx --output data/report_deck.pptx
```

## Structure

```
project3-excel-to-ppt/
├── data/
│   ├── financials.xlsx       # Sample Excel workbook with named ranges
│   ├── deck_template.pptx    # PowerPoint template with placeholders
│   └── report_deck.pptx      # Generated output deck
└── src/
    └── excel_to_ppt.py       # Main script
```
