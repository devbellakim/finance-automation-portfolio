# Finance Automation Portfolio

A collection of real-world finance automation projects demonstrating the migration from manual VBA/Alteryx workflows to scalable, maintainable Python solutions.

## About

This portfolio showcases automation built for finance and accounting teams — replacing repetitive spreadsheet work, Alteryx flows, and VBA macros with Python scripts that are faster, auditable, and easier to maintain.

Each project follows the same structure: a real problem, a Python-based solution, and measurable impact.

## Projects

| # | Project | Description | Tools |
|---|---------|-------------|-------|
| 1 | [SAP Report Automation](./project1-sap-report/) | Transforms raw SAP exports into formatted management reports | pandas, openpyxl |
| 2 | [Lease Automation (ASC 842)](./project2-lease-automation/) | Generates monthly journal entries for operating leases under ASC 842 | pandas, numpy |
| 3 | [Excel to PowerPoint](./project3-excel-to-ppt/) | Pulls live Excel ranges and charts into a PowerPoint deck | python-pptx, openpyxl |
| 4 | [Equity Tracker](./project4-equity-tracker/) | Automates RSU vesting and ESPP purchase reporting | pandas, yfinance |

## Tech Stack

- **Language:** Python 3.14.3
- **Core Libraries:** pandas, openpyxl, python-pptx, numpy
- **Previous Stack (replaced):** VBA macros, Alteryx workflows, manual Excel

## Background

These projects were built to solve real finance team pain points:
- Month-end close reports taking hours of manual formatting
- Lease schedules maintained in fragile spreadsheets
- Slide decks rebuilt from scratch every reporting cycle
- Equity compensation tracked across multiple disconnected files

## Structure

Each project folder contains:
- `README.md` — problem, solution, and impact
- `/data` — sample or anonymized input/output data
- `/src` — Python source code

## Portfolio Site

See [`/portfolio-website`](./portfolio-website/) for the Next.js site that presents these projects publicly.
