# Project 4: Equity Tracker (RSU/ESPP Automation)

## Problem

Tracking RSU vesting events and ESPP purchase cycles required cross-referencing brokerage statements, HR data, and stock price history across multiple spreadsheets — making it difficult to produce accurate compensation expense reports or answer employee questions quickly.

## Solution

A Python script that consolidates RSU grant/vesting data and ESPP enrollment data, fetches historical stock prices, and generates a unified equity compensation report showing gross income, tax withholding estimates, and cumulative vesting by employee or grant.

**Stack:** Python 3.14.3, pandas, yfinance

## Impact

- Replaced a 3-file manual process with a single automated report
- Reduced time to produce quarterly equity comp summary from ~3 hours to minutes
- Enabled on-demand lookups for HR and payroll teams

## Usage

```bash
python src/equity_report.py --grants data/rsu_grants.xlsx --espp data/espp_enrollments.xlsx --output data/equity_report.xlsx
```

## Structure

```
project4-equity-tracker/
├── data/
│   ├── rsu_grants.xlsx         # Sample RSU grant/vesting data
│   ├── espp_enrollments.xlsx   # Sample ESPP enrollment data
│   └── equity_report.xlsx      # Generated output report
└── src/
    └── equity_report.py        # Main script
```
