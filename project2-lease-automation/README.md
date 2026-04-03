# Project 2: Lease Automation (ASC 842)

## Problem

Under ASC 842, the accounting team maintained operating lease amortization schedules in a shared Excel file. Each month-end required manually calculating interest expense, amortization of the right-of-use asset, and building journal entry lines — error-prone and time-consuming across a portfolio of 20+ leases.

## Solution

A Python script that reads a lease schedule input file, computes the full amortization table per lease (present value, ROU asset, lease liability, interest/amortization split), and outputs ready-to-upload journal entry rows for each period.

**Stack:** Python 3.14.3, pandas, numpy

## Impact

- Eliminated manual JE preparation for entire lease portfolio
- Reduced month-end close time for lease accounting by ~2 hours
- Output format maps directly to ERP upload template

## Usage

```bash
python src/lease_journal_entries.py --input data/lease_schedule.xlsx --period 2026-03
```

## Structure

```
project2-lease-automation/
├── data/
│   ├── lease_schedule.xlsx       # Sample lease input data
│   └── journal_entries.csv       # Generated JE output
└── src/
    └── lease_journal_entries.py  # Main script
```
