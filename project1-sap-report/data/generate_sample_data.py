"""
Generate realistic SAP GL export sample data for Project 1.
Outputs both .xlsx and .csv to the data/ folder.
"""

import random
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

# --- Reference data ---

COMPANY_CODES = ["1000", "2000", "3000", "4000"]

DOCUMENT_TYPES = {
    "SA": "G/L Account Document",
    "KR": "Vendor Invoice",
    "KZ": "Vendor Payment",
    "DR": "Customer Invoice",
    "DZ": "Customer Payment",
    "AB": "Accounting Document",
    "WA": "Goods Issue",
    "WE": "Goods Receipt",
    "RV": "Billing Document Transfer",
    "ZP": "Payment Posting",
}

# GL Account: (account number, account name, typical sign, typical range)
GL_ACCOUNTS = [
    ("100000", "Cash and Cash Equivalents",       1,  (50000,  500000)),
    ("110000", "Accounts Receivable - Trade",      1,  (10000,  200000)),
    ("120000", "Accounts Receivable - Interco",    1,  (5000,   100000)),
    ("150000", "Prepaid Expenses",                 1,  (1000,   30000)),
    ("160000", "Inventory - Finished Goods",       1,  (20000,  300000)),
    ("200000", "Accounts Payable - Trade",        -1,  (5000,   150000)),
    ("210000", "Accounts Payable - Interco",      -1,  (2000,   80000)),
    ("220000", "Accrued Liabilities",             -1,  (3000,   60000)),
    ("230000", "Deferred Revenue",                -1,  (1000,   40000)),
    ("300000", "Common Stock",                    -1,  (100000, 500000)),
    ("310000", "Retained Earnings",               -1,  (50000,  400000)),
    ("400000", "Revenue - Product Sales",         -1,  (10000,  500000)),
    ("410000", "Revenue - Services",              -1,  (5000,   200000)),
    ("420000", "Revenue - Intercompany",          -1,  (2000,   100000)),
    ("500000", "Cost of Goods Sold",               1,  (8000,   400000)),
    ("510000", "Direct Labor",                     1,  (5000,   150000)),
    ("520000", "Manufacturing Overhead",           1,  (3000,   80000)),
    ("600000", "Salaries and Wages",               1,  (20000,  200000)),
    ("610000", "Employee Benefits",                1,  (5000,   50000)),
    ("620000", "Payroll Tax Expense",              1,  (2000,   30000)),
    ("630000", "Travel and Entertainment",         1,  (500,    15000)),
    ("640000", "Office Supplies",                  1,  (200,    5000)),
    ("650000", "Rent Expense",                     1,  (5000,   80000)),
    ("660000", "Utilities Expense",                1,  (1000,   20000)),
    ("670000", "Depreciation Expense",             1,  (2000,   40000)),
    ("680000", "Insurance Expense",                1,  (1000,   25000)),
    ("690000", "Professional Fees",                1,  (2000,   50000)),
    ("700000", "Marketing and Advertising",        1,  (1000,   60000)),
    ("710000", "Software Subscriptions",           1,  (500,    20000)),
    ("720000", "Bank Charges",                     1,  (100,    3000)),
    ("730000", "Interest Expense",                 1,  (500,    15000)),
    ("740000", "Foreign Exchange Loss",            1,  (200,    10000)),
    ("800000", "Income Tax Expense",               1,  (5000,   80000)),
]

COST_CENTERS = [
    "CC1000", "CC1100", "CC1200",   # Finance
    "CC2000", "CC2100", "CC2200",   # Operations
    "CC3000", "CC3100",             # Sales
    "CC4000", "CC4100",             # IT
    "CC5000",                       # HR
    "CC6000", "CC6100",             # Executive
]

VENDORS = [f"V{str(i).zfill(6)}" for i in range(100001, 100051)]

DESCRIPTIONS = [
    "Monthly rent payment - {month}",
    "Vendor invoice - {vendor}",
    "Payroll posting - {month}",
    "Accrual reversal - {month}",
    "Intercompany charge - {cc}",
    "Depreciation run - {month}",
    "Travel expense reimbursement",
    "Software license renewal",
    "Utilities payment - {month}",
    "Insurance premium - Q{q}",
    "Professional services - {vendor}",
    "Marketing campaign spend",
    "Office supplies purchase",
    "Bank fee posting",
    "FX revaluation - {month}",
    "Tax accrual - {month}",
    "Customer payment received",
    "Goods receipt - PO {po}",
    "Goods issue - order {po}",
    "Manual journal entry - {month}",
    "Correction entry - {doc}",
    "Bonus accrual - {month}",
    "Benefits cost allocation",
    "Interest charge - {month}",
    "Inventory adjustment",
]

MONTHS = [
    "Jan 2026", "Feb 2026", "Mar 2026",
    "Oct 2025", "Nov 2025", "Dec 2025",
]

def random_description():
    template = random.choice(DESCRIPTIONS)
    return template.format(
        month=random.choice(MONTHS),
        vendor=random.choice(VENDORS),
        cc=random.choice(COST_CENTERS),
        q=random.randint(1, 4),
        po=f"PO{random.randint(4500000, 4599999)}",
        doc=f"190{random.randint(1000000, 9999999)}",
    )

def random_posting_date():
    start = date(2025, 10, 1)
    end = date(2026, 3, 31)
    delta = (end - start).days
    d = start + timedelta(days=random.randint(0, delta))
    return d.strftime("%Y%m%d")

def generate_rows(n=500):
    rows = []
    doc_number = 1900000001

    for _ in range(n):
        account_tuple = random.choice(GL_ACCOUNTS)
        gl_acct, _, sign, (lo, hi) = account_tuple

        # Occasionally flip sign to represent reversals or adjustments
        if random.random() < 0.15:
            sign *= -1

        amount = round(sign * random.uniform(lo, hi), 2)
        doc_type = random.choice(list(DOCUMENT_TYPES.keys()))

        # Vendors only appear on payable-related doc types
        if doc_type in ("KR", "KZ", "ZP"):
            vendor = random.choice(VENDORS)
        else:
            vendor = ""

        rows.append({
            "Document_Number": str(doc_number),
            "Posting_Date":    random_posting_date(),
            "Document_Type":   doc_type,
            "Company_Code":    random.choice(COMPANY_CODES),
            "GL_Account":      gl_acct,
            "Cost_Center":     random.choice(COST_CENTERS),
            "Amount":          amount,
            "Currency":        "USD",
            "Vendor_ID":       vendor,
            "Description":     random_description(),
        })

        doc_number += random.randint(1, 5)   # SAP doc numbers aren't perfectly sequential

    return rows

def main():
    here = Path(__file__).parent
    rows = generate_rows(500)
    df = pd.DataFrame(rows)

    # Sort by Posting_Date then Document_Number (mirrors a real SAP export)
    df = df.sort_values(["Posting_Date", "Document_Number"]).reset_index(drop=True)

    xlsx_path = here / "sap_export.xlsx"
    csv_path  = here / "sap_export.csv"

    df.to_excel(xlsx_path, index=False, sheet_name="GL_Export")
    df.to_csv(csv_path, index=False)

    print(f"Generated {len(df)} rows")
    print(f"  XLSX -> {xlsx_path}")
    print(f"  CSV  -> {csv_path}")
    print(f"\nDocument types breakdown:\n{df['Document_Type'].value_counts().to_string()}")
    print(f"\nCompany codes breakdown:\n{df['Company_Code'].value_counts().to_string()}")

if __name__ == "__main__":
    main()
