"""
Generate realistic Lease Harbor quarterly reports for ASC 842 (Q3 and Q4 2025).

Two output files:
  lease_harbor_Q3.xlsx  — balances as of Sep 30, 2025
  lease_harbor_Q4.xlsx  — balances as of Dec 31, 2025 (same leases, updated figures)

Q3 → Q4 dynamics:
  - ~85 leases carry through both quarters (updated amortization / liability)
  - ~10 leases terminate between Q3 and Q4 (absent from Q4)
  - ~15 new leases commence in Q4 (absent from Q3)
"""

import random
import math
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
from pathlib import Path

random.seed(7)

# ---------------------------------------------------------------------------
# Reference data
# ---------------------------------------------------------------------------

PORTFOLIOS     = ["Region A", "Region B", "Region C"]
CURRENCIES     = ["USD", "EUR", "GBP"]
CURRENCY_WEIGHTS = [0.60, 0.25, 0.15]          # USD-heavy, realistic mix
COMPANY_CODES  = ["1000", "2000", "3000", "4000"]

Q3_END = date(2025, 9, 30)
Q4_END = date(2025, 12, 31)

# Lease types drive size ranges (in USD)
LEASE_TYPES = {
    "Office":      (200_000, 3_000_000),
    "Warehouse":   (150_000, 1_500_000),
    "Retail":      (80_000,  800_000),
    "Equipment":   (20_000,  300_000),
    "Vehicle":     (10_000,  80_000),
    "Data Center": (500_000, 5_000_000),
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def next_lease_id(counter: list) -> str:
    counter[0] += 1
    return f"CL-{counter[0]:05d}"


def random_commencement(earliest: date, latest: date) -> date:
    delta = (latest - earliest).days
    d = earliest + timedelta(days=random.randint(0, delta))
    # Snap to first of the month — leases typically start on the 1st
    return d.replace(day=1)


def random_term_months() -> int:
    """Return a realistic lease term: most 24–84 months, some short/long."""
    buckets = [
        (12, 24,  0.10),   # short-term (< 2 yr)
        (24, 60,  0.40),   # 2–5 yr
        (60, 84,  0.30),   # 5–7 yr
        (84, 120, 0.15),   # 7–10 yr
        (120, 180, 0.05),  # 10–15 yr
    ]
    r = random.random()
    cumulative = 0.0
    for lo, hi, prob in buckets:
        cumulative += prob
        if r <= cumulative:
            return random.randint(lo, hi)
    return 60


def months_between(d1: date, d2: date) -> float:
    """Approximate months between two dates."""
    return (d2.year - d1.year) * 12 + (d2.month - d1.month) + (d2.day - d1.day) / 30


def quarterly_amortization(rou_cost: float, total_months: float) -> float:
    """Straight-line quarterly amortization amount."""
    if total_months <= 0:
        return 0.0
    return (rou_cost / total_months) * 3          # 3 months per quarter


def liability_quarterly_reduction(liability_start: float, total_months: float) -> float:
    """Approximate quarterly principal reduction (simplified annuity)."""
    if total_months <= 0:
        return 0.0
    # Use a flat principal-reduction proxy; real ASC 842 uses effective interest
    return liability_start / (total_months / 3)


def add_noise(value: float, pct: float = 0.02) -> float:
    """Add ±pct random noise (simulates rounding, FX, adjustments)."""
    return round(value * (1 + random.uniform(-pct, pct)), 2)


# ---------------------------------------------------------------------------
# Lease factory
# ---------------------------------------------------------------------------

def build_lease(lease_id: str, commencement: date, term_months: int,
                lease_type: str, as_of: date) -> dict | None:
    """
    Build a single lease record valued as of `as_of`.
    Returns None if the lease is not active as of that date.
    """
    termination = commencement + relativedelta(months=term_months)

    # Only include if active as of the report date
    if commencement > as_of or termination <= as_of:
        return None

    total_months   = months_between(commencement, termination)
    elapsed_months = months_between(commencement, as_of)
    elapsed_months = max(0.0, min(elapsed_months, total_months))

    lo, hi      = LEASE_TYPES[lease_type]
    rou_cost    = round(random.uniform(lo, hi), -2)   # round to nearest 100

    accum_amort = round((rou_cost / total_months) * elapsed_months, 2) if total_months > 0 else 0.0
    rou_nbv     = round(rou_cost - accum_amort, 2)

    # Lease liability: starts equal to ROU cost, reduces with payments
    initial_liability  = round(rou_cost * random.uniform(0.92, 1.00), 2)  # slight variance
    liability_balance  = round(initial_liability * (1 - elapsed_months / total_months), 2)
    liability_balance  = max(0.0, liability_balance)

    # Remaining cash = future undiscounted payments proxy
    remaining_payments = max(0, total_months - elapsed_months)
    monthly_payment    = round(initial_liability / total_months, 2) if total_months > 0 else 0.0
    remaining_cash     = round(monthly_payment * remaining_payments, 2)

    currency = random.choices(CURRENCIES, weights=CURRENCY_WEIGHTS, k=1)[0]
    # FX conversion factor (approximate, for realism)
    fx = {"USD": 1.0, "EUR": 1.08, "GBP": 1.27}[currency]
    fx_adj = lambda v: round(v / fx, 2)

    return {
        "Capital_Lease_ID":        lease_id,
        "Commencement_Date":       commencement.strftime("%Y-%m-%d"),
        "Termination_Date":        termination.strftime("%Y-%m-%d"),
        "ROU_Asset_Cost":          fx_adj(rou_cost),
        "Lease_Liability_Balance": fx_adj(add_noise(liability_balance)),
        "Remaining_Cash_Balance":  fx_adj(add_noise(remaining_cash)),
        "ROU_Asset_NBV":           fx_adj(add_noise(rou_nbv)),
        "Accumulated_Amortization":fx_adj(add_noise(accum_amort)),
        "File_ID":                 f"FILE-{random.randint(10000, 99999)}",
        "Portfolio":               random.choice(PORTFOLIOS),
        "Currency":                currency,
        "Company_Code":            random.choice(COMPANY_CODES),
        "_lease_type":             lease_type,     # internal — dropped before export
        "_term_months":            term_months,
        "_commencement":           commencement,
        "_termination":            termination,
        "_rou_cost_usd":           rou_cost,
        "_initial_liability_usd":  initial_liability,
        "_monthly_payment_usd":    monthly_payment,
    }


def update_lease_for_q4(q3_row: dict) -> dict:
    """
    Advance a Q3 lease record by one quarter (3 months) to produce the Q4 balance.
    Preserves all metadata; only updates the financial columns.
    """
    row = q3_row.copy()

    commencement  = q3_row["_commencement"]
    termination   = q3_row["_termination"]
    rou_cost_usd  = q3_row["_rou_cost_usd"]
    init_liab_usd = q3_row["_initial_liability_usd"]
    mthly_pay_usd = q3_row["_monthly_payment_usd"]
    term_months   = q3_row["_term_months"]

    elapsed_q4     = months_between(commencement, Q4_END)
    elapsed_q4     = max(0.0, min(elapsed_q4, term_months))
    remaining_q4   = max(0, term_months - elapsed_q4)

    accum_amort_q4 = round((rou_cost_usd / term_months) * elapsed_q4, 2) if term_months > 0 else 0.0
    rou_nbv_q4     = round(rou_cost_usd - accum_amort_q4, 2)
    liability_q4   = round(init_liab_usd * (1 - elapsed_q4 / term_months), 2) if term_months > 0 else 0.0
    liability_q4   = max(0.0, liability_q4)
    remaining_cash = round(mthly_pay_usd * remaining_q4, 2)

    currency = q3_row["Currency"]
    fx = {"USD": 1.0, "EUR": 1.08, "GBP": 1.27}[currency]
    fx_adj = lambda v: round(v / fx, 2)

    row["ROU_Asset_Cost"]           = fx_adj(rou_cost_usd)
    row["Accumulated_Amortization"] = fx_adj(add_noise(accum_amort_q4))
    row["ROU_Asset_NBV"]            = fx_adj(add_noise(rou_nbv_q4))
    row["Lease_Liability_Balance"]  = fx_adj(add_noise(liability_q4))
    row["Remaining_Cash_Balance"]   = fx_adj(add_noise(remaining_cash))

    return row


# ---------------------------------------------------------------------------
# Main generation logic
# ---------------------------------------------------------------------------

EXPORT_COLS = [
    "Capital_Lease_ID", "Commencement_Date", "Termination_Date",
    "ROU_Asset_Cost", "Lease_Liability_Balance", "Remaining_Cash_Balance",
    "ROU_Asset_NBV", "Accumulated_Amortization",
    "File_ID", "Portfolio", "Currency", "Company_Code",
]

INTERNAL_COLS = [c for c in [
    "_lease_type", "_term_months", "_commencement",
    "_termination", "_rou_cost_usd", "_initial_liability_usd", "_monthly_payment_usd"
]]


def main():
    here = Path(__file__).parent
    counter = [0]

    lease_types = list(LEASE_TYPES.keys())

    # ------------------------------------------------------------------
    # Pool A: ~85 leases active in BOTH Q3 and Q4
    # Commencement before Q3, termination after Q4
    # ------------------------------------------------------------------
    pool_a_leases = []
    attempts = 0
    while len(pool_a_leases) < 85 and attempts < 500:
        attempts += 1
        lt      = random.choice(lease_types)
        term    = random_term_months()
        # Must start before Q3_END and end after Q4_END
        # Commencement latest: Q3_END - term_months (so it doesn't end before Q4)
        latest_start = Q4_END - relativedelta(months=1)    # must still be active in Q4
        earliest_start = date(2019, 1, 1)
        comm = random_commencement(earliest_start, latest_start)
        termination = comm + relativedelta(months=term)

        # Active in Q3 and Q4
        if comm <= Q3_END and termination > Q4_END:
            rec = build_lease(next_lease_id(counter), comm, term, lt, Q3_END)
            if rec:
                pool_a_leases.append(rec)

    # ------------------------------------------------------------------
    # Pool B: ~10 leases active in Q3 only (terminate Oct–Dec 2025)
    # ------------------------------------------------------------------
    pool_b_leases = []
    attempts = 0
    while len(pool_b_leases) < 10 and attempts < 300:
        attempts += 1
        lt   = random.choice(lease_types)
        # Termination between Oct 1 and Dec 31 2025
        term_date = date(2025, random.randint(10, 12), 1) + relativedelta(day=31)
        # Commencement: 12–84 months before termination
        back_months = random.randint(12, 84)
        comm = (term_date - relativedelta(months=back_months)).replace(day=1)
        term_months = round(months_between(comm, term_date))

        if comm <= Q3_END and comm >= date(2019, 1, 1):
            rec = build_lease(next_lease_id(counter), comm, term_months, lt, Q3_END)
            if rec:
                pool_b_leases.append(rec)

    # ------------------------------------------------------------------
    # Pool C: ~15 new leases starting in Q4 (Oct–Dec 2025)
    # ------------------------------------------------------------------
    pool_c_leases_q4 = []
    attempts = 0
    while len(pool_c_leases_q4) < 15 and attempts < 300:
        attempts += 1
        lt   = random.choice(lease_types)
        term = random_term_months()
        comm = random_commencement(date(2025, 10, 1), date(2025, 12, 1))
        rec  = build_lease(next_lease_id(counter), comm, term, lt, Q4_END)
        if rec:
            pool_c_leases_q4.append(rec)

    # ------------------------------------------------------------------
    # Build Q3 DataFrame  (Pool A + Pool B only)
    # ------------------------------------------------------------------
    q3_records = pool_a_leases + pool_b_leases
    random.shuffle(q3_records)

    df_q3 = pd.DataFrame(q3_records)[EXPORT_COLS]
    df_q3 = df_q3.sort_values("Capital_Lease_ID").reset_index(drop=True)

    # ------------------------------------------------------------------
    # Build Q4 DataFrame  (Pool A updated + Pool C new)
    # ------------------------------------------------------------------
    pool_a_q4 = [update_lease_for_q4(r) for r in pool_a_leases]
    q4_records = pool_a_q4 + pool_c_leases_q4
    random.shuffle(q4_records)

    df_q4 = pd.DataFrame(q4_records)[EXPORT_COLS]
    df_q4 = df_q4.sort_values("Capital_Lease_ID").reset_index(drop=True)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------
    q3_path = here / "lease_harbor_Q3.xlsx"
    q4_path = here / "lease_harbor_Q4.xlsx"
    df_q3.to_excel(q3_path, index=False, sheet_name="Lease_Harbor_Q3")
    df_q4.to_excel(q4_path, index=False, sheet_name="Lease_Harbor_Q4")

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    q3_ids = set(df_q3["Capital_Lease_ID"])
    q4_ids = set(df_q4["Capital_Lease_ID"])

    print(f"Q3 leases         : {len(df_q3):>4}   -> {q3_path.name}")
    print(f"Q4 leases         : {len(df_q4):>4}   -> {q4_path.name}")
    print(f"Continuing (A∩B)  : {len(q3_ids & q4_ids):>4}   (same ID, updated balances)")
    print(f"Terminated in Q4  : {len(q3_ids - q4_ids):>4}   (in Q3 only)")
    print(f"New in Q4         : {len(q4_ids - q3_ids):>4}   (in Q4 only)")
    print()

    for label, df in [("Q3", df_q3), ("Q4", df_q4)]:
        print(f"--- {label} breakdown ---")
        print(f"  Portfolios    : {dict(df['Portfolio'].value_counts().to_dict())}")
        print(f"  Currencies    : {dict(df['Currency'].value_counts().to_dict())}")
        print(f"  Company Codes : {dict(df['Company_Code'].value_counts().to_dict())}")
        total_rou = df["ROU_Asset_Cost"].sum()
        total_liab = df["Lease_Liability_Balance"].sum()
        print(f"  Total ROU Asset Cost      : {total_rou:>15,.2f}")
        print(f"  Total Lease Liability     : {total_liab:>15,.2f}")
        print()


if __name__ == "__main__":
    main()
