"""
Formula builders — keeps formulas consistent across sheets.
All formulas are Excel-native so the workbook works without Python.
"""

# ---------- Vehicles ----------
def net_rental(row):
    """Net Monthly Rental = Rental - Maintenance"""
    return f"=IF(N(I{row})=0,0,I{row}-J{row})"

# ---------- Revenue / Expense status logic ----------
def collected_amount(row, amount_col="F", status_col="I", pay_date_col="H"):
    """
    Collected = amount if status = مدفوع
              = 0 if status in (مستحق, معلق)
              = user-entered partial if مدفوع جزئيا (handled by extra column)
    We keep it simple: collected = full amount when paid, else 0.
    """
    return (
        f'=IF({status_col}{row}="مدفوع",N({amount_col}{row}),'
        f'IF({status_col}{row}="مدفوع جزئيا",N({amount_col}{row})*0.5,0))'
    )
    # NOTE: partial is refined later via a dedicated "المبلغ المحصّل" column.

def is_paid(status_cell):
    return f'=IF({status_cell}="مدفوع","✔","")'

# ---------- Aggregation helpers ----------
def sumif_range(criteria_range, criteria, sum_range, sheet=None):
    ref = f"'{sheet}'!" if sheet else ""
    return f"=SUMIF({ref}{criteria_range},{criteria},{ref}{sum_range})"

def sumifs_range(sum_range, crit_pairs, sheet=None):
    """crit_pairs = [(range, criteria), ...]"""
    ref = f"'{sheet}'!" if sheet else ""
    parts = ",".join([f"{ref}{r},{c}" for r, c in crit_pairs])
    return f"=SUMIFS({ref}{sum_range},{parts})"

def total_range(rng, sheet=None):
    ref = f"'{sheet}'!" if sheet else ""
    return f"=SUM({ref}{rng})"

# ---------- Date range helpers ----------
def sumifs_dated(sum_range, date_range, start_cell, end_cell, extra=None):
    extra = extra or []
    parts = [f"{sum_range},{date_range},\">=\"&{start_cell},{date_range},\"<=\"&{end_cell}"]
    for r, c in extra:
        parts.append(f"{r},{c}")
    return "=SUMIFS(" + ",".join(parts) + ")"

# ---------- Bank ----------
def bank_balance(opening, collected_range, paid_range):
    """Actual bank = opening + collected - paid"""
    return f"=N({opening})+SUM({collected_range})-SUM({paid_range})"