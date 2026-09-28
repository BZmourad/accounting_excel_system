"""Seeds realistic test data into the generated workbook."""
from config import (
    SEED_TEST_DATA, SHEETS, INSTITUTIONS, OPENING_BALANCES
)
from data.test_data import VEHICLES, REVENUES, EXPENSES, ANNUAL_FEES


# ----------------------------------------------------------------------
# Vehicle seeding
# ----------------------------------------------------------------------
def seed_vehicles(wb):
    for idx, inst in enumerate(INSTITUTIONS):
        sheet_name = SHEETS[f"vehicles_{idx+1}"]
        ws = wb[sheet_name]
        rows = VEHICLES.get(inst["name"], [])
        start_row = 5
        for i, row in enumerate(rows):
            r = start_row + i
            (customer, org, vtype, count, veh_no, plate,
             delivery, rental, maintenance, status, notes) = row
            ws.cell(row=r, column=2, value=customer)
            ws.cell(row=r, column=3, value=org)
            ws.cell(row=r, column=4, value=vtype)
            ws.cell(row=r, column=5, value=count)
            ws.cell(row=r, column=6, value=veh_no)
            ws.cell(row=r, column=7, value=plate)
            ws.cell(row=r, column=8, value=delivery)
            ws.cell(row=r, column=9, value=rental)
            ws.cell(row=r, column=10, value=maintenance)
            # Column 11 (net rental) is already a formula — don't overwrite
            ws.cell(row=r, column=12, value=status)
            ws.cell(row=r, column=13, value=notes)


# ----------------------------------------------------------------------
# Revenue seeding
# ----------------------------------------------------------------------
def seed_revenues(wb):
    ws = wb[SHEETS["revenues"]]
    start_row = 5
    for i, row in enumerate(REVENUES):
        r = start_row + i
        (inst, cat, cust, desc, amount,
         due, collected, status, bank, notes) = row
        ws.cell(row=r, column=2, value=inst)
        ws.cell(row=r, column=3, value=cat)
        ws.cell(row=r, column=4, value=cust)
        ws.cell(row=r, column=5, value=desc)
        ws.cell(row=r, column=6, value=amount)
        ws.cell(row=r, column=7, value=due)
        ws.cell(row=r, column=8, value=collected or None)
        ws.cell(row=r, column=9, value=status)
        ws.cell(row=r, column=10, value=bank)
        ws.cell(row=r, column=11, value=notes)


# ----------------------------------------------------------------------
# Expense seeding
# ----------------------------------------------------------------------
def seed_expenses(wb):
    ws = wb[SHEETS["expenses"]]
    start_row = 5
    for i, row in enumerate(EXPENSES):
        r = start_row + i
        (inst, cat, desc, benef, amount,
         due, paid, status, bank, notes) = row
        ws.cell(row=r, column=2, value=inst)
        ws.cell(row=r, column=3, value=cat)
        ws.cell(row=r, column=4, value=desc)
        ws.cell(row=r, column=5, value=benef)
        ws.cell(row=r, column=6, value=amount)
        ws.cell(row=r, column=7, value=due)
        ws.cell(row=r, column=8, value=paid or None)
        ws.cell(row=r, column=9, value=status)
        ws.cell(row=r, column=10, value=bank)
        ws.cell(row=r, column=11, value=notes)


# ----------------------------------------------------------------------
# Annual fees seeding
# ----------------------------------------------------------------------
def seed_annual_fees(wb):
    ws = wb[SHEETS["annual_fees"]]
    start_row = 5
    for i, row in enumerate(ANNUAL_FEES):
        r = start_row + i
        inst, desc, total, paid, due, status = row
        ws.cell(row=r, column=2, value=inst)
        ws.cell(row=r, column=3, value=desc)
        ws.cell(row=r, column=4, value=total)
        ws.cell(row=r, column=5, value=paid)
        # Column 6 (remaining) is a formula — don't overwrite
        ws.cell(row=r, column=7, value=due)
        ws.cell(row=r, column=8, value=status)


# ----------------------------------------------------------------------
# Opening balances in Bank sheet
# ----------------------------------------------------------------------
def seed_opening_balances(wb):
    ws = wb[SHEETS["bank"]]
    for i, inst in enumerate(INSTITUTIONS):
        r = 5 + i
        opening = OPENING_BALANCES.get(inst["name"], 0)
        ws.cell(row=r, column=2, value=opening)


# ----------------------------------------------------------------------
# Master entry point
# ----------------------------------------------------------------------
def seed_all(wb):
    if not SEED_TEST_DATA:
        print("ℹ️  SEED_TEST_DATA = False — generating empty workbook.")
        return
    print("🌱 إضافة بيانات تجريبية...")
    seed_vehicles(wb)
    seed_revenues(wb)
    seed_expenses(wb)
    seed_annual_fees(wb)
    seed_opening_balances(wb)
    print("   ✓ بيانات المركبات")
    print(f"   ✓ {len(REVENUES)} عملية إيراد")
    print(f"   ✓ {len(EXPENSES)} عملية مصروف")
    print(f"   ✓ {len(ANNUAL_FEES)} رسم سنوي")