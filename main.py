"""
Entry point — generates the complete Arabic RTL Excel accounting workbook
and seeds it with realistic test data.
Run:
    python main.py
"""
import os

from config import OUTPUT_FILE, SHEETS
from generators.workbook import create_workbook
from generators.lists_sheet import build_lists_sheet
from generators.institutions import build_institutions
from generators.dashboard import build_dashboard
from generators.vehicles import build_vehicle_sheet
from generators.revenues import build_revenues
from generators.expenses import build_expenses
from generators.bank import build_bank
from generators.annual_fees import build_annual_fees
from generators.reports import build_reports
from generators.readme_sheet import build_guide
from generators.seeder import seed_all 


def main():
    print("🔧 إنشاء ملف Excel المحاسبي...")

    wb = create_workbook()

    # ---- Structural generation ----
    build_dashboard(wb)
    build_institutions(wb)
    build_vehicle_sheet(wb, 0, SHEETS["vehicles_1"])
    build_vehicle_sheet(wb, 1, SHEETS["vehicles_2"])
    build_vehicle_sheet(wb, 2, SHEETS["vehicles_3"])
    build_revenues(wb)
    build_expenses(wb)
    build_bank(wb)
    build_annual_fees(wb)
    build_reports(wb)
    build_lists_sheet(wb)
    build_guide(wb)

    # ---- Seed test data ----
    seed_all(wb)

    # ---- Save ----
    os.makedirs(os.path.dirname(OUTPUT_FILE) or ".", exist_ok=True)
    wb.save(OUTPUT_FILE)

    # ---- Encrypt with password (optional) ----
    from config import ENCRYPT_OUTPUT, WORKBOOK_PASSWORD
    if ENCRYPT_OUTPUT:
        from generators.encrypt import encrypt_xlsx
        encrypt_xlsx(OUTPUT_FILE, WORKBOOK_PASSWORD)
        print(f"🔐 تم تشفير الملف بكلمة مرور.")
        print(f"   كلمة المرور: {WORKBOOK_PASSWORD}")
    else:
        print(f"ℹ️  التشفير معطّل — الملف بدون كلمة مرور.")

    print(f"✅ تم إنشاء الملف بنجاح: {OUTPUT_FILE}")
    print(f"   عدد الأوراق: {len(wb.sheetnames)}")


if __name__ == "__main__":
    main()