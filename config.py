"""
Global configuration for the Arabic RTL Accounting System.
"""

# -------------------------------
# Test data seeding
# -------------------------------
SEED_TEST_DATA = True          # set to False for an empty production workbook
OPENING_BALANCES = {
    "المؤسسة الأولى":  50000,
    "المؤسسة الثانية": 30000,
    "المؤسسة الثالثة": 20000,
}

# -------------------------------
# Institutions
# -------------------------------
INSTITUTIONS = [
    {"id": 1, "name": "المؤسسة الأولى",  "sheet": "المؤسسة 1", "short": "م1"},
    {"id": 2, "name": "المؤسسة الثانية", "sheet": "المؤسسة 2", "short": "م2"},
    {"id": 3, "name": "المؤسسة الثالثة", "sheet": "المؤسسة 3", "short": "م3"},
]

# -------------------------------
# Workbook output
# -------------------------------
OUTPUT_FILE = "output/accounting_system.xlsx"
CURRENCY = "ر.س"
COMPANY_NAME = "نظام الإدارة المالية والمحاسبية"
COMPANY_SUBTITLE = "لإدارة المؤسسات الثلاث"

# -------------------------------
# Sheet names (Arabic RTL)
# -------------------------------
SHEETS = {
    "dashboard":   "لوحة التحكم",
    "institutions":"المؤسسات",
    "vehicles_1":  "مركبات المؤسسة 1",
    "vehicles_2":  "مركبات المؤسسة 2",
    "vehicles_3":  "مركبات المؤسسة 3",
    "revenues":    "الإيرادات",
    "expenses":    "المصروفات",
    "bank":        "الحسابات البنكية",
    "receivables": "المبالغ المستحقة",
    "liabilities": "الالتزامات",
    "annual_fees": "الرسوم السنوية",
    "reports":     "التقارير",
    "lists":       "القوائم والإعدادات",
}

# -------------------------------
# Capacity of data tables (rows reserved for user input)
# -------------------------------
MAX_ROWS = 500          # generic data rows
MAX_VEHICLES = 300
MAX_TX = 1000           # revenue / expense transactions

# -------------------------------
# Row layout constants
# -------------------------------
HEADER_ROW = 1
TITLE_ROW  = 1
SUBTITLE_ROW = 2
DATA_START_ROW = 5
TOTALS_GAP = 2

# -------------------------------
# Workbook encryption
# -------------------------------
WORKBOOK_PASSWORD = "Client@2026!Secure"   # change per client
ENCRYPT_OUTPUT    = True                    # set to False to skip encryption