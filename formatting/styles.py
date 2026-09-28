"""Reusable openpyxl styles for the Arabic RTL workbook."""
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle, Protection
)
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

from .colors import *
from config import CURRENCY

# ---------- Fonts ----------
ARABIC_FONT = "Calibri"          # safe cross-platform Arabic-capable font
TITLE_FONT  = Font(name=ARABIC_FONT, size=20, bold=True, color=WHITE)
SUBTITLE_FONT = Font(name=ARABIC_FONT, size=12, color=WHITE)
SECTION_FONT = Font(name=ARABIC_FONT, size=14, bold=True, color=PRIMARY)
HEADER_FONT  = Font(name=ARABIC_FONT, size=11, bold=True, color=TABLE_HEADER_FG)
BODY_FONT    = Font(name=ARABIC_FONT, size=11, color=BLACK)
BODY_BOLD    = Font(name=ARABIC_FONT, size=11, bold=True, color=BLACK)
KPI_LABEL_FONT = Font(name=ARABIC_FONT, size=10, bold=True, color="404040")
KPI_VALUE_FONT = Font(name=ARABIC_FONT, size=16, bold=True, color=PRIMARY)
NOTE_FONT    = Font(name=ARABIC_FONT, size=9, italic=True, color="808080")

# ---------- Fills ----------
HEADER_FILL   = PatternFill("solid", fgColor=TABLE_HEADER_BG)
TITLE_FILL    = PatternFill("solid", fgColor=PRIMARY)
SUBTITLE_FILL = PatternFill("solid", fgColor=PRIMARY_LIGHT)
SECTION_FILL  = PatternFill("solid", fgColor=INFO)
ALT_ROW_FILL  = PatternFill("solid", fgColor=TABLE_ALT_ROW_BG)
KPI_FILLS = {
    "revenue":  PatternFill("solid", fgColor=KPI_REVENUE_BG),
    "expense":  PatternFill("solid", fgColor=KPI_EXPENSE_BG),
    "bank":     PatternFill("solid", fgColor=KPI_BANK_BG),
    "receiv":   PatternFill("solid", fgColor=KPI_RECEIV_BG),
    "liab":     PatternFill("solid", fgColor=KPI_LIAB_BG),
    "net":      PatternFill("solid", fgColor=KPI_NET_BG),
}

# ---------- Alignment (RTL) ----------
# Text cells — Right-to-Left (Arabic content reads RTL)
RTL_ALIGN  = Alignment(horizontal="right", vertical="center", readingOrder=2, wrap_text=True)
RTL_CENTER = Alignment(horizontal="center", vertical="center", readingOrder=2, wrap_text=True)
RTL_LEFT   = Alignment(horizontal="left", vertical="center", readingOrder=2, wrap_text=True)

# Numeric cells — Left-to-Right (numbers use Western digits, read LTR)
# readingOrder=1 forces Excel to render the digits in their natural order
RTL_NUMBER = Alignment(horizontal="left", vertical="center", readingOrder=1)

# ---------- Border ----------
thin = Side(style="thin", color=TABLE_BORDER)
med  = Side(style="medium", color=PRIMARY)
TABLE_BORDER_STYLE = Border(left=thin, right=thin, top=thin, bottom=thin)
BOX_BORDER         = Border(left=med, right=med, top=med, bottom=med)

# ---------- Number formats ----------
MONEY_FMT = f'#,##0.00 "{CURRENCY}"'
DATE_FMT  = "yyyy-mm-dd"
PERCENT_FMT = "0.0%"

# ---------- Protection ----------
LOCKED   = Protection(locked=True)
UNLOCKED = Protection(locked=False)