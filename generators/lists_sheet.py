"""Builds the 'القوائم والإعدادات' sheet containing dropdown source lists."""
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from config import SHEETS, INSTITUTIONS, MAX_ROWS
from data.lists import (
    INSTITUTIONS_LIST, REVENUE_CATEGORIES, EXPENSE_CATEGORIES,
    PAYMENT_STATUS, VEHICLE_TYPES, CONTRACT_STATUS, BANK_ACCOUNTS
)
from formatting.colors import PRIMARY, WHITE
from formatting.styles import HEADER_FILL, HEADER_FONT, RTL_CENTER, RTL_ALIGN, BODY_FONT


# Column -> list mapping
LISTS_LAYOUT = {
    "A": ("المؤسسات", INSTITUTIONS_LIST),
    "B": ("تصنيفات الإيرادات", REVENUE_CATEGORIES),
    "C": ("تصنيفات المصروفات", EXPENSE_CATEGORIES),
    "D": ("حالة الدفع", PAYMENT_STATUS),
    "E": ("أنواع المركبات", VEHICLE_TYPES),
    "F": ("حالة العقد", CONTRACT_STATUS),
    "G": ("الحسابات البنكية", BANK_ACCOUNTS),
}


def build_lists_sheet(wb):
    ws = wb.create_sheet(SHEETS["lists"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = "808080"

    # write header
    for col_letter, (title, values) in LISTS_LAYOUT.items():
        c = ws[f"{col_letter}1"]
        c.value = title
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = RTL_CENTER
        for i, val in enumerate(values, start=2):
            cell = ws[f"{col_letter}{i}"]
            cell.value = val
            cell.font = BODY_FONT
            cell.alignment = RTL_CENTER
        ws.column_dimensions[col_letter].width = 24

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    return ws


def list_range(col_letter: str, length: int) -> str:
    """Absolute reference to a list range in the lists sheet."""
    return f"'{SHEETS['lists']}'!${col_letter}$2:${col_letter}${1+length}"