"""Creates the workbook and applies page-level settings (RTL, tab colors)."""
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import SHEETS, INSTITUTIONS
from formatting.colors import PRIMARY, PRIMARY_LIGHT, ACCENT, INFO


def create_workbook() -> Workbook:
    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)
    wb.properties.title = "نظام الإدارة المالية"
    wb.properties.creator = "Python Accounting Generator"
    wb.properties.description = "نظام محاسبي احترافي باللغة العربية"

    return wb


def setup_sheet(ws: Worksheet, tab_color: str = PRIMARY, rtl: bool = True):
    """Apply RTL, gridlines off, tab color."""
    ws.sheet_view.rightToLeft = rtl
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color