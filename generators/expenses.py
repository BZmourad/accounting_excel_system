"""Expenses transactions sheet."""
from openpyxl.utils import get_column_letter

from config import SHEETS, MAX_TX
from formatting.styles import *
from formatting.dimensions import WIDTH_TX, TITLE_ROW_HEIGHT
from formatting.colors import PRIMARY
from validation.data_validation import dv_range_list, dv_decimal, dv_date, add_dv
from generators.lists_sheet import list_range
from data.lists import (
    INSTITUTIONS_LIST, EXPENSE_CATEGORIES, PAYMENT_STATUS, BANK_ACCOUNTS
)


HEADERS = [
    "#", "المؤسسة", "تصنيف المصروف", "الوصف", "المستفيد",
    "المبلغ", "تاريخ الاستحقاق", "تاريخ الدفع", "حالة الدفع",
    "الحساب البنكي", "ملاحظات"
]


def _header(ws):
    last_col = get_column_letter(len(HEADERS))
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]; c.value = "سجل المصروفات"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT
    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value = ("المصروف يؤثر على الرصيد البنكي فقط عند حالة الدفع = مدفوع "
               "| المستحق والمعلق يمثلان التزامات فقط")
    c.font = SUBTITLE_FONT; c.fill = SUBTITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[2].height = 22


def _columns(ws):
    for col, w in WIDTH_TX.items():
        ws.column_dimensions[col].width = w


def _table_header(ws):
    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=4, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[4].height = 34


def _rows(ws, start_row, end_row):
    for r in range(start_row, end_row + 1):
        ws.cell(row=r, column=1,
                value=f"=IF(B{r}=\"\",\"\",ROW()-{start_row-1})")
        for col in range(1, len(HEADERS) + 1):
            c = ws.cell(row=r, column=col)
            c.border = TABLE_BORDER_STYLE
            if col in (1, 2, 3, 9, 10):
                c.alignment = RTL_CENTER
            elif col == 6:
                c.alignment = RTL_NUMBER
                c.number_format = MONEY_FMT
            elif col in (7, 8):
                c.alignment = RTL_CENTER
                c.number_format = DATE_FMT
            else:
                c.alignment = RTL_ALIGN
            if r % 2 == 0:
                c.fill = ALT_ROW_FILL


def _validation(ws, start_row, end_row):
    add_dv(ws, dv_range_list(list_range("A", len(INSTITUTIONS_LIST))), f"B{start_row}:B{end_row}")
    add_dv(ws, dv_range_list(list_range("C", len(EXPENSE_CATEGORIES))), f"C{start_row}:C{end_row}")
    add_dv(ws, dv_range_list(list_range("D", len(PAYMENT_STATUS))), f"I{start_row}:I{end_row}")
    add_dv(ws, dv_range_list(list_range("G", len(BANK_ACCOUNTS))), f"J{start_row}:J{end_row}")
    add_dv(ws, dv_decimal(minimum=0), f"F{start_row}:F{end_row}")
    add_dv(ws, dv_date(), f"G{start_row}:G{end_row}")
    add_dv(ws, dv_date(), f"H{start_row}:H{end_row}")


def _totals(ws, totals_row, end_row):
    ws.merge_cells(start_row=totals_row, start_column=1, end_row=totals_row, end_column=5)
    c = ws.cell(row=totals_row, column=1, value="الإجماليات")
    c.font = Font(name=ARABIC_FONT, bold=True, color=WHITE, size=12)
    c.fill = HEADER_FILL; c.alignment = RTL_CENTER

    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=totals_row, column=col)
        cell.fill = HEADER_FILL
        cell.border = TABLE_BORDER_STYLE

    ws.cell(row=totals_row, column=6,
            value=f"=SUM(F5:F{end_row})").number_format = MONEY_FMT
    ws.cell(row=totals_row, column=9,
            value=f'=COUNTIF(I5:I{end_row},"مدفوع")&" مدفوع / "'
                  f'&COUNTIF(I5:I{end_row},"مستحق")&" مستحق / "'
                  f'&COUNTIF(I5:I{end_row},"معلق")&" معلق"'
            ).alignment = RTL_CENTER
    for col in (6, 9):
        cell = ws.cell(row=totals_row, column=col)
        cell.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)


def build_expenses(wb):
    ws = wb.create_sheet(SHEETS["expenses"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = "C00000"

    _header(ws)
    _columns(ws)
    _table_header(ws)

    start_row = 5
    end_row = start_row + MAX_TX - 1
    _rows(ws, start_row, end_row)
    _validation(ws, start_row, end_row)

    totals_row = end_row + 2
    _totals(ws, totals_row, end_row)

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:K{end_row}"

    return ws