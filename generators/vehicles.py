"""Vehicle management sheet — one per institution, identical layout."""
from openpyxl.utils import get_column_letter

from config import SHEETS, INSTITUTIONS, MAX_VEHICLES, CURRENCY
from formatting.styles import *
from formatting.dimensions import WIDTH_VEHICLES, TITLE_ROW_HEIGHT
from formatting.colors import PRIMARY, PRIMARY_LIGHT, ACCENT, INFO
from validation.data_validation import dv_range_list, dv_decimal, dv_date, add_dv
from generators.lists_sheet import list_range
from data.lists import VEHICLE_TYPES, CONTRACT_STATUS


HEADERS = [
    "#", "اسم العميل / المستأجر", "الجهة / المؤسسة", "نوع المركبة",
    "عدد المركبات", "رقم المركبة", "رقم اللوحة", "تاريخ التسليم",
    "قيمة الإيجار الشهري", "خصم الصيانة", "صافي الإيجار الشهري",
    "حالة العقد", "ملاحظات"
]


def _header(ws, title: str, subtitle: str = ""):
    last_col = get_column_letter(len(HEADERS))
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]; c.value = title
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    if subtitle:
        ws.merge_cells(f"A2:{last_col}2")
        c = ws["A2"]; c.value = subtitle
        c.font = SUBTITLE_FONT; c.fill = SUBTITLE_FILL; c.alignment = RTL_CENTER
        ws.row_dimensions[2].height = 22


def _table_header(ws, row=4):
    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[row].height = 34


def _write_columns(ws):
    for col, w in WIDTH_VEHICLES.items():
        ws.column_dimensions[col].width = w


def _apply_validation(ws, start_row, end_row):
    # Vehicle type col D
    dv = dv_range_list(list_range("E", len(VEHICLE_TYPES)), allow_blank=True)
    add_dv(ws, dv, f"D{start_row}:D{end_row}")

    # Contract status col L
    dv2 = dv_range_list(list_range("F", len(CONTRACT_STATUS)), allow_blank=True)
    add_dv(ws, dv2, f"L{start_row}:L{end_row}")

    # Count col E
    dv3 = dv_decimal(minimum=0)
    add_dv(ws, dv3, f"E{start_row}:E{end_row}")

    # Amounts I, J
    dv4 = dv_decimal(minimum=0)
    add_dv(ws, dv4, f"I{start_row}:I{end_row}")
    add_dv(ws, dv4, f"J{start_row}:J{end_row}")

    # Delivery date H
    dv5 = dv_date()
    add_dv(ws, dv5, f"H{start_row}:H{end_row}")


def _write_rows(ws, start_row, end_row):
    for r in range(start_row, end_row + 1):
        # Serial
        cell = ws.cell(row=r, column=1, value=f"=IF(B{r}=\"\",\"\",ROW()-{start_row-1})")
        cell.alignment = RTL_CENTER
        # Net rental formula (col K)
        ws.cell(row=r, column=11).value = (
            f"=IF(OR(I{r}=\"\",I{r}=0),\"\",N(I{r})-N(J{r}))"
        )
        ws.cell(row=r, column=11).number_format = MONEY_FMT
        ws.cell(row=r, column=11).font = BODY_BOLD
        ws.cell(row=r, column=11).alignment = RTL_NUMBER

        for col in range(1, len(HEADERS) + 1):
            c = ws.cell(row=r, column=col)
            c.border = TABLE_BORDER_STYLE
            if col in (1, 4, 5, 8, 12):
                c.alignment = RTL_CENTER
            elif col in (9, 10, 11):
                c.alignment = RTL_NUMBER
                c.number_format = MONEY_FMT
            else:
                c.alignment = RTL_ALIGN
            if col == 8:
                c.number_format = DATE_FMT
            if r % 2 == 0:
                c.fill = ALT_ROW_FILL
        ws.row_dimensions[r].height = 22


def _totals_row(ws, row, end_row):
    """Totals: net rental sum, vehicle count, monthly rental sum."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    c = ws.cell(row=row, column=1, value="الإجماليات")
    c.font = Font(name=ARABIC_FONT, bold=True, color=WHITE, size=12)
    c.fill = HEADER_FILL; c.alignment = RTL_CENTER

    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.border = TABLE_BORDER_STYLE

    # E = sum of counts
    ws.cell(row=row, column=5,
            value=f"=SUM(E{5}:E{end_row})").number_format = "#,##0"
    # I = sum monthly rental
    ws.cell(row=row, column=9,
            value=f"=SUM(I{5}:I{end_row})").number_format = MONEY_FMT
    # J = sum maintenance
    ws.cell(row=row, column=10,
            value=f"=SUM(J{5}:J{end_row})").number_format = MONEY_FMT
    # K = sum net
    ws.cell(row=row, column=11,
            value=f"=SUM(K{5}:K{end_row})").number_format = MONEY_FMT

    for col in (5, 9, 10, 11):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)
        cell.alignment = RTL_NUMBER


def build_vehicle_sheet(wb, inst_idx: int, sheet_name: str):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = ["2E75B6", "C55A11", "70AD47"][inst_idx]

    inst = INSTITUTIONS[inst_idx]
    _header(
        ws,
        f"إدارة المركبات - {inst['name']}",
        "جميع المبالغ بالريال السعودي | صافي الإيجار = الإيجار الشهري - خصم الصيانة"
    )
    _write_columns(ws)
    _table_header(ws, row=4)

    start_row = 5
    end_row = start_row + MAX_VEHICLES - 1
    _write_rows(ws, start_row, end_row)
    _apply_validation(ws, start_row, end_row)

    totals_row = end_row + 2
    _totals_row(ws, totals_row, end_row)

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:M{end_row}"

    return ws