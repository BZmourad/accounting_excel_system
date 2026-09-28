"""Annual fees tracker with paid/remaining per fee."""
from config import SHEETS, INSTITUTIONS
from formatting.styles import *
from formatting.colors import PRIMARY
from formatting.dimensions import TITLE_ROW_HEIGHT
from validation.data_validation import dv_range_list, dv_decimal, dv_date, add_dv
from generators.lists_sheet import list_range
from data.lists import INSTITUTIONS_LIST, PAYMENT_STATUS


HEADERS = [
    "#", "المؤسسة", "بيان الرسم", "إجمالي الرسم",
    "المبلغ المدفوع", "المبلغ المتبقي", "تاريخ الاستحقاق", "حالة الدفع"
]


def build_annual_fees(wb):
    ws = wb.create_sheet(SHEETS["annual_fees"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = "7030A0"

    ws.merge_cells("A1:H1")
    c = ws["A1"]; c.value = "متابعة الرسوم السنوية"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value = "المبلغ المتبقي = إجمالي الرسم − المبلغ المدفوع (يُحتسب تلقائياً)"
    c.font = SUBTITLE_FONT; c.fill = SUBTITLE_FILL; c.alignment = RTL_CENTER

    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=4, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[4].height = 34

    widths = [6, 22, 30, 16, 16, 16, 14, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    start = 5
    end = 5 + 200 - 1
    for r in range(start, end + 1):
        ws.cell(row=r, column=1, value=f"=IF(B{r}=\"\",\"\",ROW()-{start-1})")
        ws.cell(row=r, column=6, value=f'=IF(D{r}="","",N(D{r})-IF(E{r}="",0,N(E{r})))')
        for col in range(1, 9):
            cell = ws.cell(row=r, column=col)
            cell.border = TABLE_BORDER_STYLE
            if col in (1, 2, 8):
                cell.alignment = RTL_CENTER
            elif col in (4, 5, 6):
                cell.alignment = RTL_NUMBER
                cell.number_format = MONEY_FMT
            elif col == 7:
                cell.alignment = RTL_CENTER
                cell.number_format = DATE_FMT
            else:
                cell.alignment = RTL_ALIGN
            if r % 2 == 0:
                cell.fill = ALT_ROW_FILL

    add_dv(ws, dv_range_list(list_range("A", len(INSTITUTIONS_LIST))), f"B{start}:B{end}")
    add_dv(ws, dv_range_list(list_range("D", len(PAYMENT_STATUS))), f"H{start}:H{end}")
    add_dv(ws, dv_decimal(minimum=0), f"D{start}:E{end}")
    add_dv(ws, dv_date(), f"G{start}:G{end}")

    tr = end + 2
    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=3)
    ws.cell(row=tr, column=1, value="الإجماليات").font = Font(
        name=ARABIC_FONT, bold=True, color=WHITE)
    ws.cell(row=tr, column=1).fill = HEADER_FILL
    ws.cell(row=tr, column=1).alignment = RTL_CENTER
    for col in range(1, 9):
        ws.cell(row=tr, column=col).fill = HEADER_FILL
        ws.cell(row=tr, column=col).border = TABLE_BORDER_STYLE
    for col in (4, 5, 6):
        letter = chr(64 + col)
        c = ws.cell(row=tr, column=col, value=f"=SUM({letter}5:{letter}{end})")
        c.number_format = MONEY_FMT
        c.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)
        c.alignment = RTL_NUMBER

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H{end}"
    return ws