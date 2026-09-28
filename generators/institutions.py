"""Institutions master table."""
from config import SHEETS, INSTITUTIONS
from formatting.styles import *
from formatting.dimensions import TITLE_ROW_HEIGHT
from formatting.colors import PRIMARY


def build_institutions(wb):
    ws = wb.create_sheet(SHEETS["institutions"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = PRIMARY

    ws.merge_cells("A1:E1")
    c = ws["A1"]; c.value = "بيانات المؤسسات"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    headers = ["#", "اسم المؤسسة", "الرقم الضريبي", "العنوان", "ملاحظات"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE

    widths = [6, 26, 20, 32, 32]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    for idx, inst in enumerate(INSTITUTIONS, start=1):
        r = 3 + idx
        ws.cell(row=r, column=1, value=idx).alignment = RTL_CENTER
        ws.cell(row=r, column=2, value=inst["name"]).alignment = RTL_ALIGN
        for col in range(1, 6):
            cell = ws.cell(row=r, column=col)
            cell.border = TABLE_BORDER_STYLE
            cell.font = BODY_FONT if col > 1 else BODY_BOLD
            if r % 2 == 0:
                cell.fill = ALT_ROW_FILL

    ws.freeze_panes = "A4"
    return ws