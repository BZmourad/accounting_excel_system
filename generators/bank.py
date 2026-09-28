"""Bank accounts summary per institution."""
from config import SHEETS, INSTITUTIONS
from formatting.styles import *
from formatting.colors import PRIMARY, PRIMARY_LIGHT, INFO
from formatting.dimensions import TITLE_ROW_HEIGHT


def build_bank(wb):
    ws = wb.create_sheet(SHEETS["bank"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = PRIMARY_LIGHT

    ws.merge_cells("A1:G1")
    c = ws["A1"]; c.value = "الأرصدة البنكية الفعلية"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    ws.merge_cells("A2:G2")
    c = ws["A2"]
    c.value = ("الرصيد الفعلي = الرصيد الافتتاحي + المبالغ المحصّلة فعلياً "
               "− المصروفات المدفوعة فعلياً")
    c.font = SUBTITLE_FONT; c.fill = SUBTITLE_FILL; c.alignment = RTL_CENTER

    headers = ["المؤسسة", "الرصيد الافتتاحي",
               "إجمالي المحصّل", "إجمالي المدفوع",
               "الرصيد الفعلي", "المستحقات", "الالتزامات"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[4].height = 34

    widths = [26, 18, 18, 18, 18, 18, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    rev = SHEETS["revenues"]; exp = SHEETS["expenses"]
    for idx, inst in enumerate(INSTITUTIONS, start=1):
        r = 4 + idx
        name = inst["name"]
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=0)
        ws.cell(row=r, column=3,
                value=f'=SUMIFS(\'{rev}\'!$F$5:$F$100000,'
                      f'\'{rev}\'!$B$5:$B$100000,"{name}",'
                      f'\'{rev}\'!$I$5:$I$100000,"مدفوع")')
        ws.cell(row=r, column=4,
                value=f'=SUMIFS(\'{exp}\'!$F$5:$F$100000,'
                      f'\'{exp}\'!$B$5:$B$100000,"{name}",'
                      f'\'{exp}\'!$I$5:$I$100000,"مدفوع")')
        ws.cell(row=r, column=5, value=f"=B{r}+C{r}-D{r}")
        ws.cell(row=r, column=6,
                value=f'=SUMIFS(\'{rev}\'!$F$5:$F$100000,'
                      f'\'{rev}\'!$B$5:$B$100000,"{name}",'
                      f'\'{rev}\'!$I$5:$I$100000,"مستحق")'
                      f'+SUMIFS(\'{rev}\'!$F$5:$F$100000,'
                      f'\'{rev}\'!$B$5:$B$100000,"{name}",'
                      f'\'{rev}\'!$I$5:$I$100000,"معلق")')
        ws.cell(row=r, column=7,
                value=f'=SUMIFS(\'{exp}\'!$F$5:$F$100000,'
                      f'\'{exp}\'!$B$5:$B$100000,"{name}",'
                      f'\'{exp}\'!$I$5:$I$100000,"مستحق")'
                      f'+SUMIFS(\'{exp}\'!$F$5:$F$100000,'
                      f'\'{exp}\'!$B$5:$B$100000,"{name}",'
                      f'\'{exp}\'!$I$5:$I$100000,"معلق")')

        for col in range(1, 8):
            cell = ws.cell(row=r, column=col)
            cell.border = TABLE_BORDER_STYLE
            if col == 1:
                cell.font = BODY_BOLD; cell.alignment = RTL_ALIGN
            else:
                cell.font = BODY_FONT; cell.alignment = RTL_NUMBER
                cell.number_format = MONEY_FMT
            if r % 2 == 0:
                cell.fill = ALT_ROW_FILL

    tr = 4 + len(INSTITUTIONS) + 1
    ws.cell(row=tr, column=1, value="الإجمالي")
    for col in range(1, 8):
        cell = ws.cell(row=tr, column=col)
        cell.fill = HEADER_FILL; cell.border = TABLE_BORDER_STYLE
        if col == 1:
            cell.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)
            cell.alignment = RTL_CENTER
        else:
            letter = chr(64 + col)
            cell.value = f"=SUM({letter}5:{letter}{tr-1})"
            cell.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)
            cell.alignment = RTL_NUMBER
            cell.number_format = MONEY_FMT

    ws.freeze_panes = "A5"
    return ws