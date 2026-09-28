"""Reports sheet with date-range selector and institution comparison."""
from datetime import date

from openpyxl.styles import PatternFill

from config import SHEETS, INSTITUTIONS
from formatting.styles import *
from formatting.colors import *
from formatting.dimensions import TITLE_ROW_HEIGHT


def build_reports(wb):
    ws = wb.create_sheet(SHEETS["reports"])
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = "44546A"

    # ---------- Title ----------
    ws.merge_cells("A1:F1")
    c = ws["A1"]; c.value = "التقارير التحليلية بالفترة"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    widths = [30, 20, 20, 20, 20, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    # ---------- Date range picker (real dates!) ----------
    r = 3
    ws.cell(row=r, column=1, value="من تاريخ").font = BODY_BOLD
    ws.cell(row=r, column=1).alignment = RTL_CENTER

    ws.cell(row=r, column=2).value = date(2026, 1, 1)
    ws.cell(row=r, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
    ws.cell(row=r, column=2).alignment = RTL_CENTER
    ws.cell(row=r, column=2).font = BODY_BOLD

    ws.cell(row=r, column=3, value="إلى تاريخ").font = BODY_BOLD
    ws.cell(row=r, column=3).alignment = RTL_CENTER

    ws.cell(row=r, column=4).value = date(2026, 12, 31)
    ws.cell(row=r, column=4).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor="FFF2CC")
    ws.cell(row=r, column=4).alignment = RTL_CENTER
    ws.cell(row=r, column=4).font = BODY_BOLD

    ws.cell(row=r, column=5, value="ملاحظة:").font = NOTE_FONT
    ws.cell(row=r, column=6, value="غيّر التواريخ لتحديث كافة التقارير").font = NOTE_FONT
    ws.cell(row=r, column=6).alignment = RTL_ALIGN

    # ---------- Section title ----------
    r = 6
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value="مقارنة المؤسسات (خلال الفترة المحددة)")
    c.font = SECTION_FONT; c.fill = SECTION_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[r].height = 26

    # ---------- Comparison header ----------
    r += 1
    headers = ["المؤشر", "المؤسسة الأولى", "المؤسسة الثانية",
               "المؤسسة الثالثة", "الإجمالي الموحّد"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[r].height = 30
    header_row = r

    # ---------- Safe formula helpers ----------
    rev = SHEETS["revenues"]
    exp = SHEETS["expenses"]

    def ref(sheet, col):
        # Always quote the sheet name — safe even for Arabic names
        return f"'{sheet}'!${col}$5:${col}$100000"

    rev_amt  = ref(rev, "F")
    rev_inst = ref(rev, "B")
    rev_stat = ref(rev, "I")
    rev_date = ref(rev, "G")

    exp_amt  = ref(exp, "F")
    exp_inst = ref(exp, "B")
    exp_stat = ref(exp, "I")
    exp_date = ref(exp, "G")

    START = "$B$3"
    END   = "$D$3"

    def sumifs(amount, inst, date_range, inst_name,
               status=None, status_range=None):
        parts = [
            amount,
            inst, f'"{inst_name}"',
            date_range, f'">="&{START}',
            date_range, f'"<="&{END}',
        ]
        if status is not None and status_range is not None:
            parts += [status_range, f'"{status}"']
        return "=SUMIFS(" + ",".join(parts) + ")"

    def recv_or_liab(amount, inst, date_range, inst_name, status_range):
        a = sumifs(amount, inst, date_range, inst_name, "مستحق", status_range)
        b = sumifs(amount, inst, date_range, inst_name, "معلق", status_range)
        # a starts with '=', b starts with '=' — strip b's leading '=' to add
        return a + "+" + b[1:]

    metrics = [
        ("الإيرادات",
         lambda n: sumifs(rev_amt, rev_inst, rev_date, n)),
        ("الإيرادات المحصّلة",
         lambda n: sumifs(rev_amt, rev_inst, rev_date, n, "مدفوع", rev_stat)),
        ("المستحقات",
         lambda n: recv_or_liab(rev_amt, rev_inst, rev_date, n, rev_stat)),
        ("المصروفات",
         lambda n: sumifs(exp_amt, exp_inst, exp_date, n)),
        ("المصروفات المدفوعة",
         lambda n: sumifs(exp_amt, exp_inst, exp_date, n, "مدفوع", exp_stat)),
        ("الالتزامات",
         lambda n: recv_or_liab(exp_amt, exp_inst, exp_date, n, exp_stat)),
    ]

    for i, (label, builder) in enumerate(metrics, start=1):
        rr = header_row + i
        ws.cell(row=rr, column=1, value=label).font = BODY_BOLD
        ws.cell(row=rr, column=1).alignment = RTL_ALIGN
        for j, inst in enumerate(INSTITUTIONS):
            f = builder(inst["name"])
            ws.cell(row=rr, column=2 + j, value=f)
        ws.cell(row=rr, column=5, value=f"=SUM(B{rr}:D{rr})")
        for col in range(1, 6):
            cell = ws.cell(row=rr, column=col)
            cell.border = TABLE_BORDER_STYLE
            if col > 1:
                cell.alignment = RTL_NUMBER
                cell.number_format = MONEY_FMT
                cell.font = BODY_FONT
            if rr % 2 == 0:
                cell.fill = ALT_ROW_FILL

    ws.freeze_panes = "A5"
    return ws