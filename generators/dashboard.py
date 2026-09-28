"""Professional consolidated dashboard."""
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

from config import SHEETS, INSTITUTIONS, COMPANY_NAME, COMPANY_SUBTITLE, CURRENCY
from formatting.colors import *
from formatting.styles import *
from formatting.dimensions import TITLE_ROW_HEIGHT, KPI_ROW_HEIGHT


def _title_block(ws, last_col_letter):
    ws.merge_cells(f"A1:{last_col_letter}1")
    c = ws["A1"]; c.value = COMPANY_NAME
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT

    ws.merge_cells(f"A2:{last_col_letter}2")
    c = ws["A2"]; c.value = COMPANY_SUBTITLE
    c.font = SUBTITLE_FONT; c.fill = SUBTITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[2].height = 22


def _kpi(ws, row, col_start, label, value_formula, fill_key, value_fmt=MONEY_FMT):
    """Write a KPI card: label + value in a merged 2x1 block."""
    label_cell = ws.cell(row=row, column=col_start, value=label)
    label_cell.font = KPI_LABEL_FONT
    label_cell.alignment = RTL_CENTER
    label_cell.fill = KPI_FILLS[fill_key]

    value_cell = ws.cell(row=row, column=col_start + 1, value=value_formula)
    value_cell.font = KPI_VALUE_FONT
    value_cell.alignment = RTL_CENTER
    value_cell.fill = KPI_FILLS[fill_key]
    value_cell.number_format = value_fmt

    # top + bottom borders for polish
    for cc in (label_cell, value_cell):
        cc.border = TABLE_BORDER_STYLE


def build_dashboard(wb):
    ws = wb.create_sheet(SHEETS["dashboard"], 0)
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = ACCENT

    last_col = "H"
    _title_block(ws, last_col)

    # Column widths
    for col in "ABCDEFGH":
        ws.column_dimensions[col].width = 22

    rev_sheet  = SHEETS["revenues"]
    exp_sheet  = SHEETS["expenses"]
    inst_sheet = SHEETS["institutions"]

    # ---------- Section: institution revenues ----------
    r = 4
    ws.merge_cells(f"A{r}:H{r}")
    c = ws[f"A{r}"]; c.value = "نظرة عامة على المؤسسات"
    c.font = SECTION_FONT; c.fill = SECTION_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[r].height = 26

    r += 1
    headers = ["المؤسسة", "إجمالي الإيرادات", "الإيرادات المحصّلة",
               "المبالغ المستحقة", "إجمالي المصروفات",
               "المصروفات المدفوعة", "الالتزامات", "صافي المركز المالي"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = RTL_CENTER; cell.border = TABLE_BORDER_STYLE
    ws.row_dimensions[r].height = 30

    header_row = r
    for idx, inst in enumerate(INSTITUTIONS, start=1):
        rr = r + idx
        inst_name = inst["name"]

        # SUMIFS against revenues/expenses sheets
        rev_range = f"'{rev_sheet}'!$B${2}:$B$100000"
        amt_range = f"'{rev_sheet}'!$F${2}:$F$100000"
        col_range = f"'{rev_sheet}'!$I${2}:$I$100000"

        exp_amt = f"'{exp_sheet}'!$F${2}:$F$100000"
        exp_inst = f"'{exp_sheet}'!$B${2}:$B$100000"
        exp_stat = f"'{exp_sheet}'!$I${2}:$I$100000"

        row_vals = [
            inst_name,
            f'=SUMIF({rev_range},"{inst_name}",{amt_range})',
            f'=SUMIFS({amt_range},{rev_range},"{inst_name}",{col_range},"مدفوع")',
            f'=SUMIFS({amt_range},{rev_range},"{inst_name}",{col_range},"مستحق")'
            f'+SUMIFS({amt_range},{rev_range},"{inst_name}",{col_range},"معلق")',
            f'=SUMIF({exp_inst},"{inst_name}",{exp_amt})',
            f'=SUMIFS({exp_amt},{exp_inst},"{inst_name}",{exp_stat},"مدفوع")',
            f'=SUMIFS({exp_amt},{exp_inst},"{inst_name}",{exp_stat},"مستحق")'
            f'+SUMIFS({exp_amt},{exp_inst},"{inst_name}",{exp_stat},"معلق")',
        ]
        # Net = collected - paid
        row_vals.append(f"=C{rr}-F{rr}")

        for i, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=rr, column=i, value=val)
            cell.alignment = RTL_CENTER if i == 1 else RTL_NUMBER
            cell.border = TABLE_BORDER_STYLE
            if i == 1:
                cell.font = BODY_BOLD
            else:
                cell.font = BODY_FONT
                cell.number_format = MONEY_FMT
            if rr % 2 == 0:
                cell.fill = ALT_ROW_FILL

    # Consolidated row
    total_row = header_row + len(INSTITUTIONS) + 1
    ws.cell(row=total_row, column=1, value="الإجمالي الموحّد").font = Font(
        name=ARABIC_FONT, bold=True, color=WHITE, size=12)
    for col in range(1, 9):
        c = ws.cell(row=total_row, column=col)
        c.fill = HEADER_FILL
        c.border = TABLE_BORDER_STYLE
        c.alignment = RTL_CENTER if col == 1 else RTL_NUMBER
        if col > 1:
            col_letter = get_column_letter(col)
            c.value = f"=SUM({col_letter}{header_row+1}:{col_letter}{total_row-1})"
            c.number_format = MONEY_FMT
            c.font = Font(name=ARABIC_FONT, bold=True, color=WHITE)

    # ---------- KPI cards row ----------
    kpi_row = total_row + 3
    ws.merge_cells(f"A{kpi_row}:H{kpi_row}")
    c = ws[f"A{kpi_row}"]; c.value = "المؤشرات الرئيسية"
    c.font = SECTION_FONT; c.fill = SECTION_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[kpi_row].height = 26

    kr = kpi_row + 1
    rev_amt = f"'{rev_sheet}'!$F$2:$F$100000"
    rev_stat= f"'{rev_sheet}'!$I$2:$I$100000"
    exp_amt = f"'{exp_sheet}'!$F$2:$F$100000"
    exp_stat= f"'{exp_sheet}'!$I$2:$I$100000"

    kpis = [
        ("إجمالي الإيرادات",       f"=SUM({rev_amt})", "revenue"),
        ("الإيرادات المحصّلة",     f'=SUMIF({rev_stat},"مدفوع",{rev_amt})', "revenue"),
        ("إجمالي المستحقات",       f'=SUMIF({rev_stat},"مستحق",{rev_amt})'
                                    f'+SUMIF({rev_stat},"معلق",{rev_amt})', "receiv"),
        ("إجمالي المصروفات",       f"=SUM({exp_amt})", "expense"),
        ("المصروفات المدفوعة",     f'=SUMIF({exp_stat},"مدفوع",{exp_amt})', "expense"),
        ("إجمالي الالتزامات",      f'=SUMIF({exp_stat},"مستحق",{exp_amt})'
                                    f'+SUMIF({exp_stat},"معلق",{exp_amt})', "liab"),
        ("الرصيد البنكي الفعلي",   f'=SUMIF({rev_stat},"مدفوع",{rev_amt})'
                                    f'-SUMIF({exp_stat},"مدفوع",{exp_amt})', "bank"),
        ("صافي المركز المالي",     f'=SUMIF({rev_stat},"مدفوع",{rev_amt})'
                                    f'-SUMIF({exp_stat},"مدفوع",{exp_amt})', "net"),
    ]

    for i, (label, formula, key) in enumerate(kpis):
        col = (i % 4) * 2 + 1
        row = kr + (i // 4) * 2
        _kpi(ws, row, col, label, formula, key)
        _kpi(ws, row + 1, col, "", "", key, value_fmt="@")  # spacer


    ws.freeze_panes = "A4"
    return ws