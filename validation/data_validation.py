"""Data validation helpers using openpyxl."""
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import quote_sheetname, absolute_coordinate


def _list_formula(values, lists_sheet):
    """Return a formula string pointing at a range in the lists sheet."""
    # values are written to lists sheet columns; we reference them by range
    return None


def dv_list(values, allow_blank=True, error="القيمة غير مسموح بها", prompt=None):
    """Build an inline-list DataValidation (short lists only)."""
    joined = '"' + ",".join(values) + '"'
    dv = DataValidation(
        type="list",
        formula1=joined,
        allow_blank=allow_blank,
        showErrorMessage=True,
        errorTitle="خطأ في الإدخال",
        error=error,
        showInputMessage=bool(prompt),
        promptTitle="تلميح" if prompt else None,
        prompt=prompt,
    )
    return dv


def dv_range_list(range_ref, allow_blank=True):
    """DV from a named/absolute range, e.g. 'القوائم والإعدادات'!$A$2:$A$5"""
    dv = DataValidation(
        type="list",
        formula1=range_ref,
        allow_blank=allow_blank,
        showErrorMessage=True,
        errorTitle="خطأ في الإدخال",
        error="الرجاء اختيار قيمة من القائمة.",
    )
    return dv


def dv_decimal(minimum=0, maximum=None, allow_blank=True):
    return DataValidation(
        type="decimal",
        operator="between" if maximum is not None else "greaterThanOrEqual",
        formula1=minimum,
        formula2=maximum,
        allow_blank=allow_blank,
        showErrorMessage=True,
        errorTitle="قيمة غير صحيحة",
        error="الرجاء إدخال قيمة رقمية صحيحة.",
    )


def dv_date(allow_blank=True):
    return DataValidation(
        type="date",
        operator="greaterThanOrEqual",
        formula1="1900-01-01",
        allow_blank=allow_blank,
        showErrorMessage=True,
        errorTitle="تاريخ غير صحيح",
        error="الرجاء إدخال تاريخ صحيح.",
    )


def add_dv(ws, dv, cell_range):
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv