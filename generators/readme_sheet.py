"""A friendly 'كيفية الاستخدام' guide sheet."""
from config import SHEETS
from formatting.styles import *
from formatting.colors import PRIMARY, ACCENT


def build_guide(wb):
    ws = wb.create_sheet("دليل الاستخدام")
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = "A6A6A6"

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110

    ws.merge_cells("A1:B1")
    c = ws["A1"]; c.value = "دليل استخدام النظام المحاسبي"
    c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = RTL_CENTER
    ws.row_dimensions[1].height = 40

    guide = [
        ("خطوات البدء السريعة:", True),
        ("1. ابدأ بمراجعة بيانات المؤسسات في ورقة «المؤسسات».", False),
        ("2. سجّل إيراداتك في ورقة «الإيرادات»، مع تحديد حالة الدفع بدقة.", False),
        ("3. سجّل مصروفاتك في ورقة «المصروفات»، مع تحديد حالة الدفع بدقة.", False),
        ("4. أضف المركبات والعملاء في أوراق «مركبات المؤسسة 1/2/3».", False),
        ("5. سجّل الرسوم السنوية في ورقة «الرسوم السنوية» لمتابعة المتبقي منها.", False),
        ("", False),
        ("القاعدة المحاسبية الأساسية:", True),
        ("• الرصيد البنكي الفعلي يعتمد فقط على المبالغ المحصّلة والمصروفات المدفوعة فعلياً.", False),
        ("• المبالغ المستحقة والمعلقة تُعرض بشكل منفصل ولا تُحسب كنقدية متاحة.", False),
        ("", False),
        ("الرصيد البنكي الفعلي = الرصيد الافتتاحي + المحصّل − المدفوع", True),
        ("صافي المركز المالي = إجمالي المحصّل − إجمالي المدفوع", True),
        ("المبلغ المتبقي على العميل = المبلغ المستحق − المحصّل", True),
        ("المبلغ المتبقي من المصروف = المبلغ المستحق − المدفوع", True),
        ("", False),
        ("حالات الدفع المتاحة:", True),
        ("مدفوع  •  مستحق  •  معلق  •  مدفوع جزئيا", False),
        ("", False),
        ("نصائح:", True),
        ("• استخدم القوائم المنسدلة دائماً لضمان دقة البيانات.", False),
        ("• يمكن إضافة صفوف جديدة ضمن النطاق المخصص لكل جدول دون التأثير على المعادلات.", False),
        ("• ورقة «التقارير» تتيح تغيير الفترة الزمنية بتعديل تاريخي البداية والنهاية.", False),
    ]

    r = 3
    for text, bold in guide:
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = RTL_ALIGN
        c.font = SECTION_FONT if bold else BODY_FONT
        r += 1

    ws.sheet_view.showGridLines = False
    return ws