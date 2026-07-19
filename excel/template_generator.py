# excel_template.py
# -----------------------------------------------------------------------------
# هذه الوحدة مسؤولة عن توليد قالب ملف الإكسل لكشف الحضور والانصراف.
# تستقبل الشهر والسنة، وتحسب تلقائيًا:
#   - عدد أيام الشهر (28 حتى 31 حسب الشهر والسنة)
#   - اسم اليوم بالعربي لكل تاريخ (الأحد، الإثنين، ...)
# ثم تكتب كل ذلك في ملف Excel بنفس ترتيب وشكل الجدول الأصلي (اتجاه من اليمين لليسار).
#
# لاحقًا: وحدة OCR ستملأ فقط عمودي "وقت الدخول" و"وقت الخروج"،
# وباقي الجدول (التاريخ، اسم اليوم، المعادلات) يبقى كما هو.
# -----------------------------------------------------------------------------

import calendar
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# خريطة تحويل رقم اليوم (weekday) إلى اسمه بالعربي
# ملاحظة: في بايثون، date.weekday() تُرجع: الإثنين = 0 ... الأحد = 6
ARABIC_WEEKDAYS = {
    0: "الإثنين",
    1: "الثلاثاء",
    2: "الأربعاء",
    3: "الخميس",
    4: "الجمعة",
    5: "السبت",
    6: "الأحد",
}

ARABIC_MONTHS = {
    1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
    5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
    9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر",
}

# ألوان وأنماط مستخدمة في القالب
HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # خلايا يجب تعبئتها يدويًا
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def generate_attendance_template(month: int, year: int, employee_name: str = "",
                                  phone: str = "", output_path: str = "attendance_template.xlsx"):
    """
    يولّد ملف إكسل لكشف حضور شهر كامل.

    المعاملات:
        month: رقم الشهر (1-12)
        year: السنة (مثال: 2026)
        employee_name: اسم الموظف (اختياري، يمكن تركه فارغًا للتعبئة اليدوية)
        phone: رقم الهاتف (اختياري)
        output_path: مسار حفظ الملف الناتج
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "كشف الحضور"

    # جعل اتجاه الورقة من اليمين لليسار (RTL) ليطابق الجدول الأصلي
    ws.sheet_view.rightToLeft = True

    # ------------------------------------------------------------------
    # الصف الأول: عنوان الكشف
    # ------------------------------------------------------------------
    ws.merge_cells("A1:E1")
    ws["A1"] = f"كشف الحضور والانصراف - شهر {ARABIC_MONTHS[month]} {year}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = CENTER

    # ------------------------------------------------------------------
    # الصف الثاني: بيانات الموظف (اسم / هاتف)
    # ------------------------------------------------------------------
    ws["A2"] = "الاسم:"
    ws["A2"].font = Font(bold=True)
    ws.merge_cells("B2:C2")
    ws["B2"] = employee_name
    ws["B2"].fill = INPUT_FILL

    ws["D2"] = "رقم الهاتف:"
    ws["D2"].font = Font(bold=True)
    ws["E2"] = phone
    ws["E2"].fill = INPUT_FILL

    # ------------------------------------------------------------------
    # صف عنوان الجدول (الصف 4)
    # ترتيب الأعمدة من اليمين لليسار (لأن الورقة RTL): A هي أقصى اليمين
    # A: التاريخ | B: وقت الدخول | C: وقت الخروج | D: مجموع الساعات | E: الملاحظات
    # ------------------------------------------------------------------
    header_row = 4
    headers = ["التاريخ", "وقت الدخول", "وقت الخروج", "مجموع الساعات", "الملاحظات"]
    for col_idx, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # ------------------------------------------------------------------
    # صفوف الأيام: تُحسب تلقائيًا حسب عدد أيام الشهر المُدخل
    # ------------------------------------------------------------------
    days_in_month = calendar.monthrange(year, month)[1]  # عدد أيام الشهر (28-31)
    first_data_row = header_row + 1

    for day_num in range(1, days_in_month + 1):
        current_date = date(year, month, day_num)
        day_name = ARABIC_WEEKDAYS[current_date.weekday()]
        row = first_data_row + day_num - 1

        # عمود التاريخ: اسم اليوم + التاريخ بصيغة شهر/يوم (كما في الجدول الأصلي)
        date_cell = ws.cell(row=row, column=1, value=f"{day_name}  {month:02d}/{day_num:02d}")
        date_cell.alignment = CENTER
        date_cell.border = THIN_BORDER

        # عمودا وقت الدخول والخروج: يُتركان فارغين ليملأهما الـ OCR لاحقًا
        for col in (2, 3):
            c = ws.cell(row=row, column=col)
            c.fill = INPUT_FILL
            c.alignment = CENTER
            c.border = THIN_BORDER
            c.number_format = "HH:MM"

        # عمود مجموع الساعات: معادلة تحسب الفرق بين الخروج والدخول
        # نستخدم MOD() لمعالجة حالات الدخول والخروج عبر منتصف الليل
        hours_cell = ws.cell(row=row, column=4)
        hours_cell.value = (
            f'=IF(AND(B{row}<>"",C{row}<>""),MOD(C{row}-B{row},1)*24,"")'
        )
        hours_cell.number_format = "0.00"
        hours_cell.alignment = CENTER
        hours_cell.border = THIN_BORDER

        # عمود الملاحظات: فارغ للتعبئة اليدوية (إجازة، خصم، سلفة... إلخ)
        notes_cell = ws.cell(row=row, column=5)
        notes_cell.alignment = CENTER
        notes_cell.border = THIN_BORDER

    # ------------------------------------------------------------------
    # صف الإجمالي الشهري في نهاية الجدول
    # ------------------------------------------------------------------
    total_row = first_data_row + days_in_month
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    total_label = ws.cell(row=total_row, column=1, value="إجمالي ساعات العمل الشهرية")
    total_label.font = Font(bold=True)
    total_label.alignment = CENTER
    total_label.fill = HEADER_FILL
    total_label.border = THIN_BORDER

    total_value = ws.cell(row=total_row, column=4)
    total_value.value = f"=SUM(D{first_data_row}:D{total_row - 1})"
    total_value.number_format = "0.00"
    total_value.font = Font(bold=True)
    total_value.alignment = CENTER
    total_value.fill = HEADER_FILL
    total_value.border = THIN_BORDER

    ws.cell(row=total_row, column=5).border = THIN_BORDER
    ws.cell(row=total_row, column=5).fill = HEADER_FILL

    # ------------------------------------------------------------------
    # ضبط عرض الأعمدة
    # ------------------------------------------------------------------
    widths = {1: 20, 2: 14, 3: 14, 4: 16, 5: 22}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    # مثال تجريبي: توليد كشف لشهر أبريل 2026 (نفس الشهر الموجود في الصورة)
    path = generate_attendance_template(
        month=4, year=2026,
        employee_name="", phone="",
        output_path="output/attendance_2026_04.xlsx",
    )
    print(f"تم إنشاء الملف: {path}")
