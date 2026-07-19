# gui/app_window.py
# -----------------------------------------------------------------------------
# يشغّل الواجهة الجديدة (HTML/CSS/JS الموجودة في gui/web) داخل نافذة تطبيق
# مكتبي حقيقية باستخدام مكتبة pywebview (وليست نافذة متصفح).
#
# هذه الوحدة تكشف "Api" لجافاسكريبت: أي زر في الواجهة يستدعي دالة بايثون
# هنا عبر window.pywebview.api.اسم_الدالة() من ملف web/app.js.
# -----------------------------------------------------------------------------

import os
import sys
import base64
import subprocess
from datetime import date

import webview

from excel.template_generator import generate_attendance_template, ARABIC_MONTHS

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")
INDEX_HTML = os.path.join(WEB_DIR, "index.html")

IMAGE_FILE_TYPES = ("ملفات الصور (*.jpg;*.jpeg;*.png;*.bmp;*.webp)", "كل الملفات (*.*)")


class Api:
    """
    كل دالة هنا تُستدعى مباشرة من جافاسكريبت (gui/web/app.js).
    القاعدة: كل دالة ترجع قيمة بسيطة (نص/رقم/قاموس) قابلة للتحويل إلى JSON.
    """

    # -------------------------------------------------------------------
    def get_months(self):
        """قائمة أسماء الأشهر بالعربي، لتعبئة القائمة المنسدلة في الواجهة."""
        return list(ARABIC_MONTHS.values())

    # -------------------------------------------------------------------
    def get_defaults(self):
        """القيم الافتراضية عند فتح البرنامج: الشهر والسنة الحاليان، ومجلد الحفظ الافتراضي."""
        today = date.today()
        return {
            "month": today.month,
            "year": today.year,
            "outputDir": DEFAULT_OUTPUT_DIR,
        }

    # -------------------------------------------------------------------
    def choose_folder(self):
        """يفتح نافذة اختيار مجلد نظام التشغيل الأصلية، ويرجع المسار المختار."""
        result = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG, directory=DEFAULT_OUTPUT_DIR
        )
        if result:
            # بعض الأنظمة ترجع نصًا مباشرة، وأخرى ترجع قائمة تحتوي مسارًا واحدًا
            return result[0] if isinstance(result, (list, tuple)) else result
        return None

    # -------------------------------------------------------------------
    def choose_image(self):
        """
        يفتح نافذة اختيار صورة، ويرجع مسارها واسمها وصورة مصغّرة كـ Base64
        لعرضها في المعاينة داخل الواجهة مباشرة (بدون حفظ نسخة إضافية).
        """
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, file_types=IMAGE_FILE_TYPES
        )
        if not result:
            return None

        path = result[0] if isinstance(result, (list, tuple)) else result
        try:
            with open(path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            extension = os.path.splitext(path)[1].lower().replace(".", "") or "png"
            if extension == "jpg":
                extension = "jpeg"
            return {
                "path": path,
                "name": os.path.basename(path),
                "dataUrl": f"data:image/{extension};base64,{encoded}",
            }
        except Exception as error:
            return {"error": str(error)}

    # -------------------------------------------------------------------
    def check_exists(self, folder, year, month):
        """يتحقق مسبقًا هل يوجد ملف بنفس اسم الشهر، لتُظهر الواجهة تأكيدًا قبل الاستبدال."""
        file_name = self._build_file_name(year, month)
        full_path = os.path.join(folder or DEFAULT_OUTPUT_DIR, file_name)
        return {"exists": os.path.exists(full_path), "fileName": file_name}

    # -------------------------------------------------------------------
    def create_attendance(self, options):
        """ينشئ ملف الإكسل الفعلي بناءً على البيانات القادمة من الواجهة."""
        try:
            month_index = int(options.get("month"))
            year = int(options.get("year"))
            employee_name = (options.get("name") or "").strip()
            phone = (options.get("phone") or "").strip()
            folder = options.get("folder") or DEFAULT_OUTPUT_DIR
            open_after = bool(options.get("openAfter"))

            os.makedirs(folder, exist_ok=True)

            file_name = self._build_file_name(year, month_index)
            full_path = os.path.join(folder, file_name)

            generate_attendance_template(
                month=month_index, year=year,
                employee_name=employee_name, phone=phone,
                output_path=full_path,
            )

            if open_after:
                self._open_file(full_path)

            return {"success": True, "path": full_path, "fileName": file_name}

        except Exception as error:
            return {"success": False, "error": str(error)}

    # -------------------------------------------------------------------
    @staticmethod
    def _build_file_name(year, month_index):
        return f"attendance_{int(year)}_{int(month_index):02d}.xlsx"

    # -------------------------------------------------------------------
    @staticmethod
    def _open_file(path: str):
        """يفتح ملف الإكسل الناتج ببرنامج الإكسل الافتراضي حسب نظام التشغيل."""
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # خاص بويندوز فقط
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception:
            # فشل الفتح التلقائي لا يجب أن يوقف البرنامج
            pass


def main():
    api = Api()
    webview.create_window(
        "كشف الحضور الذكي",
        INDEX_HTML,
        js_api=api,
        width=1080,
        height=720,
        min_size=(820, 600),
    )
    webview.start()


if __name__ == "__main__":
    main()
