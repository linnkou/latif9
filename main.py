from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_from_directory,
)
import os
import shutil
from openpyxl import load_workbook

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = "your_secret_key_here"  # مفتاح سري للمصادقة
app.config["UPLOAD_FOLDER"] = "uploads"  # مجلد الملفات المرفوعة
app.config["PROCESSED_FOLDER"] = "processed"  # مجلد الملفات المعالجة

# التأكد من وجود المجلدات
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["PROCESSED_FOLDER"], exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # استلام الملف وعدد الأقسام
            file = request.files.get("file")
            num_classes = request.form.get("num_classes", type=int)

            # التحقق من صحة الإدخالات
            if not file or num_classes <= 0:
                flash("يرجى رفع ملف Excel وإدخال عدد الأقسام بشكل صحيح", "error")
                return redirect(url_for("index"))

            # حفظ النسخة الأصلية من الملف
            original_file_path = os.path.join(
                app.config["UPLOAD_FOLDER"], file.filename
            )
            file.save(original_file_path)

            # إنشاء نسخة جديدة لتعديل التقديرات
            processed_file_path = os.path.join(
                app.config["PROCESSED_FOLDER"], f"processed_{file.filename}"
            )
            shutil.copy(original_file_path, processed_file_path)

            # تعديل التقديرات على النسخة المعالجة
            add_grades_to_excel(processed_file_path, num_classes)

            flash("تم حفظ الملف الأصلي ومعالجة النسخة المعدلة بالتقديرات", "success")
            return render_template(
                "index.html",
                original_file=file.filename,
                processed_file=f"processed_{file.filename}",
                show_results=True,
            )
        except Exception as e:
            flash(f"حدث خطأ أثناء معالجة الملف: {str(e)}", "error")
            return redirect(url_for("index"))

    return render_template("index.html", show_results=False)


@app.route("/download/<file_type>/<filename>")
def download_file(file_type, filename):
    try:
        if file_type not in ["original", "processed"]:
            return "نوع الملف غير صالح", 400

        folder = (
            app.config["UPLOAD_FOLDER"]
            if file_type == "original"
            else app.config["PROCESSED_FOLDER"]
        )
        return send_from_directory(folder, filename, as_attachment=True)
    except Exception as e:
        flash(f"حدث خطأ أثناء تحميل الملف: {str(e)}", "error")
        return redirect(url_for("index"))


def add_grades_to_excel(file_path, num_classes):
    """
    وظيفة لإضافة التقديرات إلى ملف Excel.
    - file_path: مسار الملف.
    - num_classes: عدد الأقسام.
    """
    try:
        workbook = load_workbook(file_path)
        for index, sheet_name in enumerate(workbook.sheetnames[:num_classes]):
            sheet = workbook[sheet_name]
            for row in range(9, sheet.max_row + 1):  # بدءًا من الصف 9
                final_grade_cell = f"G{row}"  # على سبيل المثال
                sheet[final_grade_cell] = "تقدير مضاف"  # إضافة التقدير
        workbook.save(file_path)
    except Exception as e:
        raise Exception(f"خطأ أثناء تعديل ملف Excel: {str(e)}")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
