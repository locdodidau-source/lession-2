from flask import Flask, render_template, request, redirect, url_for, flash
import os
from read_excel import doc_tkb
from google_calendar import dang_nhap_google, tao_su_kien, xoa_su_kien_tkb

app = Flask(__name__)
app.secret_key = "secret_key_demo"  # cần cho flash message

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Lấy dữ liệu từ form
        remind_value = int(request.form.get("remind_value", 10))
        remind_unit = request.form.get("remind_unit", "phút")
        remind_method = request.form.get("remind_method", "popup")
        prefix = request.form.get("prefix", "[TKB]")

        # Đổi sang phút
        if remind_unit == "giờ":
            remind_value *= 60
        elif remind_unit == "ngày":
            remind_value *= 1440

        # Xử lý file upload
        if "file_excel" not in request.files:
            flash("Chưa chọn file Excel")
            return redirect(url_for("index"))

        file = request.files["file_excel"]
        if file.filename == "":
            flash("File Excel không hợp lệ")
            return redirect(url_for("index"))

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Gọi hàm xử lý lịch
        try:
            events = doc_tkb(file_path)
            service = dang_nhap_google()
            reminders = [{"method": remind_method, "minutes": remind_value}]

            for e in events:
                if not e["gio_bd"] or not e["gio_kt"] or not e["thu"]:
                    continue

                tao_su_kien(
                    service=service,
                    mon=e["mon"],
                    phong=e["phong"],
                    giang_vien=e["giang_vien"],
                    start_date=e["ngay_bat_dau"],
                    end_date=e["ngay_ket_thuc"],
                    weekday=e["thu"],
                    start_time=e["gio_bd"],
                    end_time=e["gio_kt"],
                    reminders=reminders,
                    prefix=prefix
                )

            flash("✅ Đã lên lịch thời khóa biểu thành công!")
        except Exception as ex:
            flash(f"❌ Lỗi: {str(ex)}")

        return redirect(url_for("index"))

    return render_template("index.html")
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Heroku cung cấp PORT
    app.run(host="0.0.0.0", port=port, debug=True)