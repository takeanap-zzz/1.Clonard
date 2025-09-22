import os
import re
import xlwings as xw
from PyPDF2 import PdfMerger
from pathlib import Path
from tkinter import Tk, filedialog
import subprocess
import platform
import datetime

# ==== Hộp thoại chọn nhiều file Excel ====
root = Tk()
root.withdraw()
files = filedialog.askopenfilenames(
    title="Chọn các file Excel",
    filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
)

if not files:
    print("❌ Không chọn file nào, thoát chương trình.")
    exit()

excel_files = [Path(f) for f in files]
folder = excel_files[0].parent  # lấy thư mục chứa file đầu tiên

# ==== Regex bắt ngày cuối cùng trong khoảng: Aug 11 - Aug 17 2025 ====
date_pattern = re.compile(r"[A-Za-z]{3}\s+\d{1,2}\s*-\s*([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})")

# Map tháng viết tắt sang số
month_map = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}

# ==== Xuất Excel -> PDF (chỉ sheet đầu tiên) ====
pdf_files = []
app = xw.App(visible=False)
for f in excel_files:
    match = date_pattern.search(f.stem)
    if not match:
        print(f"⚠️ Bỏ qua file không nhận diện ngày cuối: {f.name}")
        continue

    month_abbr, day, year = match.groups()
    month = month_map.get(month_abbr[:3].title())  # chuẩn hóa 3 ký tự
    if not month:
        print(f"⚠️ Không nhận diện được tháng trong file: {f.name}")
        continue

    # Tạo object datetime từ ngày cuối
    date_obj = datetime.date(int(year), month, int(day))
    date_key = date_obj.strftime("%Y%m%d")  # YYYYMMDD để sort

    # Xuất PDF (sheet đầu tiên)
    pdf_path = folder / f"{date_key}_{f.stem}.pdf"
    wb = app.books.open(f)
    first_sheet = wb.sheets[0]
    first_sheet.api.ExportAsFixedFormat(0, str(pdf_path))  # 0 = PDF
    wb.close()
    pdf_files.append((date_key, pdf_path))

app.quit()

# ==== Sắp xếp PDF theo ngày cuối ====
pdf_files.sort(key=lambda x: x[0])

# ==== Gộp PDF ====
if pdf_files:
    merger = PdfMerger()
    for _, pdf in pdf_files:
        merger.append(str(pdf))

    output_pdf = folder / "Labour monthly.pdf"
    merger.write(str(output_pdf))
    merger.close()

    # ==== Xóa các file PDF lẻ ====
    for _, pdf in pdf_files:
        try:
            os.remove(pdf)
        except Exception as e:
            print(f"⚠️ Không xóa được {pdf}: {e}")

    print(f"✅ Đã tạo file gộp: {output_pdf}")
    print("🗑️ Đã xóa các file PDF lẻ.")

    # ==== Mở file PDF sau khi tạo ====
    try:
        if platform.system() == "Windows":
            os.startfile(output_pdf)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", output_pdf])
        else:  # Linux
            subprocess.run(["xdg-open", output_pdf])
    except Exception as e:
        print(f"⚠️ Không mở được file PDF: {e}")

else:
    print("❌ Không có file PDF nào được tạo.")
