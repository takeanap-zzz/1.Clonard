import shutil
from pathlib import Path
from datetime import datetime, timedelta
import re
import xlwings as xw

# File Excel gốc
src = Path(r"D:\1.Clonard\8w\Dec 22 - Dec 28 2025 Weekly Timesheet Input.xlsx")

# Thư mục chứa bản sao
dst_folder = src.parent
dst_folder.mkdir(parents=True, exist_ok=True)

# Regex tách ngày + giữ nguyên suffix
match = re.search(r"([A-Za-z]+ \d{1,2})\s*-\s*([A-Za-z]+ \d{1,2}) (\d{4})(.*)", src.stem)
if not match:
    raise ValueError("❌ Không tìm thấy ngày tháng trong tên file!")

start_str, end_str, year, suffix = match.groups()

# Parse datetime (hỗ trợ viết tắt và đầy đủ của tháng)
try:
    start_date = datetime.strptime(f"{start_str} {year}", "%b %d %Y")
    end_date = datetime.strptime(f"{end_str} {year}", "%b %d %Y")
except ValueError:
    start_date = datetime.strptime(f"{start_str} {year}", "%B %d %Y")
    end_date = datetime.strptime(f"{end_str} {year}", "%B %d %Y")

# Tạo 6 bản copy (6 tuần tiếp theo)
for i in range(1, 7):
    new_start = start_date + timedelta(days=7 * i)
    new_end = end_date + timedelta(days=7 * i)

    new_filename = (
        f"{new_start.strftime('%b %d')} - {new_end.strftime('%b %d')} {new_end.year}{suffix}.xlsx"
    )
    new_path = dst_folder / new_filename

    # Copy file gốc trước
    shutil.copy2(src, new_path)

    # --- Update D5 bằng xlwings ---
    app = xw.App(visible=False)   # chạy ngầm Excel
    wb = xw.Book(new_path)
    ws = wb.sheets[0]             # sheet đầu tiên (có thể đổi tên sheet nếu cần)

    # Ghi ngày dạng long date
    #ws.range("D5").value = new_start.strftime("%A, %B %d, %Y")
    # Ghi giá trị datetime thật
    ws.range("D5").value = new_start

    # Áp dụng định dạng Long Date
    ws.range("D5").api.NumberFormat = "dddd, mmmm dd, yyyy"

    wb.save()
    wb.close()
    app.quit()

