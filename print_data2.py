import os
import xlwings as xw
from pathlib import Path
from tkinter import Tk, filedialog

# ==== Chọn thư mục chứa file Excel ====
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel")

if not folder_selected:
    print("❌ Không chọn thư mục nào, thoát chương trình.")
    exit()

folder = Path(folder_selected)

# ==== Lấy danh sách file Excel ====
excel_files = list(folder.glob("*.xls*"))  # gồm .xlsx, .xlsm, .xls

if not excel_files:
    print("❌ Không tìm thấy file Excel nào trong thư mục.")
    exit()

# ==== Tạo thư mục con để lưu PDF ====
pdf_folder = folder / "PDF_Export"
pdf_folder.mkdir(exist_ok=True)

# ==== Xuất PDF ====
app = xw.App(visible=False)
for f in excel_files:
    try:
        wb = app.books.open(f)
        first_sheet = wb.sheets[0]  # chỉ lấy sheet đầu tiên
        pdf_path = pdf_folder / (f.stem + ".pdf")  # lưu trong folder PDF_Export
        first_sheet.api.ExportAsFixedFormat(0, str(pdf_path))
        wb.close()
        print(f"✅ Đã chuyển: {f.name} → {pdf_path.name}")
    except Exception as e:
        print(f"⚠️ Lỗi khi xử lý {f.name}: {e}")

app.quit()

print(f"🎉 Hoàn tất! Tất cả PDF được lưu trong thư mục: {pdf_folder}")
