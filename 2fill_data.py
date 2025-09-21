import xlwings as xw
from tkinter import filedialog, Tk
from pathlib import Path
import datetime

def to_str_date(val):
    """Chuyển giá trị ngày sang chuỗi dd/mm/yyyy (nếu có)"""
    if isinstance(val, datetime.datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, datetime.date):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, (int, float)):
        try:
            # Excel serial date
            base = datetime.datetime(1899, 12, 30)
            return (base + datetime.timedelta(days=val)).strftime("%d/%m/%Y")
        except:
            return str(val)
    if val:
        return str(val).strip()
    return None

# ==== Chọn file summary ====
root = Tk()
root.withdraw()
summary_path = filedialog.askopenfilename(
    title="Chọn file SUMMARY",
    filetypes=[("Excel Files", "*.xlsx *.xlsm")]
)
if not summary_path:
    print("❌ Không chọn file summary.")
    exit()

# ==== Chọn file nguồn ====
source_files = filedialog.askopenfilenames(
    title="Chọn các file nguồn",
    filetypes=[("Excel Files", "*.xlsx *.xlsm")]
)
if not source_files:
    print("❌ Không chọn file nguồn.")
    exit()

# ==== Mapping cột C..G ====
summary_columns = [
    ("C9", "C10", "C12:C16"),
    ("D9", "D10", "D12:D16"),
    ("E9", "E10", "E12:E16"),
    ("F9", "F10", "F12:F16"),
    ("G9", "G10", "G12:G16"),
]

app = xw.App(visible=False)
error_files = []

try:
    wb_summary = xw.Book(summary_path)
    ws_summary = wb_summary.sheets[0]
    filled_columns = {}

    for src in source_files:
        matched = False
        wb_src = xw.Book(src)
        ws_src = wb_src.sheets[0]

        src_date1 = to_str_date(ws_src.range("D5").value)
        src_date2 = to_str_date(ws_src.range("D6").value)

        print(f"\n📄 Đang xử lý: {Path(src).name}")
        print(f"   → Ngày nguồn: {src_date1} - {src_date2}")

        for check_cell1, check_cell2, paste_range in summary_columns:
            summary_date1 = to_str_date(ws_summary.range(check_cell1).value)
            summary_date2 = to_str_date(ws_summary.range(check_cell2).value)

            print(f"   So với {check_cell1}:{summary_date1}, {check_cell2}:{summary_date2}")

            if (src_date1 == summary_date1) and (src_date2 == summary_date2):
                if paste_range in filled_columns:
                    error_files.append((Path(src).name,
                        f"trùng ngày với {paste_range} (đã copy từ {filled_columns[paste_range]})"))
                else:
                    data = ws_src.range("D8:D12").value
                    ws_summary.range(paste_range).value = data
                    filled_columns[paste_range] = Path(src).name
                    print(f"✅ Copy vào {paste_range}")
                matched = True
                break

        if not matched:
            error_files.append((Path(src).name, "❌ không khớp ngày nào"))

        wb_src.close()

    wb_summary.save()
    print("\n🎉 Hoàn thành cập nhật Summary.xlsx")

    if error_files:
        print("\n❌ Các file lỗi:")
        for f, reason in error_files:
            print(f"   - {f}: {reason}")

finally:
    wb_summary.close()
    app.quit()
