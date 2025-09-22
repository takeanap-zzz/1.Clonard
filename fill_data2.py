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
            base = datetime.datetime(1899, 12, 30)  # Excel base date
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

# ==== Chọn các file nguồn ====
source_files = filedialog.askopenfilenames(
    title="Chọn các file nguồn",
    filetypes=[("Excel Files", "*.xlsx *.xlsm")]
)
if not source_files:
    print("❌ Không chọn file nguồn.")
    exit()

# ==== Mapping C..G ====
summary_columns = [
    ("C9", "C10", "C12:C16"),
    ("D9", "D10", "D12:D16"),
    ("E9", "E10", "E12:E16"),
    ("F9", "F10", "F12:F16"),
    ("G9", "G10", "G12:G16"),
]

app = xw.App(visible=False)
error_files = []
used_blocks = {}

try:
    wb_summary = xw.Book(summary_path)
    ws_summary = wb_summary.sheets[0]

    for src in source_files:
        matched = False
        wb_src = xw.Book(src)
        ws_src = wb_src.sheets[0]

        src_date1 = to_str_date(ws_src.range("D5").value)
        src_date2 = to_str_date(ws_src.range("D6").value)

        for check_cell1, check_cell2, paste_range in summary_columns:
            sum_date1 = to_str_date(ws_summary.range(check_cell1).value)
            sum_date2 = to_str_date(ws_summary.range(check_cell2).value)

            if src_date1 == sum_date1 and src_date2 == sum_date2:
                if paste_range in used_blocks:
                    error_files.append(f"{Path(src).name}: trùng ngày với {used_blocks[paste_range]}")
                else:
                    data = ws_src.range("D8:D12").value
                    # ép thành cột dọc
                    data = [[v] for v in data]
                    ws_summary.range(paste_range).value = data
                    used_blocks[paste_range] = Path(src).name
                    print(f"✅ {Path(src).name} → {paste_range}")
                matched = True
                break

        if not matched:
            error_files.append(f"{Path(src).name}: ❌ không khớp ngày nào")

        wb_src.close()

    wb_summary.save()
    print("\n🎉 Hoàn thành cập nhật Summary.xlsx")

    if error_files:
        print("\n❌ Lỗi:")
        for f in error_files:
            print("   -", f)

finally:
    wb_summary.close()
    app.quit()
