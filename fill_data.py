import xlwings as xw
from tkinter import filedialog, Tk
from pathlib import Path

# ==== Chọn file summary trước ====
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

# ==== Mapping vị trí trong summary ====
summary_targets = [
    ("C33", "C34", "C37:F41"),  # block 1
    ("C45", "C46", "C49:F53"),  # block 2
    ("C57", "C58", "C61:F65"),  # block 3
    ("C69", "C70", "C73:F77"),  # block 4
    ("C81", "C82", "C85:F89"),  # block 5
]

# ==== Bắt đầu xử lý ====
app = xw.App(visible=False)
error_files = []
try:
    wb_summary = xw.Book(summary_path)
    ws_summary = wb_summary.sheets[0]  # sheet đầu tiên

    for src in source_files:
        matched = False  # flag kiểm tra
        wb_src = xw.Book(src)
        ws_src = wb_src.sheets[0]

        src_date1 = ws_src.range("D5").value
        src_date2 = ws_src.range("D6").value

        for check_cell1, check_cell2, paste_range in summary_targets:
            summary_date1 = ws_summary.range(check_cell1).value
            summary_date2 = ws_summary.range(check_cell2).value

            if (src_date1 == summary_date1) and (src_date2 == summary_date2):
                data = ws_src.range("E16:H20").value
                ws_summary.range(paste_range).value = data
                print(f"✅ {Path(src).name}: Copy vào {paste_range}")
                matched = True
                break  # thoát vòng for vì đã tìm đúng chỗ

        if not matched:
            error_files.append(Path(src).name)
            print(f"❌ {Path(src).name}: Ngày không khớp với bất kỳ block nào")

        wb_src.close()

    wb_summary.save()
    print("🎉 Hoàn thành cập nhật Summary.xlsx")

    if error_files:
        print("\n⚠️ Các file bị lỗi ngày:")
        for f in error_files:
            print(f"   - {f}")

finally:
    wb_summary.close()
    app.quit()
