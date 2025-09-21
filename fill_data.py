# chọn file summary, chọn các file nguồn, copy dữ liệu từ các file nguồn vào đúng vị trí trong summary dựa trên ngày tháng


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
error_files = []  # (tên file, lý do)
try:
    wb_summary = xw.Book(summary_path)
    ws_summary = wb_summary.sheets[0]

    filled_blocks = {}  # ghi nhớ block đã được dùng

    for src in source_files:
        matched = False
        wb_src = xw.Book(src)
        ws_src = wb_src.sheets[0]

        src_date1 = ws_src.range("D5").value
        src_date2 = ws_src.range("D6").value

        for check_cell1, check_cell2, paste_range in summary_targets:
            summary_date1 = ws_summary.range(check_cell1).value
            summary_date2 = ws_summary.range(check_cell2).value

            if (src_date1 == summary_date1) and (src_date2 == summary_date2):
                # kiểm tra trùng ngày
                if paste_range in filled_blocks:
                    error_files.append((
                        Path(src).name,
                        f"trùng ngày với block {paste_range} (đã copy từ {filled_blocks[paste_range]})"
                    ))
                else:
                    data = ws_src.range("E16:H20").value
                    ws_summary.range(paste_range).value = data
                    filled_blocks[paste_range] = Path(src).name
                    print(f"✅ {Path(src).name}: Copy vào {paste_range}")
                matched = True
                break

        if not matched:
            error_files.append((Path(src).name, "không khớp ngày nào trong summary"))

        wb_src.close()

    wb_summary.save()
    print("🎉 Hoàn thành cập nhật Summary.xlsx")

    if error_files:
        print("\n❌ Các file lỗi:")
        for f, reason in error_files:
            print(f"   - {f}: {reason}")

finally:
    wb_summary.close()
    app.quit()

