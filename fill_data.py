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

# ==== Mapping vị trí trong summary (Loại 1: Block lớn 4x5) ====
summary_targets_large = [
    ("C33", "C34", "C37:F41"),  # block 1
    ("C45", "C46", "C49:F53"),  # block 2
    ("C57", "C58", "C61:F65"),  # block 3
    ("C69", "C70", "C73:F77"),  # block 4
    ("C81", "C82", "C85:F89"),  # block 5
]

# ==== Mapping vị trí trong summary (Loại 2: Cột dọc) ====
summary_columns = [
    ("C9", "C10", "C12:C16"),
    ("D9", "D10", "D12:D16"),
    ("E9", "E10", "E12:E16"),
    ("F9", "F10", "F12:F16"),
    ("G9", "G10", "G12:G16"),
]

# ==== Bắt đầu xử lý ====
app = xw.App(visible=False)
error_files = []  # (tên file, lý do)
try:
    wb_summary = xw.Book(summary_path)
    ws_summary = wb_summary.sheets[0]

    filled_blocks_large = {}  # ghi nhớ block lớn đã được dùng
    used_blocks_column = {}   # ghi nhớ cột dọc đã được dùng

    for src in source_files:
        matched_large = False
        matched_column = False
        wb_src = xw.Book(src)
        ws_src = wb_src.sheets[0]

        # Đọc ngày từ file nguồn
        src_date1 = to_str_date(ws_src.range("D5").value)
        src_date2 = to_str_date(ws_src.range("D6").value)

        print(f"\n🔍 Xử lý file: {Path(src).name}")
        print(f"   Ngày 1: {src_date1}, Ngày 2: {src_date2}")

        # ==== KIỂM TRA MAPPING LOẠI 1: Block lớn 4x5 (E16:H20 → C37:F41, etc.) ====
        for check_cell1, check_cell2, paste_range in summary_targets_large:
            summary_date1 = to_str_date(ws_summary.range(check_cell1).value)
            summary_date2 = to_str_date(ws_summary.range(check_cell2).value)

            if (src_date1 == summary_date1) and (src_date2 == summary_date2):
                # kiểm tra trùng ngày
                if paste_range in filled_blocks_large:
                    error_files.append((
                        Path(src).name,
                        f"trùng ngày với block lớn {paste_range} (đã copy từ {filled_blocks_large[paste_range]})"
                    ))
                else:
                    # Copy dữ liệu từ vùng E16:H20 (4x5)
                    data = ws_src.range("E16:H20").value
                    if data:
                        ws_summary.range(paste_range).value = data
                        filled_blocks_large[paste_range] = Path(src).name
                        print(f"✅ {Path(src).name}: Copy block lớn vào {paste_range}")
                        matched_large = True
                    else:
                        error_files.append((Path(src).name, f"không có dữ liệu trong vùng E16:H20"))
                # KHÔNG break ở đây - tiếp tục kiểm tra các block khác

        # ==== KIỂM TRA MAPPING LOẠI 2: Cột dọc (D8:D12 → C12:C16, etc.) ====
        for check_cell1, check_cell2, paste_range in summary_columns:
            sum_date1 = to_str_date(ws_summary.range(check_cell1).value)
            sum_date2 = to_str_date(ws_summary.range(check_cell2).value)

            if src_date1 == sum_date1 and src_date2 == sum_date2:
                if paste_range in used_blocks_column:
                    error_files.append((
                        Path(src).name,
                        f"trùng ngày với cột {paste_range} (đã copy từ {used_blocks_column[paste_range]})"
                    ))
                else:
                    # Copy dữ liệu từ vùng D8:D12 (5x1) thành cột dọc
                    data = ws_src.range("D8:D12").value
                    if data:
                        # ép thành cột dọc
                        data = [[v] for v in data]
                        ws_summary.range(paste_range).value = data
                        used_blocks_column[paste_range] = Path(src).name
                        print(f"✅ {Path(src).name}: Copy cột dọc vào {paste_range}")
                        matched_column = True
                    else:
                        error_files.append((Path(src).name, f"không có dữ liệu trong vùng D8:D12"))
                # KHÔNG break ở đây - tiếp tục kiểm tra các cột khác

        if not matched_large and not matched_column:
            error_files.append((Path(src).name, "không khớp ngày nào trong summary"))

        wb_src.close()

    wb_summary.save()
    print("\n🎉 Hoàn thành cập nhật Summary.xlsx")

    # ==== Báo cáo kết quả ====
    total_files = len(source_files)
    error_count = len(error_files)
    success_count = total_files - error_count

    print(f"\n📊 Thống kê:")
    print(f"   📁 Tổng số file: {total_files}")
    print(f"   ✅ Thành công: {success_count}")
    print(f"   ❌ Lỗi: {error_count}")

    if filled_blocks_large:
        print(f"\n📋 Các block lớn đã được fill:")
        for paste_range, filename in filled_blocks_large.items():
            print(f"   • {paste_range} ← {filename}")

    if used_blocks_column:
        print(f"\n📋 Các cột dọc đã được fill:")
        for paste_range, filename in used_blocks_column.items():
            print(f"   • {paste_range} ← {filename}")

    if error_files:
        print("\n❌ Các file lỗi:")
        for f, reason in error_files:
            print(f"   - {f}: {reason}")

finally:
    wb_summary.close()
    app.quit()

