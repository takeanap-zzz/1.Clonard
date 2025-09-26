import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

# ==== BƯỚC 1: Đọc file nguồn ====
file_input = r"D:\1.python\1.Clonard-1\Cnard\Summary\Mirvish Village Billing Outline 01Sep-28Sep2025_v.xxSep2025.xlsx"
file_summary = r"D:\1.python\1.Clonard-1\Cnard\Summary\CGI Summary.xlsx"

# Row 4 là header → bỏ qua 3 dòng đầu
df = pd.read_excel(file_input, skiprows=3)

# Chuẩn hóa tên cột
df.columns = df.columns.str.strip()
df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

print("✅ Danh sách cột:", df.columns.tolist())

# ==== BƯỚC 2: Gom nhóm dữ liệu ====
rows = []

# Gom theo Date + Trade
grouped = df.groupby(["Date", "Trade"])

for (date, trade), g in grouped:
    # Regular
    reg_sum = g["Reg (Hrs)"].fillna(0).sum()
    if reg_sum > 0:
        rows.append({
            "Date": date,
            "Trade": f"{trade}: {len(g)}",
            "Hrs": reg_sum
        })

    # Overtime 1.5X
    ot15_sum = g["O / T 1.5X"].fillna(0).sum()
    if ot15_sum > 0:
        num_workers_ot = (g["O / T 1.5X"].fillna(0) > 0).sum()
        rows.append({
            "Date": date,
            "Trade": f"{trade}: {num_workers_ot}",
            "Hrs": ot15_sum
        })

    # Overtime 2X
    ot2_sum = g["O/T 2X"].fillna(0).sum()
    if ot2_sum > 0:
        num_workers_ot2 = (g["O/T 2X"].fillna(0) > 0).sum()
        rows.append({
            "Date": date,
            "Trade": f"{trade}: {num_workers_ot2} (OT2)",
            "Hrs": ot2_sum
        })

result = pd.DataFrame(rows)
print("✅ Kết quả chuyển đổi:\n", result.head())

# ==== BƯỚC 3: Ghi vào file Summary có sẵn ====
wb = load_workbook(file_summary)
ws = wb.active  # hoặc ws = wb["Sheet1"]

# Thiết lập vị trí ghi
start_row = 11  # bắt đầu ghi từ row 11
col_date = 1    # cột A
col_trade = 2   # cột B
col_category = 3  # cột C (mới thêm)
col_description = 4  # cột D (mới thêm)
col_hrs = 5     # cột E

# Tô đỏ các ô dữ liệu mới
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

current_row = start_row
for date, group in result.groupby("Date"):
    first_row = current_row
    is_first_row_of_date = True
    
    for _, r in group.iterrows():
        # Ghi Date chỉ 1 lần (ở dòng đầu tiên của mỗi ngày)
        if is_first_row_of_date:
            cell_date = ws.cell(row=current_row, column=col_date, value=date)
            cell_date.fill = red_fill  # Tô đỏ ô ngày

        # Ghi Trade + tô đỏ
        cell_trade = ws.cell(row=current_row, column=col_trade, value=r["Trade"])
        cell_trade.fill = red_fill

        # Ghi cột C (Category)
        if is_first_row_of_date:
            cell_category = ws.cell(row=current_row, column=col_category, value="General & Safety")
        else:
            cell_category = ws.cell(row=current_row, column=col_category, value="- ditto -")
        cell_category.fill = red_fill

        # Ghi cột D (Description)
        if is_first_row_of_date:
            cell_description = ws.cell(row=current_row, column=col_description, value="Various")
        else:
            cell_description = ws.cell(row=current_row, column=col_description, value=' " ')
        cell_description.fill = red_fill

        # Ghi Hrs + tô đỏ
        cell_hrs = ws.cell(row=current_row, column=col_hrs, value=r["Hrs"])
        cell_hrs.fill = red_fill

        current_row += 1
        is_first_row_of_date = False  # Sau dòng đầu tiên, các dòng sau sẽ dùng "ditto" và " "

    # Nếu có nhiều dòng cho cùng 1 Date → merge Date (tuỳ chọn, bạn có thể bật lại nếu cần)
    # if current_row - first_row > 1:
    #     ws.merge_cells(start_row=first_row, start_column=col_date,
    #                    end_row=current_row-1, end_column=col_date)

# Lưu lại file
wb.save(file_summary)
print("🎉 Đã ghi dữ liệu vào Summary.xlsx thành công!")


