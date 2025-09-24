
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ==== BƯỚC 1: Đọc file nguồn ====
file_input = r"D:\1.python\1.Clonard-1\Cnard\Summary\489-539.xlsx"
file_summary = r"D:\1.python\1.Clonard-1\Cnard\Summary\CGI Summary.xlsx"

# Row 4 là header → bỏ qua 3 dòng đầu
df = pd.read_excel(file_input, skiprows=3)

# Chuẩn hóa tên cột
df.columns = df.columns.str.strip()
df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

print("✅ Danh sách cột:", df.columns.tolist())

# ==== BƯỚC 2: Gom nhóm ====
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
        rows.append({
            "Date": date,
            "Trade": f"{trade}: {len(g)} (OT1.5)",
            "Hrs": ot15_sum
        })

    # Overtime 2X
    ot2_sum = g["O/T 2X"].fillna(0).sum()
    if ot2_sum > 0:
        rows.append({
            "Date": date,
            "Trade": f"{trade}: {len(g)} (OT2)",
            "Hrs": ot2_sum
        })

result = pd.DataFrame(rows)
print("✅ Kết quả chuyển đổi:\n", result.head())

# ==== BƯỚC 3: Ghi vào file Summary có sẵn ====
wb = load_workbook(file_summary)
ws = wb.active  # hoặc ws = wb["Sheet1"]

start_row = 11  # bắt đầu ghi từ row 11
col_date = 1    # cột A
col_trade = 2   # cột B
col_hrs = 5     # cột E

current_row = start_row
for date, group in result.groupby("Date"):
    first_row = current_row
    for _, r in group.iterrows():
        # Ghi Date chỉ 1 lần (ở dòng đầu tiên)
        if current_row == first_row:
            ws.cell(row=current_row, column=col_date, value=date)
            # Ngay dưới Date có dòng thứ (ddd)
            ws.cell(row=current_row+1, column=col_date, value=f"({date.strftime('%a')})")
            ws.cell(row=current_row+1, column=col_date).alignment = Alignment(horizontal="center")

        ws.cell(row=current_row, column=col_trade, value=r["Trade"])
        ws.cell(row=current_row, column=col_hrs, value=r["Hrs"])
        current_row += 1

    # Nếu có nhiều dòng cho cùng 1 Date → merge Date
    if current_row - first_row > 1:
        ws.merge_cells(start_row=first_row, start_column=col_date,
                       end_row=current_row-1, end_column=col_date)

wb.save(file_summary)
print("🎉 Đã ghi dữ liệu vào Summary.xlsx thành công!")




