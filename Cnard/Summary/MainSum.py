import pandas as pd
import openpyxl
from openpyxl import load_workbook

# ==== Đọc file nguồn ====
df = pd.read_excel(
    r"D:\1.python\1.Clonard-1\Cnard\Summary\489-539 King Street West Billing Outline 01Sep to 28Sep2025_v.xxSep2025.xlsx",
    skiprows=3
)

# Chỉ lấy cột cần thiết
df = df[["Date", "Trade", "Reg (Hrs)", "Rate ($) Trade"]]

# Gom nhóm theo Date + Trade
grouped = df.groupby(["Date", "Trade"]).agg(
    count_trade=("Trade", "count"),
    hrs=("Reg (Hrs)", "sum"),
    rate=("Rate ($) Trade", "first")
).reset_index()

# Chuẩn bị dữ liệu output
output = []
for _, row in grouped.iterrows():
    trade_label = f"{row['Trade']}: {row['count_trade']}"
    description = "General & Safety"
    ref = "Various"
    hrs = row["hrs"]
    reg = row["rate"]
    amount = hrs * reg
    output.append([
        row["Date"], trade_label, description, ref, hrs, reg, 0, 0, amount
    ])

# ==== Mở file summary có sẵn ====
summary_path = r"D:\1.python\1.Clonard-1\Cnard\Summary\CGI Summary2508xxxx_489-539 King Street West_xxAug2025.xlsx"   # 👉 đổi lại tên file summary thực tế
wb = load_workbook(summary_path)
ws = wb.active   # sheet đầu tiên (có header từ row 9)

start_row = 11  # dòng bắt đầu điền dữ liệu

# Điền dữ liệu vào file summary
for i, row in enumerate(output, start=start_row):
    date_value = row[0]
    ws.cell(row=i, column=1, value=date_value)   # cột A: Date
    ws.cell(row=i+1, column=1, value=f'=TEXT(A{i},"(ddd)")')  # dòng ngay dưới: công thức TEXT
    
    ws.cell(row=i, column=2, value=row[1])  # cột B: Trade
    ws.cell(row=i, column=3, value=row[2])  # cột C: Description
    ws.cell(row=i, column=4, value=row[3])  # cột D: Ref
    ws.cell(row=i, column=5, value=row[4])  # cột E: Hrs
    ws.cell(row=i, column=6, value=row[5])  # cột F: Reg
    ws.cell(row=i, column=7, value=row[6])  # cột G: 1.5X
    ws.cell(row=i, column=8, value=row[7])  # cột H: 2X
    ws.cell(row=i, column=9, value=row[8])  # cột I: Amount

# Lưu lại file summary
wb.save(summary_path)

print("✅ Đã điền dữ liệu vào file summary thành công!")
