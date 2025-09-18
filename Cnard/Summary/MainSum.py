import pandas as pd
from openpyxl import Workbook

# Đọc file 1 (giả sử tên file: input.xlsx, sheet1)
df = pd.read_excel("489-539 King Street West Billing Outline 01Sep to 28Sep2025_v.xxSep2025.xlsx", skiprows=3)  # bỏ qua 3 dòng đầu => bắt đầu row 4

# Chỉ lấy các cột cần thiết
df = df[["Date", "Trade", "Reg (Hrs)", "Rate ($) Trade"]]

# Gom nhóm theo ngày + trade
grouped = df.groupby(["Date", "Trade"]).agg(
    count_trade=("Trade", "count"),
    hrs=("Reg (Hrs)", "sum"),
    rate=("Rate ($) Trade", "first")  # giả sử cùng 1 rate cho trade
).reset_index()

# Tạo bảng output như file 2
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

out_df = pd.DataFrame(output, columns=["DATE", "TRADE", "DESCRIPTION", "REF", 
                                       "HRS", "REG", "1.5X", "2X", "AMOUNT"])

# Xuất ra file Excel (bắt đầu từ row 9)
wb = Workbook()
ws = wb.active

# Chèn 8 dòng trống trước
for _ in range(8):
    ws.append([])

# Thêm header
ws.append(list(out_df.columns))

# Thêm dữ liệu
for row in out_df.itertuples(index=False):
    ws.append(row)

wb.save("file2.xlsx")

print("✅ Đã tạo file2.xlsx theo format yêu cầu")
