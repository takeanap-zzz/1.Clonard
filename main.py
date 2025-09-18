import shutil
from pathlib import Path
from datetime import datetime, timedelta
import re

# File Excel gốc
src = Path(r"D:\py\252Church\Aug 04- Aug 10 2025 Weekly Timesheet Input v6.xlsx")

# Thư mục chứa bản sao
dst_folder = src.parent
dst_folder.mkdir(parents=True, exist_ok=True)

# ------------------------
# Regex cho phép khoảng trắng tùy ý quanh dấu "-"
# ------------------------
match = re.search(r"([A-Za-z]+ \d{1,2})\s*-\s*([A-Za-z]+ \d{1,2}) (\d{4})", src.stem)
if not match:
    raise ValueError("❌ Không tìm thấy ngày tháng trong tên file!")

start_str, end_str, year = match.groups()

# Parse datetime
start_date = datetime.strptime(f"{start_str} {year}", "%b %d %Y")
end_date = datetime.strptime(f"{end_str} {year}", "%b %d %Y")

# ------------------------
# Tạo 6 bản copy (6 tuần tiếp theo)
# ------------------------
for i in range(1, 6):
    new_start = start_date + timedelta(days=7 * i)
    new_end = end_date + timedelta(days=7 * i)
    new_filename = f"{new_start.strftime('%b %d')} - {new_end.strftime('%b %d')} {year} Weekly Timesheet Input v6.xlsx"
    shutil.copy2(src, dst_folder / new_filename)
    
