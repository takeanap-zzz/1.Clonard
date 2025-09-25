import os
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

DRY_RUN = False   # Đặt False để thực hiện thật
PATTERN = re.compile(r'^CGI.*\.pdf$', re.IGNORECASE)
INVALID_CHARS = r'<>:"/\\|?*'

def sanitize_name(name: str) -> str:
    s = re.sub(r'[{}]'.format(re.escape(INVALID_CHARS)), '_', name)
    return s.strip() or "CGI_pdf_name"

def choose_root_dir():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Chọn thư mục gốc")
    root.destroy()
    return folder

def find_cgi_pdf_in_dir(dirpath: Path):
    try:
        for p in sorted(dirpath.iterdir()):
            if p.is_file() and PATTERN.match(p.name):
                return p
    except PermissionError:
        pass
    return None

def make_unique_target(path: Path) -> Path:
    if not path.exists():
        return path
    base = path.name
    parent = path.parent
    i = 1
    while True:
        candidate = parent / f"{base}_{i}"
        if not candidate.exists():
            return candidate
        i += 1

def main():
    folder = choose_root_dir()
    if not folder:
        print("❌ Không chọn thư mục.")
        return
    root = Path(folder)
    print(f"Quét trong: {root}\nDRY_RUN = {DRY_RUN}\n")

    all_dirs = [Path(dp) for dp, _, _ in os.walk(root)]
    all_dirs.sort(key=lambda p: len(p.parts), reverse=True)

    actions = []
    for d in all_dirs:
        if not d.exists():
            continue
        cgi_pdf = find_cgi_pdf_in_dir(d)
        if cgi_pdf:
            new_name = sanitize_name(cgi_pdf.stem)
            target = d.parent / new_name
            if d.resolve() == target.resolve():
                continue
            target = make_unique_target(target)
            actions.append((d, target))

    if not actions:
        print("⚠️ Không tìm thấy file PDF 'CGI...' nào.")
        return

    print("👉 Các rename sẽ thực hiện:")
    for src, dst in actions:
        print(f"  {src}  -->  {dst}")

    if DRY_RUN:
        print("\n(DRY_RUN bật, chưa đổi tên. Đặt DRY_RUN=False để chạy thật.)")
        print(f"Tổng cộng {len(actions)} thư mục cần đổi tên.")
        return
    success = 0
    for src, dst in actions:
        if src.exists():
            try:
                src.rename(dst)
                success += 1
                #print(f"✅ Đã đổi: {src} -> {dst}")
            except Exception as e:
                print(f"❌ Lỗi đổi {src}: {e}")

    print(f"\n 🎉🎉🎉 Đã đổi tên thành công {success}/{len(actions)} thư mục.")

if __name__ == "__main__":
    main()
