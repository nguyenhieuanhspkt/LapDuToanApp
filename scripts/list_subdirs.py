import os

folder = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng'

print("=== RECURSIVE LISTING OF Đơn giá lâm đồng ===")
for root, dirs, files in os.walk(folder):
    rel = os.path.relpath(root, folder)
    print(f"\nDirectory: [{rel}]")
    for d in dirs:
        print(f"  [DIR]  {d}")
    for f in files:
        full = os.path.join(root, f)
        print(f"  [FILE] {f} ({os.path.getsize(full)} bytes)")
