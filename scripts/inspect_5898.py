import fitz
import os

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Nhân công\5898.pdf'

doc = fitz.open(pdf_path)
print(f"=== Nhân công/5898.pdf ({len(doc)} pages) ===")
for i in range(len(doc)):
    print(f"\n--- PAGE {i+1} ---")
    print(doc[i].get_text())
