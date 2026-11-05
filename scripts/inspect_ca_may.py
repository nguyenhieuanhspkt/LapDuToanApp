import fitz

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

doc = fitz.open(pdf_path)
print(f"=== ca máy 21 LĐ.pdf ({len(doc)} pages) ===")
for i in range(min(5, len(doc))):
    print(f"\n--- PAGE {i+1} ---")
    print(doc[i].get_text()[:1200])
