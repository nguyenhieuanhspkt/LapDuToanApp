import fitz
import re

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

doc = fitz.open(pdf_path)
print(f"Total pages: {len(doc)}")

# Look for table headers and machine rows
machine_rows = []
for pno in range(9, len(doc)):
    page = doc[pno]
    text = page.get_text()
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    # check if page has machine codes like M101.0101, etc.
    codes = re.findall(r'\b(M\d{3}\.\d{4})\b', text)
    if codes:
        machine_rows.extend(codes)

print(f"Total machine codes found across pages 10-{len(doc)}: {len(machine_rows)}")
print(f"Sample codes: {machine_rows[:20]}")
