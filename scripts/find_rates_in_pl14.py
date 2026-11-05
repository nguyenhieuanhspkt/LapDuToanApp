import fitz
import re

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Cung Ứng\Trung tu S1 2025\01_QĐ-NĐVT4\Phụ luc A_Khối lượng, dự toán CSL năm 2025\Phụ lục 1.4_Khối lượng nhân công_Phần tổ máy.pdf'
doc = fitz.open(pdf_path)

unit_prices = set()
for page in doc:
    text = page.get_text()
    # find lines with numbers like 334.778 or similar 6-digit numbers
    matches = re.findall(r'\b(\d{3}\.\d{3})\b', text)
    for m in matches:
        unit_prices.add(m)

print("Unique rates found in PL 1.4:")
for p in sorted(unit_prices):
    print("  ", p)
