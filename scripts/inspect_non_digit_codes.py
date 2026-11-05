import pdfplumber
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

non_digit = []
with pdfplumber.open(pdf_path) as pdf:
    for pno in range(9, len(pdf.pages)):
        tables = pdf.pages[pno].extract_tables()
        for t in tables:
            for r in t:
                if not r or len(r) < 3: continue
                c0 = str(r[0] or '').strip().replace('\n', ' ')
                c1 = str(r[1] or '').strip().replace('\n', ' ')
                if re.match(r'^M\d{3}\.\d{4}[a-zA-Z]?$', c1) and not c0.isdigit():
                    non_digit.append((pno+1, c0, c1, r[2]))

print(f"Total non-digit code rows: {len(non_digit)}")
for item in non_digit[:15]:
    print(item)
