import pdfplumber
import re

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

all_machines = []
current_chapter = ""
current_subgroup = ""

with pdfplumber.open(pdf_path) as pdf:
    # Pages 10 to 49 (indices 9 to 48)
    for pno in range(9, len(pdf.pages)):
        page = pdf.pages[pno]
        tables = page.extract_tables()
        if not tables:
            continue
        
        for table in tables:
            for row in table:
                if not row or len(row) < 3:
                    continue
                
                c0 = str(row[0] or '').strip().replace('\n', ' ')
                c1 = str(row[1] or '').strip().replace('\n', ' ')
                c2 = str(row[2] or '').strip().replace('\n', ' ')
                
                # Check for chapter / group header
                if 'CHƢƠNG' in c0 or 'CHƯƠNG' in c0 or 'CHƯƠNG' in c1 or 'CHƢƠNG' in c1:
                    current_chapter = f"{c0} - {c1} - {c2}".strip(' -')
                    continue
                if c1.startswith('M') and c1.endswith('.0000'):
                    current_chapter = f"{c1}: {c2}"
                    continue
                if c1.startswith('M') and c1.endswith('.0100') or (c1.startswith('M') and c1.endswith('00') and not c0.isdigit()):
                    current_subgroup = f"{c1}: {c2}"
                    continue
                
                # Check if it's a machine item row: c0 is digit or c1 is like Mxxx.xxxx
                if re.match(r'^M\d{3}\.\d{4}$', c1):
                    # It is a machine!
                    # Normalize columns
                    # Expected row len usually 11, 12, 13
                    # Row cols: STT | Mã hiệu | Loại máy | Định mức NL | Loại NL | Chi phí NL | Thợ điều khiển | Lương Vùng II | Lương Vùng III | Lương Vùng IV | Giá ca máy Vùng II | Giá ca máy Vùng III | Giá ca máy Vùng IV
                    all_machines.append({
                        'page': pno + 1,
                        'chapter': current_chapter,
                        'subgroup': current_subgroup,
                        'stt': c0,
                        'code': c1,
                        'name': c2,
                        'raw_row': [str(x or '').strip().replace('\n', ' ') for x in row]
                    })

print(f"Total extracted machines: {len(all_machines)}")
if all_machines:
    print("First 3 machines:")
    for m in all_machines[:3]:
        print(f"  [{m['code']}] {m['name']} | Row: {m['raw_row']}")
    print("Last 3 machines:")
    for m in all_machines[-3:]:
        print(f"  [{m['code']}] {m['name']} | Row: {m['raw_row']}")
