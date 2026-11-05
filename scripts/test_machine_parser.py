import pdfplumber
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

def clean_num(s):
    if not s or str(s).strip() in ['-', '', 'None']:
        return None
    s = re.sub(r'\s+', '', str(s)).replace('.', '').replace(',', '.')
    try:
        return float(s) if '.' in s else int(s)
    except:
        return s

records = []
current_chapter = ''
current_group = ''

with pdfplumber.open(pdf_path) as pdf:
    for pno in range(9, len(pdf.pages)):
        tables = pdf.pages[pno].extract_tables()
        for t in tables:
            for r in t:
                if not r or len(r) < 3:
                    continue
                c0 = str(r[0] or '').strip().replace('\n', ' ')
                c1 = str(r[1] or '').strip().replace('\n', ' ')
                c2 = str(r[2] or '').strip().replace('\n', ' ')
                
                # Header check
                if 'CHƯƠNG' in c1 or 'CHƢƠNG' in c1 or 'CHƯƠNG' in c0 or 'CHƢƠNG' in c0:
                    current_chapter = f"{c1}: {c2}".strip(': ')
                    continue
                if re.match(r'^M\d{3}\.0000$', c1):
                    current_group = f"{c1}: {c2}".strip(': ')
                    continue
                if re.match(r'^M\d{3}\.\d{2}00$', c1) or (not c0.isdigit() and c1.startswith('M')):
                    current_group = f"{c1}: {c2}".strip(': ')
                    continue
                
                if re.match(r'^M\d{3}\.\d{4}$', c1) and c0.isdigit():
                    row_data = {
                        'stt': int(c0),
                        'code': c1,
                        'spec': c2,
                        'chapter': current_chapter,
                        'group': current_group,
                        'fuel_norm': clean_num(r[3]) if len(r) > 3 else None,
                        'fuel_type': str(r[4] or '').strip().replace('\n', ' ') if len(r) > 4 else '',
                        'fuel_cost': clean_num(r[5]) if len(r) > 5 else None,
                        'driver_rank': str(r[6] or '').strip().replace('\n', ' ') if len(r) > 6 else '',
                        'driver_wage_v2': clean_num(r[7]) if len(r) > 7 else None,
                        'driver_wage_v3': clean_num(r[8]) if len(r) > 8 else None,
                        'driver_wage_v4': clean_num(r[9]) if len(r) > 9 else None,
                        'price_v2': clean_num(r[10]) if len(r) > 10 else None,
                        'price_v3': clean_num(r[11]) if len(r) > 11 else None,
                        'price_v4': clean_num(r[12]) if len(r) > 12 else None,
                    }
                    records.append(row_data)

print(f"Total machine records parsed: {len(records)}")
stt_list = [r['stt'] for r in records]
print(f"Max STT: {max(stt_list)}")
print(f"Min STT: {min(stt_list)}")
print(f"First record: {records[0]}")
print(f"Last record: {records[-1]}")
