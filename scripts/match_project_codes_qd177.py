import fitz
import re
import openpyxl

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\PL QĐ Chuyển đổi NC SCL thiết bị NMNĐ đốt than công suất tổ máy 600MW.pdf'

doc = fitz.open(pdf_path)
print(f"Reading {len(doc)} pages...")

full_text = ""
for page in doc:
    full_text += page.get_text() + "\n"

# Load our project codes from 02_Chi_Tiet_856_Dich_Vu
wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=True)
ws_02 = wb['02_Chi_Tiet_856_Dich_Vu']

our_codes = set()
for r in range(5, ws_02.max_row + 1):
    c = ws_02.cell(r, 9).value
    if c:
        c_clean = str(c).strip()
        # extract base code like E6C.3.05, E6C.3.01, E6BOP.1.03...
        base_match = re.search(r'([A-Z0-9]+(?:\.[A-Z0-9]+)+)', c_clean)
        if base_match:
            our_codes.add(base_match.group(1))

print(f"Codes to search: {sorted(our_codes)}")

# Let's search each line or table row
# The PDF structure has rows:
# STT \n Mã hiệu \n Nội dung \n Bậc thợ \n Nhóm nhân công
# Or text flow

print("\n=== MATCHING CODES IN PDF ===")
for c in sorted(our_codes):
    if not (c.startswith('E6') or c.startswith('BB') or c.startswith('BD') or c.startswith('SB') or c.startswith('DB') or c.startswith('BA')):
        continue
    # search pattern
    matches = re.findall(rf'(\b{re.escape(c)}(?:\.\d+)?\b[^\n]*\n(?:[^\n]*\n){{1,5}}?(?:Nhân công\s+Nhóm\s+\d+|Nhóm\s+\d+))', full_text, re.IGNORECASE)
    if matches:
        print(f"\nCode {c}:")
        for m in matches[:2]:
            clean_m = ' '.join(m.split())
            print(f"   -> {clean_m}")
    else:
        # try simple search
        idx = full_text.find(c)
        if idx != -1:
            snippet = ' '.join(full_text[idx:idx+200].split())
            print(f"\nCode {c} (simple find): {snippet}")
        else:
            print(f"\nCode {c}: Not found in 600MW SCL PDF (Likely TT38 or not SCL)")
