"""
Script to extract and digitize Machine Prices from:
'D:\\OneDrive - EVN\\Desktop\\phu luc kem theo qd 21.pdf'
Published under: QĐ 21/QĐ-SXD ngày 14/01/2026 Sở Xây dựng Lâm Đồng.
"""

import fitz
import re
import json
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

PDF_PATH = r"D:\OneDrive - EVN\Desktop\phu luc kem theo qd 21.pdf"
OUTPUT_DIR = Path(r"D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Hiếu\2026_LAPDUTOAN_APP\database\provinces\lam_dong")
OUTPUT_JSON = OUTPUT_DIR / "machines_2026.json"
OUTPUT_EXCEL = Path(r"D:\TaskApp_kiet\LapDuToanApp\gia_ca_may_lam_dong_2026.xlsx")

def parse_num(val):
    if not val:
        return 0
    if isinstance(val, (list, tuple)):
        val = ' '.join(str(x) for x in val)
    tokens = str(val).strip().split()
    for tok in tokens:
        clean = re.sub(r'[^\d]', '', tok)
        if clean:
            n = int(clean)
            if n < 10**11:
                return n
            return int(clean[:9])
    return 0

def normalize_code(raw):
    c = raw.upper().strip(' :.-_()[],')
    if not (c.startswith('M') or c.startswith('1') or c.startswith('I')):
        return None
    if not c.startswith('M'):
        c = 'M' + c[1:]
    body = c[1:]
    repl = {'O':'0','Q':'0','I':'1','L':'1','B':'8','S':'5','Z':'2','D':'0','G':'6'}
    clean_body = ''.join(repl.get(ch, ch) for ch in body)
    clean = 'M' + clean_body
    
    # Common OCR patterns: e.g. M101.05134 -> M101.0504, M101.Q1Q1 -> M101.0101
    m = re.match(r'^(M\d{3}\.\d{4})', clean)
    if m:
        return m.group(1)
    m2 = re.match(r'^(M\d{3}\.\d{2})', clean)
    if m2:
        return m2.group(1) + "00"
    return None

STANDARD_CATEGORIES = {
    "M101.01": "Máy đào 1 gàu bánh xích",
    "M101.02": "Máy đào 1 gàu bánh lốp",
    "M101.03": "Máy đào gàu dây",
    "M101.04": "Máy xúc lật",
    "M101.05": "Máy ủi",
    "M101.06": "Máy cạp",
    "M101.07": "Máy san",
    "M101.08": "Máy lu",
    "M101.09": "Máy đầm",
    "M102.01": "Máy búa đóng cọc",
    "M102.02": "Máy ép cọc thuỷ lực",
    "M102.03": "Máy khoan cọc",
    "M102.04": "Máy khoan cọc nhồi",
    "M103.01": "Máy trộn bê tông",
    "M103.02": "Xe vận chuyển bê tông (xe bồn)",
    "M103.03": "Máy bơm bê tông",
    "M103.04": "Máy đầm bê tông",
    "M103.05": "Máy xoa nền bê tông",
    "M104.01": "Cần trục bánh xích",
    "M104.02": "Cần trục bánh lốp / ô tô",
    "M104.03": "Cần trục tháp",
    "M104.04": "Vận thăng nâng",
    "M104.05": "Tời nâng",
    "M104.06": "Xe nâng hàng",
    "M105.01": "Máy cắt, uốn cốt thép",
    "M105.02": "Máy hàn điện, máy phát hàn",
    "M105.03": "Máy uốn ống, tiện ren",
    "M106.01": "Máy đào hầm, khoan hầm",
    "M107.01": "Ô tô tự đổ",
    "M107.02": "Ô tô tải",
    "M107.03": "Xe đầu kéo, rơ moóc",
    "M107.04": "Ô tô xitéc",
    "M108.01": "Đồng hồ vạn năng (VOM) và thiết bị đo điện",
    "M108.02": "Vôn mét điện tử",
    "M108.03": "Mê gôm mét, máy đo điện trở cách điện",
    "M108.04": "Thiết bị thử đầu báo cháy khói, nhiệt",
    "M108.05": "Thiết bị đo nội trở và dung lượng accu",
    "M108.06": "Bình khí mẫu hiệu chuẩn",
    "M109.01": "Máy thổi bụi",
    "M109.02": "Bàn cân áp suất",
    "M109.03": "Máy bơm nước",
    "M109.04": "Máy nén khí",
    "M109.05": "Máy phát điện",
    "M112.01": "Máy khoan cầm tay",
    "M112.02": "Máy mài cầm tay",
    "M112.06": "Máy đầm dùi cầm tay",
    "M112.17": "Máy khoan bê tông cầm tay",
    "M201.00": "Thiết bị thi công xây lắp chuyên ngành"
}

def extract_all():
    doc = fitz.open(PDF_PATH)
    all_machines = []
    current_chapter = ""
    current_group = ""
    current_category = ""
    
    print(f"Opened PDF with {len(doc)} pages.")
    
    for p_idx in range(4, len(doc)):
        page = doc[p_idx]
        words = page.get_text('words')
        
        # Identify row anchors: words in column range x in [105, 140]
        raw_rows = []
        for w in words:
            # Chapter header or category detection
            if w[0] < 350 and ("CHƯƠNG" in w[4].upper() or "CHU)" in w[4].upper() or "CHUONG" in w[4].upper()):
                current_chapter = w[4]
                
            if 105 <= w[0] <= 140:
                raw = w[4].upper().strip(' :.-_()[],')
                if raw.startswith('M') or raw.startswith('1') or raw.startswith('I'):
                    code = normalize_code(raw)
                    if code:
                        raw_rows.append({'code': code, 'raw_code': w[4], 'y': w[1], 'page': p_idx + 1})
                        
        # Sort rows vertically
        raw_rows.sort(key=lambda r: r['y'])
        
        # Deduplicate rows with very close y (within 4pt)
        rows = []
        for r in raw_rows:
            if not rows or abs(r['y'] - rows[-1]['y']) > 4:
                rows.append(r)
                
        for i, r in enumerate(rows):
            code = r['code']
            r_top = r['y'] - 3
            r_bottom = rows[i+1]['y'] - 2 if i + 1 < len(rows) else r['y'] + 15
            
            row_words = [w for w in words if r_top <= w[1] <= r_bottom]
            row_words.sort(key=lambda w: w[0])
            
            # Words by column zone
            name_words = [w[4] for w in row_words if 145 <= w[0] < 230]
            fuel_norm_words = [w[4] for w in row_words if 290 <= w[0] < 340]
            driver_words = [w[4] for w in row_words if 340 <= w[0] < 380]
            orig_price_words = [w[4] for w in row_words if 380 <= w[0] < 425]
            fuel_cost_words = [w[4] for w in row_words if 425 <= w[0] < 465]
            labor_v2_words = [w[4] for w in row_words if 465 <= w[0] < 492]
            labor_v3_words = [w[4] for w in row_words if 492 <= w[0] < 522]
            labor_v4_words = [w[4] for w in row_words if 522 <= w[0] < 552]
            price_v2_words = [w[4] for w in row_words if 580 <= w[0] < 630]
            price_v3_words = [w[4] for w in row_words if 630 <= w[0] < 675]
            price_v4_words = [w[4] for w in row_words if 675 <= w[0] < 720]
            price_pq_words = [w[4] for w in row_words if 720 <= w[0] < 765]
            
            name_text = ' '.join(name_words).strip()
            
            # Check if this is a group header (e.g., ending in .0000 or .0x00 with no price)
            p_v4 = parse_num(' '.join(price_v4_words))
            p_v2 = parse_num(' '.join(price_v2_words))
            p_v3 = parse_num(' '.join(price_v3_words))
            
            if code.endswith('.0000') or (p_v4 == 0 and p_v2 == 0 and name_text and not name_text[0].isdigit()):
                current_group = name_text
                continue
            elif code.endswith('00') and p_v4 == 0 and p_v2 == 0:
                current_category = name_text
                continue
                
            spec = name_text
            clean_cat = STANDARD_CATEGORIES.get(code[:7], current_category)
            if clean_cat and spec:
                full_name = f"{clean_cat} ({spec})"
            elif clean_cat:
                full_name = clean_cat
            elif spec:
                full_name = spec
            else:
                full_name = f"Máy thi công {code}"
                
            machine_entry = {
                "code": code,
                "name": full_name,
                "spec": spec,
                "category": current_category or current_group,
                "unit": "ca",
                "driver": ' '.join(driver_words).strip(),
                "fuel_norm": ' '.join(fuel_norm_words).strip(),
                "original_price_1k": parse_num(' '.join(orig_price_words)),
                "fuel_cost": parse_num(' '.join(fuel_cost_words)),
                "labor_cost": {
                    "vung_2": parse_num(' '.join(labor_v2_words)),
                    "vung_3": parse_num(' '.join(labor_v3_words)),
                    "vung_4": parse_num(' '.join(labor_v4_words))
                },
                "price": {
                    "vung_2": p_v2,
                    "vung_3": p_v3,
                    "vung_4": p_v4,
                    "phu_quy": parse_num(' '.join(price_pq_words))
                },
                "legal_basis": "QĐ 21/QĐ-SXD ngày 14/01/2026 Sở Xây dựng Lâm Đồng",
                "page": r['page']
            }
            all_machines.append(machine_entry)
            
    print(f"Extracted {len(all_machines)} machines with valid prices.")
    
    # Save to JSON
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(all_machines, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON database to: {OUTPUT_JSON}")
    
    # Save to Excel
    flat_rows = []
    for m in all_machines:
        flat_rows.append({
            "Mã hiệu": m["code"],
            "Tên loại máy / Quy cách": m["name"],
            "ĐVT": m["unit"],
            "Thợ ĐK": m["driver"],
            "ĐM Nhiên liệu": m["fuel_norm"],
            "Nguyên giá (1000đ)": m["original_price_1k"],
            "CP Nhiên liệu": m["fuel_cost"],
            "Lương thợ Vùng II": m["labor_cost"]["vung_2"],
            "Lương thợ Vùng III": m["labor_cost"]["vung_3"],
            "Lương thợ Vùng IV (Mặc định)": m["labor_cost"]["vung_4"],
            "Giá ca máy Vùng II": m["price"]["vung_2"],
            "Giá ca máy Vùng III": m["price"]["vung_3"],
            "Giá ca máy Vùng IV (Mặc định)": m["price"]["vung_4"],
            "Trang PDF": m["page"]
        })
    df = pd.DataFrame(flat_rows)
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"Saved Excel file to: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    extract_all()
