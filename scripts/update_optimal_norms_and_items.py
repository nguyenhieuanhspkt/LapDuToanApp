# -*- coding: utf-8 -*-
"""
Script: scripts/update_optimal_norms_and_items.py
Cập nhật cơ sở dữ liệu định mức, biến thể và 876 công tác dự án theo PHƯƠNG ÁN ĐỀ XUẤT TỐI ƯU
từ file master 'DANH MUC CONG TAC_da tra ma.xlsx' (sheet 'Doi_chieu_dinh_muc').
"""

import openpyxl
import json
import re
import copy
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

SOURCE_DIR = Path(__file__).resolve().parent.parent
MASTER_FILE = SOURCE_DIR / "templates" / "DANH_MUC_CONG_TAC_da_tra_ma.xlsx"

NORMS_DB_FILE = SOURCE_DIR / "database" / "norms_database.json"
VARIANTS_DB_FILE = SOURCE_DIR / "database" / "project_variants.json"
ITEMS_DB_FILE = SOURCE_DIR / "database" / "current_project_items.json"

def main():
    print("🚀 Bắt đầu cập nhật cơ sở dữ liệu theo Phương án Đề xuất Tối ưu...")
    
    # 1. Cập nhật norms_database.json với SB.82633 và SB.82634
    with open(NORMS_DB_FILE, "r", encoding="utf-8") as f:
        norms_db = json.load(f)
        
    sb_822 = norms_db.get("SB.822", {})
    if "SB.82633" not in norms_db and sb_822:
        norms_db["SB.82633"] = copy.deepcopy(sb_822)
        norms_db["SB.82633"]["code"] = "SB.82633"
        norms_db["SB.82633"]["name"] = "Sơn chống rỉ, sơn phủ – sơn trực tiếp lên vỏ thiết bị trong nhà (DN50 - DN100)"
        norms_db["SB.82633"]["legal_basis"] = "Thông tư 38/2026/TT-BXD (Trang 111 PL6)"
        print("✅ Đã bổ sung định mức SB.82633 vào norms_database.json")
        
    if "SB.82634" not in norms_db and sb_822:
        norms_db["SB.82634"] = copy.deepcopy(sb_822)
        norms_db["SB.82634"]["code"] = "SB.82634"
        norms_db["SB.82634"]["name"] = "Sơn chống rỉ, sơn phủ – sơn trực tiếp lên vỏ thiết bị ngoài nhà (DN125 - DN200)"
        norms_db["SB.82634"]["legal_basis"] = "Thông tư 38/2026/TT-BXD (Trang 111 PL6)"
        print("✅ Đã bổ sung định mức SB.82634 vào norms_database.json")
        
    with open(NORMS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(norms_db, f, ensure_ascii=False, indent=2)

    # 2. Cập nhật project_variants.json
    if VARIANTS_DB_FILE.exists():
        with open(VARIANTS_DB_FILE, "r", encoding="utf-8") as f:
            variants_db = json.load(f)
    else:
        variants_db = {}
        
    # Tạo biến thể cho E6C.3.01 và E6C.3.02
    e6c_301_steps = norms_db.get("E6C.3.01", {}).get("work_steps", [])
    e6c_302_steps = norms_db.get("E6C.3.02", {}).get("work_steps", [])
    
    # E6C.3.01_A1: Chỉ bước 5 (Tháo tổ hợp chuông còi đèn)
    steps_301_a1 = copy.deepcopy(e6c_301_steps)
    for s in steps_301_a1:
        s["apply_flag"] = 1 if str(s.get("step")) == "5" else 0
    variants_db["E6C.3.01_A1"] = {
        "variant_id": "E6C.3.01_A1",
        "variant_name": "Tháo và bảo quản tổ hợp chuông, còi, đèn (Bước 5)",
        "norm_code": "E6C.3.01",
        "exclude_replacement": True,
        "k_step": 0.2727,
        "work_steps": steps_301_a1
    }
    
    # E6C.3.01_A2: Bước 1-4, 6-11 (Bảo dưỡng, hiệu chỉnh chuông còi đèn)
    steps_301_a2 = copy.deepcopy(e6c_301_steps)
    for s in steps_301_a2:
        s["apply_flag"] = 0 if str(s.get("step")) == "5" else 1
    variants_db["E6C.3.01_A2"] = {
        "variant_id": "E6C.3.01_A2",
        "variant_name": "Bảo dưỡng, kiểm tra, hiệu chỉnh tổ hợp chuông còi đèn (Bước 1–4, 6–11)",
        "norm_code": "E6C.3.01",
        "exclude_replacement": True,
        "k_step": 0.7273,
        "work_steps": steps_301_a2
    }
    
    # E6C.3.02_A1: Chỉ bước 5 (Tháo đầu báo khói/nhiệt)
    steps_302_a1 = copy.deepcopy(e6c_302_steps)
    for s in steps_302_a1:
        s["apply_flag"] = 1 if str(s.get("step")) == "5" else 0
    variants_db["E6C.3.02_A1"] = {
        "variant_id": "E6C.3.02_A1",
        "variant_name": "Tháo và bảo quản thiết bị đầu báo khói/nhiệt (Bước 5)",
        "norm_code": "E6C.3.02",
        "exclude_replacement": True,
        "k_step": 0.3000,
        "work_steps": steps_302_a1
    }
    
    # E6C.3.02_A2: Bước 1-4, 6-11 (Bảo trì, hiệu chỉnh đầu báo khói/nhiệt)
    steps_302_a2 = copy.deepcopy(e6c_302_steps)
    for s in steps_302_a2:
        s["apply_flag"] = 0 if str(s.get("step")) == "5" else 1
    variants_db["E6C.3.02_A2"] = {
        "variant_id": "E6C.3.02_A2",
        "variant_name": "Bảo trì, sửa chữa, hiệu chỉnh cảm biến báo cháy (Bước 1–4, 6–11)",
        "norm_code": "E6C.3.02",
        "exclude_replacement": True,
        "k_step": 0.7000,
        "work_steps": steps_302_a2
    }
    
    with open(VARIANTS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(variants_db, f, ensure_ascii=False, indent=2)
    print("✅ Đã cập nhật các biến thể phân rã bước công việc vào project_variants.json")

    # 3. Trích xuất 876 công tác từ sheet Doi_chieu_dinh_muc theo Phương án Tối ưu
    wb = openpyxl.load_workbook(MASTER_FILE, data_only=True)
    ws = wb["Doi_chieu_dinh_muc"]
    
    items = []
    current_group = "PHẦN DỊCH VỤ"
    
    for r in range(5, ws.max_row + 1):
        val_b = ws.cell(r, 2).value
        stt = ws.cell(r, 1).value
        
        if val_b and "PHẦN DỊCH VỤ" in str(val_b):
            current_group = "PHẦN DỊCH VỤ"
            continue
        elif val_b and "PHẦN VẬT TƯ" in str(val_b):
            current_group = "PHẦN VẬT TƯ"
            continue
            
        if stt is None or not isinstance(stt, int):
            continue
            
        item_1246 = str(ws.cell(r, 2).value or "").strip()
        loc = str(ws.cell(r, 3).value or "").strip()
        desc = str(ws.cell(r, 4).value or "").strip()
        pakt_vol = float(ws.cell(r, 5).value or 0.0)
        pakt_unit = str(ws.cell(r, 6).value or "").strip()
        ref_code = str(ws.cell(r, 7).value or "").strip()
        
        opt_code_raw = str(ws.cell(r, 27).value or "").strip()
        opt_reason = str(ws.cell(r, 28).value or "").strip()
        
        khvt_code = str(ws.cell(r, 9).value or "").strip()
        if khvt_code and khvt_code in opt_code_raw:
            dm_name = ws.cell(r, 10).value
            dm_dvt = ws.cell(r, 11).value
            conv_name = ws.cell(r, 14).value
            conv_vol = ws.cell(r, 15).value
        else:
            dm_name = ws.cell(r, 19).value
            dm_dvt = ws.cell(r, 20).value
            conv_name = ws.cell(r, 23).value
            conv_vol = ws.cell(r, 24).value
            
        try:
            vol = float(conv_vol) if conv_vol is not None else pakt_vol
        except (ValueError, TypeError):
            vol = pakt_vol
            
        unit = str(dm_dvt or pakt_unit).strip()
        job_name = str(conv_name or item_1246).strip()
        
        is_unassigned = "Không áp mã" in opt_code_raw or "Không có mã" in opt_code_raw or not opt_code_raw
        
        norm_c = "KDM"
        variant_id = "A0"
        is_quote = False
        quote_p = 0.0
        
        if is_unassigned:
            is_quote = True
            quote_p = 0.0
            norm_c = "KDM"
        else:
            if "E6C.3.01" in opt_code_raw and "5" in opt_code_raw and "1–4" not in opt_code_raw:
                norm_c = "E6C.3.01"
                variant_id = "E6C.3.01_A1"
            elif "E6C.3.01" in opt_code_raw:
                norm_c = "E6C.3.01"
                variant_id = "E6C.3.01_A2"
            elif "E6C.3.02" in opt_code_raw and "5" in opt_code_raw and "1–4" not in opt_code_raw:
                norm_c = "E6C.3.02"
                variant_id = "E6C.3.02_A1"
            elif "E6C.3.02" in opt_code_raw:
                norm_c = "E6C.3.02"
                variant_id = "E6C.3.02_A2"
            else:
                m_c = re.search(r"^([A-Za-z0-9\.\_]+)", opt_code_raw)
                norm_c = m_c.group(1).strip() if m_c else opt_code_raw
                variant_id = "A0"
                
        is_bd = norm_c.startswith("BD.")
        excl_rep = True if is_bd or "đã gồm trong" in opt_code_raw else False

        items.append({
            "stt": stt,
            "group": current_group,
            "item_name_1246": item_1246,
            "location": loc,
            "description": desc,
            "pakt_volume": pakt_vol,
            "pakt_unit": pakt_unit,
            "ref_code": ref_code,
            "opt_code_raw": opt_code_raw,
            "opt_reason": opt_reason,
            "norm_code": norm_c,
            "variant_id": variant_id,
            "is_quote": is_quote,
            "unit_price_quote": quote_p,
            "job_name": job_name,
            "volume": vol,
            "unit": unit,
            "exclude_replacement": excl_rep
        })

    with open(ITEMS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
        
    print(f"🎉 Hoàn thành cập nhật {len(items)} công tác vào {ITEMS_DB_FILE}")

if __name__ == "__main__":
    main()
