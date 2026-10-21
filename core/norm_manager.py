"""
Norm Manager Module:
Loads, queries, updates, and aggregates norms and resource prices.
"""

import json
from pathlib import Path
from typing import Dict, List, Optional, Any
import sys

# Ensure parent directory is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from config import NORMS_DB_FILE, PRICES_DB_FILE, IMAGES_DIR

class NormManager:
    def __init__(self, norms_file: Path = NORMS_DB_FILE, prices_file: Path = PRICES_DB_FILE):
        self.norms_file = norms_file
        self.prices_file = prices_file
        self.norms: Dict[str, Any] = {}
        self.prices: Dict[str, Any] = {}
        self.load_all()

    def load_all(self):
        """Load norms and prices from JSON database with automatic header-row sanitization."""
        if self.norms_file.exists():
            with open(self.norms_file, "r", encoding="utf-8") as f:
                raw_norms = json.load(f)
                # Ensure no phantom header rows in resources
                for c, nm in raw_norms.items():
                    res = nm.get("resources", {})
                    for cat in ["materials", "labor", "machines"]:
                        if cat in res:
                            res[cat] = [
                                it for it in res[cat] 
                                if it.get("code") and it.get("name", "").strip() not in ["1. Vật liệu", "2. Nhân công", "3. Máy thi công"]
                            ]
                self.norms = raw_norms
        else:
            self.norms = {}

        if self.prices_file.exists():
            with open(self.prices_file, "r", encoding="utf-8") as f:
                self.prices = json.load(f)
        else:
            self.prices = {"materials": [], "labor": [], "machines": []}

    def get_all_norm_codes(self) -> List[str]:
        """Return list of all norm codes available in the database."""
        return list(self.norms.keys())

    def get_norm(self, code: str) -> Optional[Dict[str, Any]]:
        """Get full details of a specific norm."""
        return self.norms.get(code)

    def get_evidence_image_path(self, code: str) -> Optional[Path]:
        """Get absolute path to evidence image for a norm."""
        norm = self.get_norm(code)
        if norm and "evidence_image" in norm:
            img_rel = norm["evidence_image"]
            # Look in IMAGES_DIR
            filename = Path(img_rel).name
            img_path = IMAGES_DIR / filename
            if img_path.exists():
                return img_path
        return None

    def get_resource_prices(self) -> Dict[str, Any]:
        """Get resource prices dictionary."""
        return self.prices

    def update_resource_prices(self, updated_prices: Dict[str, Any]) -> bool:
        """Save updated resource prices back to disk."""
        self.prices = updated_prices
        with open(self.prices_file, "w", encoding="utf-8") as f:
            json.dump(self.prices, f, ensure_ascii=False, indent=2)
        return True

    def get_default_tech_profile(self, code: str) -> str:
        """Map standard norm code to technical nature profile."""
        tech_map = {
            "BD.41111": "Đầu báo khói, báo nhiệt địa chỉ (Lắp đặt mới/thay thế)",
            "BD.41121": "Đèn báo cháy chỉ thị vị trí (Lắp đặt mới/thay thế)",
            "BD.41131": "Nút nhấn khẩn báo cháy gắn tường (Lắp đặt lại)",
            "BD.41141": "Chuông báo cháy tự động âm lượng cao (Lắp đặt mới/thay thế)",
            "BD.42201": "Bộ điều khiển, bộ chuyển mạch tín hiệu PCCC (Lắp đặt)",
            "E6C.3.01": "Nút nhấn khẩn báo cháy (Bảo trì, thí nghiệm giả lập)",
            "E6C.3.02": "Đầu báo khói/nhiệt chuyên dụng (Bảo dưỡng, hiệu chuẩn độ nhạy)",
            "E6C.3.03": "Module giám sát dây báo cháy nhiệt độ cao (Bảo dưỡng)",
            "E6C.3.04": "Module cảm biến beam quang cự ly xa (Bảo dưỡng, căn chỉnh)",
            "E6C.3.05": "Tủ trung tâm báo cháy FMZ 5000 (Sửa chữa, bảo dưỡng toàn diện)",
            "E6C.3.06": "Phần điện van Deluge, Wet xả tràn (Kiểm tra, bảo dưỡng)",
            "E6B.1.13.01": "Bình ắc quy nguồn dự phòng 24VDC PCCC (Đo nội trở, nạp xả)",
            "E6C.1.31": "Đầu đo khí độc hại NH3, Clo, H2 (Bơm khí chuẩn, hiệu chuẩn Zero & Span)"
        }
        return tech_map.get(code, "Công tác bảo dưỡng kỹ thuật chung")

    def calculate_item_unit_rate(
        self,
        item: Dict[str, Any],
        coefficients: Optional[Dict[str, Any]] = None,
        custom_prices: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Calculate horizontal unit rates for a single work item:
        - dg_vl: Material unit rate (đ/unit)
        - dg_nc: Labor unit rate (đ/unit)
        - dg_m: Machine unit rate (đ/unit)
        - dg_tt: Direct unit rate T = VL + NC + M
        - dg_truoc_thue: Full pre-tax unit rate G = T * (1 + C + TT) * (1 + TL)
        - thanh_tien_truoc_thue: Volume * dg_truoc_thue
        - k_full: Full overhead and profit factor
        """
        coeffs = coefficients or {
            "c": {"rate": 7.3},
            "tt": {"rate": 2.5},
            "tl": {"rate": 5.5}
        }
        c_rate = float(coeffs.get("c", {}).get("rate", 7.3)) / 100.0
        tt_rate = float(coeffs.get("tt", {}).get("rate", 2.5)) / 100.0
        tl_rate = float(coeffs.get("tl", {}).get("rate", 5.5)) / 100.0
        k_full = (1.0 + c_rate + tt_rate) * (1.0 + tl_rate)

        vol = float(item.get("volume", 0.0))
        code = (item.get("norm_code") or "").strip()
        is_quote = item.get("is_quote", False) or code.startswith("BG") or code.startswith("TT") or not code

        if is_quote:
            quote_price = float(item.get("unit_price_quote", 0.0))
            apply_factor = item.get("apply_overhead", True)
            dg_truoc_thue = quote_price * k_full if apply_factor else quote_price
            return {
                "is_quote": True,
                "tech_profile": item.get("tech_profile", "Công tác báo giá trọn gói"),
                "dg_vl": 0.0,
                "dg_nc": quote_price,
                "dg_m": 0.0,
                "dg_tt": quote_price,
                "k_full": k_full if apply_factor else 1.0,
                "dg_truoc_thue": dg_truoc_thue,
                "thanh_tien_truoc_thue": vol * dg_truoc_thue
            }

        norm = self.get_norm(code)
        if not norm:
            return {
                "is_quote": False,
                "tech_profile": "Chưa xác định",
                "dg_vl": 0.0, "dg_nc": 0.0, "dg_m": 0.0, "dg_tt": 0.0,
                "k_full": k_full, "dg_truoc_thue": 0.0, "thanh_tien_truoc_thue": 0.0
            }

        unit_str = norm.get("unit", "")
        divisor = 10.0 if "10" in unit_str else (5.0 if "5" in unit_str else 1.0)

        # Resource price lookup maps
        prices_db = custom_prices or self.get_resource_prices()
        vl_map = {m["code"]: m.get("price_notice", 0) for m in prices_db.get("materials", [])}
        nc_map = {l["code"]: l.get("price_current", 0) for l in prices_db.get("labor", [])}
        mac_map = {mac["code"]: mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0)) for mac in prices_db.get("machines", [])}

        res = norm.get("resources", {})
        
        # 1. Material unit rate
        dg_vl = 0.0
        for m in res.get("materials", []):
            m_code = m.get("code")
            if m_code and m_code != "V00750" and m.get("unit") != "%":
                dg_vl += (float(m.get("rate", 0.0)) * float(vl_map.get(m_code, 0))) / divisor
        for m in res.get("materials", []):
            if m.get("code") == "V00750" or m.get("unit") == "%":
                pct = float(m.get("rate", 0.0)) / 100.0
                dg_vl += dg_vl * pct

        # 2. Labor unit rate
        dg_nc = 0.0
        for l in res.get("labor", []):
            l_code = l.get("code")
            if l_code:
                dg_nc += (float(l.get("rate", 0.0)) * float(nc_map.get(l_code, 0))) / divisor

        # 3. Machine unit rate
        dg_m = 0.0
        for mac in res.get("machines", []):
            m_code = mac.get("code")
            if m_code:
                dg_m += (float(mac.get("rate", 0.0)) * float(mac_map.get(m_code, 0))) / divisor

        dg_tt = dg_vl + dg_nc + dg_m
        dg_truoc_thue = dg_tt * k_full

        tech_p = item.get("tech_profile") or self.get_default_tech_profile(code)

        return {
            "is_quote": False,
            "tech_profile": tech_p,
            "dg_vl": round(dg_vl, 0),
            "dg_nc": round(dg_nc, 0),
            "dg_m": round(dg_m, 0),
            "dg_tt": round(dg_tt, 0),
            "k_full": k_full,
            "dg_truoc_thue": round(dg_truoc_thue, 0),
            "thanh_tien_truoc_thue": round(vol * dg_truoc_thue, 0)
        }

    def aggregate_resources_from_norms(self) -> Dict[str, Any]:
        """
        Dynamically aggregate unique materials, labor, and machines
        directly from the norms database.
        """
        materials_dict = {}
        labor_dict = {}
        machines_dict = {}

        # Preserve existing prices if available
        existing_mat = {m["code"]: m for m in self.prices.get("materials", [])}
        existing_lab = {l["code"]: l for l in self.prices.get("labor", [])}
        existing_mac = {mac["code"]: mac for mac in self.prices.get("machines", [])}

        for norm_code, norm_data in self.norms.items():
            res = norm_data.get("resources", {})
            for m in res.get("materials", []):
                code = m.get("code")
                if code and code not in materials_dict:
                    base = existing_mat.get(code, {})
                    materials_dict[code] = {
                        "code": code,
                        "name": m.get("name"),
                        "unit": m.get("unit"),
                        "price_base": base.get("price_base", 0),
                        "price_notice": base.get("price_notice", 0),
                        "notes": m.get("note", "")
                    }
            for l in res.get("labor", []):
                code = l.get("code")
                if code and code not in labor_dict:
                    base = existing_lab.get(code, {})
                    labor_dict[code] = {
                        "code": code,
                        "name": l.get("name"),
                        "unit": l.get("unit"),
                        "price_base": base.get("price_base", 0),
                        "price_current": base.get("price_current", 0),
                        "legal_basis": base.get("legal_basis", "")
                    }
            for mac in res.get("machines", []):
                code = mac.get("code")
                if code and code not in machines_dict:
                    base = existing_mac.get(code, {})
                    machines_dict[code] = {
                        "code": code,
                        "name": mac.get("name"),
                        "unit": mac.get("unit"),
                        "price_base": base.get("price_base", 0),
                        "price_notice": base.get("price_notice", 0),
                        "price_diff": base.get("price_diff", 0),
                        "price_current": base.get("price_current", 0),
                        "legal_basis": base.get("legal_basis", "")
                    }

        aggregated = {
            "materials": list(materials_dict.values()),
            "labor": list(labor_dict.values()),
            "machines": list(machines_dict.values())
        }
        return aggregated

if __name__ == "__main__":
    mgr = NormManager()
    print("Norm Manager Initialized:")
    print("Total norms:", len(mgr.get_all_norm_codes()))
    prices = mgr.get_resource_prices()
    print(f"Total resources: {len(prices['materials'])} VL, {len(prices['labor'])} NC, {len(prices['machines'])} May")
