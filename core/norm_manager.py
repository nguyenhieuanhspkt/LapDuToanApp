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
from config import (
    NORMS_DB_FILE, PRICES_DB_FILE, IMAGES_DIR,
    PROJECT_ITEMS_FILE, LOCAL_PROJECT_ITEMS_FILE, DATABASE_DIR,
    PROJECT_VARIANTS_FILE, LOCAL_PROJECT_VARIANTS_FILE
)

REPLACEMENT_MATERIAL_CODES = {"V82782", "V82780", "V82781", "V82783"}

def get_unit_divisor(norm_unit: str, item_unit: str) -> float:
    """
    Scale factor between norm basis unit (e.g. 10 đầu, 5 bộ) and project work item unit.
    - If norm is for 10 units and item is also 10 units (e.g. '10 đầu'): divisor = 1.0 (no division)
    - If norm is for 10 units and item is 1 unit (e.g. 'đầu', 'bộ', 'cái'): divisor = 10.0
    - If norm is for 5 units and item is 1 unit: divisor = 5.0
    - If norm is for 100 units and item is 1 unit: divisor = 100.0
    """
    nu = (norm_unit or "").strip().lower()
    iu = (item_unit or "").strip().lower()
    if "10" in nu:
        return 1.0 if "10" in iu else 10.0
    if "5" in nu:
        return 1.0 if "5" in iu else 5.0
    if "100" in nu:
        return 1.0 if "100" in iu else 100.0
    return 1.0

class NormManager:
    @staticmethod
    def get_unit_divisor(norm_unit: str, item_unit: str) -> float:
        return get_unit_divisor(norm_unit, item_unit)

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

    def get_project_items_file(self) -> Path:
        """Get the primary project items JSON file path."""
        if DATABASE_DIR.exists():
            return PROJECT_ITEMS_FILE
        LOCAL_PROJECT_ITEMS_FILE.parent.mkdir(parents=True, exist_ok=True)
        return LOCAL_PROJECT_ITEMS_FILE

    def load_work_items(self) -> List[Dict[str, Any]]:
        """Load persistent work items from JSON database. Return empty list if not present."""
        target = self.get_project_items_file()
        if target.exists():
            try:
                with open(target, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            except Exception:
                pass
        elif LOCAL_PROJECT_ITEMS_FILE.exists():
            try:
                with open(LOCAL_PROJECT_ITEMS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            except Exception:
                pass
        return []

    def save_work_items(self, items: List[Dict[str, Any]]) -> bool:
        """Persist work items to JSON database file and sync local copy."""
        target = self.get_project_items_file()
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(target, "w", encoding="utf-8") as f:
                json.dump(items, f, ensure_ascii=False, indent=2)
            if target != LOCAL_PROJECT_ITEMS_FILE:
                try:
                    LOCAL_PROJECT_ITEMS_FILE.parent.mkdir(parents=True, exist_ok=True)
                    with open(LOCAL_PROJECT_ITEMS_FILE, "w", encoding="utf-8") as f2:
                        json.dump(items, f2, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            return True
        except Exception as e:
            print(f"Error saving work items: {e}")
            return False

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

    def get_project_variants_file(self) -> Path:
        """Return the target path for project variants JSON database file."""
        if PROJECT_VARIANTS_FILE.parent.exists():
            return PROJECT_VARIANTS_FILE
        return LOCAL_PROJECT_VARIANTS_FILE

    def load_project_variants(self) -> Dict[str, Any]:
        """Load project variants profile dictionary from database."""
        target = self.get_project_variants_file()
        if target.exists():
            try:
                with open(target, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
            except Exception as e:
                print(f"Error loading project variants from {target}: {e}")
        elif LOCAL_PROJECT_VARIANTS_FILE.exists():
            try:
                with open(LOCAL_PROJECT_VARIANTS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
            except Exception:
                pass
        return {}

    def save_project_variants(self, variants: Dict[str, Any]) -> bool:
        """Persist project variants to JSON database file."""
        target = self.get_project_variants_file()
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(target, "w", encoding="utf-8") as f:
                json.dump(variants, f, ensure_ascii=False, indent=2)
            if target != LOCAL_PROJECT_VARIANTS_FILE:
                try:
                    LOCAL_PROJECT_VARIANTS_FILE.parent.mkdir(parents=True, exist_ok=True)
                    with open(LOCAL_PROJECT_VARIANTS_FILE, "w", encoding="utf-8") as f2:
                        json.dump(variants, f2, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            return True
        except Exception as e:
            print(f"Error saving project variants: {e}")
            return False

    def get_variant(self, variant_id: str) -> Optional[Dict[str, Any]]:
        """Retrieve a specific project variant profile by its ID."""
        variants = self.load_project_variants()
        return variants.get(variant_id)

    def get_variants_for_norm(self, norm_code: str) -> Dict[str, Dict[str, Any]]:
        """Retrieve all active variant profiles belonging to a given base norm code as a dict mapping variant_id to profile."""
        code = (norm_code or "").strip()
        variants = self.load_project_variants()
        results = [v for v in variants.values() if v.get("norm_code", "").strip() == code]
        # Always ensure base variant A0 exists in memory if none defined
        has_base = any(v.get("variant_id", "").endswith("A0") or v.get("variant_id") == "A0" for v in results)
        if not has_base:
            norm = self.get_norm(code)
            if norm:
                base_steps = norm.get("work_steps", [])
                results.insert(0, {
                    "variant_id": "A0",
                    "variant_name": "Định mức gốc chuẩn 100% (Bao gồm vật tư thay thế)",
                    "norm_code": code,
                    "exclude_replacement": False,
                    "k_step": 1.0,
                    "work_steps": base_steps
                })
        return {v.get("variant_id", "A0"): v for v in results}

    def create_or_update_variant(
        self,
        variant_id: str,
        variant_name: str,
        norm_code: str,
        work_steps: List[Dict[str, Any]],
        exclude_replacement: bool = False
    ) -> Dict[str, Any]:
        """Create or update a centralized variant profile and save to DB."""
        variants = self.load_project_variants()
        total_base = sum(float(s.get("rate_base", 0.0)) for s in work_steps)
        active_base = sum(float(s.get("rate_base", 0.0)) for s in work_steps if int(s.get("apply_flag", 1)) == 1)
        k_step = (active_base / total_base) if total_base > 0 else 1.0

        var_obj = {
            "variant_id": variant_id.strip(),
            "variant_name": variant_name.strip(),
            "norm_code": norm_code.strip(),
            "exclude_replacement": bool(exclude_replacement),
            "k_step": round(k_step, 4),
            "work_steps": work_steps
        }
        variants[variant_id.strip()] = var_obj
        self.save_project_variants(variants)
        return var_obj

    def get_effective_item_resources(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """
        Calculates effective resource rates for a specific work item taking into account:
        - Centralized variant profile inheritance (variant_id) or custom work_steps
        - Proportional reduction factor k_step for non-replacement materials and machines
        - Replacement material exclusion rule:
          Replacement materials (V82782, V82780, V82781, V82783) are NOT scaled by k_step.
          If exclude_replacement is True, their rate is set to 0.0.
        """
        code = (item.get("norm_code") or "").strip()
        norm = self.get_norm(code)
        if not norm:
            return {
                "k_step": 1.0, "materials": [], "labor": [], "machines": [],
                "active_steps": [], "total_steps": 0, "raw_steps": [],
                "variant_id": item.get("variant_id", ""), "variant_name": "",
                "exclude_replacement": False
            }

        import copy
        variant_id = item.get("variant_id")
        variant_info = self.get_variant(variant_id) if variant_id else None
        
        exclude_replacement = bool(item.get("exclude_replacement", False))
        variant_name = ""
        if variant_info:
            raw_steps = copy.deepcopy(variant_info.get("work_steps", []))
            exclude_replacement = bool(variant_info.get("exclude_replacement", exclude_replacement))
            variant_name = variant_info.get("variant_name", "")
        elif item.get("work_steps"):
            raw_steps = copy.deepcopy(item.get("work_steps"))
        else:
            raw_steps = copy.deepcopy(norm.get("work_steps", []))

        total_base_labor = sum(float(s.get("rate_base", 0.0)) for s in raw_steps)
        active_steps = [s for s in raw_steps if int(s.get("apply_flag", 1)) == 1]
        active_base_labor = sum(float(s.get("rate_base", 0.0)) for s in active_steps)
        k_step = (active_base_labor / total_base_labor) if total_base_labor > 0 else 1.0

        has_split = any("rate_nc" in s or "rate_ks" in s for s in raw_steps)
        active_nc_labor = sum(float(s.get("rate_nc", 0.0)) for s in active_steps) if has_split else active_base_labor
        active_ks_labor = sum(float(s.get("rate_ks", 0.0)) for s in active_steps) if has_split else 0.0

        res = norm.get("resources", {})
        
        # Effective Materials:
        # Replacement materials (V82782, V82780, V82781, V82783) are NOT scaled by k_step.
        # If exclude_replacement is True, rate is 0.0.
        eff_materials = []
        for m in res.get("materials", []):
            m_copy = dict(m)
            m_code = m_copy.get("code")
            if m_code in REPLACEMENT_MATERIAL_CODES:
                m_copy["is_replacement"] = True
                if exclude_replacement:
                    m_copy["rate"] = 0.0
                    m_copy["excluded"] = True
                else:
                    m_copy["rate"] = float(m.get("rate", 0.0))
                    m_copy["excluded"] = False
            elif m_copy.get("code") != "V00750" and m_copy.get("unit") != "%":
                m_copy["rate"] = round(float(m.get("rate", 0.0)) * k_step, 4)
                m_copy["is_replacement"] = False
                m_copy["excluded"] = False
            else:
                m_copy["is_replacement"] = False
                m_copy["excluded"] = False
            eff_materials.append(m_copy)

        # Effective Machines: base rate * k_step (proportional reduction)
        eff_machines = []
        for mac in res.get("machines", []):
            mac_copy = dict(mac)
            mac_copy["rate"] = round(float(mac.get("rate", 0.0)) * k_step, 4)
            eff_machines.append(mac_copy)

        # Effective Labor: calculated based on selected steps
        eff_labor = []
        for l in res.get("labor", []):
            l_copy = dict(l)
            l_name = l.get("name", "").lower()
            if has_split and ("kỹ sư" in l_name or "ks" in l_name):
                l_copy["rate"] = round(active_ks_labor, 4)
            elif has_split:
                l_copy["rate"] = round(active_nc_labor, 4)
            else:
                l_copy["rate"] = round(active_base_labor, 4)
            eff_labor.append(l_copy)

        return {
            "k_step": k_step,
            "raw_steps": raw_steps,
            "active_steps": active_steps,
            "total_steps": len(raw_steps),
            "materials": eff_materials,
            "labor": eff_labor,
            "machines": eff_machines,
            "variant_id": variant_id or "A0",
            "variant_name": variant_name,
            "exclude_replacement": exclude_replacement
        }

    def calculate_item_unit_rate(
        self,
        item: Dict[str, Any],
        coefficients: Optional[Dict[str, Any]] = None,
        custom_prices: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Calculate horizontal unit rates for a single work item:
        - Accurately handles unit divisor between norm unit and item unit
        - Supports centralized variant profile and exclude_replacement rule
        - Supports hybrid norm + quote items
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
        is_pure_quote = (item.get("is_quote", False) and not self.get_norm(code)) or code.startswith("BG") or code.startswith("TT") or not code
        unit_price_quote = float(item.get("unit_price_quote", 0.0))

        if is_pure_quote:
            quote_price = unit_price_quote
            apply_factor = item.get("apply_overhead", False)
            dg_truoc_thue = quote_price * k_full if apply_factor else quote_price
            return {
                "is_quote": True,
                "variant_id": "",
                "variant_name": "",
                "tech_profile": item.get("tech_profile", "Công tác báo giá trọn gói"),
                "k_step": 1.0,
                "k_pct_str": "100%",
                "active_steps_count": 0,
                "total_steps_count": 0,
                "dg_vl": 0.0,
                "dg_nc": quote_price,
                "dg_m": 0.0,
                "dg_tt": quote_price,
                "k_full": k_full if apply_factor else 1.0,
                "dg_dm_truoc_thue": 0.0,
                "unit_price_quote": quote_price,
                "dg_truoc_thue": dg_truoc_thue,
                "thanh_tien_truoc_thue": vol * dg_truoc_thue,
                "exclude_replacement": False
            }

        norm = self.get_norm(code)
        if not norm:
            return {
                "is_quote": False,
                "variant_id": "",
                "variant_name": "",
                "tech_profile": "Chưa xác định",
                "k_step": 1.0,
                "k_pct_str": "100%",
                "active_steps_count": 0,
                "total_steps_count": 0,
                "dg_vl": 0.0, "dg_nc": 0.0, "dg_m": 0.0, "dg_tt": 0.0,
                "k_full": k_full, "dg_dm_truoc_thue": 0.0,
                "unit_price_quote": unit_price_quote,
                "dg_truoc_thue": unit_price_quote,
                "thanh_tien_truoc_thue": vol * unit_price_quote,
                "exclude_replacement": False
            }

        norm_unit = norm.get("unit", "")
        item_unit = item.get("unit", norm_unit)
        divisor = get_unit_divisor(norm_unit, item_unit)

        # Resource price lookup maps
        prices_db = custom_prices or self.get_resource_prices()
        vl_map = {m["code"]: m.get("price_notice", 0) for m in prices_db.get("materials", [])}
        nc_map = {l["code"]: l.get("price_current", 0) for l in prices_db.get("labor", [])}
        mac_map = {mac["code"]: mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0)) for mac in prices_db.get("machines", [])}

        # Retrieve effective resource rates
        eff = self.get_effective_item_resources(item)
        k_step = eff["k_step"]
        exclude_replacement = eff["exclude_replacement"]
        
        # 1. Material unit rate
        dg_vl = 0.0
        for m in eff.get("materials", []):
            m_code = m.get("code")
            if m_code and m_code != "V00750" and m.get("unit") != "%":
                rate_val = float(m.get("rate", 0.0))
                price_val = 0.0 if (m_code in REPLACEMENT_MATERIAL_CODES and exclude_replacement) else float(vl_map.get(m_code, 0))
                dg_vl += (rate_val * price_val) / divisor
        for m in eff.get("materials", []):
            if m.get("code") == "V00750" or m.get("unit") == "%":
                pct = float(m.get("rate", 0.0)) / 100.0
                dg_vl += dg_vl * pct

        # 2. Labor unit rate
        dg_nc = 0.0
        for l in eff.get("labor", []):
            l_code = l.get("code")
            if l_code:
                dg_nc += (float(l.get("rate", 0.0)) * float(nc_map.get(l_code, 0))) / divisor

        # 3. Machine unit rate
        dg_m = 0.0
        for mac in eff.get("machines", []):
            m_code = mac.get("code")
            if m_code:
                dg_m += (float(mac.get("rate", 0.0)) * float(mac_map.get(m_code, 0))) / divisor

        dg_tt = dg_vl + dg_nc + dg_m
        dg_dm_truoc_thue = dg_tt * k_full

        # Hybrid model: add pre-tax quote price if any
        dg_truoc_thue = dg_dm_truoc_thue + unit_price_quote

        tech_p = item.get("tech_profile") or self.get_default_tech_profile(code)

        return {
            "is_quote": False,
            "variant_id": eff.get("variant_id", "A0"),
            "variant_name": eff.get("variant_name", ""),
            "exclude_replacement": exclude_replacement,
            "tech_profile": tech_p,
            "k_step": k_step,
            "k_pct_str": f"{round(k_step * 100)}%",
            "active_steps_count": len(eff["active_steps"]),
            "total_steps_count": eff["total_steps"],
            "divisor": divisor,
            "dg_vl": round(dg_vl, 0),
            "dg_nc": round(dg_nc, 0),
            "dg_m": round(dg_m, 0),
            "dg_tt": round(dg_tt, 0),
            "k_full": k_full,
            "dg_dm_truoc_thue": round(dg_dm_truoc_thue, 0),
            "unit_price_quote": unit_price_quote,
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

    def get_province_machines(self, province: str = "lam_dong") -> List[Dict[str, Any]]:
        """Get machine list for a given province (e.g. lam_dong)."""
        if province == "lam_dong":
            from config import LAM_DONG_MACHINES_FILE
            if LAM_DONG_MACHINES_FILE.exists():
                with open(LAM_DONG_MACHINES_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        return []

    def lookup_machine_price(self, machine_code: str, province: str = "lam_dong", region: str = "vung_4") -> Optional[Dict[str, Any]]:
        """Look up machine price and metadata from province database."""
        reg_key = region.lower().replace(" ", "_")
        if any(k in reg_key for k in ["vùng_iv", "vung_iv", "vùng_4", "vung_4", "4"]):
            reg_key = "vung_4"
        elif any(k in reg_key for k in ["vùng_iii", "vung_iii", "vùng_3", "vung_3", "3"]):
            reg_key = "vung_3"
        elif any(k in reg_key for k in ["vùng_ii", "vung_ii", "vùng_2", "vung_2", "2"]):
            reg_key = "vung_2"
        else:
            reg_key = "vung_4"

        machines = self.get_province_machines(province)
        for m in machines:
            if m.get("code") == machine_code:
                price_val = m.get("price", {}).get(reg_key, 0)
                labor_val = m.get("labor_cost", {}).get(reg_key, 0)
                return {
                    "code": m.get("code"),
                    "name": m.get("name"),
                    "unit": m.get("unit", "ca"),
                    "driver": m.get("driver", ""),
                    "price_current": price_val,
                    "labor_cost": labor_val,
                    "fuel_cost": m.get("fuel_cost", 0),
                    "legal_basis": m.get("legal_basis", ""),
                    "region": reg_key
                }
        return None

if __name__ == "__main__":
    mgr = NormManager()
    print("Norm Manager Initialized:")
    print("Total norms:", len(mgr.get_all_norm_codes()))
    prices = mgr.get_resource_prices()
    print(f"Total resources: {len(prices['materials'])} VL, {len(prices['labor'])} NC, {len(prices['machines'])} May")
