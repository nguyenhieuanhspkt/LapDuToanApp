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
        """Load norms and prices from JSON database."""
        if self.norms_file.exists():
            with open(self.norms_file, "r", encoding="utf-8") as f:
                self.norms = json.load(f)
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
