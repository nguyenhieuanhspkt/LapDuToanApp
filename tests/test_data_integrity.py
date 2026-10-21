"""
Test Data Integrity for Day 1:
Verifies that all norms, evidence images, and resource prices are 100% complete and consistent.
"""

import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import NORMS_DB_FILE, PRICES_DB_FILE, IMAGES_DIR, check_environment
from core.norm_manager import NormManager

def run_tests():
    print("==================================================")
    print("RUNNING DAY 1 DATA INTEGRITY TESTS")
    print("==================================================")
    
    # 1. Check environment
    env = check_environment()
    assert env["status"], f"Environment check failed: {env['issues']}"
    print("[PASS] 1. Environment & Directories check")

    # 2. Check Norms Database
    mgr = NormManager()
    norms = mgr.norms
    assert len(norms) == 13, f"Expected 13 norms, got {len(norms)}"
    print(f"[PASS] 2. Norms count check: 13 norms loaded")

    # 3. Check Evidence Images
    missing_images = []
    for code in norms:
        img_path = mgr.get_evidence_image_path(code)
        if not img_path or not img_path.exists() or img_path.stat().st_size == 0:
            missing_images.append(code)
    assert len(missing_images) == 0, f"Missing evidence images for: {missing_images}"
    print(f"[PASS] 3. Evidence images check: 13/13 images exist with valid size")

    # 4. Check Resource Prices coverage
    prices = mgr.get_resource_prices()
    mat_codes = {m["code"] for m in prices.get("materials", [])}
    lab_codes = {l["code"] for l in prices.get("labor", [])}
    mac_codes = {mac["code"] for mac in prices.get("machines", [])}

    missing_resources = []
    for norm_code, norm_data in norms.items():
        res = norm_data.get("resources", {})
        for m in res.get("materials", []):
            if m["code"] and m["code"] not in mat_codes:
                missing_resources.append(f"Material {m['code']} in {norm_code}")
        for l in res.get("labor", []):
            if l["code"] and l["code"] not in lab_codes:
                missing_resources.append(f"Labor {l['code']} in {norm_code}")
        for mac in res.get("machines", []):
            if mac["code"] and mac["code"] not in mac_codes:
                missing_resources.append(f"Machine {mac['code']} in {norm_code}")

    assert len(missing_resources) == 0, f"Unmapped resources: {missing_resources}"
    print(f"[PASS] 4. Resource prices coverage: 100% covered ({len(mat_codes)} VL, {len(lab_codes)} NC, {len(mac_codes)} May)")

    # 5. Check Work Steps
    for code, data in norms.items():
        steps = data.get("work_steps", [])
        assert len(steps) >= 5, f"Norm {code} has only {len(steps)} work steps"
    print("[PASS] 5. Work steps check: All norms have detailed step breakdown")

    print("\n==================================================")
    print("ALL DAY 1 TESTS PASSED SUCCESSFULLY! (100% READY)")
    print("==================================================")
    return True

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
