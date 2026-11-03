"""
Excel Generator Module for LapDuToanApp.
Orchestrates generation of complete Excel workbooks with 100% dynamic formula links
by delegating individual sheet construction to specialized builders in core.sheets package:
- Sheet TH_KinhPhi (SheetTHKinhPhiBuilder)
- Sheet Khối lượng-Định mức-ĐVT (SheetKhoiLuongBuilder)
- Sheet 00_ThongKe_DinhMuc (SheetThongKeDinhMucBuilder)
- Sheet 01_ChiTiet_HaoPhi_DinhMuc (SheetDonGiaChiTietBuilder)
- Sheet 02_GiaiTrinh_HeSo_DieuChinh (SheetGiaiTrinhHeSoBuilder)
- Sheet TH_VatLieu (SheetTHVatLieuBuilder)
- Sheet TH_NhanCong (SheetTHNhanCongBuilder)
- Sheet TH_MayThiCong (SheetTHMayThiCongBuilder)
- Individual Norm Evidence Sheets (SheetEvidenceBuilder)

Author: Tổ Thẩm Định EVN
"""

import openpyxl
from pathlib import Path
from typing import Dict, List, Any, Optional
import sys

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import IMAGES_DIR, OUTPUT_DIR
from core.norm_manager import NormManager
from core.constants import ORDERED_STANDARD_SHEET_NAMES, REPLACEMENT_MATERIAL_CODES
from core.business import get_unit_divisor
from core.sheets import (
    BaseSheetBuilder,
    SheetTHKinhPhiBuilder,
    SheetKhoiLuongBuilder,
    SheetDonGiaChiTietBuilder,
    SheetGiaiTrinhHeSoBuilder,
    SheetThongKeDinhMucBuilder,
    SheetEvidenceBuilder,
    SheetTHVatLieuBuilder,
    SheetTHNhanCongBuilder,
    SheetTHMayThiCongBuilder
)

DEFAULT_COEFFICIENTS = {
    "c": {"rate": 7.3, "basis": "Thông tư số 11/2021/TT-BXD (Định mức chi phí chung công trình công nghiệp)"},
    "tt": {"rate": 2.5, "basis": "Thông tư số 11/2021/TT-BXD (Công việc không xác định được KL từ TK)"},
    "tl": {"rate": 5.5, "basis": "Thông tư số 11/2021/TT-BXD (Thu nhập chịu thuế tính trước)"},
    "vat": {"rate": 8.0, "basis": "Nghị định 72/2024/NĐ-CP của Chính phủ"},
    "lt": {"rate": 1.1, "basis": "Thông tư số 11/2021/TT-BXD (Chi phí nhà tạm để ở và ĐHTC công trình ngoài đô thị)"}
}


class ExcelGenerator:
    """
    Main Orchestrator for Excel Workbook Generation.
    Coordinates specialized sheet builders to produce appraisal-standard workbooks.
    """

    def __init__(self, norm_manager: Optional[NormManager] = None):
        self.norm_manager = norm_manager or NormManager()

        # Initialize modular sheet builders
        self.builder_th_kinh_phi = SheetTHKinhPhiBuilder(self.norm_manager)
        self.builder_khoi_luong = SheetKhoiLuongBuilder(self.norm_manager)
        self.builder_don_gia_chi_tiet = SheetDonGiaChiTietBuilder(self.norm_manager)
        self.builder_giai_trinh_he_so = SheetGiaiTrinhHeSoBuilder(self.norm_manager)
        self.builder_thong_ke = SheetThongKeDinhMucBuilder(self.norm_manager)
        self.builder_evidence = SheetEvidenceBuilder(self.norm_manager)
        self.builder_th_vat_lieu = SheetTHVatLieuBuilder(self.norm_manager)
        self.builder_th_nhan_cong = SheetTHNhanCongBuilder(self.norm_manager)
        self.builder_th_may_thi_cong = SheetTHMayThiCongBuilder(self.norm_manager)

        # Expose styles for backward compatibility
        self.init_styles()

    def init_styles(self):
        """Expose standard styling definitions for backward compatibility."""
        base = BaseSheetBuilder(self.norm_manager)
        self.font_title = base.font_title
        self.font_subtitle = base.font_subtitle
        self.font_project = base.font_project
        self.font_unit = base.font_unit
        self.font_header = base.font_header
        self.font_code = base.font_code
        self.font_regular = base.font_regular
        self.font_bold = base.font_bold
        self.font_cat = base.font_cat
        self.font_blue = base.font_blue
        self.font_total = base.font_total

        self.font_ev_title = base.font_ev_title
        self.font_ev_header = base.font_ev_header
        self.font_ev_reg = base.font_ev_reg
        self.font_ev_bold = base.font_ev_bold
        self.font_ev_cat = base.font_ev_cat

        self.fill_code = base.fill_code
        self.fill_yellow = base.fill_yellow
        self.fill_ev_header = base.fill_ev_header

        self.border_cell = base.border_cell
        self.border_total = base.border_total

        self.align_center = base.align_center
        self.align_left = base.align_left
        self.align_right = base.align_right

    # Backwards-compatible delegation methods
    def _determine_group(self, item: Dict[str, Any]) -> str:
        return self.builder_khoi_luong.determine_group(item)

    def _build_sheet_th_kinh_phi(self, *args, **kwargs):
        return self.builder_th_kinh_phi.build(*args, **kwargs)

    def _build_sheet_giai_trinh_he_so(self, *args, **kwargs):
        return self.builder_giai_trinh_he_so.build(*args, **kwargs)

    def _build_sheet_khoi_luong(self, *args, **kwargs):
        return self.builder_khoi_luong.build(*args, **kwargs)

    def _build_sheet_thong_ke(self, *args, **kwargs):
        return self.builder_thong_ke.build(*args, **kwargs)

    def _build_evidence_sheet(self, *args, **kwargs):
        return self.builder_evidence.build(*args, **kwargs)

    def _build_sheet_don_gia_chi_tiet(self, *args, **kwargs):
        return self.builder_don_gia_chi_tiet.build(*args, **kwargs)

    def _build_sheet_th_vat_lieu(self, *args, **kwargs):
        return self.builder_th_vat_lieu.build(*args, **kwargs)

    def _build_sheet_th_nhan_cong(self, *args, **kwargs):
        return self.builder_th_nhan_cong.build(*args, **kwargs)

    def _build_sheet_th_may_thi_cong(self, *args, **kwargs):
        return self.builder_th_may_thi_cong.build(*args, **kwargs)

    def generate(
        self,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]],
        output_path: Path,
        coefficients: Optional[Dict[str, Any]] = None,
        custom_prices: Optional[Dict[str, Any]] = None
    ) -> Path:
        """
        Main generator method: builds the complete workbook and writes to output_path.
        Includes TH_KinhPhi, KhoiLuong, ThongKe, DonGiaChiTiet, TH_VL, TH_NC, TH_M, and Norm sheets.
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        coeffs = DEFAULT_COEFFICIENTS.copy()
        if coefficients:
            for k, v in coefficients.items():
                if k in coeffs:
                    coeffs[k].update(v)
                else:
                    coeffs[k] = v

        # 1. Determine all unique norm codes used in the work items
        unique_norm_codes = []
        for item in work_items:
            code = item["norm_code"]
            if code not in unique_norm_codes and code in self.norm_manager.norms:
                unique_norm_codes.append(code)

        # 2. Build 3 Resource Lookup Tables (for TH_VatLieu, TH_NhanCong, TH_MayThiCong)
        res_data = self.norm_manager.aggregate_resources_from_norms()

        # Override with custom prices if provided
        if custom_prices:
            p_mat = {m["code"]: m for m in custom_prices.get("materials", [])}
            p_lab = {l["code"]: l for l in custom_prices.get("labor", [])}
            p_mac = {mac["code"]: mac for mac in custom_prices.get("machines", [])}

            for m in res_data.get("materials", []):
                if m["code"] in p_mat:
                    m.update(p_mat[m["code"]])
            for l in res_data.get("labor", []):
                if l["code"] in p_lab:
                    l.update(p_lab[l["code"]])
            for mac in res_data.get("machines", []):
                if mac["code"] in p_mac:
                    mac.update(p_mac[mac["code"]])

        used_mat_codes = set()
        used_lab_codes = set()
        used_mac_codes = set()

        # Compute total required quantity for each resource across all work items
        mat_total_quantities = {}
        lab_total_quantities = {}
        mac_total_quantities = {}

        for item in work_items:
            code = item.get("norm_code", "").strip()
            vol = float(item.get("volume", 0.0))
            norm = self.norm_manager.get_norm(code)
            if not norm:
                continue

            norm_unit_str = norm.get("unit", "")
            divisor = 1.0
            if "10" in norm_unit_str:
                divisor = 10.0
            elif "5" in norm_unit_str:
                divisor = 5.0

            eff = self.norm_manager.get_effective_item_resources(item)
            for m in eff.get("materials", []):
                c = m.get("code")
                if c:
                    used_mat_codes.add(c)
                    mat_total_quantities[c] = mat_total_quantities.get(c, 0.0) + (vol * float(m.get("rate", 0.0)) / divisor)
            for l in eff.get("labor", []):
                c = l.get("code")
                if c:
                    used_lab_codes.add(c)
                    lab_total_quantities[c] = lab_total_quantities.get(c, 0.0) + (vol * float(l.get("rate", 0.0)) / divisor)
            for mac in eff.get("machines", []):
                c = mac.get("code")
                if c:
                    used_mac_codes.add(c)
                    mac_total_quantities[c] = mac_total_quantities.get(c, 0.0) + (vol * float(mac.get("rate", 0.0)) / divisor)

        # Retain order from database
        materials_list = [m for m in res_data["materials"] if m["code"] in used_mat_codes or not used_mat_codes]
        labor_list = [l for l in res_data["labor"] if l["code"] in used_lab_codes or not used_lab_codes]
        machines_list = [mac for mac in res_data["machines"] if mac["code"] in used_mac_codes or not used_mac_codes]

        # Attach computed total quantities
        for m in materials_list:
            m["total_qty"] = round(mat_total_quantities.get(m["code"], 0.0), 3)
        for l in labor_list:
            l["total_qty"] = round(lab_total_quantities.get(l["code"], 0.0), 3)
        for mac in machines_list:
            mac["total_qty"] = round(mac_total_quantities.get(mac["code"], 0.0), 3)

        # Filter out unused resources
        materials_list = [m for m in materials_list if m.get("total_qty", 0.0) > 0 or m.get("code") == "V00750" or m.get("unit") == "%"]
        labor_list = [l for l in labor_list if l.get("total_qty", 0.0) > 0]
        machines_list = [mac for mac in machines_list if mac.get("total_qty", 0.0) > 0]

        # Row mappings in TH sheets (headers start at row 6, data rows start at row 7)
        mat_row_map = {m["code"]: 7 + idx for idx, m in enumerate(materials_list)}
        lab_row_map = {l["code"]: 7 + idx for idx, l in enumerate(labor_list)}
        mac_row_map = {mac["code"]: 7 + idx for idx, mac in enumerate(machines_list)}

        r_total_vl = 7 + len(materials_list)
        r_total_nc = 7 + len(labor_list)
        r_total_mac = 7 + len(machines_list)

        # Compute quote items sum (including hybrid quote prices)
        quote_sum = sum(float(it.get("unit_price_quote", 0.0)) * float(it.get("volume", 0.0)) for it in work_items if float(it.get("unit_price_quote", 0.0)) > 0)
        quote_count = sum(1 for it in work_items if float(it.get("unit_price_quote", 0.0)) > 0)

        # 1. Create Sheet: 00_ThongKe_DinhMuc
        self.builder_thong_ke.build(wb, unique_norm_codes)

        # 2. Create Evidence Sheets for each unique norm code
        norm_cell_map = {}
        for code in unique_norm_codes:
            norm_cell_map[code] = self.builder_evidence.build(wb, code)

        # 3. Create Sheet: 01_ChiTiet_HaoPhi_DinhMuc (returns item_subtotal_map for linking)
        item_subtotal_map = self.builder_don_gia_chi_tiet.build(
            wb, project_name, category_name, work_items,
            norm_cell_map, mat_row_map, lab_row_map, mac_row_map
        )

        # 4. Create Sheet: 02_GiaiTrinh_HeSo_DieuChinh (thống kê biến thể A0, A1, A2 và so sánh 2 cột)
        self.builder_giai_trinh_he_so.build(wb, project_name, category_name, work_items)

        # 5. Create Sheet: Khối lượng-Định mức-ĐVT (Bảng 19 cột đầy đủ công thức link Bảng 01)
        r_grand_total_kl = self.builder_khoi_luong.build(
            wb, project_name, category_name, work_items,
            item_subtotal_map=item_subtotal_map,
            coefficients=coeffs
        )

        # 6. Create Sheet: TH_KinhPhi (BẢNG TỔNG HỢP KINH PHÍ DỰ TOÁN - link trực tiếp vào bảng Khối lượng)
        self.builder_th_kinh_phi.build(
            wb=wb,
            project_name=project_name,
            category_name=category_name,
            coefficients=coeffs,
            r_grand_total_kl=r_grand_total_kl,
            r_total_vl=r_total_vl,
            r_total_nc=r_total_nc,
            r_total_mac=r_total_mac,
            materials=materials_list,
            labor=labor_list,
            machines=machines_list,
            quote_sum=quote_sum,
            quote_count=quote_count
        )

        # 7. Create Sheet: TH_VatLieu
        self.builder_th_vat_lieu.build(wb, project_name, category_name, materials_list)

        # 8. Create Sheet: TH_NhanCong
        self.builder_th_nhan_cong.build(wb, project_name, category_name, labor_list)

        # 9. Create Sheet: TH_MayThiCong
        self.builder_th_may_thi_cong.build(wb, project_name, category_name, machines_list)

        # Reorder sheets cleanly using standard constants:
        ordered_names = list(ORDERED_STANDARD_SHEET_NAMES)
        for code in unique_norm_codes:
            ordered_names.append(code)

        wb._sheets = [wb[name] for name in ordered_names if name in wb.sheetnames]

        # Save workbook
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        return output_path
