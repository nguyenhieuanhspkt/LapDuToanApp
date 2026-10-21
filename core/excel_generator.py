"""
Excel Generator Module for LapDuToanApp.
Generates complete Excel workbooks with 100% dynamic formula links:
- Sheet Khối lượng-Định mức-ĐVT (Bảng tiên lượng khối lượng công việc)
- Sheet 00_ThongKe_DinhMuc (Hyperlinks)
- Sheet 01_ChiTiet_HaoPhi_DinhMuc (Dynamic G links to norm sheets, H links to TH sheets)
- Sheet TH_VatLieu
- Sheet TH_NhanCong
- Sheet TH_MayThiCong
- Individual Norm Evidence Sheets (Images at B3 + Work steps table + Resource table)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
from pathlib import Path
from typing import Dict, List, Any, Optional
import os
import sys

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from config import IMAGES_DIR, OUTPUT_DIR
from core.norm_manager import NormManager
from utils.num2words_vi import num2words_vi

DEFAULT_COEFFICIENTS = {
    "c": {"rate": 7.3, "basis": "Thông tư số 11/2021/TT-BXD (Định mức chi phí chung công trình công nghiệp)"},
    "tt": {"rate": 2.5, "basis": "Thông tư số 11/2021/TT-BXD (Công việc không xác định được KL từ TK)"},
    "tl": {"rate": 5.5, "basis": "Thông tư số 11/2021/TT-BXD (Thu nhập chịu thuế tính trước)"},
    "vat": {"rate": 8.0, "basis": "Nghị định 72/2024/NĐ-CP của Chính phủ"},
    "lt": {"rate": 1.1, "basis": "Thông tư số 11/2021/TT-BXD (Chi phí nhà tạm để ở và ĐHTC công trình ngoài đô thị)"}
}

class ExcelGenerator:
    def __init__(self, norm_manager: Optional[NormManager] = None):
        self.norm_manager = norm_manager or NormManager()
        self.init_styles()

    def init_styles(self):
        """Define reusable fonts, fills, alignments, and borders."""
        self.font_title = Font(name="Times New Roman", size=13, bold=True)
        self.font_subtitle = Font(name="Times New Roman", size=11, italic=True)
        self.font_project = Font(name="Times New Roman", size=11, bold=True)
        self.font_unit = Font(name="Times New Roman", size=10, italic=True)
        self.font_header = Font(name="Times New Roman", size=11, bold=True)
        self.font_code = Font(name="Times New Roman", size=11, bold=True)
        self.font_regular = Font(name="Times New Roman", size=11)
        self.font_bold = Font(name="Times New Roman", size=11, bold=True)
        self.font_cat = Font(name="Times New Roman", size=11, bold=True, italic=True, color="002060")
        self.font_blue = Font(name="Times New Roman", size=11, color="0000FF")
        self.font_total = Font(name="Times New Roman", size=11, bold=True, italic=True, color="0000FF")

        # Evidence sheet fonts
        self.font_ev_title = Font(name="Times New Roman", size=11, bold=True, color="002060")
        self.font_ev_header = Font(name="Times New Roman", size=10, bold=True)
        self.font_ev_reg = Font(name="Times New Roman", size=10)
        self.font_ev_bold = Font(name="Times New Roman", size=10, bold=True)
        self.font_ev_cat = Font(name="Times New Roman", size=10, bold=True, italic=True)

        self.fill_code = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        self.fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        self.fill_ev_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        thin_side = Side(border_style="thin", color="A6A6A6")
        self.border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        self.border_total = Border(top=thin_side, bottom=Side(border_style="double", color="0000FF"), left=thin_side, right=thin_side)

        self.align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")

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
            # Map custom prices if available
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
            code = item["norm_code"]
            vol = float(item.get("volume", 0.0))
            norm = self.norm_manager.get_norm(code)
            if not norm: continue
            
            norm_unit_str = norm.get("unit", "")
            divisor = 1.0
            if "10" in norm_unit_str: divisor = 10.0
            elif "5" in norm_unit_str: divisor = 5.0
            
            res = norm.get("resources", {})
            for m in res.get("materials", []):
                c = m.get("code")
                if c:
                    used_mat_codes.add(c)
                    mat_total_quantities[c] = mat_total_quantities.get(c, 0.0) + (vol * float(m.get("rate", 0.0)) / divisor)
            for l in res.get("labor", []):
                c = l.get("code")
                if c:
                    used_lab_codes.add(c)
                    lab_total_quantities[c] = lab_total_quantities.get(c, 0.0) + (vol * float(l.get("rate", 0.0)) / divisor)
            for mac in res.get("machines", []):
                c = mac.get("code")
                if c:
                    used_mac_codes.add(c)
                    mac_total_quantities[c] = mac_total_quantities.get(c, 0.0) + (vol * float(mac.get("rate", 0.0)) / divisor)

        # Retain order from database
        materials_list = [m for m in res_data["materials"] if m["code"] in used_mat_codes or not used_mat_codes]
        labor_list = [l for l in res_data["labor"] if l["code"] in used_lab_codes or not used_lab_codes]
        machines_list = [mac for mac in res_data["machines"] if mac["code"] in used_mac_codes or not used_mac_codes]

        # Attach computed total quantities
        for m in materials_list: m["total_qty"] = round(mat_total_quantities.get(m["code"], 0.0), 3)
        for l in labor_list: l["total_qty"] = round(lab_total_quantities.get(l["code"], 0.0), 3)
        for mac in machines_list: mac["total_qty"] = round(mac_total_quantities.get(mac["code"], 0.0), 3)

        # Row mappings in TH sheets (headers start at row 6, data rows start at row 7)
        mat_row_map = {m["code"]: 7 + idx for idx, m in enumerate(materials_list)}
        lab_row_map = {l["code"]: 7 + idx for idx, l in enumerate(labor_list)}
        mac_row_map = {mac["code"]: 7 + idx for idx, mac in enumerate(machines_list)}

        r_total_vl = 7 + len(materials_list)
        r_total_nc = 7 + len(labor_list)
        r_total_mac = 7 + len(machines_list)

        # Compute quote items sum if any
        quote_items = [it for it in work_items if it.get("is_quote")]
        quote_sum = sum(float(it.get("unit_price_quote", 0.0)) * float(it.get("volume", 0.0)) for it in quote_items)

        # 3. Create Sheet: TH_KinhPhi (BẢNG TỔNG HỢP KINH PHÍ DỰ TOÁN - SHEET ĐẦU TIÊN)
        self._build_sheet_th_kinh_phi(
            wb=wb,
            project_name=project_name,
            category_name=category_name,
            coefficients=coeffs,
            r_total_vl=r_total_vl,
            r_total_nc=r_total_nc,
            r_total_mac=r_total_mac,
            materials=materials_list,
            labor=labor_list,
            machines=machines_list,
            quote_sum=quote_sum,
            quote_count=len(quote_items)
        )

        # 4. Create Sheet: Khối lượng-Định mức-ĐVT (Bảng tiên lượng công việc)
        self._build_sheet_khoi_luong(wb, project_name, category_name, work_items)

        # 5. Create Sheet: 00_ThongKe_DinhMuc
        self._build_sheet_thong_ke(wb, unique_norm_codes)

        # 6. Create Evidence Sheets for each unique norm code
        norm_cell_map = {}
        for code in unique_norm_codes:
            norm_cell_map[code] = self._build_evidence_sheet(wb, code)

        # 7. Create Sheet: 01_ChiTiet_HaoPhi_DinhMuc
        self._build_sheet_don_gia_chi_tiet(
            wb, project_name, category_name, work_items,
            norm_cell_map, mat_row_map, lab_row_map, mac_row_map
        )

        # 8. Create Sheet: TH_VatLieu
        self._build_sheet_th_vat_lieu(wb, project_name, category_name, materials_list)

        # 9. Create Sheet: TH_NhanCong
        self._build_sheet_th_nhan_cong(wb, project_name, category_name, labor_list)

        # 10. Create Sheet: TH_MayThiCong
        self._build_sheet_th_may_thi_cong(wb, project_name, category_name, machines_list)

        # Reorder sheets cleanly:
        # 1. TH_KinhPhi
        # 2. Khối lượng-Định mức-ĐVT
        # 3. 00_ThongKe_DinhMuc
        # 4. 01_ChiTiet_HaoPhi_DinhMuc
        # 5. TH_VatLieu
        # 6. TH_NhanCong
        # 7. TH_MayThiCong
        # 8. [Evidence Sheets]
        ordered_names = [
            "TH_KinhPhi",
            "Khối lượng-Định mức-ĐVT",
            "00_ThongKe_DinhMuc",
            "01_ChiTiet_HaoPhi_DinhMuc",
            "TH_VatLieu",
            "TH_NhanCong",
            "TH_MayThiCong"
        ]
        for code in unique_norm_codes:
            ordered_names.append(code)
            
        wb._sheets = [wb[name] for name in ordered_names if name in wb.sheetnames]

        # Save workbook
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        return output_path

    # =========================================================================
    # INTERNAL BUILDERS
    # =========================================================================

    def _build_sheet_th_kinh_phi(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        coefficients: Dict[str, Any],
        r_total_vl: int,
        r_total_nc: int,
        r_total_mac: int,
        materials: List[Dict[str, Any]],
        labor: List[Dict[str, Any]],
        machines: List[Dict[str, Any]],
        quote_sum: float = 0.0,
        quote_count: int = 0
    ):
        """
        Builds the summary cost estimation sheet (Bảng tổng hợp dự toán chi phí xây dựng)
        matching 100% the layout and formulas of EVN standard format.
        """
        ws = wb.create_sheet(title="TH_KinhPhi")
        ws.views.sheetView[0].showGridLines = True

        # Titles
        ws["A1"] = "BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG"
        ws["A1"].font = self.font_title
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A2"].font = self.font_project
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["A3"].font = self.font_project
        ws["F4"] = "Đơn vị: đồng"
        ws["F4"].font = self.font_unit
        ws["F4"].alignment = self.align_right

        # Table Headers
        headers = ["STT", "NỘI DUNG CHI PHÍ", "CÁCH TÍNH", "GIÁ TRỊ", "KÝ HIỆU", "CƠ SỞ HỆ SỐ / GHI CHÚ"]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_ev_header

        # Coefficients
        c_rate = float(coefficients.get("c", {}).get("rate", 7.3))
        c_basis = coefficients.get("c", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        tt_rate = float(coefficients.get("tt", {}).get("rate", 2.5))
        tt_basis = coefficients.get("tt", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        tl_rate = float(coefficients.get("tl", {}).get("rate", 5.5))
        tl_basis = coefficients.get("tl", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        vat_rate = float(coefficients.get("vat", {}).get("rate", 8.0))
        vat_basis = coefficients.get("vat", {}).get("basis", "Nghị định 72/2024/NĐ-CP")
        lt_rate = float(coefficients.get("lt", {}).get("rate", 1.1))
        lt_basis = coefficients.get("lt", {}).get("basis", "Thông tư 11/2021/TT-BXD")

        # Row Data Definitions: (STT, Content, Calculation_Text, Formula_Value, Symbol, Note, is_bold, is_highlight)
        rows_spec = []
        r_T = 6
        if quote_sum > 0:
            formula_T = "=D7+D9+D11+D13"
            calc_T = "T = VL + NC + M + BG"
            r_GT = 15
            r_C = 16
            r_TT = 17
            r_TL = 18
            r_G = 19
            r_VAT = 20
            r_GXD = 21
            r_LT = 22
            r_TOT = 23

            rows_spec.append(("I", "Chi phí trực tiếp", formula_T, formula_T, "T", calc_T, True, True))
            rows_spec.append(("1", "Chi phí vật liệu", "=D8", "=D8", "VL", "VL = VLHT", True, False))
            rows_spec.append(("", " - Đơn giá vật liệu", "Theo bảng tổng hợp vật liệu", f"='TH_VatLieu'!H{r_total_vl}", "VLHT", "Bảng tổng hợp vật liệu", False, False))
            rows_spec.append(("2", "Chi phí nhân công", "=D10", "=D10", "NC", "NC = NCHT", True, False))
            rows_spec.append(("", " - Đơn giá nhân công", "Theo bảng tổng hợp nhân công", f"='TH_NhanCong'!G{r_total_nc}", "NCHT", "Bảng tổng hợp nhân công", False, False))
            rows_spec.append(("3", "Chi phí máy và thiết bị thi công", "=D12", "=D12", "M", "M = MHT", True, False))
            rows_spec.append(("", " - Đơn giá máy thi công", "Theo bảng tổng hợp máy thi công", f"='TH_MayThiCong'!G{r_total_mac}", "MHT", "Bảng tổng hợp máy thi công", False, False))
            rows_spec.append(("4", "Chi phí công tác báo giá / tạm tính", "=D14", "=D14", "BG", f"{quote_count} công việc báo giá trọn gói", True, False))
            rows_spec.append(("", " - Chi phí báo giá trước thuế", "Theo báo giá cạnh tranh / PAKT", f"{quote_sum:.2f}", "BGHT", "Hồ sơ PAKT & Báo giá", False, False))
        else:
            formula_T = "=D7+D9+D11"
            calc_T = "T = VL + NC + M"
            r_GT = 13
            r_C = 14
            r_TT = 15
            r_TL = 16
            r_G = 17
            r_VAT = 18
            r_GXD = 19
            r_LT = 20
            r_TOT = 21

            rows_spec.append(("I", "Chi phí trực tiếp", formula_T, formula_T, "T", calc_T, True, True))
            rows_spec.append(("1", "Chi phí vật liệu", "=D8", "=D8", "VL", "VL = VLHT", True, False))
            rows_spec.append(("", " - Đơn giá vật liệu", "Theo bảng tổng hợp vật liệu", f"='TH_VatLieu'!H{r_total_vl}", "VLHT", "Bảng tổng hợp vật liệu", False, False))
            rows_spec.append(("2", "Chi phí nhân công", "=D10", "=D10", "NC", "NC = NCHT", True, False))
            rows_spec.append(("", " - Đơn giá nhân công", "Theo bảng tổng hợp nhân công", f"='TH_NhanCong'!G{r_total_nc}", "NCHT", "Bảng tổng hợp nhân công", False, False))
            rows_spec.append(("3", "Chi phí máy và thiết bị thi công", "=D12", "=D12", "M", "M = MHT", True, False))
            rows_spec.append(("", " - Đơn giá máy thi công", "Theo bảng tổng hợp máy thi công", f"='TH_MayThiCong'!G{r_total_mac}", "MHT", "Bảng tổng hợp máy thi công", False, False))

        rows_spec.extend([
            ("II", "Chi phí gián tiếp", f"=D{r_C}+D{r_TT}", f"=D{r_C}+D{r_TT}", "GT", "GT = C + TT", True, True),
            ("1", "Chi phí chung", f"=D{r_T}*{c_rate}%", f"=D{r_T}*{c_rate/100:.6f}", "C", c_basis, False, False),
            ("2", "Chi phí một số công việc không xác định được KL từ thiết kế", f"=D{r_T}*{tt_rate}%", f"=D{r_T}*{tt_rate/100:.6f}", "TT", tt_basis, False, False),
            ("III", "Thu nhập chịu thuế tính trước", f"=(D{r_T}+D{r_GT})*{tl_rate}%", f"=(D{r_T}+D{r_GT})*{tl_rate/100:.6f}", "TL", tl_basis, True, True),
            ("", "Chi phí xây dựng trước thuế", f"=D{r_T}+D{r_GT}+D{r_TL}", f"=D{r_T}+D{r_GT}+D{r_TL}", "G", "G = T + GT + TL", True, False),
            ("IV", "Thuế giá trị gia tăng", f"=D{r_G}*{vat_rate}%", f"=D{r_G}*{vat_rate/100:.6f}", "GTGT", vat_basis, True, True),
            ("", "CHI PHÍ XÂY DỰNG SAU THUẾ", f"=D{r_G}+D{r_VAT}", f"=D{r_G}+D{r_VAT}", "Gxd", "Gxd = G + GTGT", True, True),
            ("V", "Chi phí nhà tạm để ở và điều hành thi công", f"=D{r_G}*{lt_rate}%*(1+{vat_rate}%)", f"=D{r_G}*{lt_rate/100:.6f}*(1+{vat_rate/100:.6f})", "LT", lt_basis, True, False),
            ("", "Tổng cộng", f"=D{r_GXD}+D{r_LT}", f"=D{r_GXD}+D{r_LT}", "", "", True, True),
            ("", "Làm tròn", f"=ROUND(D{r_TOT},-3)", f"=ROUND(D{r_TOT},-3)", "", "", True, True)
        ])

        curr_r = 6
        for stt, name, calc, formula, sym, note, is_b, is_hl in rows_spec:
            ws.row_dimensions[curr_r].height = 22
            c_stt = ws.cell(curr_r, 1, stt)
            c_stt.alignment = self.align_center
            if is_b: c_stt.font = self.font_bold

            c_name = ws.cell(curr_r, 2, name)
            c_name.alignment = self.align_left
            c_name.font = self.font_bold if is_b else self.font_regular

            c_calc = ws.cell(curr_r, 3, calc)
            c_calc.alignment = self.align_left
            c_calc.font = self.font_regular

            c_val = ws.cell(curr_r, 4, formula)
            c_val.alignment = self.align_right
            c_val.font = self.font_bold if is_b else self.font_regular
            c_val.number_format = "#,##0"

            c_sym = ws.cell(curr_r, 5, sym)
            c_sym.alignment = self.align_center
            c_sym.font = self.font_bold if sym else self.font_regular

            c_note = ws.cell(curr_r, 6, note)
            c_note.alignment = self.align_left
            c_note.font = self.font_regular

            for c in range(1, 7):
                ws.cell(curr_r, c).border = self.border_cell
                if is_hl and stt in ["I", "II", "III", "IV", "V"]:
                    ws.cell(curr_r, c).fill = self.fill_yellow
                elif name in ["Tổng cộng", "Làm tròn", "CHI PHÍ XÂY DỰNG SAU THUẾ"]:
                    ws.cell(curr_r, c).fill = self.fill_code

            curr_r += 1

        # Calculate estimated total for num2words_vi (Current direct prices)
        sum_vl_chinh = sum(m.get("total_qty", 0.0) * m.get("price_notice", 0) for m in materials if m.get("code") != "V00750" and m.get("unit") != "%")
        sum_vl = sum_vl_chinh * 1.02
        sum_nc = sum(l.get("total_qty", 0.0) * l.get("price_current", 0) for l in labor)
        sum_mac = sum(mac.get("total_qty", 0.0) * mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0)) for mac in machines)
        t_est = sum_vl + sum_nc + sum_mac + quote_sum
        c_est = t_est * (c_rate / 100.0)
        tt_est = t_est * (tt_rate / 100.0)
        gt_est = c_est + tt_est
        tl_est = (t_est + gt_est) * (tl_rate / 100.0)
        g_est = t_est + gt_est + tl_est
        vat_est = g_est * (vat_rate / 100.0)
        gxd_est = g_est + vat_est
        lt_est = g_est * (lt_rate / 100.0) * (1.0 + vat_rate / 100.0)
        tot_est = gxd_est + lt_est
        round_est = round(tot_est, -3)
        words_vi = num2words_vi(round_est)

        # Row 23: In words
        ws.row_dimensions[curr_r].height = 24
        ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=6)
        c_words = ws.cell(curr_r, 2, f"Bằng chữ: {words_vi}")
        c_words.font = Font(name="Times New Roman", size=11, bold=True, italic=True, color="002060")
        c_words.alignment = self.align_left

        # Signatures block
        r_sig_title = curr_r + 2
        ws.row_dimensions[r_sig_title].height = 22
        
        ws.cell(r_sig_title, 2, "NGƯỜI LẬP").font = self.font_bold
        ws.cell(r_sig_title, 2).alignment = self.align_center

        ws.cell(r_sig_title, 5, "NGƯỜI CHỦ TRÌ").font = self.font_bold
        ws.cell(r_sig_title, 5).alignment = self.align_center

        r_sig_sub = r_sig_title + 1
        ws.cell(r_sig_sub, 5, "(Chứng chỉ kỹ sư định giá XD)").font = self.font_subtitle
        ws.cell(r_sig_sub, 5).alignment = self.align_center

        col_w = {1: 8, 2: 45, 3: 28, 4: 18, 5: 12, 6: 42}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _determine_group(self, item: Dict[str, Any]) -> str:
        """Determines the standard work group for an item."""
        grp = (item.get("group") or "").strip()
        if grp:
            return grp
        code = (item.get("norm_code") or "").strip()
        if code.startswith("BD.41") or code.startswith("BD.42"):
            return "Hệ thống báo cháy tự động và thiết bị ngoại vi (Lắp đặt - Thay mới)"
        elif code.startswith("E6C.3.01") or code.startswith("E6C.3.02") or code.startswith("E6C.3.03") or code.startswith("E6C.3.04") or code.startswith("E6C.3.05"):
            return "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy"
        elif code.startswith("E6C.3.06") or code.startswith("E6B.") or code.startswith("E6C.1."):
            return "Bảo dưỡng van tràn Deluge, hệ thống ắc quy và thiết bị đo khí"
        return "Các công tác lắp đặt và bảo dưỡng khác"

    def _build_sheet_khoi_luong(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]]
    ):
        """Builds the primary work volume sheet (Bảng tiên lượng khối lượng công việc), with clean standard grouping."""
        ws = wb.create_sheet(title="Khối lượng-Định mức-ĐVT")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP KHỐI LƯỢNG CÔNG VIỆC VÀ ĐỊNH MỨC ÁP DỤNG"
        ws["A1"].font = self.font_title
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A2"].font = self.font_project
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["A3"].font = self.font_project

        headers = ["STT", "Mã hiệu\nđịnh mức", "Danh mục hàng hóa/dịch vụ (THEO PAKT)", "Mô tả công việc", "Khối lượng", "Đơn vị tính", "Căn cứ pháp lý áp dụng", "Ghi chú"]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            if c_idx == 2:
                cell.fill = self.fill_code

        # Organize items by group while preserving order of first appearance
        grouped_items = []
        groups_order = []
        groups_map = {}
        for item in work_items:
            g_name = self._determine_group(item)
            if g_name not in groups_map:
                groups_map[g_name] = []
                groups_order.append(g_name)
            groups_map[g_name].append(item)

        r = 6
        roman_nums = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
        stt_counter = 1

        for g_idx, g_name in enumerate(groups_order):
            group_roman = roman_nums[g_idx % len(roman_nums)]
            items_in_g = groups_map[g_name]

            # 1. Group Header Row
            ws.row_dimensions[r].height = 24
            ws.cell(r, 1, group_roman).alignment = self.align_center
            ws.cell(r, 1).font = self.font_bold
            
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            c_grp = ws.cell(r, 2, f"NHÓM {group_roman}: {g_name.upper()}")
            c_grp.font = Font(name="Times New Roman", size=11, bold=True, color="002060")
            c_grp.alignment = self.align_left
            
            for c in range(1, 9):
                ws.cell(r, c).border = self.border_cell
                ws.cell(r, c).fill = self.fill_ev_header
            r += 1

            # 2. Work Items in this Group
            for item in items_in_g:
                ws.row_dimensions[r].height = 22
                ws.cell(r, 1, stt_counter).alignment = self.align_center
                
                c_code = ws.cell(r, 2, item.get("norm_code", ""))
                c_code.font = self.font_code
                c_code.alignment = self.align_center
                c_code.fill = self.fill_code
                
                ws.cell(r, 3, item.get("job_name", "")).alignment = self.align_left
                ws.cell(r, 4, item.get("description", item.get("job_name", ""))).alignment = self.align_left
                
                c_kl = ws.cell(r, 5, item.get("volume", 0))
                c_kl.alignment = self.align_right
                c_kl.font = self.font_bold
                c_kl.number_format = "#,##0.00"
                
                ws.cell(r, 6, item.get("unit", "")).alignment = self.align_center
                
                norm = self.norm_manager.get_norm(item.get("norm_code", ""))
                basis_str = (norm.get("legal_basis", "") if norm else "") or item.get("quote_legal_basis", "") or item.get("basis", "")
                ws.cell(r, 7, basis_str).alignment = self.align_left
                ws.cell(r, 8, item.get("notes", "")).alignment = self.align_left

                for c in range(1, 9):
                    ws.cell(r, c).border = self.border_cell
                    if c not in [1, 2, 5]:
                        ws.cell(r, c).font = self.font_regular
                r += 1
                stt_counter += 1

        col_w = {1: 6, 2: 15, 3: 45, 4: 45, 5: 14, 6: 12, 7: 35, 8: 20}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_thong_ke(self, wb: openpyxl.Workbook, norm_codes: List[str]):
        ws = wb.create_sheet(title="00_ThongKe_DinhMuc")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG THỐNG KÊ TOÀN BỘ ĐỊNH MỨC ÁP DỤNG TRONG DỰ TOÁN PCCC 2026"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Quy chuẩn phân loại theo ĐM SCL EVN QĐ 45/QĐ-HĐTV và ĐM BXD TT 38/2026/TT-BXD)"
        ws["A2"].font = self.font_subtitle

        headers = ["STT", "Bộ định mức", "Mã hiệu chuẩn", "Tên công tác chuẩn", "ĐVT gốc", "Căn cứ pháp lý", "Link sang sheet minh chứng"]
        ws.row_dimensions[4].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(4, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            if c_idx == 3:
                cell.fill = self.fill_code

        for idx, code in enumerate(norm_codes, start=1):
            norm = self.norm_manager.get_norm(code)
            r = 4 + idx
            ws.row_dimensions[r].height = 22
            ws.cell(r, 1, idx).alignment = self.align_center
            
            bo_dm = "ĐM SCL EVN (QĐ 45)" if code.startswith("E") else "ĐM BXD (TT 38/2026)"
            ws.cell(r, 2, bo_dm).alignment = self.align_center
            
            c_code = ws.cell(r, 3, code)
            c_code.alignment = self.align_center
            c_code.font = self.font_code
            c_code.fill = self.fill_code
            
            ws.cell(r, 4, norm.get("name", "")).alignment = self.align_left
            ws.cell(r, 5, norm.get("unit", "")).alignment = self.align_center
            ws.cell(r, 6, norm.get("legal_basis", "")).alignment = self.align_left
            
            # Hyperlink to evidence sheet
            c_link = ws.cell(r, 7, f"Xem sheet {code}")
            c_link.hyperlink = f"#'{code}'!A1"
            c_link.font = Font(name="Times New Roman", size=11, color="0563C1", underline="single")
            c_link.alignment = self.align_center
            
            for c in range(1, 8):
                ws.cell(r, c).border = self.border_cell

        col_w = {1: 6, 2: 24, 3: 16, 4: 55, 5: 12, 6: 45, 7: 25}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_evidence_sheet(self, wb: openpyxl.Workbook, code: str) -> Dict[str, Any]:
        """Creates individual evidence sheet with PDF image at B3 and tables at Column P."""
        ws = wb.create_sheet(title=code)
        ws.views.sheetView[0].showGridLines = True
        norm = self.norm_manager.get_norm(code)

        # 1. Insert Image at B3 if exists
        img_path = self.norm_manager.get_evidence_image_path(code)
        if img_path and img_path.exists():
            try:
                img = OpenpyxlImage(str(img_path))
                img.width = 650
                img.height = int(650 * (img.height / img.width)) if img.width else 450
                ws.add_image(img, "B3")
            except Exception as e:
                print(f"Warning: could not insert image for {code}: {e}")

        # 2. Header Info at P1, P2
        ws["P1"] = f"{code}: {norm.get('name', '')}"
        ws["P1"].font = self.font_ev_title
        ws["P2"] = f"Căn cứ: {norm.get('legal_basis', '')}"
        ws["P2"].font = self.font_subtitle

        # 3. Work Steps Table (Rows 3 to N)
        steps = norm.get("work_steps", [])
        labor_split = norm.get("labor_split")
        is_split = labor_split is not None

        if is_split:
            # Multi-labor columns (e.g. BD norms: NC nhóm 3 + Kỹ sư)
            step_cols = [
                ("P", "Thành phần công việc"),
                ("Q", "ĐVT / Tỷ lệ"),
                ("R", "NC nhóm 3 (công)"),
                ("S", "Kỹ sư (công)"),
                ("T", "Tổng ĐM gốc (công)"),
                ("U", "Áp dụng (1/0)"),
                ("V", "Hao phí tính toán")
            ]
            for col_idx, (col_let, col_title) in enumerate(step_cols, start=16):
                cell = ws[f"{col_let}3"]
                cell.value = col_title
                cell.font = self.font_ev_header
                cell.fill = self.fill_ev_header
                cell.border = self.border_cell
                cell.alignment = self.align_center

            r_step = 4
            for s in steps:
                ws.cell(r_step, 16, s["name"]).font = self.font_ev_reg
                dvt_label = f"{s.get('dvt_step', '')} ({s.get('pct', '')})" if s.get('pct') else s.get('dvt_step', '')
                ws.cell(r_step, 17, dvt_label).alignment = self.align_center
                
                c_nc = ws.cell(r_step, 18, s.get("rate_nc", 0.0))
                c_nc.alignment = self.align_right
                c_nc.font = self.font_ev_reg
                c_nc.number_format = "0.0000"

                c_ks = ws.cell(r_step, 19, s.get("rate_ks", 0.0))
                c_ks.alignment = self.align_right
                c_ks.font = self.font_ev_reg
                c_ks.number_format = "0.0000"

                c_tot_goc = ws.cell(r_step, 20, f"=R{r_step}+S{r_step}")
                c_tot_goc.alignment = self.align_right
                c_tot_goc.font = self.font_ev_reg
                c_tot_goc.number_format = "0.0000"

                c_u = ws.cell(r_step, 21, s.get("apply_flag", 1))
                c_u.alignment = self.align_center
                c_u.font = self.font_ev_bold

                c_v = ws.cell(r_step, 22, f"=T{r_step}*U{r_step}")
                c_v.alignment = self.align_right
                c_v.font = self.font_ev_reg
                c_v.number_format = "0.0000"

                for c in range(16, 23):
                    ws.cell(r_step, c).border = self.border_cell
                r_step += 1

            r_tot_step = r_step
            ws.cell(r_tot_step, 16, "Tổng cộng hao phí các bước:").font = self.font_ev_bold
            ws.cell(r_tot_step, 16).alignment = self.align_right

            c_tot_nc = ws.cell(r_tot_step, 18, f"=SUM(R4:R{r_tot_step-1})")
            c_tot_nc.font = self.font_ev_bold
            c_tot_nc.alignment = self.align_right
            c_tot_nc.number_format = "0.000"

            c_tot_ks = ws.cell(r_tot_step, 19, f"=SUM(S4:S{r_tot_step-1})")
            c_tot_ks.font = self.font_ev_bold
            c_tot_ks.alignment = self.align_right
            c_tot_ks.number_format = "0.000"

            c_tot_goc = ws.cell(r_tot_step, 20, f"=SUM(T4:T{r_tot_step-1})")
            c_tot_goc.font = self.font_ev_bold
            c_tot_goc.alignment = self.align_right
            c_tot_goc.number_format = "0.000"

            c_tot_calc = ws.cell(r_tot_step, 22, f"=SUM(V4:V{r_tot_step-1})")
            c_tot_calc.font = self.font_ev_bold
            c_tot_calc.alignment = self.align_right
            c_tot_calc.number_format = "0.000"

            for c in range(16, 23):
                ws.cell(r_tot_step, c).border = self.border_cell

            # Clarification note row
            r_note = r_tot_step + 1
            ws.merge_cells(start_row=r_note, start_column=16, end_row=r_note, end_column=22)
            lbl_nc = labor_split.get("nc_total_per_unit", 0.087)
            lbl_ks = labor_split.get("ks_total_per_unit", 0.087)
            lbl_tot = labor_split.get("total_per_unit", 0.174)
            c_expl = ws.cell(r_note, 16, f"*(Ghi chú: {lbl_tot} công = {lbl_nc} công NC nhóm 3 + {lbl_ks} công Kỹ sư bậc 5,0/8 theo Thông tư 38/2026/TT-BXD)*")
            c_expl.font = Font(name="Times New Roman", size=10, bold=True, italic=True, color="002060")
            c_expl.alignment = self.align_left

            col_w = {"P": 38, "Q": 14, "R": 15, "S": 15, "T": 15, "U": 12, "V": 16}
            for col_l, w in col_w.items():
                ws.column_dimensions[col_l].width = w

            r_res_title = r_tot_step + 3
            table_end_col = 22

        else:
            # Single-labor standard columns (e.g. E6 norms)
            step_cols = [
                ("P", "Thành phần công việc"),
                ("Q", "ĐVT / Tỷ lệ"),
                ("R", "Định mức gốc"),
                ("S", "Áp dụng (1=Có, 0=Không)"),
                ("T", "Hao phí tính toán")
            ]
            for col_idx, (col_let, col_title) in enumerate(step_cols, start=16):
                cell = ws[f"{col_let}3"]
                cell.value = col_title
                cell.font = self.font_ev_header
                cell.fill = self.fill_ev_header
                cell.border = self.border_cell
                cell.alignment = self.align_center

            r_step = 4
            for s in steps:
                ws.cell(r_step, 16, s["name"]).font = self.font_ev_reg
                ws.cell(r_step, 17, s.get("dvt_step", "")).alignment = self.align_center
                c_r = ws.cell(r_step, 18, s.get("rate_base", 0.0))
                c_r.alignment = self.align_right
                c_r.font = self.font_ev_reg
                c_s = ws.cell(r_step, 19, s.get("apply_flag", 1))
                c_s.alignment = self.align_center
                c_s.font = self.font_ev_bold
                c_t = ws.cell(r_step, 20, f"=R{r_step}*S{r_step}")
                c_t.alignment = self.align_right
                c_t.font = self.font_ev_reg
                for c in range(16, 21):
                    ws.cell(r_step, c).border = self.border_cell
                r_step += 1

            r_tot_step = r_step
            ws.cell(r_tot_step, 16, "Tổng cộng hao phí các bước:").font = self.font_ev_bold
            ws.cell(r_tot_step, 16).alignment = self.align_right
            ws.cell(r_tot_step, 18, f"=SUM(R4:R{r_tot_step-1})").font = self.font_ev_bold
            ws.cell(r_tot_step, 20, f"=SUM(T4:T{r_tot_step-1})").font = self.font_ev_bold
            for c in range(16, 21):
                ws.cell(r_tot_step, c).border = self.border_cell

            col_w = {"P": 45, "Q": 12, "R": 18, "S": 22, "T": 25}
            for col_l, w in col_w.items():
                ws.column_dimensions[col_l].width = w

            r_res_title = r_tot_step + 2
            table_end_col = 20

        # 4. Resource Table (Vật liệu, Nhân công, Máy thi công)
        ws.cell(r_res_title, 16, f"BẢNG THÀNH PHẦN HAO PHÍ ĐỊNH MỨC (MÃ HIỆU: {code})").font = self.font_ev_title
        
        r_res_hdr = r_res_title + 1
        res_headers = ["Mã hiệu / Thành phần hao phí", "ĐVT gốc", "Định mức hao phí", "Cấp bậc / Thông số kỹ thuật", "Ghi chú"]
        for c_idx, h in enumerate(res_headers, start=16):
            cell = ws.cell(r_res_hdr, c_idx, h)
            cell.font = self.font_ev_header
            cell.fill = self.fill_ev_header
            cell.border = self.border_cell
            cell.alignment = self.align_center

        cell_coords = {"materials": {}, "labor": {}, "machines": {}}
        curr_r = r_res_hdr + 1
        res = norm.get("resources", {})

        # Materials
        mats = [m for m in res.get("materials", []) if m.get("code") and m.get("name", "").strip() != "1. Vật liệu"]
        if mats:
            ws.cell(curr_r, 16, "1. Vật liệu").font = self.font_ev_cat
            for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1
            for m in mats:
                ws.cell(curr_r, 16, f"   - {m['name']} ({m['code']})").font = self.font_ev_reg
                ws.cell(curr_r, 17, m.get("unit", "")).alignment = self.align_center
                c_val = ws.cell(curr_r, 18, m.get("rate", 0.0))
                c_val.font = self.font_ev_bold
                c_val.alignment = self.align_right
                ws.cell(curr_r, 19, m.get("spec", "")).font = self.font_ev_reg
                ws.cell(curr_r, 20, m.get("note", "")).font = self.font_ev_reg
                for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
                cell_coords["materials"][m["code"]] = curr_r
                curr_r += 1

        # Labor
        labs = [l for l in res.get("labor", []) if l.get("code") and l.get("name", "").strip() != "2. Nhân công"]
        if labs:
            cat_lbl = "2. Nhân công" if mats else "1. Nhân công trực tiếp"
            ws.cell(curr_r, 16, cat_lbl).font = self.font_ev_cat
            for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1
            for l in labs:
                ws.cell(curr_r, 16, f"   - {l['name']} ({l['code']})").font = self.font_ev_reg
                ws.cell(curr_r, 17, l.get("unit", "công")).alignment = self.align_center
                if code.startswith("E6") or l.get("formula_ref"):
                    c_val = ws.cell(curr_r, 18, f"=T{r_tot_step}")
                else:
                    c_val = ws.cell(curr_r, 18, l.get("rate", 0.0))
                c_val.font = self.font_ev_bold
                c_val.alignment = self.align_right
                ws.cell(curr_r, 19, l.get("spec", "")).font = self.font_ev_reg
                ws.cell(curr_r, 20, l.get("note", "")).font = self.font_ev_reg
                for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
                cell_coords["labor"][l["code"]] = curr_r
                curr_r += 1

        # Machines
        macs = [mac for mac in res.get("machines", []) if mac.get("code") and mac.get("name", "").strip() != "3. Máy thi công"]
        if macs:
            cat_lbl = "3. Máy thi công" if mats else "2. Máy thi công"
            ws.cell(curr_r, 16, cat_lbl).font = self.font_ev_cat
            for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1
            for mac in macs:
                ws.cell(curr_r, 16, f"   - {mac['name']} ({mac['code']})").font = self.font_ev_reg
                ws.cell(curr_r, 17, mac.get("unit", "ca")).alignment = self.align_center
                c_val = ws.cell(curr_r, 18, mac.get("rate", 0.0))
                c_val.font = self.font_ev_bold
                c_val.alignment = self.align_right
                ws.cell(curr_r, 19, mac.get("spec", "")).font = self.font_ev_reg
                ws.cell(curr_r, 20, mac.get("note", "")).font = self.font_ev_reg
                for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
                cell_coords["machines"][mac["code"]] = curr_r
                curr_r += 1

        return cell_coords

    def _build_sheet_don_gia_chi_tiet(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]],
        norm_cell_map: Dict[str, Any],
        mat_row_map: Dict[str, int],
        lab_row_map: Dict[str, int],
        mac_row_map: Dict[str, int]
    ):
        ws = wb.create_sheet(title="01_ChiTiet_HaoPhi_DinhMuc")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG ĐƠN GIÁ CHI TIẾT"
        ws["A1"].font = self.font_title
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A2"].font = self.font_project
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["A3"].font = self.font_project

        headers = ["STT", "Mã hiệu\nđơn giá", "", "", "Tên công tác", "Đơn vị", "Định mức", "Đơn giá", "Hệ số", "Thành tiền"]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell

        curr_r = 6
        norm_work_items = [it for it in work_items if not it.get("is_quote")]
        for item_idx, item in enumerate(norm_work_items, start=1):
            stt = item.get("stt", item_idx)
            code = item["norm_code"]
            norm = self.norm_manager.get_norm(code)
            if not norm:
                continue
            coords = norm_cell_map.get(code, {"materials": {}, "labor": {}, "machines": {}})

            # Job Header Row
            ws.row_dimensions[curr_r].height = 22
            ws.cell(curr_r, 1, stt).alignment = self.align_center
            ws.cell(curr_r, 1).font = self.font_bold
            
            c_code = ws.cell(curr_r, 2, code)
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            
            ws.cell(curr_r, 5, item.get("job_name", norm.get("name", ""))).font = self.font_bold
            ws.cell(curr_r, 6, item.get("unit", norm.get("unit", ""))).alignment = self.align_center
            
            job_hdr_r = curr_r
            curr_r += 1

            res = norm.get("resources", {})
            subtotal_rows = []

            # 1. Materials
            mats = [m for m in res.get("materials", []) if m.get("code") and m.get("name", "").strip() != "1. Vật liệu"]
            if mats:
                ws.cell(curr_r, 5, "Vật liệu").font = self.font_cat
                curr_r += 1
                mat_start_r = curr_r
                for m in mats:
                    m_code = m["code"]
                    ws.cell(curr_r, 2, m_code).font = self.font_regular
                    ws.cell(curr_r, 2).alignment = self.align_center
                    ws.cell(curr_r, 5, f" - {m['name']}").font = self.font_regular
                    ws.cell(curr_r, 6, m.get("unit", "")).alignment = self.align_center
                    
                    r_ev = coords["materials"].get(m_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, m.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    if m_code == "V00750" or m.get("unit") == "%":
                        ws.cell(curr_r, 10, f"=SUM(J{mat_start_r}:J{curr_r-1})*G{curr_r}/100").alignment = self.align_right
                    else:
                        r_th = mat_row_map.get(m_code)
                        c_h = ws.cell(curr_r, 8, f"='TH_VatLieu'!$G${r_th}" if r_th else 0)
                        c_h.fill = self.fill_yellow
                        c_h.alignment = self.align_right
                        c_h.number_format = "#,##0"
                        ws.cell(curr_r, 10, f"=G{curr_r}*H{curr_r}").alignment = self.align_right

                    ws.cell(curr_r, 10).number_format = "#,##0"
                    curr_r += 1

                mat_end_r = curr_r - 1
                ws.cell(job_hdr_r + 1, 10, f"=SUM(J{mat_start_r}:J{mat_end_r})").font = self.font_cat
                ws.cell(job_hdr_r + 1, 10).number_format = "#,##0"
                subtotal_rows.append(job_hdr_r + 1)

            # 2. Labor
            labs = [l for l in res.get("labor", []) if l.get("code") and l.get("name", "").strip() != "2. Nhân công"]
            if labs:
                ws.cell(curr_r, 5, "Nhân công").font = self.font_cat
                curr_r += 1
                lab_start_r = curr_r
                for l in labs:
                    l_code = l["code"]
                    ws.cell(curr_r, 2, l_code).font = self.font_regular
                    ws.cell(curr_r, 2).alignment = self.align_center
                    ws.cell(curr_r, 5, f" - {l['name']}").font = self.font_regular
                    ws.cell(curr_r, 6, l.get("unit", "công")).alignment = self.align_center
                    
                    r_ev = coords["labor"].get(l_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, l.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    r_th = lab_row_map.get(l_code)
                    c_h = ws.cell(curr_r, 8, f"='TH_NhanCong'!$F${r_th}" if r_th else 0)
                    c_h.fill = self.fill_yellow
                    c_h.alignment = self.align_right
                    c_h.number_format = "#,##0"

                    ws.cell(curr_r, 10, f"=G{curr_r}*H{curr_r}").alignment = self.align_right
                    ws.cell(curr_r, 10).number_format = "#,##0"
                    curr_r += 1

                lab_end_r = curr_r - 1
                ws.cell(lab_start_r - 1, 10, f"=SUM(J{lab_start_r}:J{lab_end_r})").font = self.font_cat
                ws.cell(lab_start_r - 1, 10).number_format = "#,##0"
                subtotal_rows.append(lab_start_r - 1)

            # 3. Machines
            macs = [mac for mac in res.get("machines", []) if mac.get("code") and mac.get("name", "").strip() != "3. Máy thi công"]
            if macs:
                ws.cell(curr_r, 5, "Máy thi công").font = self.font_cat
                curr_r += 1
                mac_start_r = curr_r
                for mac in macs:
                    m_code = mac["code"]
                    ws.cell(curr_r, 2, m_code).font = self.font_regular
                    ws.cell(curr_r, 2).alignment = self.align_center
                    ws.cell(curr_r, 5, f" - {mac['name']}").font = self.font_regular
                    ws.cell(curr_r, 6, mac.get("unit", "ca")).alignment = self.align_center
                    
                    r_ev = coords["machines"].get(m_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, mac.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    r_th = mac_row_map.get(m_code)
                    c_h = ws.cell(curr_r, 8, f"='TH_MayThiCong'!$F${r_th}" if r_th else 0)
                    c_h.fill = self.fill_yellow
                    c_h.alignment = self.align_right
                    c_h.number_format = "#,##0"

                    ws.cell(curr_r, 10, f"=G{curr_r}*H{curr_r}").alignment = self.align_right
                    ws.cell(curr_r, 10).number_format = "#,##0"
                    curr_r += 1

                mac_end_r = curr_r - 1
                ws.cell(mac_start_r - 1, 10, f"=SUM(J{mac_start_r}:J{mac_end_r})").font = self.font_cat
                ws.cell(mac_start_r - 1, 10).number_format = "#,##0"
                subtotal_rows.append(mac_start_r - 1)

            if subtotal_rows:
                ws.cell(job_hdr_r, 10, "=" + "+".join(f"J{r_sub}" for r_sub in subtotal_rows)).font = self.font_bold
                ws.cell(job_hdr_r, 10).number_format = "#,##0"

        for r in range(5, curr_r):
            for c in range(1, 11):
                ws.cell(r, c).border = self.border_cell

        col_w = {1: 6, 2: 14, 3: 4, 4: 4, 5: 50, 6: 12, 7: 15, 8: 15, 9: 10, 10: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_vat_lieu(self, wb: openpyxl.Workbook, project_name: str, category_name: str, materials: List[Dict[str, Any]]):
        """
        Builds the current price material summary sheet (Bảng tổng hợp vật liệu dạng giá hiện tại)
        matching 100% the layout in user request:
        STT | Mã hiệu | Tên vật tư | Đơn vị | Nguồn mua | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
        """
        ws = wb.create_sheet(title="TH_VatLieu")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP VẬT LIỆU"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Theo giá thông báo quý .. năm 20..)"
        ws["A2"].font = self.font_subtitle
        ws["A3"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"].font = self.font_project
        ws["A4"] = f"HẠNG MỤC: {category_name}"
        ws["A4"].font = self.font_project
        ws["I5"] = "Đơn vị: đồng"
        ws["I5"].font = self.font_unit
        ws["I5"].alignment = self.align_right

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Nguồn mua", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_ev_header
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        curr_r = r_start
        for idx, m in enumerate(materials, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(curr_r, 2, m["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(curr_r, 3, m["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, m["unit"]).alignment = self.align_center
            
            # Nguồn mua
            ws.cell(curr_r, 5, m.get("source", "")).alignment = self.align_left
            
            is_other = (m.get("code") == "V00750" or m.get("unit") == "%")
            if is_other:
                # Vật liệu khác: không có khối lượng và đơn giá, tính theo % tổng vật liệu chính
                ws.cell(curr_r, 6, "")
                ws.cell(curr_r, 7, "")
                pct = float(m.get("rate", 2.0))
                c_tt = ws.cell(curr_r, 8, f"=ROUND(SUM(H{r_start}:H{curr_r-1})*{pct}%, 0)")
                c_tt.alignment = self.align_right
                c_tt.number_format = "#,##0"
                
                c_basis = ws.cell(curr_r, 9, m.get("legal_basis", f"{pct}% vật liệu chính theo định mức"))
                c_basis.alignment = self.align_left
            else:
                c_kl = ws.cell(curr_r, 6, m.get("total_qty", 0.0))
                c_kl.alignment = self.align_right
                c_kl.number_format = "#,##0.00"
                
                c_ht = ws.cell(curr_r, 7, m.get("price_notice", 0))
                c_ht.alignment = self.align_right
                c_ht.number_format = "#,##0"
                
                c_tt = ws.cell(curr_r, 8, f"=ROUND(F{curr_r}*G{curr_r}, 0)")
                c_tt.alignment = self.align_right
                c_tt.number_format = "#,##0"
                
                basis_val = m.get("legal_basis", m.get("notes", "Công bố giá LS / Báo giá NSX"))
                c_basis = ws.cell(curr_r, 9, basis_val)
                c_basis.alignment = self.align_left

            for c in range(1, 10):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 2, 8]: ws.cell(curr_r, c).font = self.font_regular
                elif c == 8: ws.cell(curr_r, c).font = self.font_bold

            curr_r += 1

        r_total = curr_r
        ws.row_dimensions[r_total].height = 24
        for c in range(1, 10): ws.cell(r_total, c).border = self.border_total
        
        font_blue_bold = Font(name="Times New Roman", size=11, bold=True, color="002060")
        c_lbl = ws.cell(r_total, 3, "Tổng vật liệu:")
        c_lbl.font = font_blue_bold
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")
        
        c_tot = ws.cell(r_total, 8, f"=SUM(H{r_start}:H{r_total-1})")
        c_tot.font = font_blue_bold
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 45, 4: 10, 5: 20, 6: 15, 7: 16, 8: 18, 9: 35}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_nhan_cong(self, wb: openpyxl.Workbook, project_name: str, category_name: str, labor: List[Dict[str, Any]]):
        """
        Builds the current price labor summary sheet (Bảng tổng hợp nhân công dạng giá hiện tại)
        matching 100% the layout in user request:
        STT | Mã hiệu | Tên vật tư | Đơn vị | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
        """
        ws = wb.create_sheet(title="TH_NhanCong")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP NHÂN CÔNG"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Theo giá thông báo quý .. năm 20..)"
        ws["A2"].font = self.font_subtitle
        ws["A3"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"].font = self.font_project
        ws["A4"] = f"HẠNG MỤC: {category_name}"
        ws["A4"].font = self.font_project
        ws["H5"] = "Đơn vị: đồng"
        ws["H5"].font = self.font_unit
        ws["H5"].alignment = self.align_right

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_ev_header
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        curr_r = r_start
        for idx, l in enumerate(labor, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(curr_r, 2, l["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(curr_r, 3, l["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, l.get("unit", "công")).alignment = self.align_center
            
            c_kl = ws.cell(curr_r, 5, l.get("total_qty", 0.0))
            c_kl.alignment = self.align_right
            c_kl.number_format = "#,##0.00"
            
            c_ht = ws.cell(curr_r, 6, l.get("price_current", 0))
            c_ht.alignment = self.align_right
            c_ht.number_format = "#,##0"
            
            c_tt = ws.cell(curr_r, 7, f"=ROUND(E{curr_r}*F{curr_r}, 0)")
            c_tt.alignment = self.align_right
            c_tt.number_format = "#,##0"

            # Basis column
            basis_val = l.get("legal_basis", "5898/SXD-QLXD ngày 02/7/2026 của SXD tỉnh Lâm Đồng")
            c_basis = ws.cell(curr_r, 8, basis_val)
            c_basis.alignment = self.align_left
            
            for c in range(1, 9):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 2, 7]: ws.cell(curr_r, c).font = self.font_regular
                elif c == 7: ws.cell(curr_r, c).font = self.font_bold

            curr_r += 1

        r_total = curr_r
        ws.row_dimensions[r_total].height = 24
        for c in range(1, 9): ws.cell(r_total, c).border = self.border_total
        
        font_blue_bold = Font(name="Times New Roman", size=11, bold=True, color="002060")
        c_lbl = ws.cell(r_total, 3, "Cộng nhân công:")
        c_lbl.font = font_blue_bold
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")
        
        c_tot = ws.cell(r_total, 7, f"=SUM(G{r_start}:G{r_total-1})")
        c_tot.font = font_blue_bold
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 45, 4: 12, 5: 15, 6: 16, 7: 18, 8: 38}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_may_thi_cong(self, wb: openpyxl.Workbook, project_name: str, category_name: str, machines: List[Dict[str, Any]]):
        """
        Builds the current price machine summary sheet (Bảng tổng hợp máy thi công dạng giá hiện tại)
        matching 100% the layout in user request:
        STT | Mã hiệu | Tên vật tư | Đơn vị | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
        """
        ws = wb.create_sheet(title="TH_MayThiCong")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP MÁY THI CÔNG"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Theo giá thông báo quý .. năm 20..)"
        ws["A2"].font = self.font_subtitle
        ws["A3"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"].font = self.font_project
        ws["A4"] = f"HẠNG MỤC: {category_name}"
        ws["A4"].font = self.font_project
        ws["H5"] = "Đơn vị: đồng"
        ws["H5"].font = self.font_unit
        ws["H5"].alignment = self.align_right

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_ev_header
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        curr_r = r_start
        for idx, mac in enumerate(machines, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(curr_r, 2, mac["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(curr_r, 3, mac["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, mac.get("unit", "ca")).alignment = self.align_center
            
            c_kl = ws.cell(curr_r, 5, mac.get("total_qty", 0.0))
            c_kl.alignment = self.align_right
            c_kl.number_format = "#,##0.00"
            
            c_ht_val = mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0))
            c_ht = ws.cell(curr_r, 6, c_ht_val)
            c_ht.alignment = self.align_right
            c_ht.number_format = "#,##0"
            
            c_tt = ws.cell(curr_r, 7, f"=ROUND(E{curr_r}*F{curr_r}, 0)")
            c_tt.alignment = self.align_right
            c_tt.number_format = "#,##0"

            # Basis column
            basis_val = mac.get("legal_basis", "Thông tư 37/2026/TT-BXD của Bộ Xây dựng")
            c_basis = ws.cell(curr_r, 8, basis_val)
            c_basis.alignment = self.align_left
            
            for c in range(1, 9):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 2, 7]: ws.cell(curr_r, c).font = self.font_regular
                elif c == 7: ws.cell(curr_r, c).font = self.font_bold

            curr_r += 1

        r_total = curr_r
        ws.row_dimensions[r_total].height = 24
        for c in range(1, 9): ws.cell(r_total, c).border = self.border_total
        
        font_blue_bold = Font(name="Times New Roman", size=11, bold=True, color="002060")
        c_lbl = ws.cell(r_total, 3, "Cộng Máy:")
        c_lbl.font = font_blue_bold
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")
        
        c_tot = ws.cell(r_total, 7, f"=SUM(G{r_start}:G{r_total-1})")
        c_tot.font = font_blue_bold
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 48, 4: 12, 5: 15, 6: 16, 7: 18, 8: 38}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
